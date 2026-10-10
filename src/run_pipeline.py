"""End-to-end pipeline - the single command that reproduces the whole project.

    python -m src.run_pipeline

Runs Phase 1 (synthetic data) and Phase 2 (RFM segmentation, cohort retention,
CLV), then writes ``data/processed/kpi_summary.json``.

That JSON file is the authoritative source for every number quoted in the
README. Nothing in the documentation is typed by hand: if the seed or any
assumption changes, re-running this command updates the figures.
"""

from __future__ import annotations

import json
from typing import Any

import pandas as pd

from src import (
    cohort_analysis,
    config,
    customer_lifetime_value,
    customer_segmentation,
    generate_data,
    kpis,
)


def _jsonable(value: Any) -> Any:
    """Convert numpy / pandas scalars into plain JSON-serialisable types."""
    if isinstance(value, (pd.Timestamp,)):
        return value.date().isoformat()
    if hasattr(value, "item"):
        return value.item()
    return value


def build_full_summary(tables: dict[str, pd.DataFrame]) -> dict[str, Any]:
    """Assemble every headline figure the README and dashboard rely on.

    Args:
        tables: The seven Phase 1 tables.

    Returns:
        A nested dictionary of KPIs, segment economics, cohort retention,
        clienteling performance and campaign effectiveness.
    """
    transactions = tables["transactions"]
    customers = tables["customers"]
    interactions = tables["clienteling_interactions"]
    campaigns = tables["campaigns"]
    associates = tables["associates"]

    summary: dict[str, Any] = {
        "headline": kpis.build_kpi_summary(
            transactions, interactions, campaigns, customers
        )
    }

    # ---- Segments ------------------------------------------------------- #
    segmented = customer_segmentation.build_segmentation(transactions, customers)
    seg_summary = customer_segmentation.segment_summary(segmented)
    summary["segments"] = [
        {
            "segment": str(row.segment),
            "clients": int(row.clients),
            "client_share": round(float(row.client_share), 4),
            "net_sales": round(float(row.net_sales), 2),
            "revenue_share": round(float(row.revenue_share), 4),
            "avg_order_value": round(float(row.avg_order_value), 2),
            "avg_frequency": round(float(row.avg_frequency), 2),
            "avg_recency_days": round(float(row.avg_recency_days), 1),
        }
        for row in seg_summary.itertuples(index=False)
    ]

    # ---- Cohort retention ----------------------------------------------- #
    _, rates = cohort_analysis.cohort_retention_matrix(transactions, customers)
    curve = cohort_analysis.retention_curve(rates)
    summary["cohorts"] = {
        "cohort_count": int(len(rates)),
        "repeat_rate_12m": round(
            cohort_analysis.repeat_rate_within_months(transactions, customers, 12), 4
        ),
        "avg_retention_by_month": {
            str(int(offset)): round(float(value), 4)
            for offset, value in curve.items()
            if offset > 0 and pd.notna(value)
        },
    }

    # ---- CLV ------------------------------------------------------------ #
    clv = customer_lifetime_value.build_clv(transactions, customers)
    clv_summary = customer_lifetime_value.clv_by_segment(clv)
    summary["clv"] = {
        "base_retention_rate": round(
            cohort_analysis.repeat_rate_within_months(transactions, customers, 12), 4
        ),
        "gross_margin_rate": round(
            customer_lifetime_value.realised_gross_margin(transactions), 4
        ),
        "annual_discount_rate": config.CLV_ANNUAL_DISCOUNT_RATE,
        "total_clv_12m": round(float(clv["clv_12m"].sum()), 2),
        "by_segment": [
            {
                "segment": str(row.segment),
                "clients": int(row.clients),
                "avg_clv_12m": round(float(row.avg_clv_12m), 2),
                "total_clv_12m": round(float(row.total_clv_12m), 2),
                "clv_share": round(float(row.clv_share), 4),
            }
            for row in clv_summary.itertuples(index=False)
        ],
    }

    # ---- Clienteling ---------------------------------------------------- #
    attributed = kpis.attribute_clienteling(transactions, interactions)
    by_type = (
        attributed.groupby("interaction_type")
        .agg(
            interactions=("interaction_id", "count"),
            conversions=("converted_in_window", "sum"),
            attributed_revenue=("attributed_revenue", "sum"),
        )
        .reset_index()
    )
    by_type["conversion_rate"] = by_type["conversions"] / by_type["interactions"]
    by_type["revenue_per_interaction"] = (
        by_type["attributed_revenue"] / by_type["interactions"]
    )
    summary["clienteling"] = {
        "attribution_window_days": config.CLIENTELING_ATTRIBUTION_WINDOW_DAYS,
        "by_type": [
            {
                "interaction_type": str(row.interaction_type),
                "interactions": int(row.interactions),
                "conversion_rate": round(float(row.conversion_rate), 4),
                "attributed_revenue": round(float(row.attributed_revenue), 2),
                "revenue_per_interaction": round(
                    float(row.revenue_per_interaction), 2
                ),
            }
            for row in by_type.itertuples(index=False)
        ],
    }

    performance = kpis.associate_performance(attributed, associates)
    summary["clienteling"]["top_associates"] = [
        {
            "associate_id": str(row.associate_id),
            "associate_name": str(row.associate_name),
            "boutique_id": str(row.boutique_id),
            "interactions": int(row.interactions),
            "conversion_rate": round(float(row.conversion_rate), 4),
            "attributed_revenue": round(float(row.attributed_revenue), 2),
        }
        for row in performance.head(5).itertuples(index=False)
    ]
    # Associate spread is measured on appointments only. Blending appointments
    # with calls hides the gap, because associates differ in how many of each
    # they log rather than only in how well they convert.
    appointments = attributed[attributed["interaction_type"] == "Appointment"]
    appointment_perf = kpis.associate_performance(appointments, associates)
    # Ignore associates with very few appointments - small denominators produce
    # conversion rates that look extreme but mean nothing.
    appointment_perf = appointment_perf[appointment_perf["interactions"] >= 15]
    summary["clienteling"]["appointment_conversion_by_associate"] = {
        "associates_measured": int(len(appointment_perf)),
        "min_appointments": 15,
        "top_quartile": round(
            float(appointment_perf["conversion_rate"].quantile(0.75)), 4
        ),
        "median": round(float(appointment_perf["conversion_rate"].median()), 4),
        "bottom_quartile": round(
            float(appointment_perf["conversion_rate"].quantile(0.25)), 4
        ),
        "best": round(float(appointment_perf["conversion_rate"].max()), 4),
        "worst": round(float(appointment_perf["conversion_rate"].min()), 4),
        "revenue_per_appointment_top_quartile": round(
            float(appointment_perf["revenue_per_interaction"].quantile(0.75)), 2
        ),
        "revenue_per_appointment_bottom_quartile": round(
            float(appointment_perf["revenue_per_interaction"].quantile(0.25)), 2
        ),
    }

    # ---- Campaigns ------------------------------------------------------ #
    channel = kpis.campaign_roas_by_channel(campaigns)
    summary["campaigns"] = {
        "blended_roas": round(kpis.campaign_roas(campaigns), 2),
        "by_channel": [
            {
                "channel": str(row.channel),
                "campaigns": int(row.campaigns),
                "spend": round(float(row.spend), 2),
                "attributed_net_revenue": round(float(row.attributed_net_revenue), 2),
                "roas": round(float(row.roas), 2),
            }
            for row in channel.itertuples(index=False)
        ],
    }

    # Full-price behaviour: campaign-influenced purchases versus the rest.
    sales = kpis.valid_sales(transactions)
    influenced = sales["campaign_id"].notna()
    summary["campaigns"]["full_price_share_campaign_influenced"] = round(
        float(
            sales.loc[influenced & sales["is_full_price"], "net_amount"].sum()
            / max(sales.loc[influenced, "net_amount"].sum(), 1)
        ),
        4,
    )
    summary["campaigns"]["full_price_share_non_influenced"] = round(
        float(
            sales.loc[~influenced & sales["is_full_price"], "net_amount"].sum()
            / max(sales.loc[~influenced, "net_amount"].sum(), 1)
        ),
        4,
    )

    return summary


def main() -> None:
    """Run every phase in order and write all outputs to ``data/processed``."""
    print("=" * 66)
    print("Luxury Retail CRM & Clienteling Analytics - full pipeline")
    print("=" * 66)

    config.ensure_directories()

    # ---- Phase 1 --------------------------------------------------------- #
    tables = generate_data.generate_all()
    generate_data.export_tables(tables)
    print(f"[1/5] synthetic data      : {len(tables['transactions']):,} transaction lines")

    transactions = tables["transactions"]
    customers = tables["customers"]

    # ---- Phase 2a -------------------------------------------------------- #
    segmented = customer_segmentation.build_segmentation(transactions, customers)
    segmented.to_csv(
        config.PROCESSED_DIR / config.ANALYTICS_FILES["rfm_segments"], index=False
    )
    print(f"[2/5] RFM segmentation    : {len(segmented):,} clients scored")

    # ---- Phase 2b -------------------------------------------------------- #
    counts, rates = cohort_analysis.cohort_retention_matrix(transactions, customers)
    counts.to_csv(
        config.PROCESSED_DIR / config.ANALYTICS_FILES["cohort_retention_counts"]
    )
    rates.round(4).to_csv(
        config.PROCESSED_DIR / config.ANALYTICS_FILES["cohort_retention_rates"]
    )
    print(f"[3/5] cohort retention    : {len(rates)} monthly cohorts")

    # ---- Phase 2c -------------------------------------------------------- #
    clv = customer_lifetime_value.build_clv(transactions, customers)
    clv.to_csv(
        config.PROCESSED_DIR / config.ANALYTICS_FILES["customer_lifetime_value"],
        index=False,
    )
    print(
        f"[4/5] 12-month CLV        : {config.CURRENCY_SYMBOL}"
        f"{clv['clv_12m'].sum():,.0f} total projected margin"
    )

    # ---- Summary --------------------------------------------------------- #
    summary = build_full_summary(tables)
    summary_path = config.PROCESSED_DIR / config.ANALYTICS_FILES["kpi_summary"]
    with summary_path.open("w", encoding="utf-8") as handle:
        json.dump(summary, handle, indent=2, default=_jsonable)
    print(f"[5/5] KPI summary         : {summary_path.name}")

    head = summary["headline"]
    print("-" * 66)
    print(f"  Net Sales            : {config.CURRENCY_SYMBOL}{head['net_sales']:,.0f}")
    print(f"  ATV                  : {config.CURRENCY_SYMBOL}{head['average_transaction_value']:,.0f}")
    print(f"  AUR                  : {config.CURRENCY_SYMBOL}{head['average_unit_retail']:,.0f}")
    print(f"  Repeat Purchase Rate : {head['repeat_purchase_rate']:.1%}")
    print(f"  Full-Price Share     : {head['full_price_sales_share']:.1%}")
    print(f"  Appointment Conv.    : {head['appointment_conversion_rate']:.1%}")
    print(f"  Blended Campaign ROAS: {head['blended_campaign_roas']:.2f}x")
    print("-" * 66)
    print("Done. Launch the dashboard with:")
    print("  streamlit run dashboard/streamlit_app.py")


if __name__ == "__main__":
    main()
