"""Phase 2c - 12-month customer lifetime value.

A deliberately transparent model. Every input is observable in the data or
declared as an assumption in ``src/config.py``, so a commercial stakeholder can
follow the arithmetic end to end:

    CLV(12m) = AOV x Expected Orders(12m) x Retention x Gross Margin % / (1 + d)^0.5

where

* **AOV**             - the client's own average order value;
* **Expected Orders** - their annualised purchase frequency, shrunk towards the
  segment mean so a single-order client is not projected from one data point;
* **Retention**       - the empirical 12-month repeat rate from the cohort
  model, adjusted for the client's recency score;
* **Gross Margin %**  - the realised blended margin from the product master;
* **d**               - a 10% annual cost of capital, applied at the half-year
  point because purchases arrive through the year rather than on day 365.

The output is a *margin* figure, not revenue. Spending to retain a client is a
margin decision, and quoting revenue-based CLV is the most common way these
models end up justifying uneconomic discounts.

Run directly to write ``data/processed/customer_lifetime_value.csv``:

    python -m src.customer_lifetime_value
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from src import cohort_analysis, config, customer_segmentation, kpis

# How a client's RFM recency score modulates the baseline retention rate.
# A client who has gone quiet is materially less likely to return; a highly
# active one is more likely. Documented assumption, not an observed value.
RECENCY_RETENTION_FACTOR: dict[int, float] = {
    1: 0.35,
    2: 0.60,
    3: 0.85,
    4: 1.05,
    5: 1.20,
}

# Weight applied to a client's own purchase rate versus their segment mean.
# Equivalent to a simple empirical-Bayes shrinkage with a prior worth one year
# of observation: a client with two years of history is trusted twice as much
# as one with a single year.
SHRINKAGE_PRIOR_YEARS: float = 1.0

# Upper bound on projected orders, so an outlier cannot distort the book.
MAX_EXPECTED_ORDERS_12M: float = 12.0


def realised_gross_margin(transactions: pd.DataFrame) -> float:
    """Blended gross margin actually realised after discount, excluding returns.

    Uses net revenue minus cost of goods, so markdown correctly erodes the
    margin rate rather than being hidden in a list-price assumption.

    Args:
        transactions: Transaction lines from Phase 1.

    Returns:
        Gross margin as a rate between 0 and 1.
    """
    sales = kpis.valid_sales(transactions)
    revenue = float(sales["net_amount"].sum())
    if revenue == 0:
        return 0.0
    cost = float((sales["unit_cost"] * sales["quantity"]).sum())
    return (revenue - cost) / revenue


def expected_orders_next_12m(rfm: pd.DataFrame) -> pd.Series:
    """Project each client's order count over the next 12 months.

    The client's own annualised rate is blended with their segment's mean rate,
    weighted by how much history supports it. Tenure is floored at six months
    so a client who joined last week is not projected to buy 50 times a year.

    Args:
        rfm: Segmented RFM frame from :mod:`src.customer_segmentation`.

    Returns:
        Expected orders over the next 12 months, indexed like ``rfm``.
    """
    tenure_years = (rfm["tenure_days"] / 365.25).clip(lower=0.5)
    own_rate = rfm["frequency"] / tenure_years

    segment_rate = own_rate.groupby(rfm["segment"]).transform("mean")

    # Shrinkage weight: more observed history -> trust the client's own rate.
    weight = tenure_years / (tenure_years + SHRINKAGE_PRIOR_YEARS)
    blended = weight * own_rate + (1 - weight) * segment_rate
    return blended.clip(upper=MAX_EXPECTED_ORDERS_12M)


def estimate_clv(
    rfm: pd.DataFrame,
    base_retention_rate: float,
    gross_margin_rate: float,
    horizon_months: int = config.CLV_HORIZON_MONTHS,
    annual_discount_rate: float = config.CLV_ANNUAL_DISCOUNT_RATE,
) -> pd.DataFrame:
    """Estimate 12-month CLV for every scored client.

    Args:
        rfm: Segmented RFM frame with ``avg_order_value``, ``frequency``,
            ``tenure_days``, ``r_score`` and ``segment`` columns.
        base_retention_rate: Empirical 12-month repeat rate from the cohort
            model.
        gross_margin_rate: Realised blended gross margin.
        horizon_months: Projection horizon. Defaults to 12.
        annual_discount_rate: Cost of capital used to discount future margin.

    Returns:
        One row per client with the model inputs and the resulting
        ``predicted_revenue_12m`` and ``clv_12m``.
    """
    horizon_years = horizon_months / 12.0
    clv = rfm[
        [
            "customer_id",
            "segment",
            "frequency",
            "monetary",
            "avg_order_value",
            "recency_days",
            "tenure_days",
            "r_score",
            "f_score",
            "m_score",
        ]
    ].copy()

    clv["expected_orders_12m"] = expected_orders_next_12m(rfm) * horizon_years

    # Retention: baseline from the cohort model, tilted by recency behaviour.
    recency_factor = clv["r_score"].map(RECENCY_RETENTION_FACTOR).astype(float)
    clv["retention_rate"] = (base_retention_rate * recency_factor).clip(0.05, 0.95)

    clv["predicted_revenue_12m"] = (
        clv["avg_order_value"] * clv["expected_orders_12m"] * clv["retention_rate"]
    )

    clv["gross_margin_rate"] = gross_margin_rate
    # Discount at the half-horizon point: purchases arrive through the year.
    discount_factor = (1 + annual_discount_rate) ** (horizon_years / 2)
    clv["clv_12m"] = (
        clv["predicted_revenue_12m"] * gross_margin_rate / discount_factor
    )

    numeric = [
        "expected_orders_12m",
        "retention_rate",
        "predicted_revenue_12m",
        "clv_12m",
    ]
    clv[numeric] = clv[numeric].round(2)
    return clv.sort_values("clv_12m", ascending=False).reset_index(drop=True)


def clv_by_segment(clv: pd.DataFrame) -> pd.DataFrame:
    """Aggregate the CLV book by segment for executive reporting."""
    summary = (
        clv.groupby("segment")
        .agg(
            clients=("customer_id", "count"),
            avg_clv_12m=("clv_12m", "mean"),
            total_clv_12m=("clv_12m", "sum"),
            avg_order_value=("avg_order_value", "mean"),
            avg_expected_orders=("expected_orders_12m", "mean"),
            avg_retention_rate=("retention_rate", "mean"),
        )
        .reset_index()
    )
    summary["clv_share"] = summary["total_clv_12m"] / summary["total_clv_12m"].sum()
    summary["segment"] = pd.Categorical(
        summary["segment"],
        categories=customer_segmentation.SEGMENT_ORDER,
        ordered=True,
    )
    return summary.sort_values("segment").reset_index(drop=True)


def build_clv(
    transactions: pd.DataFrame,
    customers: pd.DataFrame,
) -> pd.DataFrame:
    """Run the full CLV model from the raw Phase 1 tables.

    Args:
        transactions: Transaction lines.
        customers: Customer master.

    Returns:
        The per-client CLV frame.
    """
    rfm = customer_segmentation.build_segmentation(transactions, customers)
    retention = cohort_analysis.repeat_rate_within_months(
        transactions, customers, config.CLV_HORIZON_MONTHS
    )
    margin = realised_gross_margin(transactions)
    return estimate_clv(rfm, retention, margin)


def main() -> None:
    """Estimate CLV and write ``data/processed/customer_lifetime_value.csv``."""
    transactions = kpis.load_table("transactions")
    customers = kpis.load_table("customers")

    retention = cohort_analysis.repeat_rate_within_months(
        transactions, customers, config.CLV_HORIZON_MONTHS
    )
    margin = realised_gross_margin(transactions)
    clv = build_clv(transactions, customers)

    config.ensure_directories()
    clv.to_csv(
        config.PROCESSED_DIR / config.ANALYTICS_FILES["customer_lifetime_value"],
        index=False,
    )

    summary = clv_by_segment(clv)
    print("Phase 2c - 12-month CLV")
    print(f"  base retention rate : {retention:.1%} (from cohort model)")
    print(f"  realised gross margin: {margin:.1%}")
    print(f"  discount rate       : {config.CLV_ANNUAL_DISCOUNT_RATE:.0%} p.a.")
    for row in summary.itertuples(index=False):
        print(
            f"  {row.segment:<20} {row.clients:>5,} clients  "
            f"avg {config.CURRENCY_SYMBOL}{row.avg_clv_12m:>8,.0f}  "
            f"total {config.CURRENCY_SYMBOL}{row.total_clv_12m:>12,.0f}"
        )
    print(
        f"  total 12m CLV       : {config.CURRENCY_SYMBOL}"
        f"{clv['clv_12m'].sum():,.0f}"
    )


if __name__ == "__main__":
    main()
