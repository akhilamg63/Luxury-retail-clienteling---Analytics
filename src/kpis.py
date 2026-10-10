"""Shared KPI definitions and data access.

Every headline metric in this project is defined exactly once, here. The
analytics modules, the Streamlit dashboard and the pytest suite all import
these functions, so a number shown on the dashboard cannot drift from the
number asserted in a test or quoted in the README.

The definitions mirror the Power BI DAX measures in ``bi/measures.dax`` and
the Tableau calculated fields in ``bi/calculated_fields.md`` one-for-one.
"""

from __future__ import annotations

import datetime as dt
from typing import Final

import pandas as pd

from src import config

# Columns that must be parsed as dates when a table is read back from CSV.
_DATE_COLUMNS: Final[dict[str, list[str]]] = {
    "boutiques": ["opening_date"],
    "associates": [],
    "products": [],
    "customers": ["join_date", "churn_date"],
    "transactions": ["transaction_date"],
    "clienteling_interactions": ["interaction_date"],
    "campaigns": ["start_date", "end_date"],
}


# --------------------------------------------------------------------------- #
# Data access                                                                  #
# --------------------------------------------------------------------------- #


def load_table(name: str) -> pd.DataFrame:
    """Load one generated table from ``data/processed`` by logical name.

    Args:
        name: Key from ``config.RAW_TABLE_FILES`` (e.g. ``"transactions"``).

    Returns:
        The table with date columns parsed.

    Raises:
        FileNotFoundError: If the pipeline has not been run yet.
    """
    if name not in config.RAW_TABLE_FILES:
        raise KeyError(f"Unknown table '{name}'. Expected one of {list(config.RAW_TABLE_FILES)}")

    path = config.PROCESSED_DIR / config.RAW_TABLE_FILES[name]
    if not path.exists():
        raise FileNotFoundError(
            f"{path} not found. Run `python -m src.run_pipeline` first."
        )
    return pd.read_csv(path, parse_dates=_DATE_COLUMNS[name])


def load_all_tables() -> dict[str, pd.DataFrame]:
    """Load every generated table into a dictionary keyed by table name."""
    return {name: load_table(name) for name in config.RAW_TABLE_FILES}


# --------------------------------------------------------------------------- #
# Core sales KPIs                                                              #
# --------------------------------------------------------------------------- #


def valid_sales(transactions: pd.DataFrame) -> pd.DataFrame:
    """Return only the lines that count as sales: everything that is not a return.

    This single filter is the backbone of every monetary KPI in the project.
    Returned lines stay in the dataset so that return behaviour can be analysed,
    but they never contribute to Net Sales.
    """
    return transactions.loc[~transactions["is_return"].astype(bool)]


def net_sales(transactions: pd.DataFrame) -> float:
    """Net Sales = sum of line value after discount, excluding returned lines."""
    return float(valid_sales(transactions)["net_amount"].sum())


def gross_sales(transactions: pd.DataFrame) -> float:
    """Gross Sales = sum of line value at full list price, excluding returns."""
    return float(valid_sales(transactions)["gross_amount"].sum())


def units_sold(transactions: pd.DataFrame) -> int:
    """Total units sold, excluding returned lines."""
    return int(valid_sales(transactions)["quantity"].sum())


def order_count(transactions: pd.DataFrame) -> int:
    """Number of distinct orders containing at least one non-returned line."""
    return int(valid_sales(transactions)["order_id"].nunique())


def average_transaction_value(transactions: pd.DataFrame) -> float:
    """ATV = Net Sales / number of distinct orders.

    Measured at order level, not line level: a client buying a coat and a scarf
    in one visit is one transaction.
    """
    orders = order_count(transactions)
    return net_sales(transactions) / orders if orders else 0.0


def average_unit_retail(transactions: pd.DataFrame) -> float:
    """AUR = Net Sales / units sold - the realised price per item."""
    units = units_sold(transactions)
    return net_sales(transactions) / units if units else 0.0


def repeat_purchase_rate(transactions: pd.DataFrame) -> float:
    """Share of purchasing clients who placed two or more orders in the period.

    Only clients with at least one non-returned order are in the denominator,
    so prospects who never bought do not depress the metric.
    """
    sales = valid_sales(transactions)
    if sales.empty:
        return 0.0
    orders_per_client = sales.groupby("customer_id")["order_id"].nunique()
    return float((orders_per_client >= 2).mean())


def full_price_sales_share(transactions: pd.DataFrame) -> float:
    """Share of Net Sales transacted at full list price (zero discount).

    Measured by value, not by line count - the commercially meaningful question
    is how much revenue is protected from markdown, not how many tickets were.
    """
    sales = valid_sales(transactions)
    total = float(sales["net_amount"].sum())
    if total == 0:
        return 0.0
    full_price = float(sales.loc[sales["is_full_price"].astype(bool), "net_amount"].sum())
    return full_price / total


def discount_rate(transactions: pd.DataFrame) -> float:
    """Average markdown depth = 1 - (Net Sales / Gross Sales at list price)."""
    gross = gross_sales(transactions)
    return 1.0 - (net_sales(transactions) / gross) if gross else 0.0


def return_rate_by_value(transactions: pd.DataFrame) -> float:
    """Returned value as a share of all transacted value (returns included)."""
    total = float(transactions["net_amount"].sum())
    if total == 0:
        return 0.0
    returned = float(
        transactions.loc[transactions["is_return"].astype(bool), "net_amount"].sum()
    )
    return returned / total


# --------------------------------------------------------------------------- #
# Clienteling KPIs                                                             #
# --------------------------------------------------------------------------- #


def attribute_clienteling(
    transactions: pd.DataFrame,
    interactions: pd.DataFrame,
    window_days: int = config.CLIENTELING_ATTRIBUTION_WINDOW_DAYS,
) -> pd.DataFrame:
    """Attribute purchases to clienteling interactions using a time window.

    A purchase is credited to an interaction when it is made by the same client
    within ``window_days`` days *after* the interaction. Where several
    interactions qualify, the most recent one wins (last-touch).

    This deliberately uses a windowed join rather than the ``interaction_id``
    key present in the generated data, because that is what a real CRM team can
    actually do in Power BI, Tableau or Excel. Using the same rule everywhere
    keeps all four layers of this project consistent.

    Args:
        transactions: Transaction lines.
        interactions: Clienteling activity log.
        window_days: Attribution window in days.

    Returns:
        One row per interaction with the attributed order, associate and
        revenue, plus a boolean ``converted_in_window`` flag.
    """
    sales = valid_sales(transactions)
    order_level = (
        sales.groupby(["order_id", "customer_id", "transaction_date"], as_index=False)[
            "net_amount"
        ]
        .sum()
        .rename(columns={"net_amount": "order_net_amount"})
    )

    merged = interactions.merge(order_level, on="customer_id", how="left")
    delta = (merged["transaction_date"] - merged["interaction_date"]).dt.days
    in_window = delta.between(0, window_days)
    merged = merged.loc[in_window].copy()
    merged["days_to_purchase"] = delta.loc[in_window]

    # Conversion and attribution answer two different questions, and collapsing
    # them into one number is a common and costly mistake:
    #
    #   Conversion  - "did this interaction lead to a purchase?" Every
    #                 interaction followed by a qualifying order counts, because
    #                 the question is about the interaction's effectiveness.
    #   Attribution - "which interaction earns the revenue?" Exactly one, or the
    #                 same order is banked several times over.
    #
    # Measuring conversion with the attribution rule understates it: an
    # appointment that genuinely worked loses its conversion if the client
    # happens to receive a courtesy call before collecting the piece.
    conversion = (
        merged.groupby("interaction_id")
        .agg(
            orders_in_window=("order_id", "nunique"),
            days_to_purchase=("days_to_purchase", "min"),
        )
        .reset_index()
    )

    # Last-touch attribution: each order is credited to the most recent
    # qualifying interaction, and to that one only.
    last_touch = merged.sort_values(
        ["order_id", "interaction_date"], ascending=[True, False]
    ).drop_duplicates(subset=["order_id"], keep="first")
    credited = (
        last_touch.groupby("interaction_id")
        .agg(
            attributed_orders=("order_id", "nunique"),
            attributed_revenue=("order_net_amount", "sum"),
        )
        .reset_index()
    )

    attributed = interactions.merge(conversion, on="interaction_id", how="left")
    attributed = attributed.merge(credited, on="interaction_id", how="left")

    for column in ("orders_in_window", "attributed_orders"):
        attributed[column] = attributed[column].fillna(0).astype(int)
    attributed["attributed_revenue"] = attributed["attributed_revenue"].fillna(0.0)
    attributed["converted_in_window"] = attributed["orders_in_window"] > 0
    return attributed


def clienteling_conversion_rate(
    attributed_interactions: pd.DataFrame,
    interaction_type: str | None = None,
) -> float:
    """Share of clienteling interactions followed by a purchase in the window.

    Args:
        attributed_interactions: Output of :func:`attribute_clienteling`.
        interaction_type: Optional filter, e.g. ``"Appointment"``.
    """
    frame = attributed_interactions
    if interaction_type is not None:
        frame = frame.loc[frame["interaction_type"] == interaction_type]
    if frame.empty:
        return 0.0
    return float(frame["converted_in_window"].mean())


def associate_performance(
    attributed_interactions: pd.DataFrame, associates: pd.DataFrame
) -> pd.DataFrame:
    """Rank sales associates by clienteling volume, conversion and revenue."""
    summary = (
        attributed_interactions.groupby("associate_id")
        .agg(
            interactions=("interaction_id", "count"),
            appointments=(
                "interaction_type",
                lambda s: int((s == "Appointment").sum()),
            ),
            conversions=("converted_in_window", "sum"),
            attributed_revenue=("attributed_revenue", "sum"),
        )
        .reset_index()
    )
    summary["conversion_rate"] = summary["conversions"] / summary["interactions"]
    summary["revenue_per_interaction"] = (
        summary["attributed_revenue"] / summary["interactions"]
    )
    summary = summary.merge(
        associates[["associate_id", "associate_name", "boutique_id", "role"]],
        on="associate_id",
        how="left",
    )
    return summary.sort_values("attributed_revenue", ascending=False).reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Campaign KPIs                                                                #
# --------------------------------------------------------------------------- #


def campaign_roas(campaigns: pd.DataFrame) -> float:
    """Blended ROAS = total attributed net revenue / total campaign spend."""
    spend = float(campaigns["spend"].sum())
    return float(campaigns["attributed_net_revenue"].sum()) / spend if spend else 0.0


def campaign_roas_by_channel(campaigns: pd.DataFrame) -> pd.DataFrame:
    """ROAS, spend and attributed revenue aggregated by marketing channel."""
    summary = (
        campaigns.groupby("channel")
        .agg(
            campaigns=("campaign_id", "count"),
            spend=("spend", "sum"),
            attributed_net_revenue=("attributed_net_revenue", "sum"),
            attributed_orders=("attributed_orders", "sum"),
        )
        .reset_index()
    )
    summary["roas"] = summary["attributed_net_revenue"] / summary["spend"]
    return summary.sort_values("roas", ascending=False).reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Summary                                                                      #
# --------------------------------------------------------------------------- #


def monthly_revenue(transactions: pd.DataFrame) -> pd.DataFrame:
    """Net Sales, orders and ATV by calendar month."""
    sales = valid_sales(transactions).copy()
    sales["month"] = sales["transaction_date"].dt.to_period("M").dt.to_timestamp()
    monthly = (
        sales.groupby("month")
        .agg(
            net_sales=("net_amount", "sum"),
            orders=("order_id", "nunique"),
            units=("quantity", "sum"),
        )
        .reset_index()
    )
    monthly["atv"] = monthly["net_sales"] / monthly["orders"]
    return monthly


def build_kpi_summary(
    transactions: pd.DataFrame,
    interactions: pd.DataFrame,
    campaigns: pd.DataFrame,
    customers: pd.DataFrame,
) -> dict[str, float | int | str]:
    """Compute every headline KPI in one pass.

    The result is serialised to ``data/processed/kpi_summary.json`` by the
    pipeline, and is the single source of the numbers quoted in the README.
    """
    attributed = attribute_clienteling(transactions, interactions)
    sales = valid_sales(transactions)

    return {
        "snapshot_date": config.SNAPSHOT_DATE.isoformat(),
        "period_start": config.PERIOD_START.isoformat(),
        "period_end": config.PERIOD_END.isoformat(),
        "currency": config.CURRENCY,
        "transaction_lines": int(len(transactions)),
        "orders": order_count(transactions),
        "customers_in_database": int(len(customers)),
        "purchasing_customers": int(sales["customer_id"].nunique()),
        "net_sales": round(net_sales(transactions), 2),
        "gross_sales": round(gross_sales(transactions), 2),
        "units_sold": units_sold(transactions),
        "average_transaction_value": round(average_transaction_value(transactions), 2),
        "average_unit_retail": round(average_unit_retail(transactions), 2),
        "repeat_purchase_rate": round(repeat_purchase_rate(transactions), 4),
        "full_price_sales_share": round(full_price_sales_share(transactions), 4),
        "average_markdown_depth": round(discount_rate(transactions), 4),
        "return_rate_by_value": round(return_rate_by_value(transactions), 4),
        "clienteling_interactions": int(len(interactions)),
        "appointment_conversion_rate": round(
            clienteling_conversion_rate(attributed, "Appointment"), 4
        ),
        "call_conversion_rate": round(clienteling_conversion_rate(attributed, "Call"), 4),
        "event_conversion_rate": round(clienteling_conversion_rate(attributed, "Event"), 4),
        "clienteling_attributed_revenue": round(
            float(attributed["attributed_revenue"].sum()), 2
        ),
        "campaign_spend": round(float(campaigns["spend"].sum()), 2),
        "campaign_attributed_revenue": round(
            float(campaigns["attributed_net_revenue"].sum()), 2
        ),
        "blended_campaign_roas": round(campaign_roas(campaigns), 2),
    }
