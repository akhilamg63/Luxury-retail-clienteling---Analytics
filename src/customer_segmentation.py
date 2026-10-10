"""Phase 2a - RFM scoring and luxury client segmentation.

Standard RFM is adapted to a luxury context in two ways:

1. **Recency is measured in months, not days.** A client who last bought six
   months ago is entirely normal in a category where people buy two or three
   times a year, so the lapsing threshold is set at 180 days rather than the
   30-60 days typical of mass-market retail.
2. **Monetary value dominates.** A single client buying one exceptional piece
   can out-value fifty accessory purchases, so the segmentation leads with the
   Monetary score and uses Frequency to separate loyalty from one-off spend.

Run directly to write ``data/processed/rfm_segments.csv``:

    python -m src.customer_segmentation
"""

from __future__ import annotations

import datetime as dt

import numpy as np
import pandas as pd

from src import config, kpis

# The four segments the specification asks for, in reporting order.
SEGMENT_ORDER: list[str] = [
    "High-Value Loyal",
    "Lapsing High-Value",
    "Promising New",
    "Infrequent",
]


def _score_quintiles(series: pd.Series, ascending: bool = True) -> pd.Series:
    """Score a series into 1-5 quintiles, 5 always being the best outcome.

    ``pd.qcut`` is used with ``duplicates="drop"`` so that heavily tied
    distributions (common in Frequency, where most luxury clients have 1-3
    orders) degrade gracefully instead of raising.

    Args:
        series: Values to score.
        ascending: True when a higher raw value is better (Frequency,
            Monetary); False when a lower raw value is better (Recency).

    Returns:
        Integer scores from 1 to 5.
    """
    ranked = series.rank(method="first", ascending=ascending)
    try:
        scores = pd.qcut(ranked, q=5, labels=[1, 2, 3, 4, 5], duplicates="drop")
        return scores.astype(int)
    except ValueError:
        # Fewer distinct values than quintiles - fall back to a linear stretch.
        normalised = (ranked - ranked.min()) / max(ranked.max() - ranked.min(), 1)
        return (np.ceil(normalised * 5).clip(1, 5)).astype(int)


def compute_rfm(
    transactions: pd.DataFrame,
    snapshot_date: dt.date = config.SNAPSHOT_DATE,
) -> pd.DataFrame:
    """Compute Recency, Frequency and Monetary values and their 1-5 scores.

    Definitions:
        * **Recency**  - days between the snapshot date and the client's most
          recent non-returned purchase.
        * **Frequency** - count of distinct orders containing a non-returned
          line. Order-level, not line-level.
        * **Monetary**  - net sales after discount, excluding returns.

    Clients in the database who have never purchased are excluded: they have no
    RFM profile and belong in an acquisition workflow, not a retention one.

    Args:
        transactions: Transaction lines from Phase 1.
        snapshot_date: The "as at" date recency is measured against.

    Returns:
        One row per purchasing client with raw RFM values, 1-5 scores, a
        concatenated ``rfm_cell`` string and a summed ``rfm_score``.
    """
    sales = kpis.valid_sales(transactions)
    snapshot = pd.Timestamp(snapshot_date)

    rfm = (
        sales.groupby("customer_id")
        .agg(
            last_purchase_date=("transaction_date", "max"),
            first_purchase_date=("transaction_date", "min"),
            frequency=("order_id", "nunique"),
            monetary=("net_amount", "sum"),
            units=("quantity", "sum"),
        )
        .reset_index()
    )

    rfm["recency_days"] = (snapshot - rfm["last_purchase_date"]).dt.days
    rfm["tenure_days"] = (snapshot - rfm["first_purchase_date"]).dt.days
    rfm["avg_order_value"] = rfm["monetary"] / rfm["frequency"]

    # Lower recency is better, so score it descending.
    rfm["r_score"] = _score_quintiles(rfm["recency_days"], ascending=False)
    rfm["f_score"] = _score_quintiles(rfm["frequency"], ascending=True)
    rfm["m_score"] = _score_quintiles(rfm["monetary"], ascending=True)

    rfm["rfm_cell"] = (
        rfm["r_score"].astype(str)
        + rfm["f_score"].astype(str)
        + rfm["m_score"].astype(str)
    )
    rfm["rfm_score"] = rfm["r_score"] + rfm["f_score"] + rfm["m_score"]

    return rfm.sort_values("monetary", ascending=False).reset_index(drop=True)


def assign_segments(rfm: pd.DataFrame) -> pd.DataFrame:
    """Assign each client to one of the four luxury segments.

    Rules are applied in strict precedence order, so every client receives
    exactly one label:

    1. **Lapsing High-Value** - ``M >= 4`` and no purchase for more than 180
       days. Checked first: a high-value client going quiet is the single most
       commercially urgent state in the book, and must not be masked by any
       other label.
    2. **High-Value Loyal** - ``M >= 4`` and ``F >= 4`` and active within 180
       days. Proven spend plus proven repeat behaviour.
    3. **Promising New** - first purchase within the last 365 days and active
       within 180 days. Too new to have accumulated frequency, but engaged.
    4. **Infrequent** - everyone else: low spend, low frequency, or a lapsed
       client who was never high-value.

    Thresholds are defined in ``src/config.py`` and documented in the README.

    Args:
        rfm: Output of :func:`compute_rfm`.

    Returns:
        The input frame with a ``segment`` column added.
    """
    rfm = rfm.copy()

    is_high_value = rfm["m_score"] >= config.RFM_HIGH_VALUE_MONETARY_SCORE
    is_loyal = rfm["f_score"] >= config.RFM_LOYAL_FREQUENCY_SCORE
    is_lapsed = rfm["recency_days"] > config.LAPSED_RECENCY_DAYS
    is_new = rfm["tenure_days"] <= config.NEW_CLIENT_TENURE_DAYS

    # np.select applies the first matching condition, which gives us the
    # precedence order above for free.
    rfm["segment"] = np.select(
        condlist=[
            is_high_value & is_lapsed,
            is_high_value & is_loyal & ~is_lapsed,
            is_new & ~is_lapsed,
        ],
        choicelist=[
            "Lapsing High-Value",
            "High-Value Loyal",
            "Promising New",
        ],
        default="Infrequent",
    )
    return rfm


def segment_summary(rfm: pd.DataFrame) -> pd.DataFrame:
    """Aggregate clients, revenue and behaviour by segment for reporting."""
    summary = (
        rfm.groupby("segment")
        .agg(
            clients=("customer_id", "count"),
            net_sales=("monetary", "sum"),
            avg_order_value=("avg_order_value", "mean"),
            avg_frequency=("frequency", "mean"),
            avg_recency_days=("recency_days", "mean"),
        )
        .reset_index()
    )
    summary["client_share"] = summary["clients"] / summary["clients"].sum()
    summary["revenue_share"] = summary["net_sales"] / summary["net_sales"].sum()

    # Keep a stable, commercially meaningful reporting order.
    summary["segment"] = pd.Categorical(
        summary["segment"], categories=SEGMENT_ORDER, ordered=True
    )
    return summary.sort_values("segment").reset_index(drop=True)


def build_segmentation(
    transactions: pd.DataFrame,
    customers: pd.DataFrame,
    snapshot_date: dt.date = config.SNAPSHOT_DATE,
) -> pd.DataFrame:
    """Run the full segmentation and enrich it with client master attributes.

    Args:
        transactions: Transaction lines from Phase 1.
        customers: Customer master from Phase 1.
        snapshot_date: The "as at" date recency is measured against.

    Returns:
        One row per purchasing client, ready for export and for the dashboard's
        lapsed-client outreach table.
    """
    rfm = assign_segments(compute_rfm(transactions, snapshot_date))

    master_columns = [
        "customer_id",
        "client_name",
        "join_date",
        "acquisition_source",
        "preferred_boutique_id",
        "consent_status",
        "region",
        "country",
    ]
    enriched = rfm.merge(customers[master_columns], on="customer_id", how="left")
    return enriched


def main() -> None:
    """Compute segments and write ``data/processed/rfm_segments.csv``."""
    transactions = kpis.load_table("transactions")
    customers = kpis.load_table("customers")

    segmented = build_segmentation(transactions, customers)
    config.ensure_directories()
    segmented.to_csv(
        config.PROCESSED_DIR / config.ANALYTICS_FILES["rfm_segments"], index=False
    )

    summary = segment_summary(segmented)
    print("Phase 2a - RFM segmentation")
    print(f"  clients scored: {len(segmented):,}")
    for row in summary.itertuples(index=False):
        print(
            f"  {row.segment:<20} {row.clients:>5,} clients "
            f"({row.client_share:>6.1%})  "
            f"{config.CURRENCY_SYMBOL}{row.net_sales:>12,.0f} "
            f"({row.revenue_share:>6.1%} of revenue)"
        )


if __name__ == "__main__":
    main()
