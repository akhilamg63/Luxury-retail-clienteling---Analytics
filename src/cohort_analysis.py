"""Phase 2b - Monthly retention cohort analysis.

Clients are grouped by the calendar month of their first purchase. For every
month afterwards we measure how many of that cohort bought again, which turns
"repeat purchase rate" from a single blended number into a view of whether
retention is improving or decaying for each successive intake of clients.

Luxury note: a zero in month 1 is not a failure. Clients in this category buy
seasonally, so the meaningful signals are the month-3, month-6 and month-12
columns, not the immediate follow-on month.

Run directly to write the cohort CSVs:

    python -m src.cohort_analysis
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from src import config, kpis


def build_cohort_table(
    transactions: pd.DataFrame,
    customers: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Tag every order with its client's acquisition cohort and period index.

    If ``customers`` is supplied, clients who joined before the analysis window
    are excluded. This matters: a long-standing client whose first *observed*
    purchase falls in October 2024 is not a new client, and treating them as
    one would load the earliest cohorts with proven repeat buyers and overstate
    retention. Left-censoring like this is the most common way a cohort chart
    ends up flattering the business.

    Args:
        transactions: Transaction lines from Phase 1.
        customers: Optional customer master used to restrict the analysis to
            clients genuinely acquired during the window.

    Returns:
        Order-level frame with ``cohort_month`` (month of the client's first
        purchase) and ``period_index`` (whole months elapsed since then).
    """
    sales = kpis.valid_sales(transactions).copy()
    if customers is not None:
        newly_acquired = customers.loc[
            customers["join_date"] >= pd.Timestamp(config.PERIOD_START), "customer_id"
        ]
        sales = sales.loc[sales["customer_id"].isin(set(newly_acquired))]
    sales["order_month"] = (
        sales["transaction_date"].dt.to_period("M").dt.to_timestamp()
    )

    # Collapse to order level first so a 3-line order is not counted 3 times.
    orders = sales.groupby(
        ["customer_id", "order_id", "order_month"], as_index=False
    )["net_amount"].sum()

    first_purchase = (
        orders.groupby("customer_id")["order_month"].min().rename("cohort_month")
    )
    orders = orders.merge(first_purchase, on="customer_id", how="left")

    # Whole months between the cohort month and the order month.
    orders["period_index"] = (
        (orders["order_month"].dt.year - orders["cohort_month"].dt.year) * 12
        + (orders["order_month"].dt.month - orders["cohort_month"].dt.month)
    ).astype(int)
    return orders


def cohort_retention_matrix(
    transactions: pd.DataFrame,
    customers: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Build the monthly retention cohort matrices.

    Args:
        transactions: Transaction lines from Phase 1.
        customers: Optional customer master; see :func:`build_cohort_table`.

    Returns:
        A tuple of ``(counts, rates)``:

        * ``counts`` - active clients per cohort per month offset.
        * ``rates``  - the same divided by the cohort's size at month 0, so
          every row starts at 1.0.

        Both are indexed by ``cohort_month`` with integer month-offset columns.
    """
    orders = build_cohort_table(transactions, customers)

    counts = (
        orders.groupby(["cohort_month", "period_index"])["customer_id"]
        .nunique()
        .unstack(fill_value=0)
        .sort_index()
    )

    cohort_sizes = counts.get(0, pd.Series(0, index=counts.index))
    rates = counts.div(cohort_sizes.replace(0, np.nan), axis=0)

    # Cells beyond the end of the observation window are not "zero retention",
    # they are unobserved - blank them so the heatmap does not mislead.
    max_offsets = _max_observable_offset(counts.index)
    for cohort, max_offset in max_offsets.items():
        unobservable = [c for c in rates.columns if c > max_offset]
        rates.loc[cohort, unobservable] = np.nan
        counts.loc[cohort, unobservable] = np.nan

    return counts, rates


def _max_observable_offset(cohort_months: pd.Index) -> dict[pd.Timestamp, int]:
    """Months of history actually observable for each cohort at the snapshot."""
    period_end = pd.Timestamp(config.PERIOD_END)
    return {
        cohort: (period_end.year - cohort.year) * 12
        + (period_end.month - cohort.month)
        for cohort in cohort_months
    }


def retention_curve(rates: pd.DataFrame) -> pd.Series:
    """Average retention rate at each month offset across all cohorts.

    Cohorts too young to have reached a given offset are excluded from that
    offset's average, which prevents recent cohorts dragging the curve down.
    """
    return rates.mean(axis=0, skipna=True)


def repeat_rate_within_months(
    transactions: pd.DataFrame,
    customers: pd.DataFrame | None = None,
    months: int = 12,
) -> float:
    """Share of activated clients who bought twice within ``months`` of joining.

    This is the empirical retention rate fed into the CLV model, so the CLV
    assumption is estimated from the data rather than asserted.

    The observation window is anchored on the client's **join date**, not on
    their first purchase. Anchoring on first purchase looks natural but is
    biased: requiring 12 months of history after the first order quietly drops
    every client whose single purchase happened late in their tenure, which is
    exactly the one-time-buyer population the metric is meant to capture. The
    result is a repeat rate that can overstate reality by 20 points or more.

    Denominator: clients who joined at least ``months`` before the snapshot and
    purchased at least once in their first ``months``. Numerator: those among
    them who purchased at least twice in the same period.

    Args:
        transactions: Transaction lines from Phase 1.
        customers: Customer master. Required - the metric cannot be computed
            without join dates.
        months: Observation horizon in whole months.

    Returns:
        A rate between 0 and 1.
    """
    if customers is None:
        return config.CLV_FALLBACK_RETENTION_RATE

    period_end = pd.Timestamp(config.PERIOD_END)
    # Clients acquired during the window with a full observation horizon.
    eligible_clients = customers.loc[
        (customers["join_date"] >= pd.Timestamp(config.PERIOD_START))
        & (customers["join_date"] + pd.DateOffset(months=months) <= period_end),
        ["customer_id", "join_date"],
    ]
    if eligible_clients.empty:
        return config.CLV_FALLBACK_RETENTION_RATE

    sales = kpis.valid_sales(transactions)
    orders = sales.groupby(
        ["customer_id", "order_id"], as_index=False
    )["transaction_date"].min()

    joined = orders.merge(eligible_clients, on="customer_id", how="inner")
    within = joined.loc[
        joined["transaction_date"]
        < joined["join_date"] + pd.DateOffset(months=months)
    ]

    orders_per_client = within.groupby("customer_id")["order_id"].nunique()
    activated = orders_per_client[orders_per_client >= 1]
    if activated.empty:
        return config.CLV_FALLBACK_RETENTION_RATE
    return float((activated >= 2).mean())


def main() -> None:
    """Build the cohort matrices and write them to ``data/processed``."""
    transactions = kpis.load_table("transactions")
    customers = kpis.load_table("customers")
    counts, rates = cohort_retention_matrix(transactions, customers)

    config.ensure_directories()
    counts.to_csv(config.PROCESSED_DIR / config.ANALYTICS_FILES["cohort_retention_counts"])
    rates.round(4).to_csv(
        config.PROCESSED_DIR / config.ANALYTICS_FILES["cohort_retention_rates"]
    )

    curve = retention_curve(rates)
    repeat_12m = repeat_rate_within_months(transactions, customers, 12)

    print("Phase 2b - cohort retention")
    print(f"  cohorts          : {len(rates)}")
    print(f"  12-month repeat  : {repeat_12m:.1%}")
    for offset in (1, 3, 6, 12):
        if offset in curve.index and not np.isnan(curve.loc[offset]):
            print(f"  month {offset:>2} retention: {curve.loc[offset]:.1%}")


if __name__ == "__main__":
    main()
