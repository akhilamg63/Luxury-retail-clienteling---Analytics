"""Tests for the analytics layer.

Three things matter most and are covered explicitly:

1. RFM scoring is arithmetically correct and the segmentation is exhaustive.
2. The CLV model returns the expected shape and economically sane values.
3. Net Sales excludes returned lines - the single most consequential rule in
   the project, because it flows into every monetary KPI.
"""

from __future__ import annotations

import datetime as dt

import numpy as np
import pandas as pd
import pytest

from src import (
    cohort_analysis,
    config,
    customer_lifetime_value,
    customer_segmentation,
    kpis,
)

# --------------------------------------------------------------------------- #
# Net sales / returns                                                          #
# --------------------------------------------------------------------------- #


def test_net_sales_excludes_returns(toy_transactions: pd.DataFrame) -> None:
    """Net Sales must sum only non-returned lines."""
    # 1000 + 2000 + 1200 + 400; the 500 and 9000 returns are excluded.
    assert kpis.net_sales(toy_transactions) == pytest.approx(4_600.0)


def test_net_sales_is_strictly_less_than_all_lines(
    toy_transactions: pd.DataFrame,
) -> None:
    """Returns exist in the fixture, so Net Sales must be below the raw total."""
    raw_total = float(toy_transactions["net_amount"].sum())
    assert kpis.net_sales(toy_transactions) < raw_total


def test_returning_every_line_yields_zero_net_sales(
    toy_transactions: pd.DataFrame,
) -> None:
    """If everything is returned, Net Sales collapses to zero - no leakage."""
    all_returned = toy_transactions.assign(is_return=True)
    assert kpis.net_sales(all_returned) == pytest.approx(0.0)
    assert kpis.order_count(all_returned) == 0
    assert kpis.average_transaction_value(all_returned) == pytest.approx(0.0)


def test_net_sales_excludes_returns_on_real_data(
    transactions: pd.DataFrame,
) -> None:
    """The same rule holds on the generated dataset."""
    returned_value = float(
        transactions.loc[transactions["is_return"], "net_amount"].sum()
    )
    assert returned_value > 0, "fixture should contain returns"
    assert kpis.net_sales(transactions) == pytest.approx(
        float(transactions["net_amount"].sum()) - returned_value
    )


def test_derived_sales_kpis(toy_transactions: pd.DataFrame) -> None:
    """ATV, AUR, repeat rate and full-price share follow from Net Sales."""
    assert kpis.order_count(toy_transactions) == 3
    assert kpis.units_sold(toy_transactions) == 5
    assert kpis.average_transaction_value(toy_transactions) == pytest.approx(4_600 / 3)
    assert kpis.average_unit_retail(toy_transactions) == pytest.approx(920.0)
    # C1 has 2 orders, C2 has 1, C3 has none that count -> 1 of 2 repeats.
    assert kpis.repeat_purchase_rate(toy_transactions) == pytest.approx(0.5)
    # Full-price kept lines: 1000 + 1200 + 400 = 2600 of 4600.
    assert kpis.full_price_sales_share(toy_transactions) == pytest.approx(2_600 / 4_600)


# --------------------------------------------------------------------------- #
# RFM scoring                                                                  #
# --------------------------------------------------------------------------- #


def test_rfm_values_are_computed_correctly(toy_transactions: pd.DataFrame) -> None:
    """Recency, Frequency and Monetary match hand-calculated values."""
    snapshot = dt.date(2026, 9, 30)
    rfm = customer_segmentation.compute_rfm(toy_transactions, snapshot).set_index(
        "customer_id"
    )

    # C3's only line was returned, so the client has no RFM profile at all.
    assert "C3" not in rfm.index

    # C1: last kept purchase 2026-06-15, two orders, 1000 + 2000 spent.
    assert rfm.loc["C1", "frequency"] == 2
    assert rfm.loc["C1", "monetary"] == pytest.approx(3_000.0)
    assert rfm.loc["C1", "recency_days"] == (snapshot - dt.date(2026, 6, 15)).days

    # C2: one order of two lines -> frequency 1, not 2.
    assert rfm.loc["C2", "frequency"] == 1
    assert rfm.loc["C2", "monetary"] == pytest.approx(1_600.0)
    assert rfm.loc["C2", "avg_order_value"] == pytest.approx(1_600.0)


def test_rfm_scores_are_in_range(segmented: pd.DataFrame) -> None:
    """All three scores sit on the documented 1-5 scale."""
    for column in ("r_score", "f_score", "m_score"):
        assert segmented[column].between(1, 5).all()
        assert segmented[column].dtype.kind in "iu"
    assert segmented["rfm_score"].between(3, 15).all()


def test_recency_score_is_inverted(segmented: pd.DataFrame) -> None:
    """A higher R score must mean a *lower* recency in days."""
    median_days = segmented.groupby("r_score")["recency_days"].median()
    assert median_days.is_monotonic_decreasing, median_days.to_dict()


def test_monetary_and_frequency_scores_are_ordered(segmented: pd.DataFrame) -> None:
    """Higher M and F scores must mean higher spend and more orders."""
    assert segmented.groupby("m_score")["monetary"].median().is_monotonic_increasing
    assert segmented.groupby("f_score")["frequency"].median().is_monotonic_increasing


def test_segmentation_is_exhaustive_and_exclusive(segmented: pd.DataFrame) -> None:
    """Every client gets exactly one of the four documented segments."""
    assert segmented["segment"].notna().all()
    assert set(segmented["segment"].unique()) <= set(
        customer_segmentation.SEGMENT_ORDER
    )
    assert segmented["customer_id"].is_unique


def test_lapsing_high_value_rule(segmented: pd.DataFrame) -> None:
    """Lapsing High-Value clients are high spenders who have gone quiet."""
    lapsing = segmented[segmented["segment"] == "Lapsing High-Value"]
    assert not lapsing.empty
    assert (lapsing["m_score"] >= config.RFM_HIGH_VALUE_MONETARY_SCORE).all()
    assert (lapsing["recency_days"] > config.LAPSED_RECENCY_DAYS).all()


def test_high_value_loyal_rule(segmented: pd.DataFrame) -> None:
    """High-Value Loyal clients are high spenders, frequent, and still active."""
    loyal = segmented[segmented["segment"] == "High-Value Loyal"]
    assert not loyal.empty
    assert (loyal["m_score"] >= config.RFM_HIGH_VALUE_MONETARY_SCORE).all()
    assert (loyal["f_score"] >= config.RFM_LOYAL_FREQUENCY_SCORE).all()
    assert (loyal["recency_days"] <= config.LAPSED_RECENCY_DAYS).all()


def test_promising_new_clients_are_actually_new(segmented: pd.DataFrame) -> None:
    """Promising New clients are recent acquisitions who are still engaged."""
    new = segmented[segmented["segment"] == "Promising New"]
    assert not new.empty
    assert (new["tenure_days"] <= config.NEW_CLIENT_TENURE_DAYS).all()
    assert (new["recency_days"] <= config.LAPSED_RECENCY_DAYS).all()


# --------------------------------------------------------------------------- #
# CLV                                                                          #
# --------------------------------------------------------------------------- #


def test_clv_output_shape(
    transactions: pd.DataFrame, customers: pd.DataFrame, segmented: pd.DataFrame
) -> None:
    """One row per scored client, with every documented column present."""
    clv = customer_lifetime_value.build_clv(transactions, customers)

    assert len(clv) == len(segmented)
    assert clv["customer_id"].is_unique
    assert set(clv["customer_id"]) == set(segmented["customer_id"])

    expected_columns = {
        "customer_id",
        "segment",
        "avg_order_value",
        "expected_orders_12m",
        "retention_rate",
        "predicted_revenue_12m",
        "gross_margin_rate",
        "clv_12m",
    }
    assert expected_columns <= set(clv.columns)


def test_clv_values_are_economically_sane(
    transactions: pd.DataFrame, customers: pd.DataFrame
) -> None:
    """CLV is finite, non-negative, and below the revenue it is derived from."""
    clv = customer_lifetime_value.build_clv(transactions, customers)

    assert np.isfinite(clv["clv_12m"]).all()
    assert (clv["clv_12m"] >= 0).all()
    assert clv["retention_rate"].between(0.05, 0.95).all()
    assert (clv["expected_orders_12m"] <= customer_lifetime_value.MAX_EXPECTED_ORDERS_12M).all()
    # CLV is a margin figure, so it must be below projected revenue.
    assert (clv["clv_12m"] <= clv["predicted_revenue_12m"]).all()


def test_clv_ranks_segments_as_expected(
    transactions: pd.DataFrame, customers: pd.DataFrame
) -> None:
    """High-Value Loyal clients must carry the highest average CLV."""
    clv = customer_lifetime_value.build_clv(transactions, customers)
    averages = clv.groupby("segment")["clv_12m"].mean()
    assert averages["High-Value Loyal"] == averages.max()
    assert averages["High-Value Loyal"] > averages["Infrequent"]


def test_gross_margin_is_plausible(transactions: pd.DataFrame) -> None:
    """Realised margin sits in a believable luxury band."""
    margin = customer_lifetime_value.realised_gross_margin(transactions)
    assert 0.50 < margin < 0.85


# --------------------------------------------------------------------------- #
# Cohorts and clienteling attribution                                          #
# --------------------------------------------------------------------------- #


def test_cohort_matrix_starts_at_full_retention(
    transactions: pd.DataFrame, customers: pd.DataFrame
) -> None:
    """Month 0 of every observed cohort is 100% by construction."""
    _, rates = cohort_analysis.cohort_retention_matrix(transactions, customers)
    month_zero = rates[0].dropna()
    assert not month_zero.empty
    assert np.allclose(month_zero.to_numpy(), 1.0)
    # Retention can never exceed the cohort it is measured against.
    assert (rates.fillna(0).to_numpy() <= 1.0 + 1e-9).all()


def test_clienteling_attribution_respects_the_window(
    transactions: pd.DataFrame, interactions: pd.DataFrame
) -> None:
    """No purchase outside the 14-day window may be attributed."""
    attributed = kpis.attribute_clienteling(transactions, interactions)

    # Exactly one row per interaction, however many orders followed it.
    assert len(attributed) == len(interactions)
    assert attributed["interaction_id"].is_unique

    converted = attributed[attributed["converted_in_window"]]
    assert converted["days_to_purchase"].between(
        0, config.CLIENTELING_ATTRIBUTION_WINDOW_DAYS
    ).all()
    # Last-touch: no order may be credited to more than one interaction, so the
    # total credited can never exceed the number of orders that exist.
    assert attributed["attributed_orders"].sum() <= kpis.order_count(transactions)
    # Conversion counts every qualifying order, attribution credits one
    # interaction per order, so conversion can only ever be the larger count.
    assert attributed["orders_in_window"].sum() >= attributed["attributed_orders"].sum()
    # Revenue credited to clienteling cannot exceed total Net Sales.
    assert attributed["attributed_revenue"].sum() <= kpis.net_sales(transactions)


def test_windowed_attribution_recovers_generated_conversions(
    transactions: pd.DataFrame, interactions: pd.DataFrame
) -> None:
    """The window rule finds at least the interactions known to have converted.

    The generator links each converted interaction to a real order, so a
    correct windowed join must recover them all. It will also claim some purely
    organic purchases that happen to land inside the window - that halo is a
    genuine property of time-window attribution, so the recovered rate is
    expected to sit above the true rate, never below it.
    """
    attributed = kpis.attribute_clienteling(transactions, interactions)
    truth = interactions["converted"].mean()
    measured = attributed["converted_in_window"].mean()
    assert measured >= truth
    assert measured - truth < 0.20, "conversion halo is implausibly large"


def test_pipeline_is_reproducible() -> None:
    """The same seed must produce an identical transaction table."""
    from src import generate_data

    first = generate_data.generate_all(seed=7)["transactions"]
    second = generate_data.generate_all(seed=7)["transactions"]
    pd.testing.assert_frame_equal(first, second)
