"""Shared pytest fixtures.

The generator takes ~17 seconds, so it runs at most once per test session. If
the pipeline has already been run, the fixtures read the CSVs from
``data/processed`` instead, which keeps the suite fast in normal use.
"""

from __future__ import annotations

import pandas as pd
import pytest

from src import config, customer_segmentation, generate_data, kpis


@pytest.fixture(scope="session")
def tables() -> dict[str, pd.DataFrame]:
    """The seven Phase 1 tables, loaded from disk or generated on demand."""
    transactions_path = config.PROCESSED_DIR / config.RAW_TABLE_FILES["transactions"]
    if transactions_path.exists():
        return kpis.load_all_tables()
    return generate_data.generate_all()


@pytest.fixture(scope="session")
def transactions(tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Transaction lines."""
    return tables["transactions"]


@pytest.fixture(scope="session")
def customers(tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Customer master."""
    return tables["customers"]


@pytest.fixture(scope="session")
def interactions(tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Clienteling activity log."""
    return tables["clienteling_interactions"]


@pytest.fixture(scope="session")
def segmented(
    transactions: pd.DataFrame, customers: pd.DataFrame
) -> pd.DataFrame:
    """RFM frame with segments assigned."""
    return customer_segmentation.build_segmentation(transactions, customers)


@pytest.fixture()
def toy_transactions() -> pd.DataFrame:
    """A tiny hand-built transaction table with known, checkable answers.

    Three clients:
      * C1 - two orders, one line returned.
      * C2 - one order, two lines, full price.
      * C3 - one order, fully returned (contributes nothing to Net Sales).
    """
    rows = [
        # C1, order 1: kept
        ("T1", "O1", "2026-01-10", "C1", 1, 1000.0, 0.0, 1000.0, 1000.0, False, True),
        # C1, order 2: one kept line, one returned line
        ("T2", "O2", "2026-06-15", "C1", 1, 2000.0, 0.2, 2500.0, 2000.0, False, False),
        ("T3", "O2", "2026-06-15", "C1", 1, 500.0, 0.0, 500.0, 500.0, True, True),
        # C2, order 3: two kept lines
        ("T4", "O3", "2026-09-01", "C2", 2, 600.0, 0.0, 1200.0, 1200.0, False, True),
        ("T5", "O3", "2026-09-01", "C2", 1, 400.0, 0.0, 400.0, 400.0, False, True),
        # C3, order 4: entirely returned
        ("T6", "O4", "2025-02-01", "C3", 1, 9000.0, 0.0, 9000.0, 9000.0, True, True),
    ]
    frame = pd.DataFrame(
        rows,
        columns=[
            "transaction_id",
            "order_id",
            "transaction_date",
            "customer_id",
            "quantity",
            "unit_net_price",
            "discount_pct",
            "gross_amount",
            "net_amount",
            "is_return",
            "is_full_price",
        ],
    )
    frame["transaction_date"] = pd.to_datetime(frame["transaction_date"])
    frame["unit_cost"] = frame["unit_net_price"] * 0.3
    return frame
