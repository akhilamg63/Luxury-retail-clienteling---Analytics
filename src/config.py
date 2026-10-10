"""Central configuration for the Luxury Retail CRM & Clienteling Analytics project.

Every path is derived from the repository root at runtime, so the project can be
cloned to any location and run without editing a single line of code.

All business assumptions that the source specification left open (segment
thresholds, the CLV retention assumption, the clienteling attribution window)
are declared here once and documented in the README.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Final

# --------------------------------------------------------------------------- #
# Paths                                                                        #
# --------------------------------------------------------------------------- #

# src/config.py -> src/ -> repository root
REPO_ROOT: Final[Path] = Path(__file__).resolve().parents[1]

DATA_DIR: Final[Path] = REPO_ROOT / "data"
RAW_DIR: Final[Path] = DATA_DIR / "raw"
PROCESSED_DIR: Final[Path] = DATA_DIR / "processed"

EXCEL_WORKBOOK_PATH: Final[Path] = (
    PROCESSED_DIR / "luxury_retail_clienteling_analytics.xlsx"
)

# Files written by Phase 1 (generate_data.py)
RAW_TABLE_FILES: Final[dict[str, str]] = {
    "boutiques": "boutiques.csv",
    "associates": "associates.csv",
    "products": "products.csv",
    "customers": "customers.csv",
    "transactions": "transactions.csv",
    "clienteling_interactions": "clienteling_interactions.csv",
    "campaigns": "campaigns.csv",
}

# Files written by Phase 2 (analytics modules)
ANALYTICS_FILES: Final[dict[str, str]] = {
    "rfm_segments": "rfm_segments.csv",
    "cohort_retention_rates": "cohort_retention_rates.csv",
    "cohort_retention_counts": "cohort_retention_counts.csv",
    "customer_lifetime_value": "customer_lifetime_value.csv",
    "kpi_summary": "kpi_summary.json",
}


def ensure_directories() -> None:
    """Create the data directories if they do not already exist."""
    for directory in (RAW_DIR, PROCESSED_DIR):
        directory.mkdir(parents=True, exist_ok=True)


# --------------------------------------------------------------------------- #
# Reproducibility                                                              #
# --------------------------------------------------------------------------- #

RANDOM_SEED: Final[int] = 42

# The dataset is anchored to a fixed snapshot date rather than "today" so that
# every run of the pipeline - this year or next - produces byte-identical output.
SNAPSHOT_DATE: Final[dt.date] = dt.date(2026, 9, 30)
PERIOD_START: Final[dt.date] = dt.date(2024, 10, 1)
PERIOD_END: Final[dt.date] = SNAPSHOT_DATE

# --------------------------------------------------------------------------- #
# Dataset size                                                                 #
# --------------------------------------------------------------------------- #

TARGET_TRANSACTION_LINES: Final[int] = 15_000
N_CUSTOMERS: Final[int] = 3_400
N_PRODUCTS: Final[int] = 240
N_BOUTIQUES: Final[int] = 10
N_CAMPAIGNS: Final[int] = 26

# Multiplier applied to the per-tier annual purchase rate so that the generated
# line count lands on TARGET_TRANSACTION_LINES. Calibrated empirically; changing
# the seed or the tier mix requires re-calibrating.
PURCHASE_RATE_CALIBRATION: Final[float] = 1.32

# --------------------------------------------------------------------------- #
# Business assumptions (documented defaults where the spec is silent)          #
# --------------------------------------------------------------------------- #

# A purchase is attributed to a clienteling interaction when it occurs within
# this many days of the interaction. Used identically in Python, DAX, Tableau
# and Excel so that every layer of the project reports the same number.
CLIENTELING_ATTRIBUTION_WINDOW_DAYS: Final[int] = 14

# RFM segmentation thresholds. Scores are 1-5 quintiles (5 = best).
RFM_HIGH_VALUE_MONETARY_SCORE: Final[int] = 4  # M >= 4 counts as "high value"
RFM_LOYAL_FREQUENCY_SCORE: Final[int] = 4  # F >= 4 counts as "loyal"
LAPSED_RECENCY_DAYS: Final[int] = 180  # no purchase in 180 days = lapsing
NEW_CLIENT_TENURE_DAYS: Final[int] = 365  # first purchase within 365 days = new

# 12-month CLV assumptions.
CLV_HORIZON_MONTHS: Final[int] = 12
CLV_ANNUAL_DISCOUNT_RATE: Final[float] = 0.10  # cost of capital, 10% per annum
# Fallback retention rate used only if the cohort model cannot be estimated.
CLV_FALLBACK_RETENTION_RATE: Final[float] = 0.35

# Reporting currency. The dataset is European-weighted, so EUR.
CURRENCY: Final[str] = "EUR"
CURRENCY_SYMBOL: Final[str] = "€"
