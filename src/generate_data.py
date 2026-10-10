"""Phase 1 - Reproducible synthetic data generation for a luxury retail CRM.

Generates seven related tables that mimic the behavioural signature of a luxury
maison rather than mass-market retail:

* low purchase frequency and high basket values;
* a small cohort of Very Important Clients (VICs) driving a large revenue share;
* restrained markdown activity concentrated in two private-sale windows;
* appointment-driven selling (clienteling) that converts far above walk-in;
* a realistic tail of lapsed high-value clients - the commercial problem the
  rest of this project exists to solve.

Run directly to write CSVs and a multi-sheet Excel workbook to data/processed/:

    python -m src.generate_data
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from typing import Final

import numpy as np
import pandas as pd

from src import config

# --------------------------------------------------------------------------- #
# Reference data                                                               #
# --------------------------------------------------------------------------- #

BOUTIQUE_SPECS: Final[list[dict[str, object]]] = [
    # name, city, country, region, format, relative traffic weight
    {"city": "Paris", "country": "France", "region": "Europe", "format": "Flagship", "weight": 1.00},
    {"city": "London", "country": "United Kingdom", "region": "Europe", "format": "Flagship", "weight": 0.82},
    {"city": "Milan", "country": "Italy", "region": "Europe", "format": "Boutique", "weight": 0.61},
    {"city": "New York", "country": "United States", "region": "Americas", "format": "Flagship", "weight": 0.95},
    {"city": "Los Angeles", "country": "United States", "region": "Americas", "format": "Boutique", "weight": 0.54},
    {"city": "Dubai", "country": "United Arab Emirates", "region": "Middle East", "format": "Boutique", "weight": 0.66},
    {"city": "Hong Kong", "country": "Hong Kong SAR", "region": "Asia Pacific", "format": "Boutique", "weight": 0.59},
    {"city": "Tokyo", "country": "Japan", "region": "Asia Pacific", "format": "Boutique", "weight": 0.71},
    {"city": "Singapore", "country": "Singapore", "region": "Asia Pacific", "format": "Concession", "weight": 0.44},
    {"city": "Bicester", "country": "United Kingdom", "region": "Europe", "format": "Outlet", "weight": 0.38},
]

# Category -> (list-price lognormal mu/sigma bounds, gross margin band, share of assortment)
CATEGORY_SPECS: Final[dict[str, dict[str, object]]] = {
    "Leather Goods": {
        "price_range": (1_150, 9_600),
        "price_skew": 0.55,
        "margin_band": (0.68, 0.78),
        "assortment_share": 0.30,
        "demand_share": 0.38,
        "return_rate": 0.035,
        "subcategories": ["Shoulder Bag", "Tote", "Top Handle", "Small Leather Goods", "Travel"],
    },
    "Apparel": {
        "price_range": (620, 4_800),
        "price_skew": 0.62,
        "margin_band": (0.60, 0.70),
        "assortment_share": 0.32,
        "demand_share": 0.26,
        "return_rate": 0.115,
        "subcategories": ["Outerwear", "Knitwear", "Tailoring", "Dresses", "Shirting"],
    },
    "Footwear": {
        "price_range": (540, 1_950),
        "price_skew": 0.48,
        "margin_band": (0.58, 0.68),
        "assortment_share": 0.18,
        "demand_share": 0.17,
        "return_rate": 0.125,
        "subcategories": ["Sneakers", "Loafers", "Boots", "Heels", "Sandals"],
    },
    "Accessories": {
        "price_range": (180, 980),
        "price_skew": 0.52,
        "margin_band": (0.64, 0.74),
        "assortment_share": 0.20,
        "demand_share": 0.19,
        "return_rate": 0.055,
        "subcategories": ["Silk", "Belts", "Eyewear", "Jewellery", "Fragrance"],
    },
}

COLLECTIONS: Final[list[str]] = [
    "Maison Icons",
    "Atelier Heritage",
    "Seasonal Capsule",
    "Runway Edit",
    "Essentiels",
]

SEASONS: Final[list[str]] = ["SS24", "FW24", "SS25", "FW25", "SS26"]

ACQUISITION_SOURCES: Final[dict[str, float]] = {
    "Boutique Walk-In": 0.34,
    "Client Referral": 0.14,
    "Paid Social": 0.13,
    "Organic Search": 0.11,
    "Private Event": 0.10,
    "Press & Editorial": 0.08,
    "Partnership": 0.06,
    "Concierge / Travel": 0.04,
}

CONSENT_STATUSES: Final[dict[str, float]] = {
    "Opted In": 0.71,
    "Opted Out": 0.17,
    "Pending": 0.12,
}

ASSOCIATE_ROLES: Final[dict[str, float]] = {
    "Client Advisor": 0.62,
    "Senior Client Advisor": 0.27,
    "Boutique Director": 0.11,
}

# Client value tiers. `annual_orders` is the base Poisson rate per year; the
# basket multiplier scales the price point the client shops at.
@dataclass(frozen=True)
class ClientTier:
    """Latent behavioural tier that drives a client's purchase simulation."""

    name: str
    share: float
    annual_orders: float
    basket_multiplier: float
    discount_affinity: float  # probability of buying on markdown when offered
    clienteling_propensity: float  # probability of being in the clienteling book


CLIENT_TIERS: Final[list[ClientTier]] = [
    # The long Occasional tail is deliberate: most luxury clients buy once -
    # a gift, a holiday purchase, a wedding - and are never seen again. Without
    # that tail the cohort retention curve would be implausibly flattering.
    ClientTier("VIC", 0.04, 5.8, 1.85, 0.07, 0.92),
    ClientTier("High Potential", 0.14, 3.1, 1.32, 0.18, 0.61),
    ClientTier("Core", 0.30, 1.5, 1.00, 0.38, 0.21),
    ClientTier("Occasional", 0.52, 0.40, 0.78, 0.57, 0.04),
]

# Monthly demand index: December gifting peak, August trough, sale-month volume
# bumps in January and July.
MONTH_SEASONALITY: Final[dict[int, float]] = {
    1: 0.88, 2: 0.86, 3: 0.95, 4: 0.94, 5: 1.00, 6: 0.97,
    7: 0.99, 8: 0.74, 9: 0.96, 10: 1.03, 11: 1.21, 12: 1.58,
}

# Monday .. Sunday footfall index.
WEEKDAY_SEASONALITY: Final[list[float]] = [0.82, 0.88, 0.95, 1.02, 1.18, 1.35, 0.74]

# Private sale windows (month, day ranges) - the only periods with markdown in
# full-price boutiques.
SALE_WINDOWS: Final[list[tuple[int, int, int]]] = [
    (1, 5, 31),   # January private sale
    (2, 1, 9),    # January sale run-out
    (7, 1, 31),   # July summer sale
    (8, 1, 10),   # July sale run-out
]

DISCOUNT_LADDER: Final[list[float]] = [0.10, 0.20, 0.30, 0.40]
DISCOUNT_LADDER_WEIGHTS: Final[list[float]] = [0.26, 0.38, 0.24, 0.12]

INTERACTION_TYPES: Final[dict[str, dict[str, float]]] = {
    # conversion = probability the interaction leads to a purchase within the
    # attribution window
    "Appointment": {"share": 0.46, "conversion": 0.44},
    "Call": {"share": 0.37, "conversion": 0.13},
    "Event": {"share": 0.17, "conversion": 0.27},
}

CAMPAIGN_CHANNELS: Final[dict[str, dict[str, float]]] = {
    # base_cost      : fixed production / platform cost of running the campaign
    # cost_per_head  : variable cost per client reached
    # reach_rate     : share of the eligible audience actually touched
    # lift           : share of in-window purchases by reached clients claimed
    # tail_days      : attribution tail after the campaign end date
    "Email (CRM)": {
        "share": 0.27, "base_cost": 6_000, "cost_per_head": 4.0,
        "reach_rate": 0.86, "lift": 0.44, "tail_days": 7,
    },
    "Private Event": {
        "share": 0.15, "base_cost": 62_000, "cost_per_head": 310.0,
        "reach_rate": 0.62, "lift": 0.74, "tail_days": 30,
    },
    "Paid Social": {
        "share": 0.19, "base_cost": 14_000, "cost_per_head": 9.5,
        "reach_rate": 0.64, "lift": 0.19, "tail_days": 7,
    },
    "Paid Search": {
        "share": 0.12, "base_cost": 11_000, "cost_per_head": 7.5,
        "reach_rate": 0.48, "lift": 0.23, "tail_days": 7,
    },
    "Direct Mail": {
        "share": 0.15, "base_cost": 9_000, "cost_per_head": 12.0,
        "reach_rate": 0.58, "lift": 0.31, "tail_days": 14,
    },
    "SMS / WhatsApp": {
        "share": 0.12, "base_cost": 4_000, "cost_per_head": 3.0,
        "reach_rate": 0.74, "lift": 0.26, "tail_days": 7,
    },
}

CAMPAIGN_AUDIENCES: Final[dict[str, float]] = {
    "VIC & Top Clients": 0.20,
    "High Potential": 0.19,
    "Lapsed High-Value": 0.15,
    "New & Prospective Clients": 0.23,
    "Full Client Base": 0.23,
}

# Private events are invitation-only and are never run against the full base.
PRIVATE_EVENT_AUDIENCES: Final[dict[str, float]] = {
    "VIC & Top Clients": 0.45,
    "High Potential": 0.33,
    "Lapsed High-Value": 0.22,
}

FIRST_NAMES: Final[list[str]] = [
    "Amara", "Beatrice", "Camille", "Daniela", "Elena", "Farrah", "Giulia", "Hana",
    "Ines", "Juliette", "Keiko", "Liana", "Margot", "Noor", "Olivia", "Priya",
    "Qing", "Rania", "Sofia", "Tatiana", "Ursula", "Valentina", "Wen", "Ximena",
    "Yasmin", "Zara", "Adrien", "Bastien", "Carlo", "Dmitri", "Edoardo", "Felipe",
    "Gabriel", "Hiroshi", "Idris", "Jasper", "Karim", "Lorenzo", "Mathis", "Nikolai",
    "Omar", "Pierre", "Quentin", "Rafael", "Sebastien", "Tomas", "Ulrich", "Victor",
]

LAST_NAMES: Final[list[str]] = [
    "Abbas", "Beaumont", "Castellane", "Duval", "Eriksson", "Fontaine", "Girard",
    "Haddad", "Ishikawa", "Jovanovic", "Kowalski", "Lambert", "Moreau", "Nakamura",
    "Okafor", "Pereira", "Quintero", "Rossi", "Santoro", "Takahashi", "Umeh",
    "Vasquez", "Weber", "Xu", "Yilmaz", "Zhang", "Almeida", "Bianchi", "Chevalier",
    "Delacroix", "Esposito", "Ferretti", "Grimaldi", "Hoffmann", "Ivanov", "Jenkins",
]

PRODUCT_ADJECTIVES: Final[list[str]] = [
    "Monogram", "Quilted", "Structured", "Soft", "Signature", "Archive", "Grained",
    "Smooth", "Woven", "Embossed", "Panelled", "Draped", "Tailored", "Oversized",
]

PRODUCT_NOUNS: Final[dict[str, list[str]]] = {
    "Leather Goods": ["Hobo", "Messenger", "Clutch", "Satchel", "Bucket Bag", "Card Holder"],
    "Apparel": ["Trench", "Blazer", "Cardigan", "Midi Dress", "Trouser", "Overshirt"],
    "Footwear": ["Derby", "Mule", "Ankle Boot", "Runner", "Pump", "Slide"],
    "Accessories": ["Scarf", "Belt", "Sunglasses", "Cuff", "Pendant", "Eau de Parfum"],
}


# --------------------------------------------------------------------------- #
# Helpers                                                                      #
# --------------------------------------------------------------------------- #


def _weighted_choice(
    rng: np.random.Generator, mapping: dict[str, float], size: int
) -> np.ndarray:
    """Draw ``size`` keys from ``mapping`` using its values as probabilities."""
    keys = list(mapping.keys())
    probs = np.asarray(list(mapping.values()), dtype=float)
    probs = probs / probs.sum()
    return rng.choice(keys, size=size, p=probs)


def _date_range_days(start: dt.date, end: dt.date) -> pd.DatetimeIndex:
    """Return every calendar day in the inclusive range ``start``..``end``."""
    return pd.date_range(start=start, end=end, freq="D")


def _is_in_sale_window(day: pd.Timestamp) -> bool:
    """True when the date falls inside one of the two private-sale windows."""
    for month, first, last in SALE_WINDOWS:
        if day.month == month and first <= day.day <= last:
            return True
    return False


def _build_day_weights(days: pd.DatetimeIndex) -> np.ndarray:
    """Combine month and weekday seasonality into a per-day sampling weight."""
    month_idx = np.array([MONTH_SEASONALITY[d.month] for d in days])
    weekday_idx = np.array([WEEKDAY_SEASONALITY[d.weekday()] for d in days])
    return month_idx * weekday_idx


# --------------------------------------------------------------------------- #
# Table builders                                                               #
# --------------------------------------------------------------------------- #


def build_boutiques(rng: np.random.Generator) -> pd.DataFrame:
    """Build the boutique dimension (10 locations with formats and targets)."""
    rows: list[dict[str, object]] = []
    for i, spec in enumerate(BOUTIQUE_SPECS, start=1):
        weight = float(spec["weight"])
        # Placeholder: the real annual target is calibrated against realised
        # sales once transactions exist (see `_finalise_boutique_targets`),
        # so that target achievement lands in a believable 85-115% band.
        target = float("nan")
        opening = config.PERIOD_START - dt.timedelta(days=int(rng.integers(500, 6_200)))
        rows.append(
            {
                "boutique_id": f"BTQ-{i:02d}",
                "boutique_name": f"{spec['city']} {spec['format']}",
                "city": spec["city"],
                "country": spec["country"],
                "region": spec["region"],
                "boutique_format": spec["format"],
                "opening_date": opening,
                "traffic_weight": round(weight, 3),
                "annual_sales_target": float(np.round(target, -3)),
            }
        )
    return pd.DataFrame(rows)


def build_associates(rng: np.random.Generator, boutiques: pd.DataFrame) -> pd.DataFrame:
    """Build the sales-associate dimension, sized to each boutique's traffic."""
    rows: list[dict[str, object]] = []
    counter = 1
    for _, boutique in boutiques.iterrows():
        headcount = int(np.clip(round(float(boutique["traffic_weight"]) * 8), 4, 9))
        for _ in range(headcount):
            role = _weighted_choice(rng, ASSOCIATE_ROLES, 1)[0]
            rows.append(
                {
                    "associate_id": f"ASC-{counter:03d}",
                    "associate_name": (
                        f"{rng.choice(FIRST_NAMES)} {rng.choice(LAST_NAMES)}"
                    ),
                    "boutique_id": boutique["boutique_id"],
                    "role": role,
                    "tenure_years": round(float(rng.gamma(2.2, 1.9)) + 0.3, 1),
                    # Relative skill at converting clienteling activity.
                    "performance_index": round(float(rng.normal(1.0, 0.17)), 3),
                }
            )
            counter += 1
    return pd.DataFrame(rows)


def build_products(rng: np.random.Generator) -> pd.DataFrame:
    """Build the product master across the four luxury categories."""
    rows: list[dict[str, object]] = []
    counter = 1
    for category, spec in CATEGORY_SPECS.items():
        n_sku = int(round(config.N_PRODUCTS * float(spec["assortment_share"])))
        low, high = spec["price_range"]  # type: ignore[misc]
        margin_low, margin_high = spec["margin_band"]  # type: ignore[misc]
        # Lognormal prices truncated to the category band give a realistic
        # long tail of exceptional pieces without absurd outliers.
        raw = rng.lognormal(mean=0.0, sigma=float(spec["price_skew"]), size=n_sku)
        raw = (raw - raw.min()) / (raw.max() - raw.min())
        prices = low + raw * (high - low)
        for price in prices:
            category_key = str(category)
            margin = float(rng.uniform(margin_low, margin_high))
            list_price = float(np.round(price / 10) * 10)  # round to nearest 10
            rows.append(
                {
                    "sku": f"SKU-{counter:04d}",
                    "product_name": (
                        f"{rng.choice(PRODUCT_ADJECTIVES)} "
                        f"{rng.choice(PRODUCT_NOUNS[category_key])}"
                    ),
                    "category": category_key,
                    "subcategory": str(rng.choice(spec["subcategories"])),  # type: ignore[arg-type]
                    "collection": str(rng.choice(COLLECTIONS)),
                    "season": str(rng.choice(SEASONS)),
                    "list_price": list_price,
                    "unit_cost": float(np.round(list_price * (1 - margin), 2)),
                    "gross_margin_pct": round(margin, 4),
                }
            )
            counter += 1
    return pd.DataFrame(rows)


def build_customers(
    rng: np.random.Generator, boutiques: pd.DataFrame
) -> pd.DataFrame:
    """Build the customer master with latent behavioural tiers.

    The latent tier (`client_tier`) and `churn_date` are retained in the output
    because they are the *ground truth* the analytics modules are measured
    against - RFM segmentation never reads them, but the test suite can.
    """
    n = config.N_CUSTOMERS
    tier_names = [t.name for t in CLIENT_TIERS]
    tier_probs = np.asarray([t.share for t in CLIENT_TIERS])
    tiers = rng.choice(tier_names, size=n, p=tier_probs / tier_probs.sum())

    # 62% of the book pre-dates the analysis window; the rest is acquired
    # during it, which is what makes the cohort analysis meaningful.
    existing_mask = rng.random(n) < 0.62
    days_before = rng.integers(1, 1_460, size=n)  # up to 4 years of history
    window_days = (config.PERIOD_END - config.PERIOD_START).days
    days_within = rng.integers(0, window_days - 20, size=n)
    join_dates = np.where(
        existing_mask,
        np.array([config.PERIOD_START - dt.timedelta(days=int(d)) for d in days_before]),
        np.array([config.PERIOD_START + dt.timedelta(days=int(d)) for d in days_within]),
    )

    boutique_ids = rng.choice(
        boutiques["boutique_id"].to_numpy(),
        size=n,
        p=(boutiques["traffic_weight"] / boutiques["traffic_weight"].sum()).to_numpy(),
    )
    boutique_lookup = boutiques.set_index("boutique_id")

    customers = pd.DataFrame(
        {
            "customer_id": [f"CUS-{i:05d}" for i in range(1, n + 1)],
            "client_name": [
                f"{rng.choice(FIRST_NAMES)} {rng.choice(LAST_NAMES)}" for _ in range(n)
            ],
            "join_date": pd.to_datetime(join_dates),
            "acquisition_source": _weighted_choice(rng, ACQUISITION_SOURCES, n),
            "preferred_boutique_id": boutique_ids,
            "consent_status": _weighted_choice(rng, CONSENT_STATUSES, n),
            "client_tier": tiers,
        }
    )
    customers["country"] = boutique_lookup.loc[boutique_ids, "country"].to_numpy()
    customers["region"] = boutique_lookup.loc[boutique_ids, "region"].to_numpy()
    customers["email_opt_in"] = customers["consent_status"].eq("Opted In")

    # ---- Lapsed high-value clients -------------------------------------- #
    # ~26% of VIC / High Potential clients stop purchasing part-way through the
    # window. This is the commercial problem the project is built to surface.
    high_value = customers["client_tier"].isin(["VIC", "High Potential"]).to_numpy()
    churn_draw = rng.random(config.N_CUSTOMERS) < 0.26
    churns = high_value & churn_draw
    # Churn lands between 4 and 17 months before the snapshot, so these clients
    # always exceed the 180-day lapsing threshold.
    churn_offsets = rng.integers(190, 520, size=config.N_CUSTOMERS)
    churn_dates = np.array(
        [
            config.SNAPSHOT_DATE - dt.timedelta(days=int(off)) if churn else pd.NaT
            for off, churn in zip(churn_offsets, churns)
        ],
        dtype="object",
    )
    customers["churn_date"] = pd.to_datetime(pd.Series(churn_dates))
    # A client cannot churn before they joined.
    invalid = customers["churn_date"].notna() & (
        customers["churn_date"] <= customers["join_date"]
    )
    customers.loc[invalid, "churn_date"] = pd.NaT

    # Category affinity: each client skews towards one category.
    customers["primary_category_affinity"] = rng.choice(
        list(CATEGORY_SPECS.keys()),
        size=n,
        p=[float(s["demand_share"]) for s in CATEGORY_SPECS.values()],
    )
    return customers


def build_clienteling_interactions(
    rng: np.random.Generator,
    customers: pd.DataFrame,
    associates: pd.DataFrame,
) -> pd.DataFrame:
    """Build the clienteling activity log (appointments, calls, events).

    Outcomes are provisional at this stage: `converted` is decided here, and
    `generate_transactions` creates the matching purchase. The final outcome
    label is written back once the linked orders exist.
    """
    tier_lookup = {t.name: t for t in CLIENT_TIERS}
    associates_by_boutique = {
        boutique: group for boutique, group in associates.groupby("boutique_id")
    }

    rows: list[dict[str, object]] = []
    counter = 1
    for customer in customers.itertuples(index=False):
        tier = tier_lookup[customer.client_tier]
        if rng.random() > tier.clienteling_propensity:
            continue  # not in the clienteling book

        # Window in which the client is contactable.
        start = max(pd.Timestamp(customer.join_date).date(), config.PERIOD_START)
        end = config.PERIOD_END
        if pd.notna(customer.churn_date):
            # Associates keep trying for ~90 days after a client goes quiet.
            end = min(end, pd.Timestamp(customer.churn_date).date() + dt.timedelta(days=90))
        if end <= start:
            continue

        active_years = (end - start).days / 365.25
        n_interactions = int(rng.poisson(max(tier.annual_orders * 1.35, 0.4) * active_years))
        if n_interactions == 0:
            continue

        pool = associates_by_boutique[customer.preferred_boutique_id]
        offsets = rng.integers(0, max((end - start).days, 1), size=n_interactions)
        for offset in offsets:
            interaction_date = start + dt.timedelta(days=int(offset))
            interaction_type = _weighted_choice(
                rng, {k: v["share"] for k, v in INTERACTION_TYPES.items()}, 1
            )[0]
            associate = pool.iloc[int(rng.integers(0, len(pool)))]
            base_conversion = float(INTERACTION_TYPES[interaction_type]["conversion"])
            # Conversion is modulated by associate skill and client tier.
            tier_factor = {"VIC": 1.22, "High Potential": 1.05, "Core": 0.88, "Occasional": 0.62}[
                customer.client_tier
            ]
            p_convert = float(
                np.clip(
                    base_conversion * float(associate["performance_index"]) * tier_factor,
                    0.02,
                    0.92,
                )
            )
            converted = bool(rng.random() < p_convert)
            rows.append(
                {
                    "interaction_id": f"INT-{counter:05d}",
                    "customer_id": customer.customer_id,
                    "associate_id": associate["associate_id"],
                    "boutique_id": customer.preferred_boutique_id,
                    "interaction_date": pd.Timestamp(interaction_date),
                    "interaction_type": interaction_type,
                    "converted": converted,
                }
            )
            counter += 1

    interactions = pd.DataFrame(rows).sort_values("interaction_date").reset_index(drop=True)
    return interactions


def _assign_outcomes(
    rng: np.random.Generator, interactions: pd.DataFrame
) -> pd.DataFrame:
    """Translate the boolean `converted` flag into a business outcome label."""
    outcomes = np.where(
        interactions["converted"],
        "Purchase",
        rng.choice(
            ["Follow-Up Scheduled", "No Purchase", "No Response"],
            size=len(interactions),
            p=[0.34, 0.45, 0.21],
        ),
    )
    interactions = interactions.copy()
    interactions["outcome"] = outcomes
    return interactions


def _sample_skus(
    rng: np.random.Generator,
    products: pd.DataFrame,
    affinity_category: str,
    n_lines: int,
    basket_multiplier: float,
) -> np.ndarray:
    """Pick SKUs for one order, skewed by client affinity and spending power.

    70% of lines come from the client's affinity category. Within a category,
    higher-tier clients are more likely to select higher price points.
    """
    categories = np.where(
        rng.random(n_lines) < 0.70,
        affinity_category,
        rng.choice(
            list(CATEGORY_SPECS.keys()),
            size=n_lines,
            p=[float(s["demand_share"]) for s in CATEGORY_SPECS.values()],
        ),
    )
    chosen: list[str] = []
    for category in categories:
        pool = products[products["category"] == category]
        # Price-point preference: weight ~ price ** (basket_multiplier - 1)
        weights = np.power(pool["list_price"].to_numpy(), basket_multiplier - 1.0)
        weights = weights / weights.sum()
        chosen.append(str(rng.choice(pool["sku"].to_numpy(), p=weights)))
    return np.asarray(chosen)


def generate_transactions(
    rng: np.random.Generator,
    customers: pd.DataFrame,
    products: pd.DataFrame,
    boutiques: pd.DataFrame,
    associates: pd.DataFrame,
    interactions: pd.DataFrame,
) -> pd.DataFrame:
    """Simulate order lines for every client.

    Two streams are combined:

    1. **Organic demand** - Poisson-distributed orders per client, placed on
       seasonality-weighted days.
    2. **Clienteling-driven demand** - one order for every converted
       interaction, placed 0-13 days after the interaction so that the 14-day
       attribution rule used across Python, DAX, Tableau and Excel recovers it.
    """
    tier_lookup = {t.name: t for t in CLIENT_TIERS}
    product_lookup = products.set_index("sku")
    boutique_weights = (
        boutiques.set_index("boutique_id")["traffic_weight"]
        / boutiques["traffic_weight"].sum()
    )
    outlet_ids = set(
        boutiques.loc[boutiques["boutique_format"] == "Outlet", "boutique_id"]
    )
    associates_by_boutique = {
        boutique: group["associate_id"].to_numpy()
        for boutique, group in associates.groupby("boutique_id")
    }

    all_days = _date_range_days(config.PERIOD_START, config.PERIOD_END)
    day_weights = _build_day_weights(all_days)

    converted = interactions[interactions["converted"]]
    converted_by_customer = {
        customer_id: group
        for customer_id, group in converted.groupby("customer_id")
    }

    order_records: list[dict[str, object]] = []

    for customer in customers.itertuples(index=False):
        tier = tier_lookup[customer.client_tier]

        start = max(pd.Timestamp(customer.join_date).date(), config.PERIOD_START)
        end = config.PERIOD_END
        if pd.notna(customer.churn_date):
            end = min(end, pd.Timestamp(customer.churn_date).date())
        if end <= start:
            continue

        # ---- Stream 1: organic orders ----------------------------------- #
        active_years = (end - start).days / 365.25
        expected = tier.annual_orders * active_years * config.PURCHASE_RATE_CALIBRATION
        # Clienteling orders are generated separately, so net them off here to
        # avoid double-counting the client's total propensity to buy.
        n_clienteling = len(converted_by_customer.get(customer.customer_id, []))
        expected = max(expected - n_clienteling * 0.55, 0.0)
        n_orders = int(rng.poisson(expected))

        if n_orders:
            mask = (all_days >= pd.Timestamp(start)) & (all_days <= pd.Timestamp(end))
            eligible_days = all_days[mask]
            eligible_weights = day_weights[mask]
            eligible_weights = eligible_weights / eligible_weights.sum()
            order_days = rng.choice(
                eligible_days, size=n_orders, p=eligible_weights, replace=True
            )
            for day in order_days:
                order_records.append(
                    {
                        "customer": customer,
                        "date": pd.Timestamp(day),
                        "interaction_id": None,
                        "associate_id": None,
                        "boutique_id": None,
                        "channel": "Boutique" if rng.random() < 0.83 else "Online",
                    }
                )

        # ---- Stream 2: clienteling-driven orders ------------------------ #
        for interaction in converted_by_customer.get(
            customer.customer_id, pd.DataFrame()
        ).itertuples(index=False):
            lag = int(rng.integers(0, config.CLIENTELING_ATTRIBUTION_WINDOW_DAYS))
            purchase_date = pd.Timestamp(interaction.interaction_date) + pd.Timedelta(days=lag)
            if purchase_date > pd.Timestamp(config.PERIOD_END):
                continue
            order_records.append(
                {
                    "customer": customer,
                    "date": purchase_date,
                    "interaction_id": interaction.interaction_id,
                    "associate_id": interaction.associate_id,
                    "boutique_id": interaction.boutique_id,
                    "channel": "Clienteling",
                }
            )

    # ---- Explode orders into lines ------------------------------------- #
    lines: list[dict[str, object]] = []
    order_counter = 1
    line_counter = 1

    # Deterministic ordering before line generation keeps output stable.
    order_records.sort(key=lambda r: (r["date"], r["customer"].customer_id))

    for record in order_records:
        customer = record["customer"]
        tier = tier_lookup[customer.client_tier]
        order_date: pd.Timestamp = record["date"]

        # Boutique: clienteling orders happen where the appointment was held;
        # otherwise the client's preferred boutique 78% of the time.
        if record["boutique_id"] is not None:
            boutique_id = str(record["boutique_id"])
        elif rng.random() < 0.78:
            boutique_id = str(customer.preferred_boutique_id)
        else:
            boutique_id = str(
                rng.choice(boutique_weights.index.to_numpy(), p=boutique_weights.to_numpy())
            )

        associate_id = record["associate_id"]
        if associate_id is None:
            pool = associates_by_boutique[boutique_id]
            associate_id = str(rng.choice(pool))

        order_id = f"ORD-{order_counter:06d}"
        order_counter += 1

        # Clienteling orders are larger: appointments are booked around a need.
        n_lines = int(
            rng.choice([1, 2, 3, 4], p=[0.64, 0.23, 0.09, 0.04])
            if record["channel"] != "Clienteling"
            else rng.choice([1, 2, 3, 4], p=[0.44, 0.31, 0.17, 0.08])
        )
        basket_mult = tier.basket_multiplier * (1.08 if record["channel"] == "Clienteling" else 1.0)
        skus = _sample_skus(
            rng, products, customer.primary_category_affinity, n_lines, basket_mult
        )

        # ---- Discount logic -------------------------------------------- #
        in_sale = _is_in_sale_window(order_date)
        is_outlet = boutique_id in outlet_ids
        if is_outlet:
            discount_pct = float(rng.choice([0.30, 0.40, 0.50], p=[0.42, 0.40, 0.18]))
        elif in_sale and rng.random() < tier.discount_affinity * 3.0:
            discount_pct = float(
                rng.choice(DISCOUNT_LADDER, p=DISCOUNT_LADDER_WEIGHTS)
            )
        elif rng.random() < 0.012:
            # Rare off-cycle gesture (employee / press / loyalty gesture).
            discount_pct = float(rng.choice([0.10, 0.20], p=[0.7, 0.3]))
        else:
            discount_pct = 0.0

        for sku in skus:
            product = product_lookup.loc[sku]
            category = str(product["category"])
            list_price = float(product["list_price"])
            quantity = int(
                rng.choice([1, 2, 3], p=[0.93, 0.055, 0.015])
                if category == "Accessories"
                else rng.choice([1, 2], p=[0.97, 0.03])
            )
            unit_net_price = round(list_price * (1 - discount_pct), 2)
            return_rate = float(CATEGORY_SPECS[category]["return_rate"])
            # Online and markdown purchases are returned more often.
            if record["channel"] == "Online":
                return_rate *= 1.9
            if discount_pct > 0:
                return_rate *= 1.25
            is_return = bool(rng.random() < return_rate)

            lines.append(
                {
                    "transaction_id": f"TXN-{line_counter:06d}",
                    "order_id": order_id,
                    "transaction_date": order_date,
                    "customer_id": customer.customer_id,
                    "sku": sku,
                    "boutique_id": boutique_id,
                    "associate_id": associate_id,
                    "channel": record["channel"],
                    "quantity": quantity,
                    "unit_list_price": list_price,
                    "discount_pct": round(discount_pct, 4),
                    "unit_net_price": unit_net_price,
                    "gross_amount": round(list_price * quantity, 2),
                    "net_amount": round(unit_net_price * quantity, 2),
                    "unit_cost": float(product["unit_cost"]),
                    "is_return": is_return,
                    "is_full_price": discount_pct == 0.0,
                    "interaction_id": record["interaction_id"],
                    "campaign_id": None,  # assigned in attach_campaigns
                }
            )
            line_counter += 1

    transactions = pd.DataFrame(lines)
    transactions = transactions.sort_values(
        ["transaction_date", "order_id", "transaction_id"]
    ).reset_index(drop=True)
    return transactions


def build_campaigns(
    rng: np.random.Generator, transactions: pd.DataFrame, customers: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Build the marketing campaign table and attribute transactions to it.

    Attribution rule: a transaction is credited to a campaign when the client
    is in the campaign audience and the purchase falls inside the campaign
    window plus a 7-day post-window tail. A per-channel `lift` factor controls
    what share of eligible purchases is claimed, which is what produces the
    wide ROAS spread between CRM and prospecting channels.
    """
    window_days = (config.PERIOD_END - config.PERIOD_START).days
    starts = np.sort(rng.integers(0, window_days - 30, size=config.N_CAMPAIGNS))

    # Audience membership is derived from latent tiers (what a CRM team would
    # approximate from their own segmentation).
    tier_of = customers.set_index("customer_id")["client_tier"].to_dict()
    churn_of = customers.set_index("customer_id")["churn_date"].to_dict()
    join_of = customers.set_index("customer_id")["join_date"].to_dict()

    def in_audience(customer_id: str, audience: str, start: pd.Timestamp) -> bool:
        tier = tier_of[customer_id]
        if audience == "VIC & Top Clients":
            return tier == "VIC"
        if audience == "High Potential":
            return tier in ("VIC", "High Potential")
        if audience == "Lapsed High-Value":
            churn = churn_of[customer_id]
            return pd.notna(churn) and pd.Timestamp(churn) < start
        if audience == "New & Prospective Clients":
            return pd.Timestamp(join_of[customer_id]) >= start - pd.Timedelta(days=270)
        return True

    transactions = transactions.copy()
    campaign_assignment = pd.Series(pd.NA, index=transactions.index, dtype="object")

    rows: list[dict[str, object]] = []
    for i, start_offset in enumerate(starts, start=1):
        start = pd.Timestamp(config.PERIOD_START) + pd.Timedelta(days=int(start_offset))
        duration = int(rng.integers(10, 35))
        end = start + pd.Timedelta(days=duration)
        channel = _weighted_choice(
            rng, {k: v["share"] for k, v in CAMPAIGN_CHANNELS.items()}, 1
        )[0]
        if channel == "Private Event":
            # Events are only ever run against the top of the client book.
            audience = _weighted_choice(rng, PRIVATE_EVENT_AUDIENCES, 1)[0]
        else:
            audience = _weighted_choice(rng, CAMPAIGN_AUDIENCES, 1)[0]
        channel_spec = CAMPAIGN_CHANNELS[channel]

        eligible_customers = [
            cid for cid in tier_of if in_audience(cid, audience, start)
        ]
        reach = int(len(eligible_customers) * float(channel_spec["reach_rate"]))
        # A private event has a physical guest list - it cannot scale with the
        # size of the database the way an email send can.
        if channel == "Private Event":
            reach = min(reach, int(rng.integers(90, 175)))
        targeted = (
            rng.choice(eligible_customers, size=reach, replace=False)
            if reach > 0
            else np.array([], dtype=str)
        )
        targeted_set = set(targeted.tolist())

        # Budget = fixed production cost + variable cost per client reached,
        # with +/- 12% execution variance.
        spend = float(
            np.round(
                (
                    float(channel_spec["base_cost"])
                    + float(channel_spec["cost_per_head"]) * reach
                )
                * rng.uniform(0.88, 1.12),
                2,
            )
        )
        impressions = int(reach * rng.uniform(2.1, 6.4)) if reach else 0
        clicks = int(impressions * rng.uniform(0.012, 0.085))

        campaign_id = f"CMP-{i:03d}"
        # Claim a share of in-window purchases by targeted clients. The tail
        # reflects how long the channel keeps influencing purchase behaviour.
        tail = pd.Timedelta(days=int(channel_spec["tail_days"]))
        eligible_tx = (
            transactions["customer_id"].isin(targeted_set)
            & (transactions["transaction_date"] >= start)
            & (transactions["transaction_date"] <= end + tail)
            & campaign_assignment.isna()
        )
        eligible_idx = transactions.index[eligible_tx]
        claim_mask = rng.random(len(eligible_idx)) < float(channel_spec["lift"])
        claimed_idx = eligible_idx[claim_mask]
        campaign_assignment.loc[claimed_idx] = campaign_id

        rows.append(
            {
                "campaign_id": campaign_id,
                "campaign_name": f"{channel.split(' ')[0]} - {audience} - {start:%b %Y}",
                "audience_segment": audience,
                "channel": channel,
                "start_date": start,
                "end_date": end,
                "spend": spend,
                "targeted_customers": int(reach),
                "impressions": impressions,
                "clicks": clicks,
            }
        )

    transactions["campaign_id"] = campaign_assignment
    campaigns = pd.DataFrame(rows)

    # ---- Write realised results back onto the campaign table ------------ #
    valid = transactions[~transactions["is_return"] & transactions["campaign_id"].notna()]
    results = (
        valid.groupby("campaign_id")
        .agg(
            attributed_orders=("order_id", "nunique"),
            attributed_customers=("customer_id", "nunique"),
            attributed_net_revenue=("net_amount", "sum"),
        )
        .reset_index()
    )
    campaigns = campaigns.merge(results, on="campaign_id", how="left")
    for column in ("attributed_orders", "attributed_customers", "attributed_net_revenue"):
        campaigns[column] = campaigns[column].fillna(0)
    campaigns["attributed_net_revenue"] = campaigns["attributed_net_revenue"].round(2)
    campaigns["roas"] = (
        campaigns["attributed_net_revenue"] / campaigns["spend"]
    ).round(3)
    campaigns["attributed_orders"] = campaigns["attributed_orders"].astype(int)
    campaigns["attributed_customers"] = campaigns["attributed_customers"].astype(int)

    return campaigns, transactions


def _finalise_boutique_targets(
    rng: np.random.Generator, boutiques: pd.DataFrame, transactions: pd.DataFrame
) -> pd.DataFrame:
    """Set each boutique's annual sales target from its realised performance.

    Targets are set at annualised net sales x U(0.88, 1.12), which produces a
    realistic spread of boutiques landing above and below plan rather than a
    target column that is unrelated to the data.
    """
    valid = transactions[~transactions["is_return"]]
    years = (config.PERIOD_END - config.PERIOD_START).days / 365.25
    annualised = valid.groupby("boutique_id")["net_amount"].sum() / years

    boutiques = boutiques.copy()
    achieved = boutiques["boutique_id"].map(annualised).fillna(0.0)
    noise = rng.uniform(0.88, 1.12, size=len(boutiques))
    boutiques["annual_sales_target"] = np.round(achieved.to_numpy() * noise, -3)
    return boutiques


# --------------------------------------------------------------------------- #
# Orchestration                                                                #
# --------------------------------------------------------------------------- #


def generate_all(seed: int = config.RANDOM_SEED) -> dict[str, pd.DataFrame]:
    """Generate every table and return them keyed by table name.

    Args:
        seed: Random seed controlling the whole simulation.

    Returns:
        Mapping of table name -> DataFrame, in dependency order.
    """
    rng = np.random.default_rng(seed)

    boutiques = build_boutiques(rng)
    associates = build_associates(rng, boutiques)
    products = build_products(rng)
    customers = build_customers(rng, boutiques)
    interactions = build_clienteling_interactions(rng, customers, associates)
    transactions = generate_transactions(
        rng, customers, products, boutiques, associates, interactions
    )
    campaigns, transactions = build_campaigns(rng, transactions, customers)
    interactions = _assign_outcomes(rng, interactions)
    boutiques = _finalise_boutique_targets(rng, boutiques, transactions)

    return {
        "boutiques": boutiques,
        "associates": associates,
        "products": products,
        "customers": customers,
        "transactions": transactions,
        "clienteling_interactions": interactions,
        "campaigns": campaigns,
    }


def export_tables(tables: dict[str, pd.DataFrame]) -> None:
    """Write every table to CSV and to a single multi-sheet Excel workbook."""
    config.ensure_directories()

    for name, frame in tables.items():
        frame.to_csv(config.PROCESSED_DIR / config.RAW_TABLE_FILES[name], index=False)

    sheet_names = {
        "transactions": "Transactions",
        "customers": "Customer Master",
        "products": "Product Master",
        "boutiques": "Boutique Master",
        "associates": "Associates",
        "clienteling_interactions": "Clienteling Activities",
        "campaigns": "Marketing Campaigns",
    }
    with pd.ExcelWriter(config.EXCEL_WORKBOOK_PATH, engine="openpyxl") as writer:
        # Transactions first so the workbook opens on the fact table.
        for key in (
            "transactions",
            "customers",
            "products",
            "boutiques",
            "associates",
            "clienteling_interactions",
            "campaigns",
        ):
            tables[key].to_excel(writer, sheet_name=sheet_names[key], index=False)


def main() -> None:
    """Generate and export the full synthetic dataset."""
    tables = generate_all()
    export_tables(tables)

    transactions = tables["transactions"]
    net = transactions.loc[~transactions["is_return"], "net_amount"].sum()
    print("Phase 1 - synthetic data generated")
    print(f"  seed                : {config.RANDOM_SEED}")
    print(f"  period              : {config.PERIOD_START} -> {config.PERIOD_END}")
    for name, frame in tables.items():
        print(f"  {name:<26}: {len(frame):>7,} rows")
    print(f"  net sales ({config.CURRENCY})      : {net:,.0f}")
    print(f"  written to          : {config.PROCESSED_DIR.relative_to(config.REPO_ROOT)}")


if __name__ == "__main__":
    main()
