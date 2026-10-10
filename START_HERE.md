# START HERE

Everything for the **Luxury Retail CRM & Clienteling Analytics** project is in
this one folder. The generated data is already included, so nothing has to be
run before you can look at the output.

---

## If you have 2 minutes

Open **`data/processed/luxury_retail_clienteling_analytics_workbook.xlsx`** and
read the *Executive Dashboard* tab. 19 KPIs, three charts, all live formulas.

## If you have 15 minutes

1. **`README.md`** — the commercial story: problem, method, three findings
   written as Observation → Evidence → Implication → Action, recommendations,
   and the limitations.
2. The **workbook** above — nine analytical tabs.
3. The **Streamlit dashboard**:
   ```bash
   pip install -r requirements.txt
   streamlit run dashboard/streamlit_app.py
   ```

## If you want to verify the work

```bash
pip install -r requirements.txt
python -m src.run_pipeline      # regenerates every figure from seed 42
python -m pytest tests/ -v      # 21 tests
```

---

## What is in this package

### Documentation

| File | What it is |
|---|---|
| `START_HERE.md` | This index |
| `README.md` | Full project write-up: context, methodology, KPI definitions, findings, recommendations, limitations |
| `requirements.txt` | Pinned dependency versions |
| `.gitignore` | Excludes generated data from version control |

### Excel layer

| File | What it is |
|---|---|
| `data/processed/luxury_retail_clienteling_analytics_workbook.xlsx` | **The main Excel deliverable.** 18 tabs, 11 charts, 29 conditional-formatting ranges, 325,416 live formulas. Change a value on the Config tab and everything recalculates |
| `data/processed/luxury_retail_clienteling_analytics.xlsx` | The raw pipeline export: seven data tables, no analysis. Use this if you want to build your own analysis from scratch |
| `excel/excel_workbook_guide.md` | How to build the analysis by hand: PivotTable setup and exact column-letter formulas for Recency, Frequency, Monetary and the 14-day attribution rule |
| `excel/build_workbook.py` | Regenerates the analytical workbook |

### Dashboard

| File | What it is |
|---|---|
| `dashboard/streamlit_app.py` | Four-section interactive dashboard with Plotly charts and date, boutique and category filters |

### Power BI

| File | What it is |
|---|---|
| `bi/measures.dax` | All measures: Net Sales, ATV, AUR, Repeat Purchase Rate, Full-Price Share, Campaign ROAS, Clienteling Conversion Rate, plus time intelligence and concentration |
| `bi/power_query_setup.m` | Power Query (M) scripts that load every CSV with correct data types, plus the generated Date table |
| `bi/star_schema.md` | The data model: relationships, which to leave inactive and why, the Date table, and the modelling traps |

### Tableau

| File | What it is |
|---|---|
| `bi/calculated_fields.md` | Every measure as a Tableau calculated field, including the LOD expressions needed for Repeat Purchase Rate and clienteling attribution |
| `bi/star_schema.md` | Joins and relationships (shared with Power BI) |

### Python pipeline

| File | What it is |
|---|---|
| `src/config.py` | Paths, random seed, every documented assumption |
| `src/generate_data.py` | Phase 1 — synthetic data generation |
| `src/kpis.py` | Shared KPI definitions — the single source of truth |
| `src/customer_segmentation.py` | Phase 2a — RFM scoring and segmentation |
| `src/cohort_analysis.py` | Phase 2b — monthly retention cohorts |
| `src/customer_lifetime_value.py` | Phase 2c — 12-month CLV |
| `src/run_pipeline.py` | Runs everything with one command |
| `tests/test_analytics.py` | 21 tests covering RFM, CLV and the returns rule |

### Generated data (`data/processed/`)

| File | What it is |
|---|---|
| `transactions.csv` | 15,017 order lines over 24 months |
| `customers.csv` | 3,400 clients |
| `products.csv` | 240 SKUs |
| `boutiques.csv` | 10 boutiques |
| `associates.csv` | 56 sales associates |
| `clienteling_interactions.csv` | 3,972 appointments, calls and events |
| `campaigns.csv` | 26 marketing campaigns |
| `rfm_segments.csv` | Per-client RFM scores and segment |
| `customer_lifetime_value.csv` | Per-client 12-month CLV |
| `cohort_retention_rates.csv` | Retention matrix (rates) |
| `cohort_retention_counts.csv` | Retention matrix (client counts) |
| `kpi_summary.json` | **Every figure quoted in the README** — nothing is typed by hand |

---

## About the Power BI and Tableau files

This package contains **the DAX, the M scripts, the Tableau field definitions
and the model guide — not a `.pbix` or a `.twb` file.**

That is deliberate, not an omission. A `.pbix` is a proprietary binary and a
`.twb` is a Tableau-specific XML; neither can be produced reliably without the
application itself, and I would have no way to confirm the file actually opened
before handing it over. A workbook that fails to open is worse than no workbook,
because it costs the person their time before they discover it.

What is here is everything needed to build both in about twenty minutes:

**Power BI** — open Power BI Desktop → Get Data → Blank Query → paste each query
from `bi/power_query_setup.m` → build the relationships in `bi/star_schema.md` →
paste the measures from `bi/measures.dax`.

**Tableau** — connect to `data/processed/transactions.csv`, add the other CSVs
as described in `bi/calculated_fields.md`, then create the calculated fields
from that file.

If you would rather have the analysis already built and clickable, that is what
the Excel workbook and the Streamlit dashboard are for.

---

## Headline figures

All from `data/processed/kpi_summary.json`, seed 42, 24 months to 30 Sep 2026.

| Metric | Value |
|---|---|
| Net Sales | €32,194,535 |
| Transactions | 9,036 |
| Average Transaction Value | €3,563 |
| Average Unit Retail | €2,278 |
| Repeat Purchase Rate | 67.6% |
| Full-Price Sales Share | 85.2% |
| Appointment conversion rate | 53.6% |
| Blended campaign ROAS | 7.09x |
| Lapsing high-value clients | 311, holding €6.85M |
| Top decile revenue share | 42.7% |

---

*All data is synthetic, generated locally from a fixed seed. No real customer
data is used anywhere in this project.*
