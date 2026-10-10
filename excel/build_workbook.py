"""Build the full analytical Excel workbook from the generated dataset.

Produces ``data/processed/luxury_retail_clienteling_analytics_workbook.xlsx``:
the raw data sheets plus nine analytical tabs whose every number is a **live
Excel formula**. Change an assumption on the Config tab - the attribution
window, the lapsing threshold, the discount rate - and the whole workbook
recalculates.

Analytical tabs
---------------
Executive Dashboard     headline KPIs, monthly trend, segment and decile charts
RFM & CLV               per-client scores, segment economics, decile Pareto
Cohort Retention        acquisition-cohort heatmap
Boutique Performance    net sales against target, ATV, AUR, full-price share
Category & Product      category economics, top SKUs, return rates
Monthly Mix             full price versus markdown and channel mix over time
Clienteling ROI         conversion by type, by boutique and by associate
Campaign Effectiveness  ROAS by channel and by audience
Order Attribution       order-level last-touch resolution

Design notes
------------
* No ``XLOOKUP``/``UNIQUE``/``FILTER``. Lookups use ``INDEX``/``MATCH``, and the
  de-duplicated key lists are written as values - the one thing that
  legitimately belongs in Python rather than in a cell.
* Helper columns on ``Transactions`` are formulas and exist to keep the
  analytical tabs cheap. Counting distinct orders is the expensive operation in
  any transactional workbook; here it is reduced to an O(1) running rank by
  exploiting the fact that the lines of an order are contiguous.
* Every monetary helper zeroes out returned lines, so no analytical formula
  ever needs a boolean criterion and Net Sales cannot accidentally include
  returns.

Run after the main pipeline:

    python -m src.run_pipeline
    python excel/build_workbook.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import (
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
    IconSetRule,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from src import config, kpis  # noqa: E402

OUTPUT_PATH = (
    config.PROCESSED_DIR / "luxury_retail_clienteling_analytics_workbook.xlsx"
)

# --------------------------------------------------------------------------- #
# Styling                                                                      #
# --------------------------------------------------------------------------- #

FONT = "Arial"
INK = "FF16150F"
BRASS = "FF8C6F3F"
CLARET = "FF7A2E35"

HEADER_FILL = PatternFill("solid", fgColor="FF1F1D17")
ACCENT_FILL = PatternFill("solid", fgColor=BRASS)
BAND_FILL = PatternFill("solid", fgColor="FFF2EFE9")
INPUT_FILL = PatternFill("solid", fgColor="FFFFF6D5")
RISK_FILL = PatternFill("solid", fgColor="FFF7E4E4")

TITLE_FONT = Font(name=FONT, size=16, bold=True, color=INK)
SUB_FONT = Font(name=FONT, size=9, italic=True, color="FF6B655A")
HEADER_FONT = Font(name=FONT, size=10, bold=True, color="FFFFFFFF")
BODY_FONT = Font(name=FONT, size=10, color=INK)
BOLD_FONT = Font(name=FONT, size=10, bold=True, color=INK)
INPUT_FONT = Font(name=FONT, size=10, color="FF0000FF")  # blue = hardcoded input
KPI_LABEL_FONT = Font(name=FONT, size=10, color="FF6B655A")
SECTION_FONT = Font(name=FONT, size=12, bold=True, color=BRASS)

THIN = Side(style="thin", color="FFD8D2C6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY = '#,##0;(#,##0);"-"'
MONEY2 = '#,##0.00;(#,##0.00);"-"'
PCT1 = "0.0%"
PCT2 = "0.00%"
NUM0 = "#,##0"
NUM2 = "#,##0.00"
MULT = '0.00"x"'
DATE_FMT = "yyyy-mm-dd"

TAB_ANALYSIS = "8C6F3F"
TAB_DATA = "B9B2A6"
TAB_HERO = "16150F"

SHEET_RAW = {
    "transactions": "Transactions",
    "customers": "Customer Master",
    "products": "Product Master",
    "boutiques": "Boutique Master",
    "associates": "Associates",
    "clienteling_interactions": "Clienteling Activities",
    "campaigns": "Marketing Campaigns",
}


# --------------------------------------------------------------------------- #
# Small helpers                                                                #
# --------------------------------------------------------------------------- #


def write_table(ws, frame: pd.DataFrame, start_row: int = 1) -> None:
    """Write a DataFrame as a formatted block with a styled header row."""
    for j, column in enumerate(frame.columns, start=1):
        cell = ws.cell(row=start_row, column=j, value=str(column))
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for i, row in enumerate(frame.itertuples(index=False), start=start_row + 1):
        for j, value in enumerate(row, start=1):
            if isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()
            elif value is pd.NaT:
                value = None
            elif isinstance(value, float) and pd.isna(value):
                value = None
            cell = ws.cell(row=i, column=j, value=value)
            cell.font = BODY_FONT


def autosize(ws, max_width: int = 34) -> None:
    """Set pragmatic column widths from the header text and a sample of rows."""
    for column_cells in ws.iter_cols(min_row=1, max_row=min(ws.max_row, 60)):
        longest = 0
        letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            if cell.value is not None:
                longest = max(longest, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max(10, longest + 2), max_width)


def title_block(ws, title: str, subtitle: str) -> None:
    """Write the standard two-line heading at the top of an analytical sheet."""
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A2"] = subtitle
    ws["A2"].font = SUB_FONT
    ws.row_dimensions[2].height = 26
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")


def header_row(ws, row: int, headers: list[str], start_col: int = 1) -> None:
    """Write a styled header row."""
    for j, text in enumerate(headers, start=start_col):
        cell = ws.cell(row=row, column=j, value=text)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_block(ws, first_row: int, last_row: int, first_col: int, last_col: int) -> None:
    """Apply body font and borders to a rectangular block."""
    for r in range(first_row, last_row + 1):
        for c in range(first_col, last_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = BODY_FONT
            cell.border = BOX


def set_widths(ws, widths: list[int], start_col: int = 1) -> None:
    """Set a run of column widths."""
    for j, width in enumerate(widths, start=start_col):
        ws.column_dimensions[get_column_letter(j)].width = width


# --------------------------------------------------------------------------- #
# Build                                                                        #
# --------------------------------------------------------------------------- #


def main() -> None:
    """Assemble and save the analytical workbook."""
    tables = kpis.load_all_tables()

    transactions = tables["transactions"].copy()
    transactions = transactions.sort_values(
        ["transaction_date", "order_id", "transaction_id"]
    ).reset_index(drop=True)
    assert transactions.groupby(
        "order_id", sort=False
    ).ngroup().is_monotonic_increasing, "order lines are not contiguous"

    tx_last = len(transactions) + 1
    cl_last = len(tables["clienteling_interactions"]) + 1
    cp = tables["campaigns"]
    cp_last = len(cp) + 1
    cu_last = len(tables["customers"]) + 1
    pr = tables["products"]
    pr_last = len(pr) + 1
    bq = tables["boutiques"]
    bq_last = len(bq) + 1
    assoc = tables["associates"]
    as_last = len(assoc) + 1

    clients = sorted(kpis.valid_sales(transactions)["customer_id"].unique().tolist())
    rfm_first, rfm_last = 6, 5 + len(clients)

    wb = Workbook()
    wb.remove(wb.active)

    # ====================================================================== #
    # Raw data sheets                                                        #
    # ====================================================================== #
    for key, sheet_name in SHEET_RAW.items():
        ws = wb.create_sheet(sheet_name)
        frame = transactions if key == "transactions" else tables[key]
        write_table(ws, frame)
        ws.freeze_panes = "A2"
        ws.sheet_properties.tabColor = TAB_DATA
        autosize(ws, max_width=24)

    # ---- Transactions helper columns (live formulas) --------------------- #
    ws = wb["Transactions"]
    helper_names = [
        "net_sales_value", "order_month", "order_sale_rank", "sale_date",
        "full_price_value", "counts_as_order", "units_sold_value",
        "gross_sales_value", "returned_value", "is_campaign_influenced",
        "cogs_value", "category", "discounted_value",
    ]
    for offset, name in enumerate(helper_names):
        cell = ws.cell(row=1, column=20 + offset, value=name)
        cell.font = HEADER_FONT
        cell.fill = ACCENT_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    PR_SKU = f"'Product Master'!$A$2:$A${pr_last}"
    PR_CAT = f"'Product Master'!$C$2:$C${pr_last}"
    PR_NAME = f"'Product Master'!$B$2:$B${pr_last}"

    for r in range(2, tx_last + 1):
        ws.cell(row=r, column=20, value=f"=IF(P{r}=TRUE,0,N{r})")
        ws.cell(row=r, column=21, value=f"=EOMONTH(C{r},0)").number_format = DATE_FMT
        if r == 2:
            ws.cell(row=r, column=22, value=f"=IF(P{r}=TRUE,0,1)")
        else:
            ws.cell(
                row=r, column=22,
                value=f"=IF(B{r}<>B{r-1},IF(P{r}=TRUE,0,1),V{r-1}+IF(P{r}=TRUE,0,1))",
            )
        ws.cell(row=r, column=23, value=f'=IF(P{r}=TRUE,"",C{r})').number_format = DATE_FMT
        ws.cell(row=r, column=24, value=f"=IF(P{r}=TRUE,0,IF(Q{r}=TRUE,N{r},0))")
        ws.cell(row=r, column=25, value=f"=IF(P{r}=TRUE,0,IF(V{r}=1,1,0))")
        ws.cell(row=r, column=26, value=f"=IF(P{r}=TRUE,0,I{r})")
        ws.cell(row=r, column=27, value=f"=IF(P{r}=TRUE,0,M{r})")
        ws.cell(row=r, column=28, value=f"=IF(P{r}=TRUE,N{r},0)")
        ws.cell(row=r, column=29, value=f'=IF(S{r}="",0,1)')
        ws.cell(row=r, column=30, value=f"=IF(P{r}=TRUE,0,O{r}*I{r})")
        ws.cell(
            row=r, column=31,
            value=f'=IFERROR(INDEX({PR_CAT},MATCH(E{r},{PR_SKU},0)),"Unknown")',
        )
        ws.cell(row=r, column=32, value=f"=IF(P{r}=TRUE,0,IF(Q{r}=TRUE,0,N{r}))")
        for c in (20, 24, 27, 28, 30, 32):
            ws.cell(row=r, column=c).number_format = MONEY2
    ws.freeze_panes = "C2"

    T = f"Transactions!$T$2:$T${tx_last}"
    U = f"Transactions!$U$2:$U${tx_last}"
    W = f"Transactions!$W$2:$W${tx_last}"
    X = f"Transactions!$X$2:$X${tx_last}"
    Y = f"Transactions!$Y$2:$Y${tx_last}"
    Z = f"Transactions!$Z$2:$Z${tx_last}"
    AA = f"Transactions!$AA$2:$AA${tx_last}"
    AB = f"Transactions!$AB$2:$AB${tx_last}"
    AC = f"Transactions!$AC$2:$AC${tx_last}"
    AD = f"Transactions!$AD$2:$AD${tx_last}"
    AE = f"Transactions!$AE$2:$AE${tx_last}"
    AF = f"Transactions!$AF$2:$AF${tx_last}"
    TX_CUST = f"Transactions!$D$2:$D${tx_last}"
    TX_SKU = f"Transactions!$E$2:$E${tx_last}"
    TX_BTQ = f"Transactions!$F$2:$F${tx_last}"
    TX_CHAN = f"Transactions!$H$2:$H${tx_last}"
    TX_CAMP = f"Transactions!$S$2:$S${tx_last}"
    TX_ORDER = f"Transactions!$B$2:$B${tx_last}"

    # ====================================================================== #
    # Config                                                                 #
    # ====================================================================== #
    cfg = wb.create_sheet("Config")
    cfg.sheet_properties.tabColor = TAB_ANALYSIS
    title_block(
        cfg,
        "Configuration and assumptions",
        "Blue cells are inputs - change one and the whole workbook recalculates. "
        "Black cells are formulas derived from the data.",
    )
    header_row(cfg, 4, ["Assumption", "Value", "Basis"])

    assumptions = [
        ("SnapshotDate", config.SNAPSHOT_DATE, DATE_FMT,
         "Fixed analysis date. Recency is measured against this, never TODAY()."),
        ("AttributionWindow", config.CLIENTELING_ATTRIBUTION_WINDOW_DAYS, NUM0,
         "Days after an interaction in which a purchase is credited to it."),
        ("LapseDays", config.LAPSED_RECENCY_DAYS, NUM0,
         "No purchase for this many days = lapsing. Luxury cycles are seasonal."),
        ("NewTenureDays", config.NEW_CLIENT_TENURE_DAYS, NUM0,
         "First purchase within this many days = a new client."),
        ("HighValueScore", config.RFM_HIGH_VALUE_MONETARY_SCORE, NUM0,
         "Minimum M score to count as high value (top two quintiles)."),
        ("LoyalScore", config.RFM_LOYAL_FREQUENCY_SCORE, NUM0,
         "Minimum F score to count as loyal (top two quintiles)."),
        ("BaseRetention", 0.6005, PCT2,
         "12-month repeat rate. Source: cohort model, data/processed/kpi_summary.json "
         "(cohorts.repeat_rate_12m). Cannot be derived in Excel - it needs the "
         "acquisition-cohort logic in src/cohort_analysis.py."),
        ("DiscountRate", config.CLV_ANNUAL_DISCOUNT_RATE, PCT2,
         "Annual cost of capital, applied at the half-year point."),
    ]
    row = 5
    for name, value, fmt, basis in assumptions:
        cfg[f"A{row}"] = name
        cfg[f"A{row}"].font = BODY_FONT
        cell = cfg[f"B{row}"]
        cell.value = value
        cell.font = INPUT_FONT
        cell.fill = INPUT_FILL
        cell.number_format = fmt
        cell.border = BOX
        cfg[f"C{row}"] = basis
        cfg[f"C{row}"].font = SUB_FONT
        cfg[f"C{row}"].alignment = Alignment(wrap_text=True, vertical="top")
        wb.defined_names.add(DefinedName(name, attr_text=f"Config!$B${row}"))
        row += 1

    cfg[f"A{row}"] = "GrossMargin"
    cfg[f"A{row}"].font = BODY_FONT
    cfg[f"B{row}"] = f"=(SUM({T})-SUM({AD}))/SUM({T})"
    cfg[f"B{row}"].font = BODY_FONT
    cfg[f"B{row}"].number_format = PCT2
    cfg[f"B{row}"].border = BOX
    cfg[f"C{row}"] = (
        "Realised blended margin: (Net Sales - COGS) / Net Sales, after markdown. "
        "Calculated from the data, not assumed."
    )
    cfg[f"C{row}"].font = SUB_FONT
    cfg[f"C{row}"].alignment = Alignment(wrap_text=True, vertical="top")
    wb.defined_names.add(DefinedName("GrossMargin", attr_text=f"Config!$B${row}"))

    count_row = row + 3
    cfg[f"A{count_row}"] = "RFM scoring basis"
    cfg[f"A{count_row}"].font = SECTION_FONT
    cfg[f"A{count_row + 1}"] = "ClientCount"
    cfg[f"A{count_row + 1}"].font = BODY_FONT
    cfg[f"B{count_row + 1}"] = f"=COUNTA('RFM & CLV'!$A${rfm_first}:$A${rfm_last})"
    cfg[f"B{count_row + 1}"].font = BODY_FONT
    cfg[f"B{count_row + 1}"].number_format = NUM0
    cfg[f"B{count_row + 1}"].border = BOX
    cfg[f"C{count_row + 1}"] = (
        "Clients with a purchase history. Scores are rank-based quintiles: each client "
        "is ranked against the whole book, ties broken by position, then split into "
        "five equal groups. This matches the pandas qcut logic in "
        "src/customer_segmentation.py."
    )
    cfg[f"C{count_row + 1}"].font = SUB_FONT
    cfg[f"C{count_row + 1}"].alignment = Alignment(wrap_text=True, vertical="top")
    wb.defined_names.add(
        DefinedName("ClientCount", attr_text=f"Config!$B${count_row + 1}")
    )
    set_widths(cfg, [22, 14, 86])

    # ====================================================================== #
    # RFM & CLV                                                              #
    # ====================================================================== #
    rfm = wb.create_sheet("RFM & CLV")
    rfm.sheet_properties.tabColor = TAB_ANALYSIS
    title_block(
        rfm,
        "RFM scoring, segmentation and 12-month CLV",
        "One row per client with a purchase history. Every column is a live formula "
        "driven by the Config tab. Lapsing High-Value rows are shaded claret - that is "
        "the call list.",
    )
    header_row(rfm, 5, [
        "customer_id", "client_name", "recency_days", "frequency", "monetary",
        "tenure_days", "avg_order_value", "R", "F", "M", "rfm_score", "segment",
        "expected_orders_12m", "retention_rate", "predicted_revenue_12m", "clv_12m",
        "cohort_month", "monetary_rank", "value_decile",
    ])

    CUST_KEYS = f"'Customer Master'!$A$2:$A${cu_last}"
    CUST_NAMES = f"'Customer Master'!$B$2:$B${cu_last}"
    rec_all = f"$C${rfm_first}:$C${rfm_last}"
    frq_all = f"$D${rfm_first}:$D${rfm_last}"
    mon_all = f"$E${rfm_first}:$E${rfm_last}"

    for i, customer_id in enumerate(clients):
        r = rfm_first + i
        rfm.cell(row=r, column=1, value=customer_id)
        rfm.cell(
            row=r, column=2,
            value=f'=IFERROR(INDEX({CUST_NAMES},MATCH($A{r},{CUST_KEYS},0)),"Unknown")',
        )
        rfm.cell(
            row=r, column=3, value=f"=SnapshotDate-_xlfn.MAXIFS({W},{TX_CUST},$A{r})"
        ).number_format = NUM0
        rfm.cell(row=r, column=4, value=f"=SUMIFS({Y},{TX_CUST},$A{r})").number_format = NUM0
        rfm.cell(row=r, column=5, value=f"=SUMIFS({T},{TX_CUST},$A{r})").number_format = MONEY
        rfm.cell(
            row=r, column=6,
            value=f'=SnapshotDate-_xlfn.MINIFS({W},{TX_CUST},$A{r},{W},">0")',
        ).number_format = NUM0
        rfm.cell(row=r, column=7, value=f"=IFERROR($E{r}/$D{r},0)").number_format = MONEY
        # Rank-based quintiles: "how many score worse" + "ties seen so far",
        # which breaks ties by position, as pandas rank(method="first") does.
        rfm.cell(
            row=r, column=8,
            value=(
                f'=MIN(5,INT((COUNTIF({rec_all},">"&$C{r})'
                f'+COUNTIF($C${rfm_first}:$C{r},$C{r})-1)*5/ClientCount)+1)'
            ),
        )
        rfm.cell(
            row=r, column=9,
            value=(
                f'=MIN(5,INT((COUNTIF({frq_all},"<"&$D{r})'
                f'+COUNTIF($D${rfm_first}:$D{r},$D{r})-1)*5/ClientCount)+1)'
            ),
        )
        rfm.cell(
            row=r, column=18,
            value=f'=COUNTIF({mon_all},"<"&$E{r})+COUNTIF($E${rfm_first}:$E{r},$E{r})',
        ).number_format = NUM0
        rfm.cell(row=r, column=10, value=f"=MIN(5,INT(($R{r}-1)*5/ClientCount)+1)")
        rfm.cell(row=r, column=11, value=f"=$H{r}+$I{r}+$J{r}")
        rfm.cell(
            row=r, column=12,
            value=(
                f'=IF(AND($J{r}>=HighValueScore,$C{r}>LapseDays),"Lapsing High-Value",'
                f'IF(AND($J{r}>=HighValueScore,$I{r}>=LoyalScore,$C{r}<=LapseDays),'
                f'"High-Value Loyal",'
                f'IF(AND($F{r}<=NewTenureDays,$C{r}<=LapseDays),"Promising New",'
                f'"Infrequent")))'
            ),
        )
        rfm.cell(
            row=r, column=13, value=f"=MIN(12,$D{r}/MAX(0.5,$F{r}/365.25))"
        ).number_format = NUM2
        rfm.cell(
            row=r, column=14,
            value=f"=MEDIAN(0.05,BaseRetention*CHOOSE($H{r},0.35,0.6,0.85,1.05,1.2),0.95)",
        ).number_format = PCT1
        rfm.cell(row=r, column=15, value=f"=$G{r}*$M{r}*$N{r}").number_format = MONEY
        rfm.cell(
            row=r, column=16, value=f"=$O{r}*GrossMargin/(1+DiscountRate)^0.5"
        ).number_format = MONEY
        rfm.cell(
            row=r, column=17, value=f"=EOMONTH(SnapshotDate-$F{r},0)"
        ).number_format = "mmm-yy"
        rfm.cell(
            row=r, column=19, value=f"=MIN(10,INT((ClientCount-$R{r})*10/ClientCount)+1)"
        )
        for j in range(1, 20):
            rfm.cell(row=r, column=j).font = BODY_FONT

    rfm.freeze_panes = "C6"
    rfm.auto_filter.ref = f"A5:S{rfm_last}"
    set_widths(
        rfm,
        [14, 20, 11, 10, 13, 11, 14, 5, 5, 5, 10, 19, 15, 12, 17, 13, 12, 13, 11],
    )

    rfm.conditional_formatting.add(
        f"A{rfm_first}:S{rfm_last}",
        FormulaRule(
            formula=[f'$L{rfm_first}="Lapsing High-Value"'],
            fill=RISK_FILL,
            font=Font(name=FONT, size=10, color=CLARET),
            stopIfTrue=False,
        ),
    )
    for col in ("H", "I", "J"):
        rfm.conditional_formatting.add(
            f"{col}{rfm_first}:{col}{rfm_last}",
            ColorScaleRule(
                start_type="num", start_value=1, start_color="FFFFFFFF",
                mid_type="num", mid_value=3, mid_color="FFE3D7BE",
                end_type="num", end_value=5, end_color=BRASS,
            ),
        )
    rfm.conditional_formatting.add(
        f"E{rfm_first}:E{rfm_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="8C6F3F"),
    )
    rfm.conditional_formatting.add(
        f"P{rfm_first}:P{rfm_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="4F6650"),
    )

    SEG = f"'RFM & CLV'!$L${rfm_first}:$L${rfm_last}"
    MON = f"'RFM & CLV'!$E${rfm_first}:$E${rfm_last}"
    CLV = f"'RFM & CLV'!$P${rfm_first}:$P${rfm_last}"
    FRQ = f"'RFM & CLV'!$D${rfm_first}:$D${rfm_last}"
    AOV = f"'RFM & CLV'!$G${rfm_first}:$G${rfm_last}"
    REC = f"'RFM & CLV'!$C${rfm_first}:$C${rfm_last}"
    DEC = f"'RFM & CLV'!$S${rfm_first}:$S${rfm_last}"

    # ---- Segment economics (columns U..AB) ------------------------------- #
    sr = 5
    rfm["U3"] = "Segment economics"
    rfm["U3"].font = SECTION_FONT
    header_row(rfm, sr, [
        "segment", "clients", "% of clients", "net sales", "% of revenue",
        "avg order value", "avg recency (days)", "total 12m CLV",
    ], start_col=21)

    segments = ["High-Value Loyal", "Lapsing High-Value", "Promising New", "Infrequent"]
    seg_first, seg_last = sr + 1, sr + 4
    seg_total = sr + 5
    for i, seg in enumerate(segments):
        r = seg_first + i
        rfm[f"U{r}"] = seg
        rfm[f"V{r}"] = f"=COUNTIF({SEG},$U{r})"
        rfm[f"W{r}"] = f"=IFERROR($V{r}/$V${seg_total},0)"
        rfm[f"X{r}"] = f"=SUMIF({SEG},$U{r},{MON})"
        rfm[f"Y{r}"] = f"=IFERROR($X{r}/$X${seg_total},0)"
        rfm[f"Z{r}"] = f"=IFERROR(AVERAGEIF({SEG},$U{r},{AOV}),0)"
        rfm[f"AA{r}"] = f"=IFERROR(AVERAGEIF({SEG},$U{r},{REC}),0)"
        rfm[f"AB{r}"] = f"=SUMIF({SEG},$U{r},{CLV})"
        for col, fmt in (("V", NUM0), ("W", PCT1), ("X", MONEY), ("Y", PCT1),
                         ("Z", MONEY), ("AA", NUM0), ("AB", MONEY)):
            rfm[f"{col}{r}"].number_format = fmt
    rfm[f"U{seg_total}"] = "Total"
    for col, fmt in (("V", NUM0), ("W", PCT1), ("X", MONEY), ("Y", PCT1), ("AB", MONEY)):
        rfm[f"{col}{seg_total}"] = f"=SUM({col}{seg_first}:{col}{seg_last})"
        rfm[f"{col}{seg_total}"].number_format = fmt
    style_block(rfm, seg_first, seg_total, 21, 28)
    for c in range(21, 29):
        rfm.cell(row=seg_total, column=c).font = BOLD_FONT
    rfm.conditional_formatting.add(
        f"Y{seg_first}:Y{seg_last}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                    color="8C6F3F"),
    )
    set_widths(rfm, [20, 9, 12, 14, 12, 15, 16, 15], start_col=21)

    # ---- Decile Pareto (columns AD..AH) ---------------------------------- #
    rfm["AD3"] = "Revenue concentration by client decile"
    rfm["AD3"].font = SECTION_FONT
    header_row(
        rfm, sr,
        ["decile", "clients", "net sales", "% of revenue", "cumulative %"],
        start_col=30,
    )
    dec_first, dec_last = sr + 1, sr + 10
    for i in range(10):
        r = dec_first + i
        rfm[f"AD{r}"] = i + 1
        rfm[f"AE{r}"] = f"=COUNTIF({DEC},$AD{r})"
        rfm[f"AF{r}"] = f"=SUMIF({DEC},$AD{r},{MON})"
        rfm[f"AG{r}"] = f"=IFERROR($AF{r}/SUM($AF${dec_first}:$AF${dec_last}),0)"
        rfm[f"AH{r}"] = f"=SUM($AG${dec_first}:$AG{r})"
        for col, fmt in (("AE", NUM0), ("AF", MONEY), ("AG", PCT1), ("AH", PCT1)):
            rfm[f"{col}{r}"].number_format = fmt
    style_block(rfm, dec_first, dec_last, 30, 34)
    rfm.conditional_formatting.add(
        f"AG{dec_first}:AG{dec_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="7A2E35"),
    )
    set_widths(rfm, [9, 10, 14, 13, 13], start_col=30)

    # ====================================================================== #
    # Order Attribution                                                      #
    # ====================================================================== #
    orders = (
        kpis.valid_sales(transactions)
        .groupby("order_id", as_index=False)
        .agg(customer_id=("customer_id", "first"))
        .sort_values("order_id")
    )
    ords = wb.create_sheet("Order Attribution")
    ords.sheet_properties.tabColor = TAB_DATA
    title_block(
        ords,
        "Order-level attribution",
        "One row per order with a sold line. Each order is credited to the single most "
        "recent clienteling interaction within AttributionWindow days before it. "
        "Resolving this at order grain is what stops clienteling revenue being "
        "double-counted.",
    )
    header_row(ords, 5, [
        "order_id", "customer_id", "order_date", "order_net_amount", "last_touch_date",
    ])
    ord_first = 6
    ord_last = ord_first + len(orders) - 1
    CL_DATE_R = f"'Clienteling Activities'!$E$2:$E${cl_last}"
    CL_CUST_R = f"'Clienteling Activities'!$B$2:$B${cl_last}"

    for i, row_data in enumerate(orders.itertuples(index=False)):
        r = ord_first + i
        ords.cell(row=r, column=1, value=row_data.order_id)
        ords.cell(row=r, column=2, value=row_data.customer_id)
        ords.cell(
            row=r, column=3, value=f"=_xlfn.MINIFS({W},{TX_ORDER},$A{r})"
        ).number_format = DATE_FMT
        ords.cell(
            row=r, column=4, value=f"=SUMIFS({T},{TX_ORDER},$A{r})"
        ).number_format = MONEY2
        ords.cell(
            row=r, column=5,
            value=(
                f'=IF(COUNTIFS({CL_CUST_R},$B{r},{CL_DATE_R},"<="&$C{r},'
                f'{CL_DATE_R},">="&$C{r}-AttributionWindow)=0,"",'
                f'_xlfn.MAXIFS({CL_DATE_R},{CL_CUST_R},$B{r},{CL_DATE_R},"<="&$C{r},'
                f'{CL_DATE_R},">="&$C{r}-AttributionWindow))'
            ),
        ).number_format = DATE_FMT
        for j in range(1, 6):
            ords.cell(row=r, column=j).font = BODY_FONT
    ords.freeze_panes = "A6"
    set_widths(ords, [16, 15, 14, 18, 16])

    ORD_CUST = f"'Order Attribution'!$B${ord_first}:$B${ord_last}"
    ORD_NET = f"'Order Attribution'!$D${ord_first}:$D${ord_last}"
    ORD_LT = f"'Order Attribution'!$E${ord_first}:$E${ord_last}"

    # ---- Clienteling helper columns -------------------------------------- #
    cws = wb["Clienteling Activities"]
    for j, name in enumerate(
        ["orders_in_window", "converted_in_window", "attributed_orders",
         "attributed_revenue", "is_last_touch"],
        start=9,
    ):
        c = cws.cell(row=1, column=j, value=name)
        c.font = HEADER_FONT
        c.fill = ACCENT_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    for r in range(2, cl_last + 1):
        window = f"$E{r}+AttributionWindow"
        cws.cell(
            row=r, column=9,
            value=f'=SUMIFS({Y},{TX_CUST},$B{r},{W},">="&$E{r},{W},"<="&{window})',
        ).number_format = NUM0
        cws.cell(row=r, column=10, value=f"=IF($I{r}>0,1,0)")
        cws.cell(
            row=r, column=11, value=f"=COUNTIFS({ORD_CUST},$B{r},{ORD_LT},$E{r})"
        ).number_format = NUM0
        cws.cell(
            row=r, column=12, value=f"=SUMIFS({ORD_NET},{ORD_CUST},$B{r},{ORD_LT},$E{r})"
        ).number_format = MONEY2
        cws.cell(row=r, column=13, value=f"=IF($K{r}>0,1,0)")
        for j in range(9, 14):
            cws.cell(row=r, column=j).font = BODY_FONT
    cws.freeze_panes = "C2"

    CL_TYPE = f"'Clienteling Activities'!$F$2:$F${cl_last}"
    CL_ASSOC = f"'Clienteling Activities'!$C$2:$C${cl_last}"
    CL_BTQ = f"'Clienteling Activities'!$D$2:$D${cl_last}"
    CL_CONV = f"'Clienteling Activities'!$J$2:$J${cl_last}"
    CL_REV = f"'Clienteling Activities'!$L$2:$L${cl_last}"

    # ====================================================================== #
    # Cohort Retention                                                       #
    # ====================================================================== #
    cohort_rates = pd.read_csv(
        config.PROCESSED_DIR / config.ANALYTICS_FILES["cohort_retention_rates"],
        index_col=0,
    )
    cohort_counts = pd.read_csv(
        config.PROCESSED_DIR / config.ANALYTICS_FILES["cohort_retention_counts"],
        index_col=0,
    )
    coh = wb.create_sheet("Cohort Retention")
    coh.sheet_properties.tabColor = TAB_ANALYSIS
    title_block(
        coh,
        "Acquisition cohort retention",
        "Each row is a month of newly acquired clients; each column is the share still "
        "purchasing that many months later. Clients who joined before the analysis "
        "window are excluded, so established buyers cannot inflate the early cohorts. "
        "Blank cells are unobserved, not zero.",
    )
    coh["A4"] = (
        "Source: data/processed/cohort_retention_rates.csv, produced by "
        "src/cohort_analysis.py. This is the one analytical tab carrying imported "
        "values rather than live formulas - the acquisition-cohort logic needs a "
        "distinct-count-per-client-per-month matrix that would make the file unusable."
    )
    coh["A4"].font = SUB_FONT
    coh["A4"].alignment = Alignment(wrap_text=True, vertical="top")
    coh.row_dimensions[4].height = 26

    offsets = [c for c in cohort_rates.columns if int(c) <= 18]
    header_row(coh, 6, ["cohort", "clients"] + [f"M{int(c)}" for c in offsets])
    for i, (cohort_month, row_data) in enumerate(cohort_rates.iterrows()):
        r = 7 + i
        cell = coh.cell(row=r, column=1, value=pd.Timestamp(cohort_month).to_pydatetime())
        cell.number_format = "mmm-yy"
        size = cohort_counts.loc[cohort_month].get("0")
        coh.cell(
            row=r, column=2, value=int(size) if pd.notna(size) else None
        ).number_format = NUM0
        for j, offset in enumerate(offsets, start=3):
            value = row_data[offset]
            c = coh.cell(
                row=r, column=j, value=None if pd.isna(value) else round(float(value), 4)
            )
            c.number_format = PCT1
            c.font = BODY_FONT
        for j in range(1, 3):
            coh.cell(row=r, column=j).font = BODY_FONT
    coh_last = 6 + len(cohort_rates)
    coh.freeze_panes = "C7"
    coh.conditional_formatting.add(
        f"D7:{get_column_letter(2 + len(offsets))}{coh_last}",
        ColorScaleRule(
            start_type="num", start_value=0, start_color="FFFFFFFF",
            mid_type="percentile", mid_value=50, mid_color="FFE3D7BE",
            end_type="max", end_color=BRASS,
        ),
    )
    coh.conditional_formatting.add(
        f"B7:B{coh_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="B9B2A6"),
    )
    set_widths(coh, [12, 10] + [8] * len(offsets))
    coh.cell(
        row=coh_last + 2, column=1,
        value="Month 0 is 100% by construction. These cohorts are small (single- and "
              "double-digit client counts), so individual cells are noisy - read the "
              "shape of the columns, not any one cell.",
    ).font = SUB_FONT

    # ====================================================================== #
    # Boutique Performance                                                   #
    # ====================================================================== #
    btq = wb.create_sheet("Boutique Performance")
    btq.sheet_properties.tabColor = TAB_ANALYSIS
    title_block(
        btq,
        "Boutique performance",
        "Net sales against annual target, pro-rated for the analysis window, plus the "
        "trading metrics that explain the gap.",
    )
    header_row(btq, 5, [
        "boutique", "city", "format", "net sales", "annual target", "period target",
        "achievement", "orders", "ATV", "AUR", "full-price share",
        "clienteling share", "return rate",
    ])
    BQ_KEYS = f"'Boutique Master'!$A$2:$A${bq_last}"
    years = (config.PERIOD_END - config.PERIOD_START).days / 365.25
    b_first = 6
    b_last = b_first + len(bq) - 1
    b_total = b_last + 1
    for i, row_data in enumerate(bq.itertuples(index=False)):
        r = b_first + i
        btq.cell(row=r, column=1, value=row_data.boutique_id)
        btq.cell(
            row=r, column=2,
            value=f'=IFERROR(INDEX(\'Boutique Master\'!$C$2:$C${bq_last},'
                  f'MATCH($A{r},{BQ_KEYS},0)),"")',
        )
        btq.cell(
            row=r, column=3,
            value=f'=IFERROR(INDEX(\'Boutique Master\'!$F$2:$F${bq_last},'
                  f'MATCH($A{r},{BQ_KEYS},0)),"")',
        )
        btq.cell(row=r, column=4, value=f"=SUMIFS({T},{TX_BTQ},$A{r})").number_format = MONEY
        btq.cell(
            row=r, column=5,
            value=f"=IFERROR(INDEX('Boutique Master'!$I$2:$I${bq_last},"
                  f"MATCH($A{r},{BQ_KEYS},0)),0)",
        ).number_format = MONEY
        btq.cell(row=r, column=6, value=f"=$E{r}*{years:.4f}").number_format = MONEY
        btq.cell(row=r, column=7, value=f"=IFERROR($D{r}/$F{r},0)").number_format = PCT1
        btq.cell(row=r, column=8, value=f"=SUMIFS({Y},{TX_BTQ},$A{r})").number_format = NUM0
        btq.cell(row=r, column=9, value=f"=IFERROR($D{r}/$H{r},0)").number_format = MONEY
        btq.cell(
            row=r, column=10, value=f"=IFERROR($D{r}/SUMIFS({Z},{TX_BTQ},$A{r}),0)"
        ).number_format = MONEY
        btq.cell(
            row=r, column=11, value=f"=IFERROR(SUMIFS({X},{TX_BTQ},$A{r})/$D{r},0)"
        ).number_format = PCT1
        btq.cell(
            row=r, column=12,
            value=f'=IFERROR(SUMIFS({T},{TX_BTQ},$A{r},{TX_CHAN},"Clienteling")/$D{r},0)',
        ).number_format = PCT1
        btq.cell(
            row=r, column=13,
            value=f"=IFERROR(SUMIFS({AB},{TX_BTQ},$A{r})/"
                  f"(SUMIFS({AB},{TX_BTQ},$A{r})+$D{r}),0)",
        ).number_format = PCT1
    btq.cell(row=b_total, column=1, value="Total")
    for col, fmt in (("D", MONEY), ("E", MONEY), ("F", MONEY), ("H", NUM0)):
        btq[f"{col}{b_total}"] = f"=SUM({col}{b_first}:{col}{b_last})"
        btq[f"{col}{b_total}"].number_format = fmt
    btq[f"G{b_total}"] = f"=IFERROR($D{b_total}/$F{b_total},0)"
    btq[f"G{b_total}"].number_format = PCT1
    btq[f"I{b_total}"] = f"=IFERROR($D{b_total}/$H{b_total},0)"
    btq[f"I{b_total}"].number_format = MONEY
    style_block(btq, b_first, b_total, 1, 13)
    for c in range(1, 14):
        btq.cell(row=b_total, column=c).font = BOLD_FONT
    btq.conditional_formatting.add(
        f"G{b_first}:G{b_last}",
        IconSetRule("3TrafficLights1", "num", [0, 0.95, 1.0], showValue=True),
    )
    btq.conditional_formatting.add(
        f"D{b_first}:D{b_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="8C6F3F"),
    )
    btq.conditional_formatting.add(
        f"K{b_first}:K{b_last}",
        ColorScaleRule(start_type="min", start_color="FFF7E4E4",
                       end_type="max", end_color="FFDCE8DC"),
    )
    set_widths(btq, [12, 14, 12, 14, 14, 14, 13, 10, 12, 12, 15, 16, 12])

    chart = BarChart()
    chart.type = "col"
    chart.title = "Net sales against period target by boutique"
    chart.height, chart.width = 9, 22
    chart.add_data(Reference(btq, min_col=4, min_row=5, max_row=b_last),
                   titles_from_data=True)
    chart.add_data(Reference(btq, min_col=6, min_row=5, max_row=b_last),
                   titles_from_data=True)
    chart.set_categories(Reference(btq, min_col=1, min_row=b_first, max_row=b_last))
    chart.y_axis.title = "EUR"
    btq.add_chart(chart, f"A{b_total + 3}")

    # ====================================================================== #
    # Category & Product                                                     #
    # ====================================================================== #
    cat = wb.create_sheet("Category & Product")
    cat.sheet_properties.tabColor = TAB_ANALYSIS
    title_block(
        cat,
        "Category and product performance",
        "Where the revenue and the margin actually sit, and which categories carry the "
        "return and markdown problems.",
    )
    header_row(cat, 5, [
        "category", "net sales", "% of net sales", "units", "AUR", "full-price share",
        "markdown depth", "return rate", "gross margin",
    ])
    categories = sorted(pr["category"].unique().tolist())
    c_first = 6
    c_last = c_first + len(categories) - 1
    c_total = c_last + 1
    for i, category in enumerate(categories):
        r = c_first + i
        cat.cell(row=r, column=1, value=category)
        cat.cell(row=r, column=2, value=f"=SUMIFS({T},{AE},$A{r})").number_format = MONEY
        cat.cell(
            row=r, column=3, value=f"=IFERROR($B{r}/$B${c_total},0)"
        ).number_format = PCT1
        cat.cell(row=r, column=4, value=f"=SUMIFS({Z},{AE},$A{r})").number_format = NUM0
        cat.cell(row=r, column=5, value=f"=IFERROR($B{r}/$D{r},0)").number_format = MONEY
        cat.cell(
            row=r, column=6, value=f"=IFERROR(SUMIFS({X},{AE},$A{r})/$B{r},0)"
        ).number_format = PCT1
        cat.cell(
            row=r, column=7, value=f"=IFERROR(1-$B{r}/SUMIFS({AA},{AE},$A{r}),0)"
        ).number_format = PCT1
        cat.cell(
            row=r, column=8,
            value=f"=IFERROR(SUMIFS({AB},{AE},$A{r})/"
                  f"(SUMIFS({AB},{AE},$A{r})+$B{r}),0)",
        ).number_format = PCT1
        cat.cell(
            row=r, column=9,
            value=f"=IFERROR(($B{r}-SUMIFS({AD},{AE},$A{r}))/$B{r},0)",
        ).number_format = PCT1
    cat.cell(row=c_total, column=1, value="Total")
    for col, fmt in (("B", MONEY), ("C", PCT1), ("D", NUM0)):
        cat[f"{col}{c_total}"] = f"=SUM({col}{c_first}:{col}{c_last})"
        cat[f"{col}{c_total}"].number_format = fmt
    cat[f"E{c_total}"] = f"=IFERROR($B{c_total}/$D{c_total},0)"
    cat[f"E{c_total}"].number_format = MONEY
    style_block(cat, c_first, c_total, 1, 9)
    for c in range(1, 10):
        cat.cell(row=c_total, column=c).font = BOLD_FONT
    cat.conditional_formatting.add(
        f"B{c_first}:B{c_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="8C6F3F"),
    )
    cat.conditional_formatting.add(
        f"H{c_first}:H{c_last}",
        ColorScaleRule(start_type="min", start_color="FFDCE8DC",
                       end_type="max", end_color="FFF2C9C9"),
    )
    set_widths(cat, [16, 14, 15, 10, 12, 15, 15, 12, 13])

    cat_chart = BarChart()
    cat_chart.type = "bar"
    cat_chart.title = "Net sales by category"
    cat_chart.height, cat_chart.width = 7, 16
    cat_chart.add_data(Reference(cat, min_col=2, min_row=5, max_row=c_last),
                       titles_from_data=True)
    cat_chart.set_categories(Reference(cat, min_col=1, min_row=c_first, max_row=c_last))
    cat_chart.legend = None
    cat.add_chart(cat_chart, "K5")

    valid = kpis.valid_sales(transactions)
    top_skus = valid.groupby("sku")["net_amount"].sum().nlargest(15).index.tolist()
    top_row = c_total + 3
    cat.cell(row=top_row, column=1, value="Top 15 SKUs by net sales").font = SECTION_FONT
    cat.cell(
        row=top_row + 1, column=1,
        value="SKU list selected at build time; every figure beside it is a live formula.",
    ).font = SUB_FONT
    header_row(cat, top_row + 2, [
        "sku", "product", "category", "net sales", "units", "AUR", "full-price share",
    ])
    t_first = top_row + 3
    for i, sku in enumerate(top_skus):
        r = t_first + i
        cat.cell(row=r, column=1, value=sku)
        cat.cell(
            row=r, column=2,
            value=f'=IFERROR(INDEX({PR_NAME},MATCH($A{r},{PR_SKU},0)),"")',
        )
        cat.cell(
            row=r, column=3,
            value=f'=IFERROR(INDEX({PR_CAT},MATCH($A{r},{PR_SKU},0)),"")',
        )
        cat.cell(row=r, column=4, value=f"=SUMIFS({T},{TX_SKU},$A{r})").number_format = MONEY
        cat.cell(row=r, column=5, value=f"=SUMIFS({Z},{TX_SKU},$A{r})").number_format = NUM0
        cat.cell(row=r, column=6, value=f"=IFERROR($D{r}/$E{r},0)").number_format = MONEY
        cat.cell(
            row=r, column=7, value=f"=IFERROR(SUMIFS({X},{TX_SKU},$A{r})/$D{r},0)"
        ).number_format = PCT1
    t_last = t_first + len(top_skus) - 1
    style_block(cat, t_first, t_last, 1, 7)
    cat.conditional_formatting.add(
        f"D{t_first}:D{t_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="4F6650"),
    )

    # ====================================================================== #
    # Monthly Mix                                                            #
    # ====================================================================== #
    mix = wb.create_sheet("Monthly Mix")
    mix.sheet_properties.tabColor = TAB_ANALYSIS
    title_block(
        mix,
        "Monthly mix: full price, markdown and channel",
        "Where the markdown months are, and how much of the business the clienteling "
        "channel carries each month.",
    )
    header_row(mix, 5, [
        "month", "net sales", "full-price sales", "discounted sales",
        "full-price share", "orders", "ATV", "boutique", "online", "clienteling",
        "clienteling share",
    ])
    months = (
        pd.period_range(config.PERIOD_START, config.PERIOD_END, freq="M")
        .to_timestamp(how="end")
        .normalize()
    )
    m_first = 6
    m_last = m_first + len(months) - 1
    for i, month in enumerate(months):
        r = m_first + i
        mix.cell(row=r, column=1, value=month.to_pydatetime()).number_format = "mmm-yy"
        mix.cell(row=r, column=2, value=f"=SUMIFS({T},{U},$A{r})").number_format = MONEY
        mix.cell(row=r, column=3, value=f"=SUMIFS({X},{U},$A{r})").number_format = MONEY
        mix.cell(row=r, column=4, value=f"=SUMIFS({AF},{U},$A{r})").number_format = MONEY
        mix.cell(row=r, column=5, value=f"=IFERROR($C{r}/$B{r},0)").number_format = PCT1
        mix.cell(row=r, column=6, value=f"=SUMIFS({Y},{U},$A{r})").number_format = NUM0
        mix.cell(row=r, column=7, value=f"=IFERROR($B{r}/$F{r},0)").number_format = MONEY
        for j, channel in enumerate(["Boutique", "Online", "Clienteling"], start=8):
            mix.cell(
                row=r, column=j, value=f'=SUMIFS({T},{U},$A{r},{TX_CHAN},"{channel}")'
            ).number_format = MONEY
        mix.cell(row=r, column=11, value=f"=IFERROR($J{r}/$B{r},0)").number_format = PCT1
    style_block(mix, m_first, m_last, 1, 11)
    mix.conditional_formatting.add(
        f"E{m_first}:E{m_last}",
        ColorScaleRule(start_type="min", start_color="FFF2C9C9",
                       mid_type="percentile", mid_value=50, mid_color="FFFFF4D6",
                       end_type="max", end_color="FFDCE8DC"),
    )
    mix.conditional_formatting.add(
        f"B{m_first}:B{m_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="8C6F3F"),
    )
    set_widths(mix, [11, 14, 15, 15, 14, 9, 11, 13, 12, 13, 15])

    stack = BarChart()
    stack.type = "col"
    stack.grouping = "stacked"
    stack.overlap = 100
    stack.title = "Full-price versus discounted net sales by month"
    stack.height, stack.width = 9, 24
    stack.add_data(Reference(mix, min_col=3, min_row=5, max_row=m_last),
                   titles_from_data=True)
    stack.add_data(Reference(mix, min_col=4, min_row=5, max_row=m_last),
                   titles_from_data=True)
    stack.set_categories(Reference(mix, min_col=1, min_row=m_first, max_row=m_last))
    stack.y_axis.title = "EUR"
    mix.add_chart(stack, f"A{m_last + 3}")

    chan_chart = BarChart()
    chan_chart.type = "col"
    chan_chart.grouping = "stacked"
    chan_chart.overlap = 100
    chan_chart.title = "Net sales by channel by month"
    chan_chart.height, chan_chart.width = 9, 24
    for col in (8, 9, 10):
        chan_chart.add_data(Reference(mix, min_col=col, min_row=5, max_row=m_last),
                            titles_from_data=True)
    chan_chart.set_categories(Reference(mix, min_col=1, min_row=m_first, max_row=m_last))
    mix.add_chart(chan_chart, f"A{m_last + 22}")

    # ====================================================================== #
    # Clienteling ROI                                                        #
    # ====================================================================== #
    cl_sheet = wb.create_sheet("Clienteling ROI")
    cl_sheet.sheet_properties.tabColor = TAB_ANALYSIS
    title_block(
        cl_sheet,
        "Clienteling ROI",
        "An interaction CONVERTS if any qualifying purchase follows it within "
        "AttributionWindow days. Revenue is ATTRIBUTED to one interaction per order, "
        "last touch. Those are different questions with different answers.",
    )
    cl_sheet["A4"] = "Conversion and revenue by interaction type"
    cl_sheet["A4"].font = SECTION_FONT
    header_row(cl_sheet, 5, [
        "interaction type", "interactions", "conversions", "conversion rate",
        "attributed revenue", "revenue per interaction",
    ])
    for i, itype in enumerate(["Appointment", "Call", "Event"]):
        r = 6 + i
        cl_sheet[f"A{r}"] = itype
        cl_sheet[f"B{r}"] = f"=COUNTIF({CL_TYPE},$A{r})"
        cl_sheet[f"C{r}"] = f"=SUMIF({CL_TYPE},$A{r},{CL_CONV})"
        cl_sheet[f"D{r}"] = f"=IFERROR($C{r}/$B{r},0)"
        cl_sheet[f"E{r}"] = f"=SUMIF({CL_TYPE},$A{r},{CL_REV})"
        cl_sheet[f"F{r}"] = f"=IFERROR($E{r}/$B{r},0)"
        for col, fmt in (("B", NUM0), ("C", NUM0), ("D", PCT1), ("E", MONEY), ("F", MONEY)):
            cl_sheet[f"{col}{r}"].number_format = fmt
    style_block(cl_sheet, 6, 8, 1, 6)
    cl_sheet.conditional_formatting.add(
        "D6:D8",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                    color="8C6F3F"),
    )
    cl_sheet.conditional_formatting.add(
        "F6:F8",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="4F6650"),
    )

    conv_chart = BarChart()
    conv_chart.type = "bar"
    conv_chart.title = "Conversion rate by interaction type"
    conv_chart.height, conv_chart.width = 6, 14
    conv_chart.add_data(Reference(cl_sheet, min_col=4, min_row=5, max_row=8),
                        titles_from_data=True)
    conv_chart.set_categories(Reference(cl_sheet, min_col=1, min_row=6, max_row=8))
    conv_chart.legend = None
    cl_sheet.add_chart(conv_chart, "H4")

    rev_chart = BarChart()
    rev_chart.type = "bar"
    rev_chart.title = "Attributed revenue per interaction"
    rev_chart.height, rev_chart.width = 6, 14
    rev_chart.add_data(Reference(cl_sheet, min_col=6, min_row=5, max_row=8),
                       titles_from_data=True)
    rev_chart.set_categories(Reference(cl_sheet, min_col=1, min_row=6, max_row=8))
    rev_chart.legend = None
    cl_sheet.add_chart(rev_chart, "H17")

    bsec = 11
    cl_sheet[f"A{bsec}"] = "Clienteling by boutique"
    cl_sheet[f"A{bsec}"].font = SECTION_FONT
    header_row(cl_sheet, bsec + 1, [
        "boutique", "interactions", "appointments", "conversion rate",
        "attributed revenue",
    ])
    for i, row_data in enumerate(bq.itertuples(index=False)):
        r = bsec + 2 + i
        cl_sheet[f"A{r}"] = row_data.boutique_id
        cl_sheet[f"B{r}"] = f"=COUNTIF({CL_BTQ},$A{r})"
        cl_sheet[f"C{r}"] = f'=COUNTIFS({CL_BTQ},$A{r},{CL_TYPE},"Appointment")'
        cl_sheet[f"D{r}"] = f"=IFERROR(SUMIF({CL_BTQ},$A{r},{CL_CONV})/$B{r},0)"
        cl_sheet[f"E{r}"] = f"=SUMIF({CL_BTQ},$A{r},{CL_REV})"
        for col, fmt in (("B", NUM0), ("C", NUM0), ("D", PCT1), ("E", MONEY)):
            cl_sheet[f"{col}{r}"].number_format = fmt
    b_sec_last = bsec + 1 + len(bq)
    style_block(cl_sheet, bsec + 2, b_sec_last, 1, 5)
    cl_sheet.conditional_formatting.add(
        f"D{bsec + 2}:D{b_sec_last}",
        ColorScaleRule(start_type="min", start_color="FFF2C9C9",
                       end_type="max", end_color="FFDCE8DC"),
    )
    cl_sheet.conditional_formatting.add(
        f"E{bsec + 2}:E{b_sec_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="8C6F3F"),
    )

    asec = b_sec_last + 3
    cl_sheet[f"A{asec}"] = "Performance by sales associate"
    cl_sheet[f"A{asec}"].font = SECTION_FONT
    cl_sheet[f"A{asec + 1}"] = (
        "Associates below the minimum are shown but must not be ranked: small "
        "denominators produce dramatic percentages that mean nothing."
    )
    cl_sheet[f"A{asec + 1}"].font = SUB_FONT
    cl_sheet[f"G{asec}"] = "Min appointments to rank"
    cl_sheet[f"G{asec}"].font = BODY_FONT
    cl_sheet[f"I{asec}"] = 15
    cl_sheet[f"I{asec}"].font = INPUT_FONT
    cl_sheet[f"I{asec}"].fill = INPUT_FILL
    cl_sheet[f"I{asec}"].border = BOX
    wb.defined_names.add(
        DefinedName("MinAppointments", attr_text=f"'Clienteling ROI'!$I${asec}")
    )
    header_row(cl_sheet, asec + 2, [
        "associate_id", "associate_name", "boutique", "interactions", "appointments",
        "appointment conversions", "appointment conversion rate", "attributed revenue",
        "revenue per appointment", "ranked?",
    ])
    ASSOC_KEYS = f"Associates!$A$2:$A${as_last}"
    ASSOC_NAMES = f"Associates!$B$2:$B${as_last}"
    ASSOC_BTQ = f"Associates!$C$2:$C${as_last}"
    a_first = asec + 3
    for i, associate_id in enumerate(assoc["associate_id"].tolist()):
        r = a_first + i
        cl_sheet[f"A{r}"] = associate_id
        cl_sheet[f"B{r}"] = f'=IFERROR(INDEX({ASSOC_NAMES},MATCH($A{r},{ASSOC_KEYS},0)),"")'
        cl_sheet[f"C{r}"] = f'=IFERROR(INDEX({ASSOC_BTQ},MATCH($A{r},{ASSOC_KEYS},0)),"")'
        cl_sheet[f"D{r}"] = f"=COUNTIF({CL_ASSOC},$A{r})"
        cl_sheet[f"E{r}"] = f'=COUNTIFS({CL_ASSOC},$A{r},{CL_TYPE},"Appointment")'
        cl_sheet[f"F{r}"] = f'=SUMIFS({CL_CONV},{CL_ASSOC},$A{r},{CL_TYPE},"Appointment")'
        cl_sheet[f"G{r}"] = f'=IFERROR($F{r}/$E{r},"")'
        cl_sheet[f"H{r}"] = f"=SUMIF({CL_ASSOC},$A{r},{CL_REV})"
        cl_sheet[f"I{r}"] = f'=IFERROR($H{r}/$E{r},"")'
        cl_sheet[f"J{r}"] = f'=IF($E{r}>=MinAppointments,"Yes","Too few")'
        for col, fmt in (("D", NUM0), ("E", NUM0), ("F", NUM0), ("G", PCT1),
                         ("H", MONEY), ("I", MONEY)):
            cl_sheet[f"{col}{r}"].number_format = fmt
    a_last = a_first + len(assoc) - 1
    style_block(cl_sheet, a_first, a_last, 1, 10)
    cl_sheet.auto_filter.ref = f"A{asec + 2}:J{a_last}"
    cl_sheet.conditional_formatting.add(
        f"G{a_first}:G{a_last}",
        ColorScaleRule(start_type="min", start_color="FFF2C9C9",
                       mid_type="percentile", mid_value=50, mid_color="FFFFF4D6",
                       end_type="max", end_color="FFDCE8DC"),
    )
    cl_sheet.conditional_formatting.add(
        f"H{a_first}:H{a_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="8C6F3F"),
    )
    cl_sheet.conditional_formatting.add(
        f"J{a_first}:J{a_last}",
        FormulaRule(formula=[f'$J{a_first}="Too few"'],
                    font=Font(name=FONT, size=10, italic=True, color="FF9A948A")),
    )
    set_widths(cl_sheet, [14, 20, 11, 12, 13, 16, 17, 17, 16, 11])

    # ====================================================================== #
    # Campaign Effectiveness                                                 #
    # ====================================================================== #
    camp_ws = wb["Marketing Campaigns"]
    for j, name in enumerate(
        ["live_attributed_orders", "live_attributed_revenue", "live_roas"], start=15
    ):
        c = camp_ws.cell(row=1, column=j, value=name)
        c.font = HEADER_FONT
        c.fill = ACCENT_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for r in range(2, cp_last + 1):
        camp_ws.cell(row=r, column=15, value=f"=SUMIFS({Y},{TX_CAMP},$A{r})"
                     ).number_format = NUM0
        camp_ws.cell(row=r, column=16, value=f"=SUMIFS({T},{TX_CAMP},$A{r})"
                     ).number_format = MONEY
        camp_ws.cell(row=r, column=17, value=f"=IFERROR($P{r}/$G{r},0)"
                     ).number_format = MULT
        for j in (15, 16, 17):
            camp_ws.cell(row=r, column=j).font = BODY_FONT

    CP_CHANNEL = f"'Marketing Campaigns'!$D$2:$D${cp_last}"
    CP_AUD = f"'Marketing Campaigns'!$C$2:$C${cp_last}"
    CP_SPEND = f"'Marketing Campaigns'!$G$2:$G${cp_last}"
    CP_ORDERS = f"'Marketing Campaigns'!$O$2:$O${cp_last}"
    CP_REV = f"'Marketing Campaigns'!$P$2:$P${cp_last}"

    cmp_sheet = wb.create_sheet("Campaign Effectiveness")
    cmp_sheet.sheet_properties.tabColor = TAB_ANALYSIS
    title_block(
        cmp_sheet,
        "Campaign effectiveness",
        "Attributed revenue is recalculated live from the campaign_id on each "
        "transaction line, so ROAS responds to the data rather than to a stored result.",
    )
    cmp_sheet["A4"] = "Return on ad spend by channel"
    cmp_sheet["A4"].font = SECTION_FONT
    header_row(cmp_sheet, 5, [
        "channel", "campaigns", "spend", "% of spend", "attributed orders",
        "attributed net revenue", "% of revenue", "ROAS", "cost per order",
    ])
    channels = sorted(cp["channel"].unique().tolist())
    ch_first = 6
    ch_last = ch_first + len(channels) - 1
    ch_total = ch_last + 1
    for i, channel in enumerate(channels):
        r = ch_first + i
        cmp_sheet[f"A{r}"] = channel
        cmp_sheet[f"B{r}"] = f"=COUNTIF({CP_CHANNEL},$A{r})"
        cmp_sheet[f"C{r}"] = f"=SUMIF({CP_CHANNEL},$A{r},{CP_SPEND})"
        cmp_sheet[f"D{r}"] = f"=IFERROR($C{r}/$C${ch_total},0)"
        cmp_sheet[f"E{r}"] = f"=SUMIF({CP_CHANNEL},$A{r},{CP_ORDERS})"
        cmp_sheet[f"F{r}"] = f"=SUMIF({CP_CHANNEL},$A{r},{CP_REV})"
        cmp_sheet[f"G{r}"] = f"=IFERROR($F{r}/$F${ch_total},0)"
        cmp_sheet[f"H{r}"] = f"=IFERROR($F{r}/$C{r},0)"
        cmp_sheet[f"I{r}"] = f"=IFERROR($C{r}/$E{r},0)"
        for col, fmt in (("B", NUM0), ("C", MONEY), ("D", PCT1), ("E", NUM0),
                         ("F", MONEY), ("G", PCT1), ("H", MULT), ("I", MONEY)):
            cmp_sheet[f"{col}{r}"].number_format = fmt
    cmp_sheet[f"A{ch_total}"] = "Blended"
    for col, fmt in (("B", NUM0), ("C", MONEY), ("D", PCT1), ("E", NUM0),
                     ("F", MONEY), ("G", PCT1)):
        cmp_sheet[f"{col}{ch_total}"] = f"=SUM({col}{ch_first}:{col}{ch_last})"
        cmp_sheet[f"{col}{ch_total}"].number_format = fmt
    cmp_sheet[f"H{ch_total}"] = f"=IFERROR($F{ch_total}/$C{ch_total},0)"
    cmp_sheet[f"H{ch_total}"].number_format = MULT
    cmp_sheet[f"I{ch_total}"] = f"=IFERROR($C{ch_total}/$E{ch_total},0)"
    cmp_sheet[f"I{ch_total}"].number_format = MONEY
    style_block(cmp_sheet, ch_first, ch_total, 1, 9)
    for c in range(1, 10):
        cmp_sheet.cell(row=ch_total, column=c).font = BOLD_FONT
    cmp_sheet.conditional_formatting.add(
        f"H{ch_first}:H{ch_last}",
        ColorScaleRule(start_type="min", start_color="FFF2C9C9",
                       mid_type="percentile", mid_value=50, mid_color="FFFFF4D6",
                       end_type="max", end_color="FFDCE8DC"),
    )
    cmp_sheet.conditional_formatting.add(
        f"C{ch_first}:C{ch_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="B9B2A6"),
    )
    set_widths(cmp_sheet, [18, 11, 13, 12, 16, 19, 13, 10, 14])

    roas_chart = BarChart()
    roas_chart.type = "bar"
    roas_chart.title = "ROAS by channel"
    roas_chart.height, roas_chart.width = 8, 15
    roas_chart.add_data(Reference(cmp_sheet, min_col=8, min_row=5, max_row=ch_last),
                        titles_from_data=True)
    roas_chart.set_categories(Reference(cmp_sheet, min_col=1, min_row=ch_first,
                                        max_row=ch_last))
    roas_chart.legend = None
    cmp_sheet.add_chart(roas_chart, "K4")

    spend_chart = BarChart()
    spend_chart.type = "col"
    spend_chart.title = "Spend against attributed revenue by channel"
    spend_chart.height, spend_chart.width = 8, 15
    spend_chart.add_data(Reference(cmp_sheet, min_col=3, min_row=5, max_row=ch_last),
                         titles_from_data=True)
    spend_chart.add_data(Reference(cmp_sheet, min_col=6, min_row=5, max_row=ch_last),
                         titles_from_data=True)
    spend_chart.set_categories(Reference(cmp_sheet, min_col=1, min_row=ch_first,
                                         max_row=ch_last))
    cmp_sheet.add_chart(spend_chart, "K21")

    aud_row = ch_total + 3
    cmp_sheet[f"A{aud_row}"] = "Return by target audience"
    cmp_sheet[f"A{aud_row}"].font = SECTION_FONT
    header_row(cmp_sheet, aud_row + 1, [
        "audience", "campaigns", "spend", "attributed net revenue", "ROAS",
    ])
    audiences = sorted(cp["audience_segment"].unique().tolist())
    au_first = aud_row + 2
    for i, audience in enumerate(audiences):
        r = au_first + i
        cmp_sheet[f"A{r}"] = audience
        cmp_sheet[f"B{r}"] = f"=COUNTIF({CP_AUD},$A{r})"
        cmp_sheet[f"C{r}"] = f"=SUMIF({CP_AUD},$A{r},{CP_SPEND})"
        cmp_sheet[f"D{r}"] = f"=SUMIF({CP_AUD},$A{r},{CP_REV})"
        cmp_sheet[f"E{r}"] = f"=IFERROR($D{r}/$C{r},0)"
        for col, fmt in (("B", NUM0), ("C", MONEY), ("D", MONEY), ("E", MULT)):
            cmp_sheet[f"{col}{r}"].number_format = fmt
    au_last = au_first + len(audiences) - 1
    style_block(cmp_sheet, au_first, au_last, 1, 5)
    cmp_sheet.conditional_formatting.add(
        f"E{au_first}:E{au_last}",
        ColorScaleRule(start_type="min", start_color="FFF2C9C9",
                       end_type="max", end_color="FFDCE8DC"),
    )

    fp_row = au_last + 3
    cmp_sheet[f"A{fp_row}"] = "Full-price versus discounted mix"
    cmp_sheet[f"A{fp_row}"].font = SECTION_FONT
    header_row(cmp_sheet, fp_row + 1, [
        "group", "net sales", "full-price sales", "full-price share", "discounted share",
    ])
    for i, (label, flag) in enumerate(
        [("Campaign-influenced", 1), ("Not campaign-influenced", 0)]
    ):
        r = fp_row + 2 + i
        cmp_sheet[f"A{r}"] = label
        cmp_sheet[f"B{r}"] = f"=SUMIFS({T},{AC},{flag})"
        cmp_sheet[f"C{r}"] = f"=SUMIFS({X},{AC},{flag})"
        cmp_sheet[f"D{r}"] = f"=IFERROR($C{r}/$B{r},0)"
        cmp_sheet[f"E{r}"] = f"=1-$D{r}"
        for col, fmt in (("B", MONEY), ("C", MONEY), ("D", PCT1), ("E", PCT1)):
            cmp_sheet[f"{col}{r}"].number_format = fmt
    style_block(cmp_sheet, fp_row + 2, fp_row + 3, 1, 5)

    # ====================================================================== #
    # Executive Dashboard                                                    #
    # ====================================================================== #
    dash = wb.create_sheet("Executive Dashboard")
    dash.sheet_properties.tabColor = TAB_HERO
    title_block(
        dash,
        "Executive dashboard",
        f"Luxury Retail CRM & Clienteling Analytics  |  {config.PERIOD_START} to "
        f"{config.PERIOD_END}  |  all figures exclude returned lines.",
    )
    dash["A4"] = "Headline KPIs"
    dash["A4"].font = SECTION_FONT
    header_row(dash, 5, ["KPI", "Value", "Definition"])

    kpi_rows = [
        ("Net Sales", f"=SUM({T})", MONEY,
         "Line value after discount, excluding returns", False),
        ("Transactions (orders)", f"=SUM({Y})", NUM0,
         "Distinct orders with at least one sold line", False),
        ("Average Transaction Value", "=IFERROR($B$6/$B$7,0)", MONEY,
         "Net Sales / orders - per order, not per line", False),
        ("Units sold", f"=SUM({Z})", NUM0, "Units on sold lines", False),
        ("Average Unit Retail", "=IFERROR($B$6/$B$9,0)", MONEY,
         "Net Sales / units - realised price per item", False),
        ("Purchasing clients", f"=COUNT({FRQ})", NUM0,
         "Clients with a purchase history", False),
        ("Repeat Purchase Rate", f'=IFERROR(COUNTIF({FRQ},">=2")/COUNT({FRQ}),0)', PCT1,
         "Share of purchasing clients with 2+ orders", False),
        ("Full-Price Sales Share", f"=IFERROR(SUM({X})/$B$6,0)", PCT1,
         "Share of Net Sales at zero discount, by value", False),
        ("Markdown depth", f"=IFERROR(1-$B$6/SUM({AA}),0)", PCT1,
         "1 - Net Sales / Gross Sales at list price", False),
        ("Return rate by value", f"=IFERROR(SUM({AB})/(SUM({T})+SUM({AB})),0)", PCT1,
         "Returned value / all transacted value", True),
        ("Gross margin (realised)", "=GrossMargin", PCT1,
         "After markdown; drives the CLV model", False),
        ("Total 12-month CLV", f"=SUM({CLV})", MONEY,
         "Projected gross margin over the next 12 months", False),
        ("Lapsing high-value clients", f'=COUNTIF({SEG},"Lapsing High-Value")', NUM0,
         "High spenders with no purchase in LapseDays", True),
        ("Revenue held by lapsing clients",
         f'=SUMIF({SEG},"Lapsing High-Value",{MON})', MONEY,
         "Historic net sales of the lapsing cohort", True),
        ("Top decile revenue share", f"=IFERROR(SUMIF({DEC},1,{MON})/SUM({MON}),0)", PCT1,
         "Share of revenue from the top 10% of clients", False),
        ("Campaign spend", f"=SUM({CP_SPEND})", MONEY, "Across all campaigns", False),
        ("Blended campaign ROAS", f"=IFERROR(SUM({CP_REV})/SUM({CP_SPEND}),0)", MULT,
         "Attributed net revenue / spend", False),
        ("Appointment conversion rate",
         f'=IFERROR(SUMIF({CL_TYPE},"Appointment",{CL_CONV})/'
         f'COUNTIF({CL_TYPE},"Appointment"),0)', PCT1,
         "Appointments followed by a purchase in the window", False),
        ("Clienteling attributed revenue", f"=SUM({CL_REV})", MONEY,
         "Last-touch revenue credited to clienteling", False),
    ]
    for i, (label, formula, fmt, note, is_risk) in enumerate(kpi_rows):
        r = 6 + i
        dash[f"A{r}"] = label
        dash[f"A{r}"].font = KPI_LABEL_FONT
        dash[f"B{r}"] = formula
        dash[f"B{r}"].font = Font(name=FONT, size=14, bold=True,
                                  color=CLARET if is_risk else INK)
        dash[f"B{r}"].number_format = fmt
        dash[f"B{r}"].border = BOX
        dash[f"C{r}"] = note
        dash[f"C{r}"].font = SUB_FONT
        if is_risk:
            for col in "ABC":
                dash[f"{col}{r}"].fill = RISK_FILL
        elif i % 2 == 1:
            for col in "ABC":
                dash[f"{col}{r}"].fill = BAND_FILL

    trend_first = 6
    dash["E4"] = "Monthly trend"
    dash["E4"].font = SECTION_FONT
    header_row(dash, 5, ["month", "net sales", "orders", "ATV"], start_col=5)
    for i, month in enumerate(months):
        r = trend_first + i
        cell = dash.cell(row=r, column=5, value=month.to_pydatetime())
        cell.number_format = "mmm-yy"
        cell.font = BODY_FONT
        dash.cell(row=r, column=6, value=f"=SUMIFS({T},{U},$E{r})").number_format = MONEY
        dash.cell(row=r, column=7, value=f"=SUMIFS({Y},{U},$E{r})").number_format = NUM0
        dash.cell(row=r, column=8, value=f"=IFERROR($F{r}/$G{r},0)").number_format = MONEY
        for j in range(6, 9):
            dash.cell(row=r, column=j).font = BODY_FONT
    trend_last = trend_first + len(months) - 1
    dash.conditional_formatting.add(
        f"F{trend_first}:F{trend_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="8C6F3F"),
    )

    bar = BarChart()
    bar.type = "col"
    bar.title = "Monthly net sales and average transaction value"
    bar.y_axis.title = "Net Sales"
    bar.height, bar.width = 9, 26
    bar.add_data(Reference(dash, min_col=6, min_row=5, max_row=trend_last),
                 titles_from_data=True)
    bar.set_categories(Reference(dash, min_col=5, min_row=trend_first, max_row=trend_last))
    line = LineChart()
    line.add_data(Reference(dash, min_col=8, min_row=5, max_row=trend_last),
                  titles_from_data=True)
    line.y_axis.axId = 200
    line.y_axis.title = "ATV"
    line.y_axis.crosses = "max"
    bar += line
    dash.add_chart(bar, f"E{trend_last + 3}")

    seg_chart = BarChart()
    seg_chart.type = "col"
    seg_chart.title = "Net sales by RFM segment"
    seg_chart.height, seg_chart.width = 8, 13
    seg_chart.add_data(Reference(rfm, min_col=24, min_row=sr, max_row=seg_last),
                       titles_from_data=True)
    seg_chart.set_categories(Reference(rfm, min_col=21, min_row=seg_first, max_row=seg_last))
    seg_chart.legend = None
    dash.add_chart(seg_chart, f"E{trend_last + 22}")

    pareto = BarChart()
    pareto.type = "col"
    pareto.title = "Revenue concentration by client decile"
    pareto.height, pareto.width = 8, 13
    pareto.add_data(Reference(rfm, min_col=32, min_row=sr, max_row=dec_last),
                    titles_from_data=True)
    pareto.set_categories(Reference(rfm, min_col=30, min_row=dec_first, max_row=dec_last))
    cum = LineChart()
    cum.add_data(Reference(rfm, min_col=34, min_row=sr, max_row=dec_last),
                 titles_from_data=True)
    cum.y_axis.axId = 300
    cum.y_axis.crosses = "max"
    pareto += cum
    dash.add_chart(pareto, f"M{trend_last + 22}")

    dash.freeze_panes = "A6"
    set_widths(dash, [30, 18, 50])
    for letter in "EFGH":
        dash.column_dimensions[letter].width = 14

    # ====================================================================== #
    # Read Me                                                                #
    # ====================================================================== #
    readme = wb.create_sheet("Read Me")
    readme.sheet_properties.tabColor = TAB_HERO
    title_block(
        readme,
        "Luxury Retail CRM & Clienteling Analytics",
        "Analytical workbook - every figure on the analysis tabs is a live formula.",
    )
    notes = [
        ("How to use this workbook", True),
        ("Open the Config tab and change any blue cell. The whole workbook "
         "recalculates: segment thresholds, the clienteling attribution window, the "
         "CLV discount rate. Nothing is a pasted value.", False),
        ("", False),
        ("Analysis tabs", True),
        ("Executive Dashboard - 18 KPIs, monthly trend, segment and decile charts.", False),
        ("RFM & CLV - per-client scores, segment economics and the decile "
         "concentration table. Lapsing High-Value rows are shaded claret.", False),
        ("Cohort Retention - acquisition-cohort heatmap.", False),
        ("Boutique Performance - net sales against pro-rated target, with a traffic "
         "light on achievement.", False),
        ("Category & Product - category economics and the top 15 SKUs.", False),
        ("Monthly Mix - full price versus markdown, and channel mix, by month.", False),
        ("Clienteling ROI - conversion by type, by boutique and by associate.", False),
        ("Campaign Effectiveness - ROAS by channel and by audience.", False),
        ("Order Attribution - order-level last-touch resolution.", False),
        ("Config - every assumption, in one place.", False),
        ("", False),
        ("The one rule that matters", True),
        ("Net Sales excludes returned lines. The helper columns on Transactions "
         "(T to AF) encode this once: every monetary helper is zero on a returned "
         "line, so no analytical formula has to remember the rule and none of them "
         "can get it wrong.", False),
        ("", False),
        ("Conversion is not attribution", True),
        ("An interaction CONVERTS if any qualifying purchase follows it. Revenue is "
         "ATTRIBUTED to exactly one interaction per order, last touch, resolved on the "
         "Order Attribution tab. Measuring conversion with the attribution rule "
         "understates it - an appointment that worked loses its conversion if a "
         "courtesy call happens to follow it.", False),
        ("", False),
        ("Why distinct orders are counted the way they are", True),
        ("Counting distinct orders is the expensive operation in any transactional "
         "workbook. Because the lines of an order are contiguous, column V keeps a "
         "running count of sold lines within each order and column Y marks the first "
         "one. Summing column Y counts orders - no array formula, no slow SUMPRODUCT "
         "across 15,000 rows.", False),
        ("", False),
        ("Known limitations", True),
        ("Cohort Retention is the one analytical tab carrying imported values rather "
         "than live formulas. Its source is cited on the tab.", False),
        ("Quintile scores are rank-based and reproduce the Python pipeline to within a "
         "couple of clients at the exact quintile boundaries, where Excel integer "
         "division and pandas qcut round differently.", False),
        ("Attributed revenue can be double-counted for the 36 interactions that share "
         "a client and a date with another interaction - roughly 1% of clienteling "
         "revenue. Python breaks those ties deterministically; Excel cannot.", False),
        ("BaseRetention on the Config tab is the one figure that cannot be derived in "
         "Excel - it needs the acquisition-cohort logic in src/cohort_analysis.py. It "
         "is marked as an input and its source is cited.", False),
        ("The top-15 SKU list is selected at build time; its figures are live.", False),
        ("", False),
        ("All data is synthetic, generated from seed 42. No real client data is used.",
         False),
    ]
    r = 4
    for text, is_heading in notes:
        cell = readme.cell(row=r, column=1, value=text)
        cell.font = SECTION_FONT if is_heading else BODY_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        readme.row_dimensions[r].height = 30 if len(text) > 95 else 15
        r += 1
    readme.column_dimensions["A"].width = 110

    order = [
        "Read Me", "Executive Dashboard", "RFM & CLV", "Cohort Retention",
        "Boutique Performance", "Category & Product", "Monthly Mix",
        "Clienteling ROI", "Campaign Effectiveness", "Config", "Order Attribution",
        "Transactions", "Customer Master", "Product Master", "Boutique Master",
        "Associates", "Clienteling Activities", "Marketing Campaigns",
    ]
    wb._sheets = [wb[name] for name in order]

    wb.save(OUTPUT_PATH)
    print(f"Workbook written to {OUTPUT_PATH.relative_to(config.REPO_ROOT)}")
    print(f"  transaction rows : {tx_last - 1:,}")
    print(f"  clients scored   : {len(clients):,}")
    print(f"  orders attributed: {len(orders):,}")
    print(f"  sheets           : {len(wb.sheetnames)}")


if __name__ == "__main__":
    main()
