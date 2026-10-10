"""Phase 3 - Luxury Retail CRM & Clienteling dashboard.

Four sections, reachable from the sidebar:

1. Executive Performance - headline KPIs and the monthly revenue trend.
2. Customer Intelligence - RFM segments, the retention cohort heatmap and the
   lapsed high-value outreach list.
3. Clienteling ROI       - appointment conversion and revenue by associate.
4. Campaign Effectiveness- ROAS by channel and full-price buying behaviour.

Date, boutique and category filters apply across every section. Segments,
cohorts and CLV are recomputed against the filtered transactions rather than
read from a pre-aggregated file, so the numbers always reflect the current
selection.

Launch with:

    streamlit run dashboard/streamlit_app.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Make `src` importable when Streamlit runs this file directly.
REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from src import (  # noqa: E402  (import after sys.path setup)
    cohort_analysis,
    config,
    customer_lifetime_value,
    customer_segmentation,
    kpis,
)

# --------------------------------------------------------------------------- #
# Visual identity                                                              #
# --------------------------------------------------------------------------- #

# A restrained palette drawn from the category: ink and bone rather than black
# and white, brass as the single accent, claret reserved for commercial risk.
INK = "#16150F"
BONE = "#F7F5F2"
SURFACE = "#FFFFFF"
BRASS = "#8C6F3F"
CLARET = "#7A2E35"
SAGE = "#4F6650"
STONE = "#B9B2A6"

SEGMENT_COLOURS = {
    "High-Value Loyal": BRASS,
    "Lapsing High-Value": CLARET,
    "Promising New": SAGE,
    "Infrequent": STONE,
}

CHART_FONT = "Inter, 'Helvetica Neue', Arial, sans-serif"


def _style_figure(fig: go.Figure, height: int = 380) -> go.Figure:
    """Apply the project's chart styling to a Plotly figure."""
    fig.update_layout(
        height=height,
        margin=dict(l=8, r=8, t=34, b=8),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=CHART_FONT, size=13, color=INK),
        title_font=dict(size=15, color=INK),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0
        ),
        hoverlabel=dict(font_family=CHART_FONT, bgcolor=SURFACE, font_color=INK),
    )
    fig.update_xaxes(showgrid=False, linecolor=STONE, ticks="outside", tickcolor=STONE)
    fig.update_yaxes(gridcolor="#E8E4DD", zeroline=False)
    return fig


def _inject_css() -> None:
    """Replace Streamlit's default chrome with the project's own type and colour."""
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;600&family=Inter:wght@400;500;600&display=swap');

        .stApp {{ background: {BONE}; }}
        html, body, [class*="css"] {{
            font-family: {CHART_FONT};
            color: {INK};
        }}
        h1, h2, h3 {{
            font-family: 'Cormorant Garamond', Georgia, 'Times New Roman', serif;
            font-weight: 600;
            letter-spacing: 0.01em;
            color: {INK};
        }}
        h1 {{ font-size: 2.6rem; line-height: 1.1; margin-bottom: 0.1rem; }}
        h2 {{ font-size: 1.7rem; margin-top: 1.6rem; }}

        .lede {{
            color: #5D584E;
            font-size: 0.95rem;
            max-width: 68ch;
            margin-bottom: 1.4rem;
        }}

        /* KPI tiles: the one place the design is allowed to be loud. */
        .kpi-row {{ display: flex; gap: 14px; flex-wrap: wrap; margin: 0.5rem 0 1.6rem; }}
        .kpi {{
            flex: 1 1 170px;
            background: {SURFACE};
            border-top: 2px solid {BRASS};
            padding: 16px 18px 14px;
        }}
        .kpi .value {{
            font-family: 'Cormorant Garamond', Georgia, serif;
            font-size: 2.1rem;
            line-height: 1.05;
            color: {INK};
        }}
        .kpi .label {{
            font-size: 0.78rem;
            color: #6B655A;
            margin-top: 6px;
        }}
        .kpi.risk {{ border-top-color: {CLARET}; }}

        section[data-testid="stSidebar"] {{ background: {SURFACE}; }}
        div[data-testid="stMetricValue"] {{ font-size: 1.5rem; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def _kpi_tiles(tiles: list[tuple[str, str, bool]]) -> None:
    """Render a row of KPI tiles.

    Args:
        tiles: ``(value, label, is_risk)`` triples. ``is_risk`` switches the
            accent rule to claret for figures that represent exposure.
    """
    cards = "".join(
        f'<div class="kpi{" risk" if risk else ""}">'
        f'<div class="value">{value}</div>'
        f'<div class="label">{label}</div></div>'
        for value, label, risk in tiles
    )
    st.markdown(f'<div class="kpi-row">{cards}</div>', unsafe_allow_html=True)


def money(value: float, decimals: int = 0) -> str:
    """Format a number as reporting-currency text."""
    return f"{config.CURRENCY_SYMBOL}{value:,.{decimals}f}"


# --------------------------------------------------------------------------- #
# Data loading                                                                 #
# --------------------------------------------------------------------------- #


@st.cache_data(show_spinner=False)
def load_data() -> dict[str, pd.DataFrame]:
    """Load every generated table, with product and boutique attributes joined on."""
    tables = kpis.load_all_tables()

    transactions = tables["transactions"].merge(
        tables["products"][["sku", "category", "subcategory", "collection", "season"]],
        on="sku",
        how="left",
    )
    transactions = transactions.merge(
        tables["boutiques"][["boutique_id", "boutique_name", "region", "boutique_format"]],
        on="boutique_id",
        how="left",
    )
    tables["transactions"] = transactions
    return tables


@st.cache_data(show_spinner=False)
def compute_segments(transactions: pd.DataFrame, customers: pd.DataFrame) -> pd.DataFrame:
    """Recompute RFM segments against the current filter selection."""
    return customer_segmentation.build_segmentation(transactions, customers)


@st.cache_data(show_spinner=False)
def compute_cohorts(
    transactions: pd.DataFrame, customers: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Recompute the retention cohort matrices against the current selection."""
    return cohort_analysis.cohort_retention_matrix(transactions, customers)


@st.cache_data(show_spinner=False)
def compute_clv(transactions: pd.DataFrame, customers: pd.DataFrame) -> pd.DataFrame:
    """Recompute 12-month CLV against the current selection."""
    return customer_lifetime_value.build_clv(transactions, customers)


# --------------------------------------------------------------------------- #
# Sections                                                                     #
# --------------------------------------------------------------------------- #


def page_executive(transactions: pd.DataFrame) -> None:
    """Section 1 - headline trading performance."""
    st.markdown("# Executive performance")
    st.markdown(
        '<p class="lede">Trading health for the selected period. Net Sales '
        "excludes returned lines throughout, so every figure below is the "
        "revenue the business actually kept.</p>",
        unsafe_allow_html=True,
    )

    if transactions.empty:
        st.info("No transactions match the current filters. Widen the selection to continue.")
        return

    _kpi_tiles(
        [
            (money(kpis.net_sales(transactions)), "Net Sales", False),
            (money(kpis.average_transaction_value(transactions)), "Average transaction value", False),
            (money(kpis.average_unit_retail(transactions)), "Average unit retail", False),
            (f"{kpis.repeat_purchase_rate(transactions):.1%}", "Repeat purchase rate", False),
            (f"{kpis.full_price_sales_share(transactions):.1%}", "Full-price share of sales", False),
            (f"{kpis.return_rate_by_value(transactions):.1%}", "Returns by value", True),
        ]
    )

    monthly = kpis.monthly_revenue(transactions)
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=monthly["month"],
            y=monthly["net_sales"],
            name="Net Sales",
            marker_color=BRASS,
            opacity=0.85,
            hovertemplate="%{x|%b %Y}<br>Net Sales " + config.CURRENCY_SYMBOL + "%{y:,.0f}<extra></extra>",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=monthly["month"],
            y=monthly["atv"],
            name="Average transaction value",
            yaxis="y2",
            mode="lines+markers",
            line=dict(color=INK, width=2),
            marker=dict(size=5),
            hovertemplate="%{x|%b %Y}<br>ATV " + config.CURRENCY_SYMBOL + "%{y:,.0f}<extra></extra>",
        )
    )
    fig.update_layout(
        title="Monthly net sales and average transaction value",
        yaxis=dict(title="Net Sales"),
        yaxis2=dict(title="ATV", overlaying="y", side="right", showgrid=False),
    )
    st.plotly_chart(_style_figure(fig, 420), width="stretch")

    left, right = st.columns(2)

    by_boutique = (
        kpis.valid_sales(transactions)
        .groupby("boutique_name", as_index=False)["net_amount"]
        .sum()
        .sort_values("net_amount")
    )
    fig_b = px.bar(
        by_boutique,
        x="net_amount",
        y="boutique_name",
        orientation="h",
        title="Net sales by boutique",
        labels={"net_amount": "Net Sales", "boutique_name": ""},
    )
    fig_b.update_traces(marker_color=BRASS, hovertemplate="%{y}<br>" + config.CURRENCY_SYMBOL + "%{x:,.0f}<extra></extra>")
    left.plotly_chart(_style_figure(fig_b, 400), width="stretch")

    by_category = (
        kpis.valid_sales(transactions)
        .groupby("category", as_index=False)
        .agg(net_amount=("net_amount", "sum"), units=("quantity", "sum"))
        .sort_values("net_amount")
    )
    fig_c = px.bar(
        by_category,
        x="net_amount",
        y="category",
        orientation="h",
        title="Net sales by category",
        labels={"net_amount": "Net Sales", "category": ""},
    )
    fig_c.update_traces(marker_color=INK, hovertemplate="%{y}<br>" + config.CURRENCY_SYMBOL + "%{x:,.0f}<extra></extra>")
    right.plotly_chart(_style_figure(fig_c, 400), width="stretch")


def page_customer_intelligence(
    transactions: pd.DataFrame, customers: pd.DataFrame
) -> None:
    """Section 2 - segments, retention cohorts and the outreach list."""
    st.markdown("# Customer intelligence")
    st.markdown(
        '<p class="lede">Who the clients are, whether they come back, and which '
        "high-value relationships are going quiet.</p>",
        unsafe_allow_html=True,
    )

    if transactions.empty:
        st.info("No transactions match the current filters. Widen the selection to continue.")
        return

    segmented = compute_segments(transactions, customers)
    summary = customer_segmentation.segment_summary(segmented)
    clv = compute_clv(transactions, customers)

    lapsing = summary.loc[summary["segment"] == "Lapsing High-Value"]
    at_risk_revenue = float(lapsing["net_sales"].iloc[0]) if not lapsing.empty else 0.0
    at_risk_clients = int(lapsing["clients"].iloc[0]) if not lapsing.empty else 0

    _kpi_tiles(
        [
            (f"{len(segmented):,}", "Clients with a purchase history", False),
            (money(clv["clv_12m"].sum()), "Projected 12-month CLV (margin)", False),
            (f"{at_risk_clients:,}", "Lapsing high-value clients", True),
            (money(at_risk_revenue), "Revenue held by lapsing clients", True),
        ]
    )

    left, right = st.columns([1, 1])

    plot_summary = summary.copy()
    plot_summary["segment"] = plot_summary["segment"].astype(str)
    fig_seg = px.bar(
        plot_summary,
        x="segment",
        y="clients",
        title="Clients by segment",
        labels={"clients": "Clients", "segment": ""},
        color="segment",
        color_discrete_map=SEGMENT_COLOURS,
        custom_data=["revenue_share", "net_sales"],
    )
    fig_seg.update_traces(
        showlegend=False,
        hovertemplate=(
            "%{x}<br>%{y:,} clients<br>"
            "%{customdata[0]:.1%} of revenue<extra></extra>"
        ),
    )
    left.plotly_chart(_style_figure(fig_seg), width="stretch")

    fig_rev = px.bar(
        plot_summary,
        x="segment",
        y="revenue_share",
        title="Share of net sales by segment",
        labels={"revenue_share": "Share of Net Sales", "segment": ""},
        color="segment",
        color_discrete_map=SEGMENT_COLOURS,
    )
    fig_rev.update_traces(showlegend=False, hovertemplate="%{x}<br>%{y:.1%}<extra></extra>")
    fig_rev.update_yaxes(tickformat=".0%")
    right.plotly_chart(_style_figure(fig_rev), width="stretch")

    # ---- Cohort heatmap ------------------------------------------------- #
    st.markdown("## Retention cohorts")
    st.markdown(
        '<p class="lede">Each row is a month of newly acquired clients; each '
        "column is the share still purchasing that many months later. Clients "
        "who joined before the analysis window are excluded, so established "
        "buyers cannot inflate the early cohorts.</p>",
        unsafe_allow_html=True,
    )

    _, rates = compute_cohorts(transactions, customers)
    if rates.empty:
        st.info("Not enough newly acquired clients in this selection to build cohorts.")
    else:
        display = rates.drop(columns=[0], errors="ignore")
        display = display.loc[:, [c for c in display.columns if c <= 18]]
        fig_cohort = px.imshow(
            display.to_numpy() * 100,
            x=[f"M{int(c)}" for c in display.columns],
            y=[d.strftime("%b %Y") for d in display.index],
            color_continuous_scale=["#FFFFFF", "#E3D7BE", BRASS, "#5C4727"],
            aspect="auto",
            labels=dict(color="Retention %"),
            title="Monthly retention by acquisition cohort (%)",
        )
        fig_cohort.update_traces(
            hovertemplate="Cohort %{y}<br>%{x}<br>%{z:.1f}% active<extra></extra>"
        )
        fig_cohort.update_xaxes(side="top")
        st.plotly_chart(_style_figure(fig_cohort, 520), width="stretch")

    # ---- Outreach list -------------------------------------------------- #
    st.markdown("## Lapsed high-value clients")
    st.markdown(
        '<p class="lede">High spenders with no purchase in the last '
        f"{config.LAPSED_RECENCY_DAYS} days, ranked by historic value. This is "
        "the call list, not a report: sort it, filter it, and hand it to the "
        "boutique.</p>",
        unsafe_allow_html=True,
    )

    lapsed = segmented[segmented["segment"] == "Lapsing High-Value"].merge(
        clv[["customer_id", "clv_12m"]], on="customer_id", how="left"
    )
    contactable_only = st.checkbox(
        "Show only clients who have opted in to contact", value=True
    )
    if contactable_only:
        lapsed = lapsed[lapsed["consent_status"] == "Opted In"]

    columns = [
        "customer_id",
        "client_name",
        "preferred_boutique_id",
        "consent_status",
        "recency_days",
        "frequency",
        "monetary",
        "avg_order_value",
        "clv_12m",
    ]
    lapsed = lapsed[columns].sort_values("monetary", ascending=False)

    st.dataframe(
        lapsed,
        width="stretch",
        hide_index=True,
        column_config={
            "customer_id": st.column_config.TextColumn("Client ID"),
            "client_name": st.column_config.TextColumn("Client"),
            "preferred_boutique_id": st.column_config.TextColumn("Boutique"),
            "consent_status": st.column_config.TextColumn("Consent"),
            "recency_days": st.column_config.NumberColumn("Days since purchase", format="%d"),
            "frequency": st.column_config.NumberColumn("Orders", format="%d"),
            "monetary": st.column_config.NumberColumn("Lifetime net sales", format="%.0f"),
            "avg_order_value": st.column_config.NumberColumn("Avg order value", format="%.0f"),
            "clv_12m": st.column_config.NumberColumn("Projected 12m CLV", format="%.0f"),
        },
    )
    st.download_button(
        "Download outreach list",
        data=lapsed.to_csv(index=False).encode("utf-8"),
        file_name="lapsed_high_value_clients.csv",
        mime="text/csv",
    )


def page_clienteling(
    transactions: pd.DataFrame,
    interactions: pd.DataFrame,
    associates: pd.DataFrame,
) -> None:
    """Section 3 - clienteling conversion and associate performance."""
    st.markdown("# Clienteling ROI")
    st.markdown(
        '<p class="lede">A purchase is credited to an interaction when the '
        f"client buys within {config.CLIENTELING_ATTRIBUTION_WINDOW_DAYS} days "
        "of it. Where several interactions qualify, the most recent one takes "
        "the credit.</p>",
        unsafe_allow_html=True,
    )

    if transactions.empty or interactions.empty:
        st.info("No clienteling activity matches the current filters.")
        return

    attributed = kpis.attribute_clienteling(transactions, interactions)

    by_type = (
        attributed.groupby("interaction_type")
        .agg(
            interactions=("interaction_id", "count"),
            conversions=("converted_in_window", "sum"),
            attributed_revenue=("attributed_revenue", "sum"),
        )
        .reset_index()
    )
    by_type["conversion_rate"] = by_type["conversions"] / by_type["interactions"]
    by_type["revenue_per_interaction"] = (
        by_type["attributed_revenue"] / by_type["interactions"]
    )

    appointment = by_type[by_type["interaction_type"] == "Appointment"]
    _kpi_tiles(
        [
            (f"{len(attributed):,}", "Logged interactions", False),
            (
                f"{float(appointment['conversion_rate'].iloc[0]):.1%}"
                if not appointment.empty
                else "-",
                "Appointment conversion rate",
                False,
            ),
            (
                money(float(attributed["attributed_revenue"].sum())),
                "Clienteling-attributed net sales",
                False,
            ),
            (
                money(
                    float(attributed["attributed_revenue"].sum()) / max(len(attributed), 1)
                ),
                "Revenue per interaction",
                False,
            ),
        ]
    )

    left, right = st.columns(2)

    fig_conv = px.bar(
        by_type.sort_values("conversion_rate"),
        x="conversion_rate",
        y="interaction_type",
        orientation="h",
        title="Conversion rate by interaction type",
        labels={"conversion_rate": "Converted within window", "interaction_type": ""},
    )
    fig_conv.update_traces(marker_color=BRASS, hovertemplate="%{y}<br>%{x:.1%}<extra></extra>")
    fig_conv.update_xaxes(tickformat=".0%")
    left.plotly_chart(_style_figure(fig_conv, 320), width="stretch")

    fig_rpi = px.bar(
        by_type.sort_values("revenue_per_interaction"),
        x="revenue_per_interaction",
        y="interaction_type",
        orientation="h",
        title="Attributed revenue per interaction",
        labels={"revenue_per_interaction": "Net Sales per interaction", "interaction_type": ""},
    )
    fig_rpi.update_traces(
        marker_color=INK,
        hovertemplate="%{y}<br>" + config.CURRENCY_SYMBOL + "%{x:,.0f}<extra></extra>",
    )
    right.plotly_chart(_style_figure(fig_rpi, 320), width="stretch")

    st.markdown("## Performance by sales associate")
    min_interactions = st.slider(
        "Minimum logged interactions", min_value=0, max_value=60, value=15, step=5
    )
    performance = kpis.associate_performance(attributed, associates)
    performance = performance[performance["interactions"] >= min_interactions]

    if performance.empty:
        st.info("No associates meet that threshold in the current selection.")
        return

    fig_assoc = px.scatter(
        performance,
        x="conversion_rate",
        y="revenue_per_interaction",
        size="interactions",
        color="boutique_id",
        hover_name="associate_name",
        title="Conversion rate against revenue per interaction",
        labels={
            "conversion_rate": "Conversion rate",
            "revenue_per_interaction": "Net Sales per interaction",
            "boutique_id": "Boutique",
        },
    )
    fig_assoc.update_xaxes(tickformat=".0%")
    st.plotly_chart(_style_figure(fig_assoc, 440), width="stretch")

    top = performance.head(15)[
        [
            "associate_name",
            "boutique_id",
            "role",
            "interactions",
            "appointments",
            "conversion_rate",
            "attributed_revenue",
        ]
    ]
    st.dataframe(
        top,
        width="stretch",
        hide_index=True,
        column_config={
            "associate_name": st.column_config.TextColumn("Associate"),
            "boutique_id": st.column_config.TextColumn("Boutique"),
            "role": st.column_config.TextColumn("Role"),
            "interactions": st.column_config.NumberColumn("Interactions", format="%d"),
            "appointments": st.column_config.NumberColumn("Appointments", format="%d"),
            "conversion_rate": st.column_config.NumberColumn("Conversion", format="%.1f%%"),
            "attributed_revenue": st.column_config.NumberColumn(
                "Attributed net sales", format="%.0f"
            ),
        },
    )


def page_campaigns(transactions: pd.DataFrame, campaigns: pd.DataFrame) -> None:
    """Section 4 - campaign return and full-price behaviour."""
    st.markdown("# Campaign effectiveness")
    st.markdown(
        '<p class="lede">Return on ad spend by channel, and whether marketing '
        "is pulling clients towards full-price product or towards markdown. "
        "Campaign figures cover the whole programme and are not affected by the "
        "boutique and category filters.</p>",
        unsafe_allow_html=True,
    )

    if campaigns.empty:
        st.info("No campaigns to report.")
        return

    _kpi_tiles(
        [
            (money(float(campaigns["spend"].sum())), "Campaign spend", False),
            (
                money(float(campaigns["attributed_net_revenue"].sum())),
                "Attributed net sales",
                False,
            ),
            (f"{kpis.campaign_roas(campaigns):.2f}x", "Blended ROAS", False),
            (f"{len(campaigns)}", "Campaigns in period", False),
        ]
    )

    by_channel = kpis.campaign_roas_by_channel(campaigns)
    fig_roas = px.bar(
        by_channel.sort_values("roas"),
        x="roas",
        y="channel",
        orientation="h",
        title="Return on ad spend by channel",
        labels={"roas": "ROAS (net sales per unit of spend)", "channel": ""},
        custom_data=["spend", "attributed_net_revenue", "campaigns"],
    )
    fig_roas.update_traces(
        marker_color=BRASS,
        hovertemplate=(
            "%{y}<br>ROAS %{x:.2f}x<br>"
            "Spend " + config.CURRENCY_SYMBOL + "%{customdata[0]:,.0f}<br>"
            "Attributed " + config.CURRENCY_SYMBOL + "%{customdata[1]:,.0f}"
            "<extra></extra>"
        ),
    )
    st.plotly_chart(_style_figure(fig_roas, 360), width="stretch")

    left, right = st.columns(2)

    fig_scatter = px.scatter(
        campaigns,
        x="spend",
        y="attributed_net_revenue",
        size="targeted_customers",
        color="channel",
        hover_name="campaign_name",
        title="Spend against attributed net sales",
        labels={"spend": "Spend", "attributed_net_revenue": "Attributed Net Sales"},
    )
    left.plotly_chart(_style_figure(fig_scatter, 400), width="stretch")

    sales = kpis.valid_sales(transactions)
    if sales.empty:
        right.info("No transactions match the current filters.")
    else:
        influenced = sales["campaign_id"].notna()
        rows = []
        for label, mask in (
            ("Campaign-influenced", influenced),
            ("Not campaign-influenced", ~influenced),
        ):
            subset = sales[mask]
            total = float(subset["net_amount"].sum())
            if total == 0:
                continue
            full = float(subset.loc[subset["is_full_price"], "net_amount"].sum())
            rows.append({"group": label, "basis": "Full price", "value": full / total})
            rows.append(
                {"group": label, "basis": "Discounted", "value": 1 - full / total}
            )
        mix = pd.DataFrame(rows)
        fig_mix = px.bar(
            mix,
            x="group",
            y="value",
            color="basis",
            title="Full-price versus discounted mix",
            labels={"value": "Share of Net Sales", "group": "", "basis": ""},
            color_discrete_map={"Full price": BRASS, "Discounted": STONE},
        )
        fig_mix.update_yaxes(tickformat=".0%")
        fig_mix.update_traces(hovertemplate="%{x}<br>%{fullData.name}: %{y:.1%}<extra></extra>")
        right.plotly_chart(_style_figure(fig_mix, 400), width="stretch")

    st.markdown("## Campaign detail")
    detail = campaigns[
        [
            "campaign_id",
            "campaign_name",
            "channel",
            "audience_segment",
            "start_date",
            "spend",
            "targeted_customers",
            "attributed_orders",
            "attributed_net_revenue",
            "roas",
        ]
    ].sort_values("roas", ascending=False)
    st.dataframe(
        detail,
        width="stretch",
        hide_index=True,
        column_config={
            "campaign_id": st.column_config.TextColumn("ID"),
            "campaign_name": st.column_config.TextColumn("Campaign"),
            "channel": st.column_config.TextColumn("Channel"),
            "audience_segment": st.column_config.TextColumn("Audience"),
            "start_date": st.column_config.DateColumn("Start", format="DD MMM YYYY"),
            "spend": st.column_config.NumberColumn("Spend", format="%.0f"),
            "targeted_customers": st.column_config.NumberColumn("Reached", format="%d"),
            "attributed_orders": st.column_config.NumberColumn("Orders", format="%d"),
            "attributed_net_revenue": st.column_config.NumberColumn(
                "Attributed net sales", format="%.0f"
            ),
            "roas": st.column_config.NumberColumn("ROAS", format="%.2fx"),
        },
    )


# --------------------------------------------------------------------------- #
# Application                                                                  #
# --------------------------------------------------------------------------- #


def main() -> None:
    """Build the sidebar, apply filters and render the selected section."""
    st.set_page_config(
        page_title="Luxury Retail CRM & Clienteling Analytics",
        page_icon="◈",
        layout="wide",
    )
    _inject_css()

    try:
        tables = load_data()
    except FileNotFoundError:
        st.error(
            "The dataset has not been generated yet. Run `python -m src.run_pipeline` "
            "from the repository root, then reload this page."
        )
        st.stop()

    transactions = tables["transactions"]
    customers = tables["customers"]
    interactions = tables["clienteling_interactions"]
    campaigns = tables["campaigns"]
    associates = tables["associates"]
    boutiques = tables["boutiques"]

    # ---- Sidebar ------------------------------------------------------- #
    st.sidebar.markdown("### Luxury Retail CRM")
    st.sidebar.caption(
        f"Clienteling analytics · snapshot {config.SNAPSHOT_DATE:%d %b %Y}"
    )

    section = st.sidebar.radio(
        "Section",
        [
            "Executive performance",
            "Customer intelligence",
            "Clienteling ROI",
            "Campaign effectiveness",
        ],
    )

    st.sidebar.markdown("---")
    st.sidebar.markdown("#### Filters")

    min_date = transactions["transaction_date"].min().date()
    max_date = transactions["transaction_date"].max().date()
    date_range = st.sidebar.date_input(
        "Date range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )

    boutique_names = sorted(boutiques["boutique_name"].unique())
    selected_boutiques = st.sidebar.multiselect(
        "Boutique", boutique_names, default=boutique_names
    )

    categories = sorted(transactions["category"].dropna().unique())
    selected_categories = st.sidebar.multiselect(
        "Product category", categories, default=categories
    )

    # ---- Apply filters -------------------------------------------------- #
    filtered = transactions
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start, end = date_range
        filtered = filtered[
            (filtered["transaction_date"] >= pd.Timestamp(start))
            & (filtered["transaction_date"] <= pd.Timestamp(end))
        ]
    if selected_boutiques:
        filtered = filtered[filtered["boutique_name"].isin(selected_boutiques)]
    if selected_categories:
        filtered = filtered[filtered["category"].isin(selected_categories)]

    # Clienteling activity is filtered to the boutiques and dates in scope so
    # that conversion rates are measured against the same population as sales.
    selected_boutique_ids = boutiques.loc[
        boutiques["boutique_name"].isin(selected_boutiques), "boutique_id"
    ]
    filtered_interactions = interactions[
        interactions["boutique_id"].isin(set(selected_boutique_ids))
    ]
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start, end = date_range
        filtered_interactions = filtered_interactions[
            (filtered_interactions["interaction_date"] >= pd.Timestamp(start))
            & (filtered_interactions["interaction_date"] <= pd.Timestamp(end))
        ]

    st.sidebar.markdown("---")
    st.sidebar.caption(
        f"{len(filtered):,} of {len(transactions):,} transaction lines in scope"
    )

    # ---- Render --------------------------------------------------------- #
    if section == "Executive performance":
        page_executive(filtered)
    elif section == "Customer intelligence":
        page_customer_intelligence(filtered, customers)
    elif section == "Clienteling ROI":
        page_clienteling(filtered, filtered_interactions, associates)
    else:
        page_campaigns(filtered, campaigns)


if __name__ == "__main__":
    main()
