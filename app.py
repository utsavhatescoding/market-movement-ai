import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

st.set_page_config(
    page_title="Market Movement AI Dashboard",
    page_icon="📊",
    layout="wide"
)

# -----------------------------
# Design / CSS
# -----------------------------

st.markdown("""
<style>
    .block-container {
        padding-top: 1.4rem;
        max-width: 1280px;
    }

    .hero {
        background: linear-gradient(135deg, #102A43 0%, #243B53 55%, #B7791F 130%);
        padding: 34px 36px;
        border-radius: 26px;
        color: white;
        margin-bottom: 22px;
        box-shadow: 0 18px 45px rgba(16, 42, 67, 0.18);
    }

    .hero-title {
        font-size: 42px;
        font-weight: 800;
        letter-spacing: -1.2px;
        margin-bottom: 8px;
    }

    .hero-subtitle {
        font-size: 17px;
        opacity: 0.88;
        max-width: 820px;
        line-height: 1.55;
    }

    .pill {
        display: inline-block;
        padding: 7px 12px;
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.14);
        border: 1px solid rgba(255, 255, 255, 0.20);
        font-size: 12px;
        letter-spacing: 0.6px;
        text-transform: uppercase;
        margin-bottom: 14px;
    }

    .kpi-card {
        background: #FFFFFF;
        border: 1px solid rgba(16, 42, 67, 0.08);
        border-radius: 22px;
        padding: 20px 22px;
        box-shadow: 0 12px 28px rgba(16, 42, 67, 0.08);
        min-height: 132px;
    }

    .kpi-label {
        font-size: 13px;
        color: #6B7280;
        text-transform: uppercase;
        letter-spacing: 0.7px;
        font-weight: 700;
        margin-bottom: 10px;
    }

    .kpi-value {
        font-size: 31px;
        color: #102A43;
        font-weight: 800;
        letter-spacing: -0.8px;
        margin-bottom: 6px;
    }

    .kpi-note {
        font-size: 13px;
        color: #7C6F64;
    }

    .section-card {
        background: #FFFFFF;
        border-radius: 24px;
        padding: 22px;
        border: 1px solid rgba(16, 42, 67, 0.08);
        box-shadow: 0 12px 28px rgba(16, 42, 67, 0.06);
        margin-bottom: 16px;
        line-height: 1.55;
    }

    .insight-card {
        background: #FFF9EC;
        border-left: 5px solid #B7791F;
        border-radius: 18px;
        padding: 18px 20px;
        margin-bottom: 14px;
        color: #102A43;
        line-height: 1.55;
    }

    .risk-card {
        background: #FFF1F2;
        border-left: 5px solid #BE123C;
        border-radius: 18px;
        padding: 18px 20px;
        margin-bottom: 14px;
        color: #102A43;
        line-height: 1.55;
    }

    .opportunity-card {
        background: #ECFDF5;
        border-left: 5px solid #059669;
        border-radius: 18px;
        padding: 18px 20px;
        margin-bottom: 14px;
        color: #102A43;
        line-height: 1.55;
    }

    h1, h2, h3 {
        color: #102A43;
        letter-spacing: -0.4px;
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Sidebar
# -----------------------------

st.sidebar.title("Controls")
st.sidebar.caption("Upload a Department of Customs Excel file or use the default file in your data folder.")

uploaded_file = st.sidebar.file_uploader(
    "Upload customs Excel file",
    type=["xlsx"]
)

default_file_path = Path(__file__).parent / "customs.xlsx"

if uploaded_file is not None:
    file_source = uploaded_file
elif default_file_path.exists():
    file_source = default_file_path
else:
    st.warning("Please upload a customs Excel file from the sidebar to start the dashboard.")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.info("Source values are in Rs. thousands. Dashboard values are shown in Rs. billion.")
st.sidebar.markdown("**Version:** Professional MVP 0.3")

# -----------------------------
# Load Excel
# -----------------------------

def load_data(file):
    excel = pd.ExcelFile(file)

    trade = pd.read_excel(excel, sheet_name="1_Trade_Direction", header=2)
    countries = pd.read_excel(excel, sheet_name="3_Trade_Balance_Country", header=2)
    import_partner = pd.read_excel(excel, sheet_name="4_Imports_By_Commodity_Partner", header=2)
    imports = pd.read_excel(excel, sheet_name="5_Imports_By_Commodity", header=2)
    export_partner = pd.read_excel(excel, sheet_name="6_Exports_By_Commodity_Partner", header=2)
    exports = pd.read_excel(excel, sheet_name="7_Exports_By_Commodity", header=2)
    customs = pd.read_excel(excel, sheet_name="9_Customswise_Trade", header=2)

    return trade, countries, import_partner, imports, export_partner, exports, customs

try:
    trade, countries, import_partner, imports, export_partner, exports, customs = load_data(file_source)
except Exception as e:
    st.error("Could not read the Excel file. Make sure it is the Department of Customs workbook and the sheet names are unchanged.")
    st.exception(e)
    st.stop()

# -----------------------------
# Helper functions
# -----------------------------

def rs_thousand_to_billion(value):
    return value / 1_000_000


def safe_number(value):
    try:
        return float(value)
    except Exception:
        return 0.0


def get_trade_value(pattern, value_col):
    mask = trade["Trade.Indicators"].astype(str).str.contains(
        pattern,
        case=False,
        na=False,
        regex=True
    )

    if mask.sum() == 0:
        return 0.0

    return safe_number(trade.loc[mask, value_col].iloc[0])


def short_text(text, max_len=58):
    text = str(text)
    if len(text) <= max_len:
        return text
    return text[:max_len] + "..."


def kpi_card(label, value, note):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def horizontal_bar(df, x, y, title, x_label, y_label, height=470):
    chart_df = df.copy()
    chart_df["Label"] = chart_df[y].apply(short_text)

    fig = px.bar(
        chart_df.sort_values(x),
        x=x,
        y="Label",
        orientation="h",
        title=title,
        labels={x: x_label, "Label": y_label},
        text=x
    )

    fig.update_traces(
        texttemplate="%{x:,.1f}",
        textposition="outside",
        marker_line_width=0
    )

    fig.update_layout(
        height=height,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        title_font_size=20,
        title_font_color="#102A43",
        font=dict(size=13, color="#102A43"),
        xaxis=dict(showgrid=True, gridcolor="rgba(16,42,67,0.08)"),
        yaxis=dict(title=None),
        margin=dict(l=8, r=30, t=60, b=30)
    )

    return fig


def clean_hscode(series):
    return series.astype(str).str.replace(".0", "", regex=False)
def remove_total_rows(df, columns):
    clean_df = df.copy()

    for col in columns:
        if col in clean_df.columns:
            clean_df = clean_df[
                ~clean_df[col]
                .astype(str)
                .str.strip()
                .str.lower()
                .isin(["total", "grand total"])
            ]

    return clean_df

# -----------------------------
# Prepare data
# -----------------------------

possible_year_cols = [
    col for col in trade.columns
    if str(col).startswith("FY") and "Change" not in str(col)
]

current_col = (
    "FY 2082/83 (First 11 Months)"
    if "FY 2082/83 (First 11 Months)" in trade.columns
    else possible_year_cols[-1]
)

imports_total_raw = get_trade_value(r"^Imports\s*\(", current_col)
exports_total_raw = get_trade_value(r"^Exports\s*\(", current_col)
deficit_total_raw = get_trade_value(r"Trade\s+Deficit", current_col)
total_trade_raw = get_trade_value(r"Total\s+Foreign\s+Trade", current_col)

imports_total = rs_thousand_to_billion(imports_total_raw)
exports_total = rs_thousand_to_billion(exports_total_raw)
deficit_total = rs_thousand_to_billion(deficit_total_raw)
total_trade = rs_thousand_to_billion(total_trade_raw)

import_export_ratio = imports_total_raw / exports_total_raw if exports_total_raw else 0

imports = imports.copy()
exports = exports.copy()
countries = countries.copy()
customs = customs.copy()
import_partner = import_partner.copy()
export_partner = export_partner.copy()

imports["HSCode"] = clean_hscode(imports["HSCode"])
exports["HSCode"] = clean_hscode(exports["HSCode"])
import_partner["HSCode"] = clean_hscode(import_partner["HSCode"])
export_partner["HSCode"] = clean_hscode(export_partner["HSCode"])


imports["Imports_Billion"] = imports["Imports_Value"] / 1_000_000
imports["Revenue_Billion"] = imports["Imports_Revenue"] / 1_000_000
imports["Approx_Duty_Rate"] = (imports["Imports_Revenue"] / imports["Imports_Value"]) * 100
imports["Approx_Duty_Rate"] = imports["Approx_Duty_Rate"].replace([float("inf"), -float("inf")], 0).fillna(0)

exports["Exports_Billion"] = exports["Exports_Value"] / 1_000_000

countries["Imports_Billion"] = countries["Imports_Value"] / 1_000_000
countries["Exports_Billion"] = countries["Exports_Value"] / 1_000_000
countries["Trade_Balance_Billion"] = countries["Trade_Balance"] / 1_000_000

customs["Imports_Billion"] = customs["Imports_Value"] / 1_000_000
customs["Exports_Billion"] = customs["Exports_Value"] / 1_000_000

import_partner["Imports_Billion"] = import_partner["Imports_Value"] / 1_000_000
import_partner["Revenue_Billion"] = import_partner["Imports_Revenue"] / 1_000_000

export_partner["Exports_Billion"] = export_partner["Exports_Value"] / 1_000_000
# Remove total rows so dashboard shows real products/countries/routes
imports = remove_total_rows(imports, ["Description"])
exports = remove_total_rows(exports, ["Description"])
countries = remove_total_rows(countries, ["Partner Countries"])
customs = remove_total_rows(customs, ["Customs"])

import_partner = remove_total_rows(import_partner, ["Description", "Partner Countries"])
export_partner = remove_total_rows(export_partner, ["Description", "Partner Countries"])

top_imports = imports.sort_values("Imports_Value", ascending=False).head(10)
top_exports = exports.sort_values("Exports_Value", ascending=False).head(10)
top_import_countries = countries.sort_values("Imports_Value", ascending=False).head(10)
top_export_countries = countries.sort_values("Exports_Value", ascending=False).head(10)
top_deficit_countries = countries.sort_values("Trade_Balance").head(10)
top_surplus_countries = countries.sort_values("Trade_Balance", ascending=False).head(10)
top_customs = customs.sort_values("Imports_Value", ascending=False).head(10)

top_import_product = top_imports.iloc[0]["Description"]
top_import_value = top_imports.iloc[0]["Imports_Billion"]

top_export_product = top_exports.iloc[0]["Description"]
top_export_value = top_exports.iloc[0]["Exports_Billion"]

top_import_country = top_import_countries.iloc[0]["Partner Countries"]
top_export_country = top_export_countries.iloc[0]["Partner Countries"]

top_customs_office = top_customs.iloc[0]["Customs"]
top_customs_value = top_customs.iloc[0]["Imports_Billion"]
top5_route_share = top_customs["Import_Share"].head(5).sum()

# -----------------------------
# Hero section
# -----------------------------

st.markdown(
    """
    <div class="hero">
        <div class="pill">Nepal Trade Intelligence · Customs Data</div>
        <div class="hero-title">Market Movement AI Dashboard</div>
        <div class="hero-subtitle">
            A professional dashboard that converts monthly Department of Customs data into market trends,
            product movement, country dependence, customs route signals, business opportunities, and policy risks.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# -----------------------------
# KPI cards
# -----------------------------

k1, k2, k3, k4 = st.columns(4)

with k1:
    kpi_card("Total Imports", f"Rs {imports_total:,.2f}B", "Import bill for selected period")

with k2:
    kpi_card("Total Exports", f"Rs {exports_total:,.2f}B", "Export earnings for selected period")

with k3:
    kpi_card("Trade Deficit", f"Rs {deficit_total:,.2f}B", "External trade gap")

with k4:
    kpi_card("Import / Export Ratio", f"{import_export_ratio:,.1f}x", "Imports compared to exports")

st.caption(f"Current period used: {current_col}. Source values are in Rs. thousands and shown here in Rs. billion.")

# -----------------------------
# Tabs
# -----------------------------

overview_tab, product_tab, opportunity_tab, country_tab, route_tab, insight_tab = st.tabs(
    ["Overview", "Products", "Opportunity Finder", "Countries", "Customs Routes", "Insights"]
)

# -----------------------------
# Overview tab
# -----------------------------

with overview_tab:
    st.subheader("Market Snapshot")

    c1, c2 = st.columns([1.2, 1])

    with c1:
        summary_df = pd.DataFrame({
            "Indicator": ["Imports", "Exports", "Trade Deficit", "Total Foreign Trade"],
            "Rs Billion": [imports_total, exports_total, deficit_total, total_trade]
        })

        fig = px.bar(
            summary_df,
            x="Indicator",
            y="Rs Billion",
            text="Rs Billion",
            title="Foreign Trade Snapshot"
        )

        fig.update_traces(
            texttemplate="%{text:,.1f}",
            textposition="outside"
        )

        fig.update_layout(
            height=430,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            title_font_size=20,
            font=dict(color="#102A43"),
            yaxis=dict(gridcolor="rgba(16,42,67,0.08)")
        )

        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.markdown(
            f"""
            <div class="section-card">
                <h3>What this means</h3>
                <p>Nepal's imports are around <b>{import_export_ratio:,.1f} times</b> larger than exports.</p>
                <p>The trade deficit of <b>Rs {deficit_total:,.2f} billion</b> is the main pressure point in the trade account.</p>
                <p>The dashboard helps identify <b>which products, countries, and customs routes</b> explain this movement.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.subheader("Top Signals")

    s1, s2, s3 = st.columns(3)

    with s1:
        st.info(f"Top import product: **{top_import_product}** — Rs {top_import_value:,.2f}B")

    with s2:
        st.success(f"Top export product: **{top_export_product}** — Rs {top_export_value:,.2f}B")

    with s3:
        st.warning(f"Top import country: **{top_import_country}**")

# -----------------------------
# Product tab
# -----------------------------

with product_tab:
    st.subheader("Product Movement")

    p1, p2 = st.columns(2)

    with p1:
        fig1 = horizontal_bar(
            top_imports,
            "Imports_Billion",
            "Description",
            "Top 10 Imported Products",
            "Rs Billion",
            "Product"
        )
        st.plotly_chart(fig1, use_container_width=True)

    with p2:
        fig2 = horizontal_bar(
            top_exports,
            "Exports_Billion",
            "Description",
            "Top 10 Exported Products",
            "Rs Billion",
            "Product"
        )
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("### Product Search")

    search_text = st.text_input(
        "Search commodity name or HS code",
        placeholder="Example: gold, petroleum, rice, 2710..."
    )

    if search_text:
        mask = (
            imports["Description"].astype(str).str.contains(search_text, case=False, na=False)
            | imports["HSCode"].astype(str).str.contains(search_text, case=False, na=False)
        )
        result = imports.loc[mask].sort_values("Imports_Value", ascending=False).head(30)
    else:
        result = top_imports

    st.dataframe(
        result[[
            "HSCode",
            "Description",
            "Unit",
            "Quantity",
            "Imports_Billion",
            "Revenue_Billion",
            "Approx_Duty_Rate"
        ]],
        use_container_width=True
    )

    st.download_button(
        "Download product search / top imports as CSV",
        data=result.to_csv(index=False).encode("utf-8"),
        file_name="product_search_results.csv",
        mime="text/csv"
    )

# -----------------------------
# Opportunity Finder tab
# -----------------------------

with opportunity_tab:
    st.subheader("Opportunity Finder")

    st.markdown(
        """
        This section ranks imported products by market opportunity using customs data.
        The score is based on import value, customs revenue, and approximate duty signal.
        """
    )

    opportunity_df = imports.copy()

    opportunity_df = opportunity_df[
        opportunity_df["Imports_Billion"].notna() &
        (opportunity_df["Imports_Billion"] > 0)
    ].copy()

    opportunity_df["Market_Size_Score"] = opportunity_df["Imports_Billion"].rank(pct=True) * 100
    opportunity_df["Revenue_Signal_Score"] = opportunity_df["Revenue_Billion"].rank(pct=True) * 100
    opportunity_df["Duty_Signal_Score"] = opportunity_df["Approx_Duty_Rate"].rank(pct=True) * 100

    opportunity_df["Opportunity_Score"] = (
        opportunity_df["Market_Size_Score"] * 0.60 +
        opportunity_df["Revenue_Signal_Score"] * 0.25 +
        opportunity_df["Duty_Signal_Score"] * 0.15
    )

    high_import_threshold = opportunity_df["Imports_Billion"].quantile(0.90)

    def classify_opportunity(row):
        if row["Imports_Billion"] >= high_import_threshold and row["Approx_Duty_Rate"] >= 5:
            return "High import + duty signal"
        elif row["Imports_Billion"] >= high_import_threshold:
            return "Large import market"
        elif row["Opportunity_Score"] >= 85:
            return "Strong watchlist"
        elif row["Opportunity_Score"] >= 70:
            return "Medium watchlist"
        else:
            return "Low priority"

    opportunity_df["Opportunity_Type"] = opportunity_df.apply(classify_opportunity, axis=1)

    f1, f2, f3 = st.columns(3)

    with f1:
        min_import_value = st.number_input(
            "Minimum import value, Rs billion",
            min_value=0.0,
            value=1.0,
            step=0.5
        )

    with f2:
        top_n = st.slider(
            "Number of products to show",
            min_value=10,
            max_value=100,
            value=25,
            step=5
        )

    with f3:
        opportunity_search = st.text_input(
            "Search product / HS code",
            placeholder="gold, petroleum, rice, 2710...",
            key="opportunity_search"
        )

    filtered_opportunity = opportunity_df[
        opportunity_df["Imports_Billion"] >= min_import_value
    ].copy()

    if opportunity_search:
        filtered_opportunity = filtered_opportunity[
            filtered_opportunity["Description"].astype(str).str.contains(opportunity_search, case=False, na=False)
            | filtered_opportunity["HSCode"].astype(str).str.contains(opportunity_search, case=False, na=False)
        ]

    filtered_opportunity = filtered_opportunity.sort_values(
        "Opportunity_Score",
        ascending=False
    ).head(top_n)

    o1, o2, o3, o4 = st.columns(4)

    with o1:
        kpi_card("Products Screened", f"{len(opportunity_df):,}", "Imported products with positive value")

    with o2:
        kpi_card("Shown Products", f"{len(filtered_opportunity):,}", "After filters")

    with o3:
        if len(filtered_opportunity) > 0:
            kpi_card("Top Score", f"{filtered_opportunity['Opportunity_Score'].max():,.1f}", "Out of 100")
        else:
            kpi_card("Top Score", "0", "No result")

    with o4:
        if len(filtered_opportunity) > 0:
            kpi_card("Largest Import", f"Rs {filtered_opportunity['Imports_Billion'].max():,.2f}B", "Within filtered list")
        else:
            kpi_card("Largest Import", "Rs 0B", "No result")

    st.markdown("### Top Opportunity Products")

    display_cols = [
        "HSCode",
        "Description",
        "Unit",
        "Imports_Billion",
        "Revenue_Billion",
        "Approx_Duty_Rate",
        "Market_Size_Score",
        "Revenue_Signal_Score",
        "Duty_Signal_Score",
        "Opportunity_Score",
        "Opportunity_Type"
    ]

    if len(filtered_opportunity) == 0:
        st.warning("No products found. Try reducing the minimum import value or clearing the search box.")
    else:
        st.dataframe(
            filtered_opportunity[display_cols].style.format({
                "Imports_Billion": "{:,.2f}",
                "Revenue_Billion": "{:,.2f}",
                "Approx_Duty_Rate": "{:,.2f}",
                "Market_Size_Score": "{:,.1f}",
                "Revenue_Signal_Score": "{:,.1f}",
                "Duty_Signal_Score": "{:,.1f}",
                "Opportunity_Score": "{:,.1f}"
            }),
            use_container_width=True
        )

        st.markdown("### Opportunity Map")

        fig_opportunity = px.scatter(
            filtered_opportunity,
            x="Imports_Billion",
            y="Approx_Duty_Rate",
            size="Opportunity_Score",
            hover_name="Description",
            hover_data={
                "HSCode": True,
                "Imports_Billion": ":,.2f",
                "Revenue_Billion": ":,.2f",
                "Approx_Duty_Rate": ":,.2f",
                "Opportunity_Score": ":,.1f"
            },
            title="Import Opportunity Map",
            labels={
                "Imports_Billion": "Import Value, Rs Billion",
                "Approx_Duty_Rate": "Approx. Duty / Revenue Rate, %"
            }
        )

        fig_opportunity.update_layout(
            height=540,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            title_font_size=20,
            font=dict(color="#102A43"),
            xaxis=dict(gridcolor="rgba(16,42,67,0.08)"),
            yaxis=dict(gridcolor="rgba(16,42,67,0.08)")
        )

        st.plotly_chart(fig_opportunity, use_container_width=True)

        st.markdown(
            """
            <div class="opportunity-card">
                <b>How to read this:</b><br>
                Higher import value means the product already has a large market in Nepal.<br>
                Higher revenue or duty signal may indicate policy sensitivity, tariff protection, or strong taxable import flow.<br>
                Higher opportunity score means the product deserves deeper business, production, import-substitution, or policy analysis.
            </div>
            """,
            unsafe_allow_html=True
        )

        st.download_button(
            "Download opportunity products as CSV",
            data=filtered_opportunity[display_cols].to_csv(index=False).encode("utf-8"),
            file_name="opportunity_finder_products.csv",
            mime="text/csv"
        )

# -----------------------------
# Country tab
# -----------------------------

with country_tab:
    st.subheader("Country Intelligence")

    country_list = (
        countries["Partner Countries"]
        .dropna()
        .sort_values()
        .unique()
    )

    country_list_as_list = list(country_list)

    selected_country = st.selectbox(
        "Select a partner country",
        country_list,
        index=country_list_as_list.index("India") if "India" in country_list_as_list else 0
    )

    selected_country_row = countries[
        countries["Partner Countries"] == selected_country
    ].iloc[0]

    selected_imports = selected_country_row["Imports_Billion"]
    selected_exports = selected_country_row["Exports_Billion"]
    selected_balance = selected_country_row["Trade_Balance_Billion"]

    selected_import_share = (selected_imports / imports_total) * 100 if imports_total else 0
    selected_export_share = (selected_exports / exports_total) * 100 if exports_total else 0

    if selected_balance < 0:
        balance_status = "Trade Deficit"
        balance_note = f"Nepal imports more from {selected_country} than it exports to that country."
    else:
        balance_status = "Trade Surplus"
        balance_note = f"Nepal exports more to {selected_country} than it imports from that country."

    if selected_import_share >= 30:
        dependency_level = "High dependency"
    elif selected_import_share >= 10:
        dependency_level = "Medium dependency"
    else:
        dependency_level = "Low dependency"

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi_card("Imports From Country", f"Rs {selected_imports:,.2f}B", f"{selected_import_share:,.1f}% of total imports")

    with c2:
        kpi_card("Exports To Country", f"Rs {selected_exports:,.2f}B", f"{selected_export_share:,.1f}% of total exports")

    with c3:
        kpi_card(balance_status, f"Rs {abs(selected_balance):,.2f}B", "Absolute trade balance")

    with c4:
        kpi_card("Dependency Level", dependency_level, "Based on import share")

    st.markdown("### Country Profile")

    profile_col1, profile_col2 = st.columns([1.1, 1])

    with profile_col1:
        country_mix = pd.DataFrame({
            "Flow": ["Imports", "Exports"],
            "Rs Billion": [selected_imports, selected_exports]
        })

        fig_country_mix = px.bar(
            country_mix,
            x="Flow",
            y="Rs Billion",
            text="Rs Billion",
            title=f"Nepal's Trade With {selected_country}"
        )

        fig_country_mix.update_traces(
            texttemplate="%{text:,.2f}",
            textposition="outside"
        )

        fig_country_mix.update_layout(
            height=420,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            title_font_size=20,
            font=dict(color="#102A43"),
            yaxis=dict(gridcolor="rgba(16,42,67,0.08)")
        )

        st.plotly_chart(fig_country_mix, use_container_width=True)

    with profile_col2:
        st.markdown(
            f"""
            <div class="section-card">
                <h3>{selected_country} — Trade Interpretation</h3>
                <p><b>{balance_status}:</b> {balance_note}</p>
                <p><b>Import dependence:</b> {selected_country} accounts for around <b>{selected_import_share:,.1f}%</b> of Nepal's total imports.</p>
                <p><b>Export dependence:</b> {selected_country} accounts for around <b>{selected_export_share:,.1f}%</b> of Nepal's total exports.</p>
                <p><b>Risk signal:</b> {dependency_level}.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("### Top Products Linked With Selected Country")

    selected_import_products = import_partner[
        import_partner["Partner Countries"] == selected_country
    ].sort_values("Imports_Value", ascending=False).head(15)

    selected_export_products = export_partner[
        export_partner["Partner Countries"] == selected_country
    ].sort_values("Exports_Value", ascending=False).head(15)

    p1, p2 = st.columns(2)

    with p1:
        st.markdown(f"#### Top Imports From {selected_country}")

        if len(selected_import_products) == 0:
            st.warning("No import product data found for this country.")
        else:
            fig_import_country_products = horizontal_bar(
                selected_import_products,
                "Imports_Billion",
                "Description",
                f"Top Imports From {selected_country}",
                "Rs Billion",
                "Product",
                height=520
            )

            st.plotly_chart(fig_import_country_products, use_container_width=True)

            st.dataframe(
                selected_import_products[[
                    "HSCode",
                    "Description",
                    "Unit",
                    "Quantity",
                    "Imports_Billion",
                    "Revenue_Billion"
                ]],
                use_container_width=True
            )

    with p2:
        st.markdown(f"#### Top Exports To {selected_country}")

        if len(selected_export_products) == 0:
            st.warning("No export product data found for this country.")
        else:
            fig_export_country_products = horizontal_bar(
                selected_export_products,
                "Exports_Billion",
                "Description",
                f"Top Exports To {selected_country}",
                "Rs Billion",
                "Product",
                height=520
            )

            st.plotly_chart(fig_export_country_products, use_container_width=True)

            st.dataframe(
                selected_export_products[[
                    "HSCode",
                    "Description",
                    "Unit",
                    "Quantity",
                    "Exports_Billion"
                ]],
                use_container_width=True
            )

    st.markdown("### Country-Level Opportunity and Risk Signals")

    signal_col1, signal_col2 = st.columns(2)

    with signal_col1:
        if len(selected_import_products) > 0:
            top_country_import_product = selected_import_products.iloc[0]["Description"]
            top_country_import_value = selected_import_products.iloc[0]["Imports_Billion"]

            st.markdown(
                f"""
                <div class="opportunity-card">
                    <h3>Business Signal</h3>
                    <p>The largest import from <b>{selected_country}</b> is <b>{top_country_import_product}</b>,
                    worth around <b>Rs {top_country_import_value:,.2f} billion</b>.</p>
                    <p>This product may deserve deeper study for sourcing, distribution, import financing,
                    or import-substitution potential.</p>
                </div>
                """,
                unsafe_allow_html=True
            )
        else:
            st.info("No business signal available because import product data was not found.")

    with signal_col2:
        st.markdown(
            f"""
            <div class="risk-card">
                <h3>Risk Signal</h3>
                <p>{selected_country} represents around <b>{selected_import_share:,.1f}%</b> of Nepal's import bill.</p>
                <p>If this share is high, Nepal may face supplier-country dependence, price shock exposure,
                exchange-rate pressure, or logistics concentration risk.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("---")
    st.markdown("### Overall Country Rankings")

    c1, c2 = st.columns(2)

    with c1:
        fig3 = horizontal_bar(
            top_import_countries,
            "Imports_Billion",
            "Partner Countries",
            "Top Import Countries",
            "Rs Billion",
            "Country"
        )
        st.plotly_chart(fig3, use_container_width=True)

    with c2:
        fig4 = horizontal_bar(
            top_export_countries,
            "Exports_Billion",
            "Partner Countries",
            "Top Export Countries",
            "Rs Billion",
            "Country"
        )
        st.plotly_chart(fig4, use_container_width=True)

    st.markdown("### Trade Balance Pressure")

    b1, b2 = st.columns(2)

    with b1:
        st.markdown("#### Biggest Deficit Countries")
        st.dataframe(
            top_deficit_countries[[
                "Partner Countries",
                "Imports_Billion",
                "Exports_Billion",
                "Trade_Balance_Billion"
            ]],
            use_container_width=True
        )

    with b2:
        st.markdown("#### Surplus Countries")
        st.dataframe(
            top_surplus_countries[[
                "Partner Countries",
                "Imports_Billion",
                "Exports_Billion",
                "Trade_Balance_Billion"
            ]],
            use_container_width=True
        )

# -----------------------------
# Customs route tab
# -----------------------------

with route_tab:
    st.subheader("Customs Route Intelligence")

    r1, r2 = st.columns([1.3, 1])

    with r1:
        fig5 = horizontal_bar(
            top_customs,
            "Imports_Billion",
            "Customs",
            "Top Customs Offices by Import Value",
            "Rs Billion",
            "Customs Office"
        )
        st.plotly_chart(fig5, use_container_width=True)

    with r2:
        st.markdown(
            f"""
            <div class="section-card">
                <h3>Route Concentration</h3>
                <p>The top 5 customs offices handle around <b>{top5_route_share:,.1f}%</b> of import value.</p>
                <p>The largest import route is <b>{top_customs_office}</b>, handling around <b>Rs {top_customs_value:,.2f} billion</b>.</p>
                <p>This can be used as a logistics risk and infrastructure pressure indicator.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.dataframe(
        customs.sort_values("Imports_Value", ascending=False)[[
            "Customs",
            "Imports_Billion",
            "Import_Share",
            "Exports_Billion",
            "Export_Share"
        ]],
        use_container_width=True
    )

# -----------------------------
# Insights tab
# -----------------------------

with insight_tab:
    st.subheader("AI-Style Trade Brief")

    brief_text = f"""Market Movement AI Dashboard — Monthly Trade Brief

Current period: {current_col}
Source: Department of Customs monthly foreign trade data
Unit: Rs. billion, converted from source values in Rs. thousands

1. Market Snapshot
Nepal recorded total imports of Rs {imports_total:,.2f} billion and total exports of Rs {exports_total:,.2f} billion.
The trade deficit stood at Rs {deficit_total:,.2f} billion.
Imports were around {import_export_ratio:,.1f} times larger than exports.

2. Product Movement
The leading imported product was {top_import_product}, with imports worth around Rs {top_import_value:,.2f} billion.
The leading exported product was {top_export_product}, with exports worth around Rs {top_export_value:,.2f} billion.

3. Country Movement
The largest import partner was {top_import_country}.
The largest export destination was {top_export_country}.

4. Customs Route Movement
The most important customs office by import value was {top_customs_office}, handling around Rs {top_customs_value:,.2f} billion.

5. Interpretation
The data shows Nepal's continued import dependence and a large trade deficit. High-import products may deserve deeper study for import-substitution, sourcing, trade finance, or policy analysis.
"""

    brief_html = brief_text.replace("\n", "<br>")

    st.markdown(
        f"""
        <div class="insight-card">
            <h3>Monthly Trade Brief</h3>
            <p>{brief_html}</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    i1, i2 = st.columns(2)

    with i1:
        st.markdown(
            f"""
            <div class="opportunity-card">
                <h3>Business Opportunity Signals</h3>
                <ul>
                    <li>High import products can be studied for import-substitution possibilities.</li>
                    <li><b>{top_import_product}</b> is a major import item and deserves deeper product-level analysis.</li>
                    <li>Export product <b>{top_export_product}</b> can be studied for export expansion.</li>
                    <li>Large supplier countries like <b>{top_import_country}</b> may reveal sourcing and trade-finance opportunities.</li>
                </ul>
            </div>
            """,
            unsafe_allow_html=True
        )

    with i2:
        st.markdown(
            f"""
            <div class="risk-card">
                <h3>Policy Risk Signals</h3>
                <ul>
                    <li>The trade deficit remains large compared with export earnings.</li>
                    <li>Imports are around <b>{import_export_ratio:,.1f} times</b> larger than exports.</li>
                    <li>Heavy dependence on major partner countries can create external vulnerability.</li>
                    <li>Route concentration through <b>{top_customs_office}</b> can create logistics pressure.</li>
                </ul>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("### Media / Research Story Ideas")

    st.markdown(
        """
        1. Why Nepal's imports are still far larger than exports.  
        2. Which products are driving Nepal's import bill?  
        3. What Nepal exports most — and why the export basket remains narrow.  
        4. How dependent is Nepal on major customs routes?  
        5. Can high-import products create domestic production opportunities?  
        """
    )

    st.download_button(
        "Download trade brief",
        data=brief_text.encode("utf-8"),
        file_name="monthly_trade_brief.txt",
        mime="text/plain"
    )