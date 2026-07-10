from io import BytesIO
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch

import streamlit as st
import pandas as pd
import plotly.express as px

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
    type=["xlsx"],
    key="customs_excel_uploader"
)

default_file_path = Path(__file__).parent / "customs.xlsx"
monthly_data_path = Path(__file__).parent / "monthly_data"
developer_photo_path = Path(__file__).parent / "utsav.png"

# Main dashboard needs one workbook as its default view.
# Priority: user upload → customs.xlsx → latest file inside monthly_data.
latest_monthly_file = None
if monthly_data_path.exists():
    monthly_files_for_default = sorted(monthly_data_path.glob("*.xlsx"))
    if monthly_files_for_default:
        latest_monthly_file = monthly_files_for_default[-1]

if uploaded_file is not None:
    file_source = uploaded_file
elif default_file_path.exists():
    file_source = default_file_path
elif latest_monthly_file is not None:
    file_source = latest_monthly_file
else:
    st.warning("Please add customs.xlsx to the app folder, or add monthly Excel files inside monthly_data.")
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
    trade.columns = trade.columns.astype(str).str.strip()

    # Some customs files use "Trade Indicators" while others use "Trade.Indicators".
    # Standardize it so the rest of the dashboard works.
    for col in trade.columns:
        clean_col = str(col).lower().replace(".", " ").replace("_", " ").strip()
        if "trade" in clean_col and "indicator" in clean_col:
            trade = trade.rename(columns={col: "Trade.Indicators"})
            break

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




def get_trade_value_from_df(trade_df, pattern, value_col):
    indicator_series = (
        trade_df["Trade.Indicators"]
        .astype(str)
        .str.lower()
        .str.replace(".", " ", regex=False)
        .str.replace("_", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    mask = indicator_series.str.contains(
        pattern,
        case=False,
        na=False,
        regex=True
    )

    if mask.sum() == 0:
        return 0.0

    return safe_number(trade_df.loc[mask, value_col].iloc[0])


def read_trade_direction_smart(file):
    # Read without assumptions so monthly files with slightly different labels still work.
    raw = pd.read_excel(file, sheet_name="1_Trade_Direction", header=None)

    header_row = None

    for i in range(len(raw)):
        row_text = " ".join(
            raw.iloc[i]
            .dropna()
            .astype(str)
            .str.strip()
            .tolist()
        ).lower()

        normalized = row_text.replace(".", " ").replace("_", " ")

        if "trade" in normalized and "indicator" in normalized:
            header_row = i
            break

    if header_row is None:
        raise ValueError("Could not detect header row in 1_Trade_Direction sheet.")

    trade_df = pd.read_excel(file, sheet_name="1_Trade_Direction", header=header_row)
    trade_df.columns = trade_df.columns.astype(str).str.strip()

    # Rename Trade Indicators column to one standard name
    for col in trade_df.columns:
        clean_col = str(col).lower().replace(".", " ").replace("_", " ").strip()
        if "trade" in clean_col and "indicator" in clean_col:
            trade_df = trade_df.rename(columns={col: "Trade.Indicators"})
            break

    if "Trade.Indicators" not in trade_df.columns:
        raise ValueError("Trade indicator column found, but could not standardize it.")

    return trade_df


def extract_trade_snapshot_from_file(file, file_name):
    trade_df = read_trade_direction_smart(file)

    # Find latest/current value column.
    # We ignore SN and Change columns, then use the last numeric column.
    possible_cols = []

    for col in trade_df.columns:
        col_text = str(col).strip().lower()

        if col == "Trade.Indicators":
            continue
        if "change" in col_text:
            continue
        if col_text in ["sn", "s n", "s.n", "s.no", "s no"]:
            continue

        numeric_series = pd.to_numeric(trade_df[col], errors="coerce")
        if numeric_series.notna().sum() >= 3:
            possible_cols.append(col)

    if len(possible_cols) == 0:
        raise ValueError("No usable numeric value column found.")

    selected_col = possible_cols[-1]

    imports_raw = get_trade_value_from_df(trade_df, r"^imports\b", selected_col)
    exports_raw = get_trade_value_from_df(trade_df, r"^exports\b", selected_col)
    deficit_raw = get_trade_value_from_df(trade_df, r"^trade\s+deficit\b", selected_col)
    total_trade_raw = get_trade_value_from_df(trade_df, r"^total\s+foreign\s+trade\b", selected_col)

    period_name = (
        file_name
        .replace(".xlsx", "")
        .replace(".xls", "")
        .replace("_", " ")
        .replace("-", " ")
    )

    return {
        "File": file_name,
        "Period": period_name,
        "Source Column": selected_col,
        "Imports_Billion": rs_thousand_to_billion(imports_raw),
        "Exports_Billion": rs_thousand_to_billion(exports_raw),
        "Trade_Deficit_Billion": rs_thousand_to_billion(deficit_raw),
        "Total_Trade_Billion": rs_thousand_to_billion(total_trade_raw),
        "Import_Export_Ratio": imports_raw / exports_raw if exports_raw else 0
    }



def read_product_table(file, sheet_name, value_col):
    """Read commodity import/export table from a customs workbook."""
    df = pd.read_excel(file, sheet_name=sheet_name, header=2)
    df.columns = df.columns.astype(str).str.strip()

    required_cols = ["HSCode", "Description", value_col]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"{col} not found in {sheet_name}")

    df = df.copy()

    df["HSCode"] = (
        df["HSCode"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
        .str.zfill(8)
    )

    df["Description"] = df["Description"].astype(str).str.strip()

    df = df[
        ~df["Description"]
        .str.lower()
        .isin(["total", "grand total", "nan"])
    ]

    df[value_col] = pd.to_numeric(df[value_col], errors="coerce").fillna(0)
    df["Value_Billion"] = df[value_col] / 1_000_000

    df = (
        df.groupby(["HSCode", "Description"], as_index=False)["Value_Billion"]
        .sum()
    )

    return df


def build_product_movement_table(trend_files, sheet_name, value_col):
    """Build latest monthly product movement from cumulative monthly customs files."""
    monthly_tables = []
    periods = []

    for file_path in sorted(trend_files):
        period = file_path.stem.replace("_", " ")
        periods.append(period)

        temp = read_product_table(file_path, sheet_name, value_col)
        temp["Period"] = period
        monthly_tables.append(temp)

    if len(monthly_tables) < 2:
        raise ValueError("At least two files are required for product movement analysis.")

    all_products = pd.concat(monthly_tables, ignore_index=True)

    cumulative_wide = all_products.pivot_table(
        index=["HSCode", "Description"],
        columns="Period",
        values="Value_Billion",
        aggfunc="sum",
        fill_value=0
    )

    cumulative_wide = cumulative_wide[periods]

    # Customs files are cumulative, so monthly values are differences between cumulative files.
    monthly_wide = cumulative_wide.diff(axis=1)
    monthly_wide[periods[0]] = cumulative_wide[periods[0]]

    latest_period = periods[-1]
    previous_period = periods[-2]

    result = cumulative_wide.reset_index()[["HSCode", "Description"]].copy()
    result["Latest_Cumulative_Billion"] = cumulative_wide[latest_period].values
    result["Previous_Month_Billion"] = monthly_wide[previous_period].values
    result["Latest_Month_Billion"] = monthly_wide[latest_period].values
    result["Change_Billion"] = result["Latest_Month_Billion"] - result["Previous_Month_Billion"]

    result["Growth_Percent"] = result.apply(
        lambda row: (row["Change_Billion"] / row["Previous_Month_Billion"] * 100)
        if row["Previous_Month_Billion"] != 0 else None,
        axis=1
    )

    result["Latest_Period"] = latest_period
    result["Previous_Period"] = previous_period

    return result


def build_trend_summary_for_report(monthly_data_path):
    """Create a compact trend summary for the app and PDF report from static monthly_data files."""
    try:
        monthly_data_path = Path(monthly_data_path)
        if not monthly_data_path.exists():
            return None

        trend_files = sorted(monthly_data_path.glob("*.xlsx"))
        if len(trend_files) < 2:
            return None

        trend_rows = []
        for file_path in trend_files:
            try:
                trend_rows.append(extract_trade_snapshot_from_file(file_path, file_path.name))
            except Exception:
                continue

        if len(trend_rows) < 2:
            return None

        trend_df = pd.DataFrame(trend_rows)
        trend_df["Order"] = range(1, len(trend_df) + 1)

        trend_df["Monthly_Imports_Billion"] = trend_df["Imports_Billion"].diff()
        trend_df["Monthly_Exports_Billion"] = trend_df["Exports_Billion"].diff()
        trend_df["Monthly_Deficit_Billion"] = trend_df["Trade_Deficit_Billion"].diff()
        trend_df["Monthly_Total_Trade_Billion"] = trend_df["Total_Trade_Billion"].diff()

        trend_df.loc[trend_df.index[0], "Monthly_Imports_Billion"] = trend_df.loc[trend_df.index[0], "Imports_Billion"]
        trend_df.loc[trend_df.index[0], "Monthly_Exports_Billion"] = trend_df.loc[trend_df.index[0], "Exports_Billion"]
        trend_df.loc[trend_df.index[0], "Monthly_Deficit_Billion"] = trend_df.loc[trend_df.index[0], "Trade_Deficit_Billion"]
        trend_df.loc[trend_df.index[0], "Monthly_Total_Trade_Billion"] = trend_df.loc[trend_df.index[0], "Total_Trade_Billion"]

        latest = trend_df.iloc[-1]
        previous = trend_df.iloc[-2]

        import_change = latest["Imports_Billion"] - previous["Imports_Billion"]
        export_change = latest["Exports_Billion"] - previous["Exports_Billion"]
        deficit_change = latest["Trade_Deficit_Billion"] - previous["Trade_Deficit_Billion"]

        if import_change > export_change:
            movement_note = "Imports increased more than exports in the latest period, which may widen trade-deficit pressure."
        elif export_change > import_change:
            movement_note = "Exports increased more than imports in the latest period, which may slightly ease trade-deficit pressure."
        else:
            movement_note = "Imports and exports moved by a similar amount in the latest period."

        summary = {
            "files_used": len(trend_df),
            "latest_period": latest["Period"],
            "previous_period": previous["Period"],
            "latest_imports": latest["Imports_Billion"],
            "latest_exports": latest["Exports_Billion"],
            "latest_deficit": latest["Trade_Deficit_Billion"],
            "latest_total_trade": latest["Total_Trade_Billion"],
            "import_change": import_change,
            "export_change": export_change,
            "deficit_change": deficit_change,
            "movement_note": movement_note,
            "top_rising_import": None,
            "top_rising_import_change": None,
            "top_rising_export": None,
            "top_rising_export_change": None,
            "top_falling_import": None,
            "top_falling_import_change": None,
            "top_falling_export": None,
            "top_falling_export_change": None,
        }

        try:
            import_product_movement = build_product_movement_table(
                trend_files=trend_files,
                sheet_name="5_Imports_By_Commodity",
                value_col="Imports_Value"
            )
            export_product_movement = build_product_movement_table(
                trend_files=trend_files,
                sheet_name="7_Exports_By_Commodity",
                value_col="Exports_Value"
            )

            import_filtered = import_product_movement[import_product_movement["Latest_Month_Billion"] >= 0.1].copy()
            export_filtered = export_product_movement[export_product_movement["Latest_Month_Billion"] >= 0.1].copy()

            if len(import_filtered) > 0:
                top_rising_import = import_filtered.sort_values("Change_Billion", ascending=False).iloc[0]
                top_falling_import = import_filtered.sort_values("Change_Billion", ascending=True).iloc[0]
                summary["top_rising_import"] = top_rising_import["Description"]
                summary["top_rising_import_change"] = top_rising_import["Change_Billion"]
                summary["top_falling_import"] = top_falling_import["Description"]
                summary["top_falling_import_change"] = top_falling_import["Change_Billion"]

            if len(export_filtered) > 0:
                top_rising_export = export_filtered.sort_values("Change_Billion", ascending=False).iloc[0]
                top_falling_export = export_filtered.sort_values("Change_Billion", ascending=True).iloc[0]
                summary["top_rising_export"] = top_rising_export["Description"]
                summary["top_rising_export_change"] = top_rising_export["Change_Billion"]
                summary["top_falling_export"] = top_falling_export["Description"]
                summary["top_falling_export_change"] = top_falling_export["Change_Billion"]
        except Exception:
            pass

        return summary

    except Exception:
        return None

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

def create_pdf_report(
    current_col,
    imports_total,
    exports_total,
    deficit_total,
    total_trade,
    import_export_ratio,
    top_import_product,
    top_import_value,
    top_export_product,
    top_export_value,
    top_import_country,
    top_export_country,
    top_customs_office,
    top_customs_value,
    trend_summary=None
):
    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=42,
        leftMargin=42,
        topMargin=42,
        bottomMargin=42
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TradePulseTitle",
        parent=styles["Title"],
        fontSize=24,
        leading=30,
        textColor=colors.HexColor("#0F172A"),
        spaceAfter=12
    )

    subtitle_style = ParagraphStyle(
        "TradePulseSubtitle",
        parent=styles["BodyText"],
        fontSize=11,
        leading=16,
        textColor=colors.HexColor("#475569"),
        spaceAfter=18
    )

    section_style = ParagraphStyle(
        "TradePulseSection",
        parent=styles["Heading2"],
        fontSize=15,
        leading=20,
        textColor=colors.HexColor("#102A43"),
        spaceBefore=14,
        spaceAfter=8
    )

    body_style = ParagraphStyle(
        "TradePulseBody",
        parent=styles["BodyText"],
        fontSize=10.5,
        leading=16,
        textColor=colors.HexColor("#1E293B"),
        spaceAfter=8
    )

    small_style = ParagraphStyle(
        "TradePulseSmall",
        parent=styles["BodyText"],
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor("#64748B")
    )

    story = []

    story.append(Paragraph("TRADEPULSE NEPAL", small_style))
    story.append(Paragraph("Monthly Trade Intelligence Brief", title_style))

    story.append(Paragraph(
        """
        A data-driven summary of Nepal's foreign trade movement based on Department of Customs data.
        This brief converts customs statistics into market signals, business opportunities, policy risks,
        and report-ready insights.
        """,
        subtitle_style
    ))

    meta_data = [
        ["Current period", current_col],
        ["Source", "Department of Customs, Government of Nepal"],
        ["Unit", "Rs. billion"],
        ["Conversion", "Source values in Rs. thousands divided by 1,000,000"]
    ]

    meta_table = Table(meta_data, colWidths=[1.55 * inch, 4.65 * inch])
    meta_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
        ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#0F172A")),
        ("TEXTCOLOR", (1, 0), (1, -1), colors.HexColor("#334155")),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("PADDING", (0, 0), (-1, -1), 8),
    ]))

    story.append(meta_table)
    story.append(Spacer(1, 18))

    story.append(Paragraph("1. Market Snapshot", section_style))

    kpi_data = [
        ["Indicator", "Value"],
        ["Total Imports", f"Rs {imports_total:,.2f} billion"],
        ["Total Exports", f"Rs {exports_total:,.2f} billion"],
        ["Trade Deficit", f"Rs {deficit_total:,.2f} billion"],
        ["Total Foreign Trade", f"Rs {total_trade:,.2f} billion"],
        ["Import / Export Ratio", f"{import_export_ratio:,.1f}x"]
    ]

    kpi_table = Table(kpi_data, colWidths=[2.6 * inch, 3.6 * inch])
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F172A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor("#1E293B")),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("PADDING", (0, 0), (-1, -1), 8),
    ]))

    story.append(kpi_table)
    story.append(Spacer(1, 12))

    story.append(Paragraph(
        f"""
        Nepal recorded total imports of <b>Rs {imports_total:,.2f} billion</b> and exports of
        <b>Rs {exports_total:,.2f} billion</b>. Imports were around
        <b>{import_export_ratio:,.1f} times</b> larger than exports, indicating continued import dependence.
        """,
        body_style
    ))

    story.append(Paragraph("2. Product Movement", section_style))

    product_data = [
        ["Signal", "Product", "Value"],
        ["Top import product", str(top_import_product), f"Rs {top_import_value:,.2f} billion"],
        ["Top export product", str(top_export_product), f"Rs {top_export_value:,.2f} billion"]
    ]

    product_table = Table(product_data, colWidths=[1.65 * inch, 3.15 * inch, 1.4 * inch])
    product_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#102A43")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9.3),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("PADDING", (0, 0), (-1, -1), 7),
    ]))

    story.append(product_table)
    story.append(Spacer(1, 10))

    story.append(Paragraph(
        f"""
        The leading imported product was <b>{top_import_product}</b>, while the leading exported product was
        <b>{top_export_product}</b>. These products can be used as starting points for product-level
        research, market mapping, sourcing analysis, and export-diversification discussion.
        """,
        body_style
    ))

    story.append(Paragraph("3. Country and Customs Route Movement", section_style))

    country_route_data = [
        ["Signal", "Result"],
        ["Largest import partner", str(top_import_country)],
        ["Largest export destination", str(top_export_country)],
        ["Largest customs route by import value", str(top_customs_office)],
        ["Import value handled by top route", f"Rs {top_customs_value:,.2f} billion"]
    ]

    country_route_table = Table(country_route_data, colWidths=[2.4 * inch, 3.8 * inch])
    country_route_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F172A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("PADDING", (0, 0), (-1, -1), 8),
    ]))

    story.append(country_route_table)
    story.append(Spacer(1, 10))

    story.append(Paragraph(
        f"""
        The largest import partner was <b>{top_import_country}</b>, while the largest export destination was
        <b>{top_export_country}</b>. The most important customs office by import value was
        <b>{top_customs_office}</b>. These indicators help identify country dependence and logistics concentration.
        """,
        body_style
    ))


    if trend_summary:
        story.append(Paragraph("4. Multi-Month Trend Movement", section_style))

        trend_data = [
            ["Signal", "Result"],
            ["Files compared", f"{trend_summary.get('files_used', 0)} monthly files"],
            ["Latest period", str(trend_summary.get('latest_period', ''))],
            ["Latest cumulative imports", f"Rs {trend_summary.get('latest_imports', 0):,.2f} billion"],
            ["Latest cumulative exports", f"Rs {trend_summary.get('latest_exports', 0):,.2f} billion"],
            ["Latest trade deficit", f"Rs {trend_summary.get('latest_deficit', 0):,.2f} billion"],
            ["Latest import change", f"Rs {trend_summary.get('import_change', 0):,.2f} billion"],
            ["Latest export change", f"Rs {trend_summary.get('export_change', 0):,.2f} billion"],
            ["Deficit change", f"Rs {trend_summary.get('deficit_change', 0):,.2f} billion"],
        ]

        trend_table = Table(trend_data, colWidths=[2.5 * inch, 3.7 * inch])
        trend_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#102A43")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
            ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9.3),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            ("PADDING", (0, 0), (-1, -1), 7),
        ]))

        story.append(trend_table)
        story.append(Spacer(1, 10))
        story.append(Paragraph(str(trend_summary.get("movement_note", "")), body_style))

        if trend_summary.get("top_rising_import") or trend_summary.get("top_rising_export"):
            product_signal_rows = [["Product signal", "Latest movement"]]
            if trend_summary.get("top_rising_import"):
                product_signal_rows.append([
                    "Top rising import",
                    f"{trend_summary.get('top_rising_import')} (+Rs {trend_summary.get('top_rising_import_change', 0):,.2f} billion)"
                ])
            if trend_summary.get("top_rising_export"):
                product_signal_rows.append([
                    "Top rising export",
                    f"{trend_summary.get('top_rising_export')} (+Rs {trend_summary.get('top_rising_export_change', 0):,.2f} billion)"
                ])

            product_signal_table = Table(product_signal_rows, colWidths=[2.0 * inch, 4.2 * inch])
            product_signal_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F172A")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9.2),
                ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("PADDING", (0, 0), (-1, -1), 7),
            ]))
            story.append(product_signal_table)
            story.append(Spacer(1, 10))

    story.append(Paragraph("5. Business Opportunity Signals", section_style))

    opportunity_points = [
        f"High-import products such as {top_import_product} may deserve deeper import-substitution research.",
        "Large import values can indicate existing domestic demand, but profitability must be checked separately.",
        f"Export products such as {top_export_product} can be studied for export expansion and destination diversification.",
        f"Major supplier countries such as {top_import_country} may reveal sourcing, distribution, and trade-finance opportunities."
    ]

    for point in opportunity_points:
        story.append(Paragraph(f"• {point}", body_style))

    story.append(Paragraph("6. Policy Risk Signals", section_style))

    risk_points = [
        f"The trade deficit remains large at Rs {deficit_total:,.2f} billion.",
        f"Imports are around {import_export_ratio:,.1f} times larger than exports.",
        "Heavy dependence on major partner countries can create external vulnerability.",
        f"Customs concentration through {top_customs_office} can create logistics and infrastructure pressure."
    ]

    for point in risk_points:
        story.append(Paragraph(f"• {point}", body_style))

    story.append(Spacer(1, 16))
    disclaimer_table = Table(
        [[Paragraph(
            """
            <b>Methodology note:</b> This report is generated from the uploaded Department of Customs workbook.
            Monetary values are converted from Rs. thousands to Rs. billion. The opportunity and risk signals
            are screening indicators for research and discussion, not investment advice.
            """,
            small_style
        )]],
        colWidths=[6.2 * inch]
    )

    disclaimer_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F1F5F9")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("PADDING", (0, 0), (-1, -1), 10),
    ]))

    story.append(disclaimer_table)

    doc.build(story)
    buffer.seek(0)
    return buffer


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

# Find the current/latest value column for the main dashboard.
# Some files use a full FY column name, while Shrawan uses "current".
possible_year_cols = []

for col in trade.columns:
    col_text = str(col).strip().lower()

    if col == "Trade.Indicators":
        continue
    if "change" in col_text:
        continue
    if col_text in ["sn", "s n", "s.n", "s.no", "s no"]:
        continue

    numeric_series = pd.to_numeric(trade[col], errors="coerce")
    if numeric_series.notna().sum() >= 3:
        possible_year_cols.append(col)

preferred_cols = [
    col for col in possible_year_cols
    if "2082/83" in str(col) or str(col).strip().lower() == "current"
]

current_col = (
    "FY 2082/83 (First 11 Months)"
    if "FY 2082/83 (First 11 Months)" in trade.columns
    else preferred_cols[-1]
    if preferred_cols
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

overview_tab, product_tab, opportunity_tab, country_tab, route_tab, trend_tab, about_tab, insight_tab = st.tabs(
    ["Overview", "Products", "Opportunity Finder", "Countries", "Customs Routes", "Trends", "About / Methodology", "Insights"]
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
# Trends tab
# -----------------------------

with trend_tab:
    st.subheader("Multi-Month Trade Trends")

    st.markdown(
        """
        <div class="insight-card">
            <b>How this works:</b> This tab automatically reads all monthly customs Excel files
            stored inside the <b>monthly_data</b> folder in this app. No upload is needed.
            For correct ordering, name files like <b>01_Shravan.xlsx</b>, <b>02_Bhadra.xlsx</b>,
            ... <b>11_Jestha.xlsx</b>.
        </div>
        """,
        unsafe_allow_html=True
    )

    monthly_data_path = Path(__file__).parent / "monthly_data"

    if not monthly_data_path.exists():
        st.warning("monthly_data folder not found. Please add monthly Excel files inside a folder named monthly_data.")
    else:
        trend_files = sorted(monthly_data_path.glob("*.xlsx"))

        if len(trend_files) < 2:
            st.warning("Please add at least two monthly Excel files inside the monthly_data folder.")
        else:
            trend_rows = []

            for file_path in trend_files:
                try:
                    row = extract_trade_snapshot_from_file(file_path, file_path.name)
                    trend_rows.append(row)
                except Exception as e:
                    st.warning(f"Could not process {file_path.name}: {e}")

            if len(trend_rows) < 2:
                st.warning("Could not process enough files for trend comparison.")
            else:
                trend_df = pd.DataFrame(trend_rows)
                trend_df["Order"] = range(1, len(trend_df) + 1)

                # Department of Customs monthly files are often cumulative.
                # Monthly movement is estimated by differencing consecutive cumulative files.
                trend_df["Monthly_Imports_Billion"] = trend_df["Imports_Billion"].diff()
                trend_df["Monthly_Exports_Billion"] = trend_df["Exports_Billion"].diff()
                trend_df["Monthly_Deficit_Billion"] = trend_df["Trade_Deficit_Billion"].diff()
                trend_df["Monthly_Total_Trade_Billion"] = trend_df["Total_Trade_Billion"].diff()

                trend_df.loc[trend_df.index[0], "Monthly_Imports_Billion"] = trend_df.loc[trend_df.index[0], "Imports_Billion"]
                trend_df.loc[trend_df.index[0], "Monthly_Exports_Billion"] = trend_df.loc[trend_df.index[0], "Exports_Billion"]
                trend_df.loc[trend_df.index[0], "Monthly_Deficit_Billion"] = trend_df.loc[trend_df.index[0], "Trade_Deficit_Billion"]
                trend_df.loc[trend_df.index[0], "Monthly_Total_Trade_Billion"] = trend_df.loc[trend_df.index[0], "Total_Trade_Billion"]

                latest = trend_df.iloc[-1]
                previous = trend_df.iloc[-2]

                import_change = latest["Imports_Billion"] - previous["Imports_Billion"]
                export_change = latest["Exports_Billion"] - previous["Exports_Billion"]
                deficit_change = latest["Trade_Deficit_Billion"] - previous["Trade_Deficit_Billion"]

                t1, t2, t3, t4 = st.columns(4)

                with t1:
                    kpi_card(
                        "Latest Imports",
                        f"Rs {latest['Imports_Billion']:,.2f}B",
                        f"Previous change: Rs {import_change:,.2f}B"
                    )

                with t2:
                    kpi_card(
                        "Latest Exports",
                        f"Rs {latest['Exports_Billion']:,.2f}B",
                        f"Previous change: Rs {export_change:,.2f}B"
                    )

                with t3:
                    kpi_card(
                        "Latest Deficit",
                        f"Rs {latest['Trade_Deficit_Billion']:,.2f}B",
                        f"Previous change: Rs {deficit_change:,.2f}B"
                    )

                with t4:
                    kpi_card(
                        "Files Used",
                        f"{len(trend_df)}",
                        "Monthly customs files"
                    )

                st.markdown("### Cumulative Trade Trend")

                cumulative_long = trend_df.melt(
                    id_vars=["Order", "Period"],
                    value_vars=[
                        "Imports_Billion",
                        "Exports_Billion",
                        "Trade_Deficit_Billion",
                        "Total_Trade_Billion"
                    ],
                    var_name="Indicator",
                    value_name="Rs_Billion"
                )

                fig_cumulative = px.line(
                    cumulative_long,
                    x="Order",
                    y="Rs_Billion",
                    color="Indicator",
                    markers=True,
                    title="Cumulative Trade Movement",
                    labels={
                        "Order": "Month Order",
                        "Rs_Billion": "Rs Billion",
                        "Indicator": "Indicator"
                    },
                    hover_data=["Period"]
                )

                fig_cumulative.update_layout(
                    height=520,
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)",
                    title_font_size=22,
                    font=dict(color="#0F172A"),
                    xaxis=dict(gridcolor="rgba(15,23,42,0.08)"),
                    yaxis=dict(gridcolor="rgba(15,23,42,0.08)")
                )

                st.plotly_chart(fig_cumulative, use_container_width=True)

                st.markdown("### Estimated Monthly Movement")

                monthly_long = trend_df.melt(
                    id_vars=["Order", "Period"],
                    value_vars=[
                        "Monthly_Imports_Billion",
                        "Monthly_Exports_Billion",
                        "Monthly_Deficit_Billion",
                        "Monthly_Total_Trade_Billion"
                    ],
                    var_name="Indicator",
                    value_name="Rs_Billion"
                )

                fig_monthly = px.bar(
                    monthly_long,
                    x="Order",
                    y="Rs_Billion",
                    color="Indicator",
                    barmode="group",
                    title="Estimated Monthly Movement from Cumulative Data",
                    labels={
                        "Order": "Month Order",
                        "Rs_Billion": "Rs Billion",
                        "Indicator": "Indicator"
                    },
                    hover_data=["Period"]
                )

                fig_monthly.update_layout(
                    height=540,
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)",
                    title_font_size=22,
                    font=dict(color="#0F172A"),
                    xaxis=dict(gridcolor="rgba(15,23,42,0.08)"),
                    yaxis=dict(gridcolor="rgba(15,23,42,0.08)")
                )

                st.plotly_chart(fig_monthly, use_container_width=True)

                st.markdown("### Latest Movement Signal")

                if import_change > export_change:
                    movement_note = "Imports increased more than exports in the latest period, which may widen trade-deficit pressure."
                elif export_change > import_change:
                    movement_note = "Exports increased more than imports in the latest period, which may slightly ease trade-deficit pressure."
                else:
                    movement_note = "Imports and exports moved by a similar amount in the latest period."

                st.markdown(
                    f"""
                    <div class="insight-card">
                        <p>{movement_note}</p>
                        <p>
                        Latest cumulative imports changed by <b>Rs {import_change:,.2f} billion</b>,
                        while cumulative exports changed by <b>Rs {export_change:,.2f} billion</b>.
                        The trade deficit changed by <b>Rs {deficit_change:,.2f} billion</b>.
                        </p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )


                st.markdown("### Trend Insight Summary")

                st.markdown(
                    f"""
                    <div class="insight-card">
                        <h3>What changed in the latest period?</h3>
                        <p>{movement_note}</p>
                        <p>
                        Across <b>{len(trend_df)}</b> monthly customs files, the latest period is
                        <b>{latest['Period']}</b>. Cumulative imports reached
                        <b>Rs {latest['Imports_Billion']:,.2f} billion</b>, exports reached
                        <b>Rs {latest['Exports_Billion']:,.2f} billion</b>, and the trade deficit stood at
                        <b>Rs {latest['Trade_Deficit_Billion']:,.2f} billion</b>.
                        </p>
                        <p>
                        Compared with the previous cumulative file, imports changed by
                        <b>Rs {import_change:,.2f} billion</b>, exports changed by
                        <b>Rs {export_change:,.2f} billion</b>, and the trade deficit changed by
                        <b>Rs {deficit_change:,.2f} billion</b>.
                        </p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                st.markdown("### Trend Data Table")

                show_cols = [
                    "Order",
                    "Period",
                    "Imports_Billion",
                    "Exports_Billion",
                    "Trade_Deficit_Billion",
                    "Total_Trade_Billion",
                    "Monthly_Imports_Billion",
                    "Monthly_Exports_Billion",
                    "Monthly_Deficit_Billion",
                    "Monthly_Total_Trade_Billion",
                    "Import_Export_Ratio"
                ]

                st.dataframe(
                    trend_df[show_cols].style.format({
                        "Imports_Billion": "{:,.2f}",
                        "Exports_Billion": "{:,.2f}",
                        "Trade_Deficit_Billion": "{:,.2f}",
                        "Total_Trade_Billion": "{:,.2f}",
                        "Monthly_Imports_Billion": "{:,.2f}",
                        "Monthly_Exports_Billion": "{:,.2f}",
                        "Monthly_Deficit_Billion": "{:,.2f}",
                        "Monthly_Total_Trade_Billion": "{:,.2f}",
                        "Import_Export_Ratio": "{:,.1f}"
                    }),
                    use_container_width=True
                )

                st.download_button(
                    "Download trend data as CSV",
                    data=trend_df[show_cols].to_csv(index=False).encode("utf-8"),
                    file_name="tradepulse_multi_month_trends.csv",
                    mime="text/csv"
                )


                # -----------------------------
                # Product Movement Over Time
                # -----------------------------

                st.markdown("---")
                st.markdown("### Product Movement Over Time")

                st.markdown(
                    """
                    <div class="insight-card">
                        <b>How this works:</b> The monthly customs files are cumulative, so latest monthly
                        product movement is estimated by subtracting the previous cumulative file from the latest cumulative file.
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                try:
                    import_product_movement = build_product_movement_table(
                        trend_files=trend_files,
                        sheet_name="5_Imports_By_Commodity",
                        value_col="Imports_Value"
                    )

                    export_product_movement = build_product_movement_table(
                        trend_files=trend_files,
                        sheet_name="7_Exports_By_Commodity",
                        value_col="Exports_Value"
                    )

                    min_latest_value = st.number_input(
                        "Minimum latest monthly product value, Rs billion",
                        min_value=0.0,
                        value=0.1,
                        step=0.1,
                        key="product_movement_min_value"
                    )

                    top_product_n = st.slider(
                        "Number of products to show",
                        min_value=5,
                        max_value=30,
                        value=10,
                        step=5,
                        key="product_movement_top_n"
                    )

                    filtered_import_movement = import_product_movement[
                        import_product_movement["Latest_Month_Billion"] >= min_latest_value
                    ].copy()

                    filtered_export_movement = export_product_movement[
                        export_product_movement["Latest_Month_Billion"] >= min_latest_value
                    ].copy()

                    rising_imports = (
                        filtered_import_movement
                        .sort_values("Change_Billion", ascending=False)
                        .head(top_product_n)
                    )

                    falling_imports = (
                        filtered_import_movement
                        .sort_values("Change_Billion", ascending=True)
                        .head(top_product_n)
                    )

                    rising_exports = (
                        filtered_export_movement
                        .sort_values("Change_Billion", ascending=False)
                        .head(top_product_n)
                    )

                    falling_exports = (
                        filtered_export_movement
                        .sort_values("Change_Billion", ascending=True)
                        .head(top_product_n)
                    )

                    st.markdown("#### Top Product Movement Signals")

                    m1, m2, m3, m4 = st.columns(4)

                    with m1:
                        if len(rising_imports) > 0:
                            kpi_card(
                                "Top Rising Import",
                                short_text(rising_imports.iloc[0]["Description"], 22),
                                f"+Rs {rising_imports.iloc[0]['Change_Billion']:,.2f}B"
                            )

                    with m2:
                        if len(falling_imports) > 0:
                            kpi_card(
                                "Top Falling Import",
                                short_text(falling_imports.iloc[0]["Description"], 22),
                                f"Rs {falling_imports.iloc[0]['Change_Billion']:,.2f}B"
                            )

                    with m3:
                        if len(rising_exports) > 0:
                            kpi_card(
                                "Top Rising Export",
                                short_text(rising_exports.iloc[0]["Description"], 22),
                                f"+Rs {rising_exports.iloc[0]['Change_Billion']:,.2f}B"
                            )

                    with m4:
                        if len(falling_exports) > 0:
                            kpi_card(
                                "Top Falling Export",
                                short_text(falling_exports.iloc[0]["Description"], 22),
                                f"Rs {falling_exports.iloc[0]['Change_Billion']:,.2f}B"
                            )

                    st.markdown("### Rising and Falling Imports")

                    im1, im2 = st.columns(2)

                    with im1:
                        fig_rising_imports = horizontal_bar(
                            rising_imports,
                            "Change_Billion",
                            "Description",
                            "Top Rising Imports",
                            "Change vs Previous Month, Rs Billion",
                            "Product",
                            height=520
                        )
                        st.plotly_chart(fig_rising_imports, use_container_width=True)

                    with im2:
                        fig_falling_imports = horizontal_bar(
                            falling_imports,
                            "Change_Billion",
                            "Description",
                            "Top Falling Imports",
                            "Change vs Previous Month, Rs Billion",
                            "Product",
                            height=520
                        )
                        st.plotly_chart(fig_falling_imports, use_container_width=True)

                    st.markdown("### Rising and Falling Exports")

                    ex1, ex2 = st.columns(2)

                    with ex1:
                        fig_rising_exports = horizontal_bar(
                            rising_exports,
                            "Change_Billion",
                            "Description",
                            "Top Rising Exports",
                            "Change vs Previous Month, Rs Billion",
                            "Product",
                            height=520
                        )
                        st.plotly_chart(fig_rising_exports, use_container_width=True)

                    with ex2:
                        fig_falling_exports = horizontal_bar(
                            falling_exports,
                            "Change_Billion",
                            "Description",
                            "Top Falling Exports",
                            "Change vs Previous Month, Rs Billion",
                            "Product",
                            height=520
                        )
                        st.plotly_chart(fig_falling_exports, use_container_width=True)

                    st.markdown("### Product Movement Tables")

                    movement_cols = [
                        "HSCode",
                        "Description",
                        "Previous_Month_Billion",
                        "Latest_Month_Billion",
                        "Change_Billion",
                        "Growth_Percent",
                        "Latest_Cumulative_Billion"
                    ]

                    table1, table2 = st.columns(2)

                    with table1:
                        st.markdown("#### Import Movement Table")
                        st.dataframe(
                            rising_imports[movement_cols].style.format({
                                "Previous_Month_Billion": "{:,.2f}",
                                "Latest_Month_Billion": "{:,.2f}",
                                "Change_Billion": "{:,.2f}",
                                "Growth_Percent": "{:,.1f}",
                                "Latest_Cumulative_Billion": "{:,.2f}"
                            }),
                            use_container_width=True
                        )

                        st.download_button(
                            "Download import product movement",
                            data=import_product_movement.to_csv(index=False).encode("utf-8"),
                            file_name="tradepulse_import_product_movement.csv",
                            mime="text/csv"
                        )

                    with table2:
                        st.markdown("#### Export Movement Table")
                        st.dataframe(
                            rising_exports[movement_cols].style.format({
                                "Previous_Month_Billion": "{:,.2f}",
                                "Latest_Month_Billion": "{:,.2f}",
                                "Change_Billion": "{:,.2f}",
                                "Growth_Percent": "{:,.1f}",
                                "Latest_Cumulative_Billion": "{:,.2f}"
                            }),
                            use_container_width=True
                        )

                        st.download_button(
                            "Download export product movement",
                            data=export_product_movement.to_csv(index=False).encode("utf-8"),
                            file_name="tradepulse_export_product_movement.csv",
                            mime="text/csv"
                        )

                    st.markdown("### Product Movement Interpretation")

                    if len(rising_imports) > 0 and len(rising_exports) > 0:
                        top_rising_import_name = rising_imports.iloc[0]["Description"]
                        top_rising_import_change = rising_imports.iloc[0]["Change_Billion"]
                        top_rising_export_name = rising_exports.iloc[0]["Description"]
                        top_rising_export_change = rising_exports.iloc[0]["Change_Billion"]

                        st.markdown(
                            f"""
                            <div class="insight-card">
                                <p>
                                The strongest import increase in the latest month came from
                                <b>{top_rising_import_name}</b>, rising by around
                                <b>Rs {top_rising_import_change:,.2f} billion</b> compared with the previous monthly movement.
                                </p>
                                <p>
                                The strongest export increase came from
                                <b>{top_rising_export_name}</b>, rising by around
                                <b>Rs {top_rising_export_change:,.2f} billion</b>.
                                </p>
                                <p>
                                These movements can be used to identify emerging demand, trade shocks,
                                export momentum, import-substitution candidates, or products that deserve deeper investigation.
                                </p>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )


                    st.markdown("### Combined Trend Analyst Summary")

                    if len(rising_imports) > 0 and len(rising_exports) > 0:
                        top_rising_import_name = rising_imports.iloc[0]["Description"]
                        top_rising_import_change = rising_imports.iloc[0]["Change_Billion"]
                        top_rising_export_name = rising_exports.iloc[0]["Description"]
                        top_rising_export_change = rising_exports.iloc[0]["Change_Billion"]

                        if len(falling_imports) > 0:
                            top_falling_import_name = falling_imports.iloc[0]["Description"]
                            top_falling_import_change = falling_imports.iloc[0]["Change_Billion"]
                        else:
                            top_falling_import_name = "Not available"
                            top_falling_import_change = 0

                        if len(falling_exports) > 0:
                            top_falling_export_name = falling_exports.iloc[0]["Description"]
                            top_falling_export_change = falling_exports.iloc[0]["Change_Billion"]
                        else:
                            top_falling_export_name = "Not available"
                            top_falling_export_change = 0

                        st.markdown(
                            f"""
                            <div class="opportunity-card">
                                <h3>TradePulse Trend Read</h3>
                                <p>
                                The latest monthly movement shows where Nepal's trade flow is changing most sharply.
                                The strongest import rise came from <b>{top_rising_import_name}</b>
                                with a change of <b>Rs {top_rising_import_change:,.2f} billion</b>.
                                The strongest export rise came from <b>{top_rising_export_name}</b>
                                with a change of <b>Rs {top_rising_export_change:,.2f} billion</b>.
                                </p>
                                <p>
                                On the downside, the largest import slowdown was linked to
                                <b>{top_falling_import_name}</b> with a change of
                                <b>Rs {top_falling_import_change:,.2f} billion</b>, while the largest export slowdown was linked to
                                <b>{top_falling_export_name}</b> with a change of
                                <b>Rs {top_falling_export_change:,.2f} billion</b>.
                                </p>
                                <p>
                                These product movements can help identify emerging demand, seasonal shocks,
                                export momentum, import-substitution candidates, and possible media or research story angles.
                                </p>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )

                except Exception as e:
                    st.warning(f"Could not generate product movement analysis: {e}")

# -----------------------------
# About / Methodology tab
# -----------------------------

with about_tab:
    st.subheader("About TradePulse Nepal")

    st.markdown(
        """
        <div class="section-card">
            <h3>What is TradePulse Nepal?</h3>
            <p>
            TradePulse Nepal is a data dashboard that converts Nepal's Department of Customs trade data
            into simple market intelligence. It helps users understand imports, exports, trade deficit,
            product movement, country dependence, customs route concentration, business opportunities,
            and policy risks.
            </p>
            <p>
            The goal is to make public trade data easier to understand for students, researchers,
            journalists, businesses, policymakers, and analysts.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    m1, m2 = st.columns(2)

    with m1:
        st.markdown(
            """
            <div class="section-card">
                <h3>Data Source</h3>
                <p><b>Primary source:</b> Department of Customs, Government of Nepal.</p>
                <p><b>Dataset type:</b> Monthly foreign trade statistics.</p>
                <p><b>Coverage:</b> Imports, exports, partner countries, commodities, HS codes, customs offices, quantity, value, and revenue.</p>
                <p><b>Current dashboard period:</b> Based on the uploaded customs Excel workbook.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    with m2:
        st.markdown(
            """
            <div class="section-card">
                <h3>Units and Conversion</h3>
                <p>The original customs workbook reports many monetary values in <b>Rs. thousands</b>.</p>
                <p>This dashboard converts values into <b>Rs. billion</b> for easier reading.</p>
                <p><b>Formula used:</b></p>
                <p>Rs. billion = Source value ÷ 1,000,000</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("### Methodology")

    st.markdown(
        """
        <div class="section-card">
            <h3>How the Dashboard Works</h3>
            <ol>
                <li>The user uploads a Department of Customs Excel workbook.</li>
                <li>The dashboard reads trade, product, country, partner, export, import, and customs-office sheets.</li>
                <li>Values are cleaned and converted from Rs. thousands to Rs. billion.</li>
                <li>Total rows are removed so rankings show actual products, countries, and customs routes.</li>
                <li>The dashboard calculates key indicators such as total imports, exports, trade deficit, import-export ratio, country shares, product rankings, and route concentration.</li>
                <li>The Opportunity Finder ranks products using import value, customs revenue, and approximate duty signal.</li>
                <li>The Insights tab generates a simple trade brief using calculated dashboard numbers.</li>
            </ol>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("### Opportunity Finder Score")

    st.markdown(
        """
        <div class="opportunity-card">
            <p><b>The Opportunity Finder is a screening tool, not an investment recommendation.</b></p>
            <p>It ranks imported products using three signals:</p>
            <ul>
                <li><b>Market Size Score:</b> based on import value.</li>
                <li><b>Revenue Signal Score:</b> based on customs revenue.</li>
                <li><b>Duty Signal Score:</b> based on approximate duty/revenue rate.</li>
            </ul>
            <p>The score helps identify products that may deserve deeper research for import-substitution,
            sourcing, distribution, trade finance, or policy analysis.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("### Limitations")

    st.markdown(
        """
        <div class="risk-card">
            <ul>
                <li>The dashboard depends on the structure and accuracy of the uploaded customs workbook.</li>
                <li>The Opportunity Score does not prove profitability or feasibility.</li>
                <li>Import-substitution potential requires additional data on domestic production, demand, costs, technology, and regulation.</li>
                <li>AI-style insights are generated only from available dashboard numbers and should be verified before publication.</li>
                <li>This dashboard is for research, business intelligence, and policy discussion — not financial or investment advice.</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("### Suggested Citation")

    st.code(
        "TradePulse Nepal Dashboard. Based on monthly foreign trade statistics published by the Department of Customs, Government of Nepal.",
        language="text"
    )

    st.markdown("---")

    st.subheader("Developer")

    dev_col1, dev_col2 = st.columns([0.32, 0.68])

    with dev_col1:
        if developer_photo_path.exists():
            st.image(str(developer_photo_path), caption="Utsav Phuyal", use_container_width=True)
        else:
            st.info("Upload your photo as utsav.png in the same folder as app.py to show it here.")

    with dev_col2:
        st.markdown(
            """
            <div class="section-card">
                <h3>Utsav Phuyal</h3>
                <p><b>Developer & Researcher, TradePulse Nepal</b></p>

                <p>
                I am a business and economics graduate student interested in data analytics,
                economic research, trade intelligence, financial stability, and AI-powered
                public-data tools.
                </p>

                <p>
                I built TradePulse Nepal to make Nepal's Department of Customs data easier
                to understand through dashboards, product-level analysis, country intelligence,
                opportunity signals, and automated trade briefs.
                </p>

                <p>
                The goal is to turn raw public data into clear market insights, business signals,
                policy risks, and report-ready analysis.
                </p>

                <p>
                <b>Contact</b><br>
                Email: utsavkphuyal@gmail.com<br>
                LinkedIn: linkedin.com/in/utsav-phuyal<br>
                GitHub: github.com/utsavhatescoding
                </p>
            </div>
            """,
            unsafe_allow_html=True
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

    st.markdown("---")
    st.markdown("## Download Report")

    st.markdown(
        """
        <div class="insight-card">
            Download the current TradePulse Nepal brief as a simple text file or as a professionally formatted PDF report.
            The PDF includes the market snapshot, product signals, country/customs movement, policy risks, business opportunities,
            and multi-month trend summary where monthly data is available.
        </div>
        """,
        unsafe_allow_html=True
    )

    trend_summary_for_pdf = build_trend_summary_for_report(Path(__file__).parent / "monthly_data")

    pdf_buffer = create_pdf_report(
        current_col=current_col,
        imports_total=imports_total,
        exports_total=exports_total,
        deficit_total=deficit_total,
        total_trade=total_trade,
        import_export_ratio=import_export_ratio,
        top_import_product=top_import_product,
        top_import_value=top_import_value,
        top_export_product=top_export_product,
        top_export_value=top_export_value,
        top_import_country=top_import_country,
        top_export_country=top_export_country,
        top_customs_office=top_customs_office,
        top_customs_value=top_customs_value,
        trend_summary=trend_summary_for_pdf
    )

    d1, d2 = st.columns(2)

    with d1:
        st.download_button(
            label="Download TXT Brief",
            data=brief_text.encode("utf-8"),
            file_name="tradepulse_nepal_monthly_brief.txt",
            mime="text/plain",
            key="download_txt_brief_main"
        )

    with d2:
        st.download_button(
            label="Download PDF Brief",
            data=pdf_buffer,
            file_name="tradepulse_nepal_monthly_brief.pdf",
            mime="application/pdf",
            key="download_pdf_brief_main"
        )

    st.caption("PDF button is in this Download Report section inside the Insights tab.")
