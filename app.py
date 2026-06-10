from pathlib import Path
import base64
import html
import re
from difflib import SequenceMatcher
from sqlite3 import connect
from typing import Any

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from src.categorizer import (
    DEFAULT_CATEGORY,
    DEFAULT_PROJECT_NAME,
    apply_manual_correction_overrides,
    assign_project_names,
    categorize_transactions,
    ensure_rules_database,
    load_category_rules,
    load_manual_category_rules,
    load_manual_project_rules,
    load_project_rules,
    save_manual_corrections,
)
from src.cleaner import clean_bank_statement
from src.loader import load_excel_file
from src.rule_miner import (
    analyze_historical_file,
    apply_exact_historical_labels,
    load_historical_category_rules,
    load_historical_project_rules,
)


APP_ROOT = Path(__file__).parent
ASSETS_DIR = APP_ROOT / "assets"
LOGO_FILE = ASSETS_DIR / "logo.png"
LOGO_URL = "https://static.tildacdn.net/tild3430-3534-4232-b237-623139346565/YoungFolks-circle-42.png"
STYLES_FILE = APP_ROOT / "styles.css"
INPUT_FILE = APP_ROOT / "data" / "input" / "YF-Jan-For visual-f_demo.xlsx"
HISTORICAL_RULE_SOURCE_FILE = INPUT_FILE
CATEGORY_RULES_FILE = APP_ROOT / "config" / "category_rules.yaml"
PROJECT_RULES_FILE = APP_ROOT / "config" / "project_rules.yaml"
RULES_DATABASE_FILE = APP_ROOT / "database" / "rules.db"

TRANSACTION_COLUMNS = [
    "date",
    "transaction_type",
    "category",
    "project_name",
    "signed_amount",
    "description",
    "counterparty_raw",
]
EDITOR_COLUMNS = [
    "row_id",
    "transaction_id",
    "date",
    "transaction_type",
    "category",
    "project_name",
    "commentary",
    "signed_amount",
    "description",
    "counterparty_raw",
]
DISPLAY_NAMES = {
    "date": "Date",
    "transaction_type": "Type",
    "category": "Category",
    "project_name": "Project Name",
    "commentary": "Commentary",
    "signed_amount": "Signed Amount",
    "description": "Description",
    "counterparty_raw": "Counterparty",
    "transaction_id": "Transaction ID",
    "row_id": "Row ID",
}


st.set_page_config(page_title="NGO Finance Dashboard", layout="wide")

APP_CSS = """
:root {
    --yf-green: #0a8527;
    --yf-mint: #dff4d8;
    --yf-cream: #eeeedd;
    --yf-yellow: #ffd84d;
    --yf-pink: #ff6f91;
    --yf-blue: #55b6ff;
    --yf-ink: #171717;
    --yf-muted: #6b6b5f;
    --yf-card: rgba(255, 255, 255, 0.92);
    --yf-border: rgba(23, 23, 23, 0.10);
    --yf-shadow: 0 18px 45px rgba(28, 61, 31, 0.10);
}

html, body, [data-testid="stAppViewContainer"] {
    background:
        radial-gradient(circle at top left, rgba(255, 216, 77, 0.22), transparent 26rem),
        radial-gradient(circle at top right, rgba(85, 182, 255, 0.16), transparent 28rem),
        linear-gradient(180deg, #fffdf8 0%, #f8f8ee 48%, #ffffff 100%);
    color: var(--yf-ink);
}

[data-testid="stHeader"] {
    background: transparent;
}

.block-container {
    max-width: 1280px;
    padding-top: 2rem;
    padding-bottom: 4rem;
}

.yf-brand-header {
    position: relative;
    overflow: hidden;
    margin: 0 auto 1.4rem;
    padding: clamp(1.75rem, 4vw, 3.2rem) 1.25rem;
    border: 1px solid var(--yf-border);
    border-radius: 28px;
    background:
        linear-gradient(135deg, rgba(238, 238, 221, 0.90), rgba(255, 255, 255, 0.94)),
        linear-gradient(90deg, rgba(10, 133, 39, 0.12), rgba(255, 216, 77, 0.15));
    box-shadow: var(--yf-shadow);
    text-align: center;
}

.yf-brand-header:before,
.yf-brand-header:after {
    content: "";
    position: absolute;
    border-radius: 999px;
    opacity: 0.62;
    pointer-events: none;
}

.yf-brand-header:before {
    width: 11rem;
    height: 11rem;
    left: -4rem;
    top: -4rem;
    background: var(--yf-yellow);
}

.yf-brand-header:after {
    width: 9rem;
    height: 9rem;
    right: -3rem;
    bottom: -3rem;
    background: var(--yf-mint);
}

.yf-logo {
    position: relative;
    z-index: 1;
    display: block;
    width: clamp(82px, 12vw, 142px);
    max-width: 38%;
    height: auto;
    margin: 0 auto 1rem;
    filter: drop-shadow(0 10px 18px rgba(10, 133, 39, 0.16));
}

.yf-brand-kicker {
    position: relative;
    z-index: 1;
    margin: 0 0 0.25rem;
    color: var(--yf-green);
    font-size: 0.78rem;
    font-weight: 800;
    letter-spacing: 0.12em;
    text-transform: uppercase;
}

.yf-brand-title {
    position: relative;
    z-index: 1;
    margin: 0;
    color: var(--yf-ink);
    font-size: clamp(2.1rem, 5vw, 4.4rem);
    font-weight: 900;
    line-height: 0.96;
}

.yf-brand-subtitle {
    position: relative;
    z-index: 1;
    max-width: 760px;
    margin: 0.9rem auto 0;
    color: var(--yf-muted);
    font-size: clamp(1rem, 1.8vw, 1.22rem);
    line-height: 1.55;
}

.yf-source-pill {
    position: relative;
    z-index: 1;
    display: inline-flex;
    align-items: center;
    gap: 0.35rem;
    margin-top: 1.15rem;
    padding: 0.5rem 0.85rem;
    border: 1px solid rgba(10, 133, 39, 0.18);
    border-radius: 999px;
    background: rgba(255, 255, 255, 0.75);
    color: var(--yf-green);
    font-size: 0.82rem;
    font-weight: 700;
}

h1, h2, h3, [data-testid="stMarkdownContainer"] h1, [data-testid="stMarkdownContainer"] h2, [data-testid="stMarkdownContainer"] h3 {
    color: var(--yf-ink);
    letter-spacing: 0;
}

[data-testid="stMarkdownContainer"] h2,
[data-testid="stMarkdownContainer"] h3 {
    font-weight: 850;
}

.yf-section-spacer {
    height: 1.15rem;
}

div[data-testid="stMetric"] {
    min-height: 116px;
    padding: 1.05rem 1.05rem 0.95rem;
    border: 1px solid var(--yf-border);
    border-radius: 22px;
    background: var(--yf-card);
    box-shadow: 0 14px 30px rgba(25, 36, 22, 0.07);
    transition: transform 160ms ease, box-shadow 160ms ease, border-color 160ms ease;
}

div[data-testid="stMetric"]:hover {
    transform: translateY(-2px);
    border-color: rgba(10, 133, 39, 0.26);
    box-shadow: 0 18px 38px rgba(10, 133, 39, 0.13);
}

div[data-testid="stMetricLabel"] p {
    color: var(--yf-muted);
    font-weight: 800;
    letter-spacing: 0.02em;
}

div[data-testid="stMetricValue"] {
    color: var(--yf-ink);
    font-weight: 900;
}

.yf-kpi-card {
    min-height: 116px;
    padding: 1.05rem 1.05rem 0.95rem;
    border: 1px solid var(--yf-border);
    border-radius: 22px;
    background: var(--yf-card);
    box-shadow: 0 14px 30px rgba(25, 36, 22, 0.07);
    transition: transform 160ms ease, box-shadow 160ms ease, border-color 160ms ease;
}

.yf-kpi-card:hover {
    transform: translateY(-2px);
    border-color: rgba(10, 133, 39, 0.26);
    box-shadow: 0 18px 38px rgba(10, 133, 39, 0.13);
}

.yf-kpi-card span {
    display: block;
    min-height: 2.2rem;
    color: var(--yf-muted);
    font-size: 0.88rem;
    font-weight: 800;
    letter-spacing: 0.02em;
}

.yf-kpi-card strong {
    display: block;
    margin-top: 0.75rem;
    color: var(--yf-ink);
    font-size: clamp(1.35rem, 2.4vw, 2rem);
    font-weight: 900;
    line-height: 1.15;
}

.yf-kpi-card em {
    display: block;
    margin-top: 0.55rem;
    color: var(--yf-muted);
    font-size: 0.78rem;
    font-style: normal;
    line-height: 1.25;
}

.yf-kpi-positive strong {
    color: var(--yf-green);
}

.yf-kpi-negative strong {
    color: #d83a3a;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #ffffff 0%, var(--yf-cream) 100%);
    border-right: 1px solid rgba(10, 133, 39, 0.12);
}

[data-testid="stSidebar"] [data-testid="stMetric"] {
    min-height: 94px;
    border-radius: 18px;
    background: rgba(255, 255, 255, 0.82);
}

[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: var(--yf-green);
    font-weight: 900;
}

button[kind="primary"],
.stButton > button[kind="primary"] {
    border: 0;
    border-radius: 999px;
    background: var(--yf-green);
    color: #ffffff;
    font-weight: 800;
    box-shadow: 0 12px 26px rgba(10, 133, 39, 0.22);
}

.stButton > button {
    border-radius: 999px;
    border-color: rgba(10, 133, 39, 0.24);
    font-weight: 750;
}

.stButton > button:hover {
    border-color: var(--yf-green);
    color: var(--yf-green);
}

[data-baseweb="tab-list"] {
    gap: 0.65rem;
    padding: 0.35rem;
    border: 1px solid var(--yf-border);
    border-radius: 999px;
    background: rgba(255, 255, 255, 0.74);
    box-shadow: 0 10px 28px rgba(0, 0, 0, 0.04);
}

[data-baseweb="tab"] {
    min-height: 42px;
    padding: 0.4rem 1rem;
    border-radius: 999px;
    color: var(--yf-muted);
    font-weight: 800;
}

[data-baseweb="tab"][aria-selected="true"] {
    background: var(--yf-green);
    color: #ffffff;
}

[data-baseweb="tab-highlight"] {
    display: none;
}

[data-testid="stPlotlyChart"] {
    margin: 1rem 0 1.35rem;
    padding: clamp(0.75rem, 2vw, 1.35rem);
    border: 1px solid var(--yf-border);
    border-radius: 24px;
    background: rgba(255, 255, 255, 0.92);
    box-shadow: var(--yf-shadow);
}

[data-testid="stDataFrame"],
[data-testid="stDataEditor"] {
    padding: 0.35rem;
    border: 1px solid var(--yf-border);
    border-radius: 20px;
    background: rgba(255, 255, 255, 0.92);
    box-shadow: 0 12px 28px rgba(0, 0, 0, 0.045);
    overflow: hidden;
}

[data-testid="stExpander"] {
    border: 1px solid rgba(10, 133, 39, 0.14);
    border-radius: 18px;
    background: rgba(255, 255, 255, 0.78);
    box-shadow: 0 10px 24px rgba(0, 0, 0, 0.035);
}

[data-testid="stAlert"] {
    border-radius: 18px;
}

div[data-baseweb="select"] > div,
div[data-baseweb="base-input"] > div,
div[data-baseweb="input"] > div,
textarea,
input {
    border-radius: 14px !important;
}

.stDateInput, .stMultiSelect, .stTextInput, .stSelectbox, .stSlider {
    padding: 0.35rem 0;
}

hr {
    border-color: rgba(10, 133, 39, 0.14);
}

@media (max-width: 700px) {
    .block-container {
        padding-left: 1rem;
        padding-right: 1rem;
    }

    .yf-brand-header {
        border-radius: 22px;
    }

    [data-baseweb="tab-list"] {
        border-radius: 22px;
        align-items: stretch;
    }

    [data-baseweb="tab"] {
        padding-left: 0.65rem;
        padding-right: 0.65rem;
    }
}
"""


def load_custom_css() -> None:
    css = STYLES_FILE.read_text() if STYLES_FILE.exists() else APP_CSS
    css += """
.yf-kpi-grid {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 0.85rem;
}

@media (max-width: 1100px) {
    .yf-kpi-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }
}

@media (max-width: 700px) {
    .block-container {
        padding-left: 0.75rem;
        padding-right: 0.75rem;
    }

    [data-testid="stPlotlyChart"] {
        overflow-x: auto;
        -webkit-overflow-scrolling: touch;
        padding: 0.75rem;
        border-radius: 18px;
    }

    [data-testid="stPlotlyChart"] .js-plotly-plot,
    [data-testid="stPlotlyChart"] .plot-container,
    [data-testid="stPlotlyChart"] .svg-container {
        min-width: 760px !important;
    }

    [data-testid="stPlotlyChart"] .modebar-container {
        display: none !important;
    }

    .yf-kpi-card {
        min-height: auto;
        padding: 0.9rem;
    }

    .yf-kpi-card strong {
        font-size: 1.55rem;
    }

    .yf-kpi-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 0.65rem;
    }
}

@media (max-width: 420px) {
    .yf-kpi-grid {
        grid-template-columns: 1fr;
    }
}

html, body, [data-testid="stAppViewContainer"],
[data-testid="stMarkdownContainer"],
[data-testid="stWidgetLabel"],
[data-testid="stMetricLabel"],
[data-testid="stMetricValue"] {
    color: #111827 !important;
}

[data-testid="stMarkdownContainer"] p,
[data-testid="stCaptionContainer"],
label,
p,
span {
    color: #374151;
}

.yf-kpi-card span,
div[data-testid="stMetricLabel"] p {
    color: #374151 !important;
}

.yf-kpi-card strong,
div[data-testid="stMetricValue"] {
    color: #111827 !important;
}

.yf-kpi-positive strong {
    color: #0a8527 !important;
}

.yf-kpi-negative strong {
    color: #d83a3a !important;
}

div[data-baseweb="select"] > div {
    background-color: #ffffff !important;
    color: #111827 !important;
    border: 1px solid rgba(17, 24, 39, 0.18) !important;
}

div[data-baseweb="select"] span,
div[data-baseweb="select"] input,
div[data-baseweb="popover"] span,
div[data-baseweb="popover"] li {
    color: #111827 !important;
}

[data-baseweb="tag"] {
    background-color: #ef4444 !important;
    color: #ffffff !important;
}

[data-baseweb="tag"] span {
    color: #ffffff !important;
}

[data-testid="stSidebar"] * {
    color: #111827;
}

[data-testid="stSidebar"] div[data-testid="stMetricLabel"] p {
    color: #374151 !important;
}

[data-testid="stSidebar"] div[data-testid="stMetricValue"] {
    color: #111827 !important;
}

[data-testid="stDataFrame"],
[data-testid="stDataFrame"] div,
[data-testid="stDataFrame"] span,
[data-testid="stDataFrame"] p {
    color: #111827 !important;
}

[data-testid="stDataFrame"] {
    background: #ffffff !important;
}

[data-testid="stDataFrame"] [role="grid"],
[data-testid="stDataFrame"] [role="row"],
[data-testid="stDataFrame"] [role="gridcell"],
[data-testid="stDataFrame"] [role="columnheader"] {
    background-color: #ffffff !important;
    color: #111827 !important;
    border-color: rgba(17,24,39,0.10) !important;
}

[data-testid="stDataFrame"] [role="columnheader"] {
    background-color: #f3f4f6 !important;
    color: #111827 !important;
    font-weight: 700 !important;
}
"""
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def image_data_uri(path: Path) -> str:
    suffix = path.suffix.lower().lstrip(".") or "png"
    image_type = "jpeg" if suffix == "jpg" else suffix
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/{image_type};base64,{encoded}"


def render_brand_header() -> None:
    logo_src = image_data_uri(LOGO_FILE) if LOGO_FILE.exists() else LOGO_URL
    st.markdown(
        f"""
        <section class="yf-brand-header">
            <img src="{logo_src}" class="yf-logo" alt="Young Folks logo">
            <p class="yf-brand-kicker">Young Folks LV</p>
            <h1 class="yf-brand-title">NGO Finance Dashboard</h1>
            <p class="yf-brand-subtitle">
                A clear, community-minded view of income, expenses, projects, and classification quality.
            </p>
            <span class="yf-source-pill">Source file: {html.escape(INPUT_FILE.name)}</span>
        </section>
        """,
        unsafe_allow_html=True,
    )


def section_gap() -> None:
    st.markdown('<div class="yf-section-spacer"></div>', unsafe_allow_html=True)


@st.cache_data
def get_historical_analysis() -> dict:
    if not HISTORICAL_RULE_SOURCE_FILE.exists():
        return {}
    return analyze_historical_file(HISTORICAL_RULE_SOURCE_FILE)


def main() -> None:
    load_custom_css()
    render_brand_header()

    if not INPUT_FILE.exists():
        st.error(f"Input file not found: {INPUT_FILE}")
        return

    try:
        transactions, validation_issues, debug_summary = load_classified_transactions()
    except Exception as exc:
        st.error(f"Could not load the input file: {exc}")
        return

    render_sidebar(transactions)

    transactions_tab, manual_tab, import_qa_tab, dashboard_tab = st.tabs(
        ["Transactions", "Manual Review", "Import QA", "Dashboard"]
    )
    with transactions_tab:
        render_transactions_tab(transactions)
    with manual_tab:
        render_manual_review(transactions)
    with import_qa_tab:
        render_import_review(transactions, validation_issues, debug_summary)
    with dashboard_tab:
        render_dashboard(transactions)


def load_classified_transactions() -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    raw_transactions = load_excel_file(INPUT_FILE)
    transactions, validation_issues, debug_summary = clean_bank_statement(raw_transactions)
    transactions = apply_exact_historical_labels(transactions, HISTORICAL_RULE_SOURCE_FILE)

    category_rules = (
        load_manual_category_rules(RULES_DATABASE_FILE)
        + load_historical_category_rules(HISTORICAL_RULE_SOURCE_FILE)
        + load_category_rules(CATEGORY_RULES_FILE)
    )
    category_rules = apply_rule_overrides(category_rules)
    transactions = categorize_transactions(transactions, category_rules)
    transactions = apply_manual_correction_overrides(
        transactions,
        RULES_DATABASE_FILE,
        include_category=True,
        include_project=False,
    )

    project_rules = (
        load_manual_project_rules(RULES_DATABASE_FILE)
        + load_historical_project_rules(HISTORICAL_RULE_SOURCE_FILE)
        + load_project_rules(PROJECT_RULES_FILE)
    )
    project_rules = apply_rule_overrides(project_rules)
    transactions = assign_project_names(transactions, project_rules)
    transactions = apply_manual_correction_overrides(transactions, RULES_DATABASE_FILE)
    transactions = apply_commentary_overrides(transactions)
    return transactions, validation_issues, debug_summary


def render_sidebar(transactions: pd.DataFrame) -> None:
    st.sidebar.header("Summary")
    st.sidebar.metric("Total income", format_money(transactions["amount_income"].sum()))
    st.sidebar.metric("Total expenses", format_money(transactions["amount_expense"].sum()))
    st.sidebar.metric("Net cash flow", format_money(transactions["signed_amount"].sum()))
    st.sidebar.metric("Unclassified", f"{(transactions['category'] == DEFAULT_CATEGORY).sum():,}")


def render_transactions_tab(transactions: pd.DataFrame) -> None:
    st.subheader("All Transactions")
    filtered = filter_transactions(transactions, key_prefix="transactions")
    st.caption(f"Showing {len(filtered):,} of {len(transactions):,} rows.")
    st.dataframe(
        format_transactions(filtered, TRANSACTION_COLUMNS),
        use_container_width=True,
        hide_index=True,
        height=760,
    )


def render_import_review(transactions: pd.DataFrame, validation_issues: pd.DataFrame, debug_summary: dict) -> None:
    st.subheader("Summary Checks")
    min_date = transactions["date"].min()
    max_date = transactions["date"].max()
    income_rows = int((transactions["transaction_type"] == "Income").sum())
    expense_rows = int((transactions["transaction_type"] == "Expense").sum())

    summary_cols = st.columns(6)
    summary_cols[0].metric("Min date", format_date_value(min_date))
    summary_cols[1].metric("Max date", format_date_value(max_date))
    summary_cols[2].metric("Rows loaded", f"{len(transactions):,}")
    summary_cols[3].metric("Income rows", f"{income_rows:,}")
    summary_cols[4].metric("Expense rows", f"{expense_rows:,}")
    summary_cols[5].metric("Removed rows", f"{debug_summary['removed_rows_count']:,}")

    section_gap()
    st.subheader("Debug Summary")
    debug_cols = st.columns(4)
    debug_cols[0].metric("Original Excel row count", f"{debug_summary['original_excel_row_count']:,}")
    debug_cols[1].metric("Cleaned dataframe row count", f"{debug_summary['cleaned_dataframe_row_count']:,}")
    debug_cols[2].metric("Removed rows count", f"{debug_summary['removed_rows_count']:,}")
    debug_cols[3].write("")
    st.caption(f"Removal reasons: {debug_summary['removed_rows_reason']}")

    section_gap()
    render_historical_source_summary(transactions)
    section_gap()
    render_historical_match_check(transactions)

    with st.expander("Validation Issues", expanded=False):
        if validation_issues.empty:
            st.success("No validation issues found.")
        else:
            st.dataframe(validation_issues.fillna(""), use_container_width=True, hide_index=True)


def render_historical_source_summary(transactions: pd.DataFrame) -> None:
    st.subheader("Historical Source of Truth")
    if not HISTORICAL_RULE_SOURCE_FILE.exists():
        st.warning(f"Historical source file not found: {HISTORICAL_RULE_SOURCE_FILE}")
        return

    analysis = get_historical_analysis()
    if not analysis:
        st.warning("Historical source could not be analyzed.")
        return

    st.write(analysis.get("source_file_name", HISTORICAL_RULE_SOURCE_FILE.name))
    project_source = analysis.get("project_source_column", "")
    if project_source == "Division":
        st.caption("Project Name is read from the historical file column: Division.")
    elif project_source:
        st.caption(f"Project Name is read from the historical file column: {project_source}.")

    missing_columns = analysis.get("missing_columns", [])
    if missing_columns:
        st.warning(
            "Historical source is missing required taxonomy column(s): "
            + ", ".join(missing_columns)
            + ". Missing labels will fall back to Unclassified / Unknown."
        )

    summary = analysis.get("summary", {})
    source_cols = st.columns(4)
    source_cols[0].metric("Historical file rows", f"{summary.get('historical_rows', 0):,}")
    source_cols[1].metric("Rows with Category", f"{summary.get('category_filled_rows', 0):,}")
    source_cols[2].metric("Rows with Project / Division", f"{summary.get('project_filled_rows', 0):,}")
    source_cols[3].metric("Historical rules rebuilt", f"{summary.get('extracted_historical_rules', 0):,}")

    result_cols = st.columns(4)
    result_cols[0].metric("Unique Categories", f"{summary.get('unique_categories', 0):,}")
    result_cols[1].metric("Unique Project Names", f"{summary.get('unique_project_names', 0):,}")
    result_cols[2].metric("Unclassified after rules", f"{(transactions['category'] == DEFAULT_CATEGORY).sum():,}")
    result_cols[3].metric("Unknown projects after rules", f"{(transactions['project_name'] == DEFAULT_PROJECT_NAME).sum():,}")

    with st.expander("Official historical taxonomy", expanded=False):
        taxonomy_cols = st.columns(2)
        taxonomy_cols[0].dataframe(
            analysis.get("taxonomy_categories", pd.DataFrame()),
            use_container_width=True,
            hide_index=True,
        )
        taxonomy_cols[1].dataframe(
            analysis.get("taxonomy_projects", pd.DataFrame()),
            use_container_width=True,
            hide_index=True,
        )


def render_historical_match_check(transactions: pd.DataFrame) -> None:
    st.subheader("Historical Match Check")
    exact_matches = transactions.get("historical_exact_match", pd.Series(False, index=transactions.index)).fillna(False)
    category_applied = transactions.get("historical_category_applied", pd.Series(False, index=transactions.index)).fillna(False)
    project_applied = transactions.get("historical_project_applied", pd.Series(False, index=transactions.index)).fillna(False)

    match_cols = st.columns(5)
    match_cols[0].metric("Exact historical matches", f"{int(exact_matches.sum()):,}")
    match_cols[1].metric("Historical Category applied", f"{int(category_applied.sum()):,}")
    match_cols[2].metric("Historical Project / Division applied", f"{int(project_applied.sum()):,}")
    match_cols[3].metric("Still Unclassified", f"{(transactions['category'] == DEFAULT_CATEGORY).sum():,}")
    match_cols[4].metric("Still Unknown", f"{(transactions['project_name'] == DEFAULT_PROJECT_NAME).sum():,}")

    mismatch = build_historical_project_mismatch_table(transactions)
    with st.expander("Unknown projects with historical Project / Division available", expanded=not mismatch.empty):
        if mismatch.empty:
            st.success("No exact historical project/division misses found.")
        else:
            st.dataframe(mismatch, use_container_width=True, hide_index=True, height=360)


def build_historical_project_mismatch_table(transactions: pd.DataFrame) -> pd.DataFrame:
    required = {"historical_project_name", "historical_match_key", "historical_category"}
    if not required.issubset(transactions.columns):
        return pd.DataFrame(columns=historical_project_mismatch_columns())

    historical_project_filled = transactions["historical_project_name"].fillna("").astype(str).str.strip().ne("")
    current_unknown = transactions["project_name"].eq(DEFAULT_PROJECT_NAME)
    rows = transactions[current_unknown & historical_project_filled].copy()
    if rows.empty:
        return pd.DataFrame(columns=historical_project_mismatch_columns())

    output = pd.DataFrame(
        {
            "date": pd.to_datetime(rows["date"], errors="coerce").dt.strftime("%Y-%m-%d"),
            "description": rows["description"],
            "signed_amount": rows["signed_amount"],
            "current_category": rows["category"],
            "current_project_name": rows["project_name"],
            "historical_category": rows["historical_category"],
            "historical_project_name": rows["historical_project_name"],
            "match_key": rows["historical_match_key"],
        }
    )
    return output[historical_project_mismatch_columns()].fillna("")


def historical_project_mismatch_columns() -> list[str]:
    return [
        "date",
        "description",
        "signed_amount",
        "current_category",
        "current_project_name",
        "historical_category",
        "historical_project_name",
        "match_key",
    ]


def render_manual_review(transactions: pd.DataFrame) -> None:
    st.subheader("Review Filters")
    filtered = filter_transactions(
        transactions,
        key_prefix="manual",
        include_category=True,
        include_project=True,
        include_text=True,
    )

    section_gap()
    st.subheader("Manual Classification Editor")
    render_manual_editor(filtered, transactions)

    st.divider()
    render_active_rules_overview()


def render_manual_editor(filtered: pd.DataFrame, all_transactions: pd.DataFrame) -> None:
    if filtered.empty:
        st.info("No transactions match the current filters.")
        return

    editor_df = format_editor_rows(filtered)
    edited = st.data_editor(
        editor_df,
        use_container_width=True,
        hide_index=True,
        height=620,
        key="manual-review-editor",
        disabled=["row_id", "transaction_id", "date", "transaction_type", "signed_amount", "description", "counterparty_raw"],
        column_config={
            "category": st.column_config.TextColumn("Category", width="medium"),
            "project_name": st.column_config.TextColumn("Project Name", width="medium"),
            "commentary": st.column_config.TextColumn("Commentary", width="large"),
        },
    )

    corrections = extract_editor_corrections(edited, all_transactions)
    duplicate_report = build_duplicate_report(corrections, all_transactions)
    render_duplicate_warnings(duplicate_report)

    needs_confirmation = bool(
        duplicate_report["new_categories"]
        or duplicate_report["new_projects"]
        or duplicate_report["category_duplicates"]
        or duplicate_report["project_duplicates"]
    )
    confirmed = True
    if needs_confirmation:
        confirmed = st.checkbox("Are you sure? This may split analytics.", key="manual-review-confirm-split")

    if st.button("Save Changes", type="primary", disabled=corrections.empty or not confirmed):
        saved = save_manual_corrections(RULES_DATABASE_FILE, corrections)
        st.success(f"Saved {saved:,} manual change(s).")
        st.rerun()


def render_dashboard(transactions: pd.DataFrame) -> None:
    dashboard_df = prepare_dashboard_data(transactions)

    st.subheader("Dashboard Filters")
    filtered = filter_dashboard_data(dashboard_df)
    if filtered.empty:
        st.info("No transactions match the selected dashboard filters.")
        return

    render_management_kpis(filtered)

    section_gap()
    monthly_chart = build_monthly_management_chart(filtered)
    st.plotly_chart(monthly_chart, use_container_width=True)
    render_chart_source_check("Monthly Income vs Expenses", filtered.assign(chart_group="Monthly total"), "chart_group")

    section_gap()
    yf_rows = filtered.copy()
    yf_rows["dashboard_yf_area"] = yf_rows.apply(assign_yf_area, axis=1)
    yf_chart = build_yf_main_extra_chart(yf_rows)
    st.plotly_chart(yf_chart, use_container_width=True)
    render_chart_source_check("YF Main vs YF Extra Overview", yf_rows, "dashboard_yf_area")

    section_gap()
    core_rows = filtered.copy()
    core_rows["dashboard_core_operation"] = core_rows["dashboard_category"].apply(map_core_operation)
    core_rows = core_rows[core_rows["dashboard_core_operation"].notna()]
    core_chart = build_core_operations_chart(core_rows)
    if core_chart is None:
        st.info("No core operations data available for the current filters.")
    else:
        st.plotly_chart(core_chart, use_container_width=True)
        render_chart_source_check("Core Operations Overview", core_rows, "dashboard_core_operation")

    section_gap()
    render_focus_breakdown(
        filtered,
        title="Membership Breakdown",
        key_prefix="membership",
        row_filter=membership_rows,
        label_assigner=assign_membership_label,
    )

    section_gap()
    render_focus_breakdown(
        filtered,
        title="Services Breakdown",
        key_prefix="services",
        row_filter=services_rows,
        label_assigner=assign_service_label,
    )

    section_gap()
    render_focus_breakdown(
        filtered,
        title="Erasmus+ Breakdown",
        key_prefix="erasmus",
        row_filter=erasmus_rows,
        label_assigner=assign_erasmus_label,
    )

    section_gap()
    render_custom_division_breakdown(filtered)

    section_gap()
    expense_categories = categories_with_amount(filtered, "amount_expense")
    selected_expense_categories = st.multiselect(
        "Expense categories to show",
        expense_categories,
        default=expense_categories,
        key="dashboard-expense-composition-categories",
    )
    if selected_expense_categories:
        expense_rows = filtered[filtered["dashboard_category"].isin(selected_expense_categories)]
        expense_other_options = other_item_options(expense_rows, "dashboard_category", top_n=5)
        selected_expense_separate = st.multiselect(
            "Expense items to show separately",
            expense_other_options,
            default=[],
            key="dashboard-expense-other-items",
        )
        expense_rows = apply_other_grouping(
            expense_rows,
            "dashboard_category",
            "dashboard_category_display",
            top_n=5,
            selected_separate=selected_expense_separate,
        )
        expense_chart = build_expense_composition_chart(expense_rows)
        st.plotly_chart(expense_chart, use_container_width=True)
        render_other_items_table("Expense Composition", expense_rows, "dashboard_category_display")
        render_chart_source_check("Expense Composition", expense_rows, "dashboard_category_display")
    else:
        st.info("Select at least one expense category to show Expense Composition.")

    section_gap()
    income_categories = categories_with_amount(filtered, "amount_income")
    selected_income_categories = st.multiselect(
        "Income categories to show",
        income_categories,
        default=income_categories,
        key="dashboard-revenue-dependency-categories",
    )
    if selected_income_categories:
        income_rows = filtered[filtered["dashboard_category"].isin(selected_income_categories)]
        income_other_options = other_item_options(income_rows, "dashboard_category", top_n=5)
        selected_income_separate = st.multiselect(
            "Income items to show separately",
            income_other_options,
            default=[],
            key="dashboard-income-other-items",
        )
        income_rows = apply_other_grouping(
            income_rows,
            "dashboard_category",
            "dashboard_category_display",
            top_n=5,
            selected_separate=selected_income_separate,
        )
        revenue_dependency_chart = build_revenue_dependency_chart(income_rows)
        st.plotly_chart(revenue_dependency_chart, use_container_width=True)
        render_other_items_table("Revenue Dependency", income_rows, "dashboard_category_display")
        render_chart_source_check("Revenue Dependency", income_rows, "dashboard_category_display")
    else:
        st.info("Select at least one income category to show Revenue Dependency.")

    section_gap()
    with st.expander("Detailed Category / Transaction Drilldown", expanded=False):
        selected_category = st.selectbox(
            "Select category for detailed transactions",
            ["All categories"] + sorted(filtered["dashboard_category"].dropna().unique().tolist()),
            key="dashboard-category-drilldown",
        )
        drilldown = filtered if selected_category == "All categories" else filtered[filtered["dashboard_category"] == selected_category]
        render_drilldown_summary(drilldown)
        st.dataframe(
            format_transactions(drilldown, TRANSACTION_COLUMNS),
            use_container_width=True,
            hide_index=True,
            height=420,
        )


def style_chart(chart) -> None:
    chart.update_layout(
        colorway=["#0a8527", "#ff6f91", "#55b6ff", "#ffd84d", "#171717"],
        font={"family": "Inter, Arial, sans-serif", "color": "#111827", "size": 13},
        title={"font": {"size": 22, "color": "#111827"}, "x": 0.02},
        title_font={"color": "#111827"},
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        legend={
            "orientation": "h",
            "yanchor": "top",
            "y": -0.18,
            "xanchor": "center",
            "x": 0.5,
            "font": {"color": "#111827", "size": 13},
            "bgcolor": "rgba(255,255,255,0.92)",
            "bordercolor": "rgba(17,24,39,0.12)",
            "borderwidth": 1,
        },
        margin={"l": 64, "r": 32, "t": 88, "b": 72},
    )
    chart.update_xaxes(
        gridcolor="rgba(17,24,39,0.08)",
        zerolinecolor="rgba(17,24,39,0.18)",
        title_font={"color": "#374151", "size": 13},
        tickfont={"color": "#4B5563", "size": 12},
    )
    chart.update_yaxes(
        gridcolor="rgba(17,24,39,0.08)",
        zerolinecolor="rgba(17,24,39,0.18)",
        title_font={"color": "#374151", "size": 13},
        tickfont={"color": "#4B5563", "size": 12},
    )


def readable_legend(y: float = -0.18, title: str = "") -> dict[str, Any]:
    return {
        "title": {"text": title},
        "orientation": "h",
        "yanchor": "top",
        "y": y,
        "xanchor": "center",
        "x": 0.5,
        "font": {"color": "#111827", "size": 13},
        "bgcolor": "rgba(255,255,255,0.92)",
        "bordercolor": "rgba(17,24,39,0.12)",
        "borderwidth": 1,
    }


def prepare_dashboard_data(transactions: pd.DataFrame) -> pd.DataFrame:
    data = transactions.copy()
    data["date"] = pd.to_datetime(data["date"], errors="coerce")
    data["dashboard_month_sort"] = data["date"].dt.strftime("%Y-%m").fillna("No date")
    data["dashboard_month_label"] = data["date"].dt.strftime("%b").fillna("No date")
    data["amount_income"] = pd.to_numeric(data["amount_income"], errors="coerce").fillna(0)
    data["amount_expense"] = pd.to_numeric(data["amount_expense"], errors="coerce").fillna(0)
    data["signed_amount"] = pd.to_numeric(data["signed_amount"], errors="coerce").fillna(0)
    data["dashboard_category"] = clean_dimension_series(data.get("Category", data.get("category")), "Unclassified")
    data["original_category"] = data["dashboard_category"]
    if "category" in data:
        data["dashboard_category"] = data["dashboard_category"].where(
            data["dashboard_category"].ne("Unclassified"),
            clean_dimension_series(data["category"], "Unclassified"),
        )
        data["original_category"] = data["dashboard_category"]
    data["dashboard_division"] = clean_dimension_series(data.get("Division", data.get("project_name")), "Unknown division")
    data["original_division"] = data["dashboard_division"]
    if "project_name" in data:
        data["dashboard_division"] = data["dashboard_division"].where(
            data["dashboard_division"].ne("Unknown division"),
            clean_dimension_series(data["project_name"], "Unknown division"),
        )
        data["original_division"] = data["dashboard_division"]
    data["original_sub"] = clean_dimension_series(data.get("Sub", pd.Series("", index=data.index)), "")
    data["original_subdivision"] = clean_dimension_series(data.get("Subdivision", data.get("Sub", pd.Series("", index=data.index))), "")
    data["dashboard_subdivision"] = clean_dimension_series(data.get("Sub", pd.Series("", index=data.index)), "")
    data["dashboard_subdivision"] = data["dashboard_subdivision"].where(
        data["dashboard_subdivision"].ne(""),
        data["dashboard_category"],
    )
    return data


def clean_dimension_series(values: Any, fallback: str) -> pd.Series:
    if values is None:
        return pd.Series(dtype=object)
    series = values if isinstance(values, pd.Series) else pd.Series(values)
    cleaned = series.fillna("").astype(str).str.strip()
    return cleaned.replace("", fallback)


def filter_dashboard_data(transactions: pd.DataFrame) -> pd.DataFrame:
    filtered = transactions.copy()
    valid_dates = filtered["date"].dropna()

    filter_cols = st.columns(4)
    if valid_dates.empty:
        date_range = None
        filter_cols[0].info("No valid dates found.")
    else:
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()
        date_range = filter_cols[0].date_input(
            "Month / date range",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            key="dashboard-management-date-range",
        )

    categories = sorted(filtered["dashboard_category"].dropna().unique())
    divisions = sorted(filtered["dashboard_division"].dropna().unique())
    subdivisions = sorted(filtered["dashboard_subdivision"].dropna().unique())

    selected_categories = filter_cols[1].multiselect(
        "Category",
        categories,
        default=categories,
        key="dashboard-management-category",
    )
    selected_divisions = filter_cols[2].multiselect(
        "Division",
        divisions,
        default=divisions,
        key="dashboard-management-division",
    )
    selected_subdivisions = filter_cols[3].multiselect(
        "Subdivision",
        subdivisions,
        default=subdivisions,
        key="dashboard-management-subdivision",
    )

    if date_range and len(date_range) == 2:
        start_date = pd.to_datetime(date_range[0])
        end_date = pd.to_datetime(date_range[1])
        filtered = filtered[filtered["date"].between(start_date, end_date, inclusive="both") | filtered["date"].isna()]
    if selected_categories:
        filtered = filtered[filtered["dashboard_category"].isin(selected_categories)]
    if selected_divisions:
        filtered = filtered[filtered["dashboard_division"].isin(selected_divisions)]
    if selected_subdivisions:
        filtered = filtered[filtered["dashboard_subdivision"].isin(selected_subdivisions)]
    return filtered


def ordered_month_labels(df: pd.DataFrame) -> list[str]:
    months = (
        df[["dashboard_month_sort", "dashboard_month_label"]]
        .drop_duplicates()
        .sort_values("dashboard_month_sort")
    )
    labels = months["dashboard_month_label"].tolist()
    return labels or ["No date"]


def render_management_kpis(filtered: pd.DataFrame) -> None:
    summary = financial_summary(filtered, "dashboard_category")
    total_income = float(filtered["amount_income"].sum())
    total_expenses = float(filtered["amount_expense"].sum())
    net_result = total_income - total_expenses
    largest_revenue = top_label(summary, "Income", "No income")
    largest_expense = top_label(summary, "Expenses", "No expenses")
    largest_growth = largest_monthly_change(filtered, "amount_income", "No income growth")
    largest_cost_increase = largest_monthly_change(filtered, "amount_expense", "No cost increase")
    dependency_value, dependency_hint = grant_donation_dependency(filtered)

    cards = [
        kpi_card("Total Income", format_money(total_income), "positive"),
        kpi_card("Total Expenses", format_money(total_expenses), "negative"),
        kpi_card("Net Result", format_money(net_result), "positive" if net_result >= 0 else "negative"),
        kpi_card("Largest Revenue Category", largest_revenue, "neutral"),
        kpi_card("Largest Expense Category", largest_expense, "neutral"),
        kpi_card("Largest Growing Category", largest_growth, "positive"),
        kpi_card("Largest Cost Increase", largest_cost_increase, "negative"),
        kpi_card("Grant / Donation Dependency", dependency_value, "neutral", dependency_hint),
    ]
    st.markdown(f"<div class='yf-kpi-grid'>{''.join(cards)}</div>", unsafe_allow_html=True)


def kpi_card(label: str, value: str, tone: str, hint: str = "") -> str:
    hint_markup = f"<em>{html.escape(hint)}</em>" if hint else ""
    return (
        f'<div class="yf-kpi-card yf-kpi-{tone}">'
        f"<span>{html.escape(label)}</span>"
        f"<strong>{html.escape(value)}</strong>"
        f"{hint_markup}"
        "</div>"
    )


def top_label(summary: pd.DataFrame, column: str, fallback: str) -> str:
    rows = summary[summary[column].gt(0)].sort_values(column, ascending=False)
    if rows.empty:
        return fallback
    row = rows.iloc[0]
    return f"{row['label']} ({format_money(row[column])})"


def largest_monthly_change(filtered: pd.DataFrame, amount_column: str, fallback: str) -> str:
    months = (
        filtered[["dashboard_month_sort", "dashboard_month_label"]]
        .drop_duplicates()
        .sort_values("dashboard_month_sort")
    )
    if len(months) < 2:
        return "N/A"

    previous_month = months.iloc[-2]["dashboard_month_label"]
    latest_month = months.iloc[-1]["dashboard_month_label"]
    monthly = (
        filtered[filtered["dashboard_month_label"].isin([previous_month, latest_month])]
        .groupby(["dashboard_category", "dashboard_month_label"], dropna=False)[amount_column]
        .sum()
        .unstack(fill_value=0)
    )
    monthly["change"] = monthly.get(latest_month, 0) - monthly.get(previous_month, 0)
    positive = monthly[monthly["change"].gt(0)].sort_values("change", ascending=False)
    if positive.empty:
        return fallback
    category = positive.index[0]
    return f"{category} (+{format_money(float(positive.iloc[0]['change']))})"


def grant_donation_dependency(filtered: pd.DataFrame) -> tuple[str, str]:
    income_rows = filtered[filtered["amount_income"].gt(0)].copy()
    total_income = float(income_rows["amount_income"].sum())
    if total_income == 0:
        return "N/A", "No income in current filter."

    dependency_mask = income_rows["dashboard_category"].apply(is_grant_or_donation_label)
    if not dependency_mask.any():
        return "N/A", "Map grant/donation categories manually."

    dependency_income = float(income_rows.loc[dependency_mask, "amount_income"].sum())
    return f"{dependency_income / total_income:.1%}", ""


def is_grant_or_donation_label(value: Any) -> bool:
    normalized = re.sub(r"[^\w\s]", " ", str(value).strip().lower())
    normalized = re.sub(r"\s+", " ", normalized)
    keywords = {
        "grant",
        "grants",
        "donation",
        "donations",
        "ziedojums",
        "ziedojumi",
        "dotacija",
        "subsidy",
        "funding",
        "valsts kase",
        "esf",
        "erasmus",
    }
    return any(keyword in normalized for keyword in keywords)


def build_monthly_management_chart(filtered: pd.DataFrame):
    monthly = (
        filtered.groupby(["dashboard_month_sort", "dashboard_month_label"], dropna=False)[["amount_income", "amount_expense"]]
        .sum()
        .reset_index()
        .sort_values("dashboard_month_sort")
    )
    x_values = monthly["dashboard_month_label"].fillna("No date")

    chart = go.Figure()
    chart.add_bar(name="Income", x=x_values, y=monthly["amount_income"], marker_color="#0a8527")
    chart.add_bar(name="Expenses", x=x_values, y=monthly["amount_expense"], marker_color="#ff6f91")
    chart.add_scatter(
        name="Income trend",
        x=x_values,
        y=monthly["amount_income"],
        mode="lines+markers",
        line={"color": "#0a8527", "width": 3, "dash": "dot"},
    )
    chart.add_scatter(
        name="Expenses trend",
        x=x_values,
        y=monthly["amount_expense"],
        mode="lines+markers",
        line={"color": "#ff6f91", "width": 3, "dash": "dot"},
    )
    chart.update_layout(
        barmode="group",
        title="Monthly Income vs Expenses",
        xaxis_title="Month",
        yaxis_title="Amount",
    )
    chart.update_yaxes(tickformat=",.2f")
    style_chart(chart)
    return chart


def build_yf_main_extra_chart(filtered: pd.DataFrame):
    chart_rows = filtered.copy()
    chart_rows["dashboard_yf_area"] = chart_rows.apply(assign_yf_area, axis=1)
    return build_income_expense_grouped_chart(
        chart_rows,
        "dashboard_yf_area",
        "YF Main vs YF Extra Overview",
        group_title="Area",
    )


def build_core_operations_chart(filtered: pd.DataFrame):
    chart_rows = filtered.copy()
    chart_rows["dashboard_core_operation"] = chart_rows["dashboard_category"].apply(map_core_operation)
    chart_rows = chart_rows[chart_rows["dashboard_core_operation"].notna()]
    if chart_rows.empty:
        return None
    return build_income_expense_grouped_chart(
        chart_rows,
        "dashboard_core_operation",
        "Core Operations Overview",
        group_title="Core operation",
    )


def render_focus_breakdown(
    filtered: pd.DataFrame,
    title: str,
    key_prefix: str,
    row_filter: Any,
    label_assigner: Any,
) -> None:
    st.subheader(title)
    rows = row_filter(filtered).copy()
    if rows.empty:
        st.info(f"No data available for {title}.")
        return

    rows["dashboard_focus_label"] = rows.apply(label_assigner, axis=1)
    rows = rows[rows["dashboard_focus_label"].notna()]
    if rows.empty:
        st.info(f"No matching breakdown items found for {title}.")
        return

    rows = filter_section_month_view(rows, key_prefix)
    if rows.empty:
        st.info(f"No data available for the selected month in {title}.")
        return

    display_column = "dashboard_focus_label"
    if key_prefix in {"services", "erasmus"}:
        other_options = other_item_options(rows, "dashboard_focus_label", top_n=5)
        selected_separate = st.multiselect(
            "Items to show separately",
            other_options,
            default=[],
            key=f"{key_prefix}-items-show-separately",
        )
        rows = apply_other_grouping(
            rows,
            "dashboard_focus_label",
            "dashboard_focus_display",
            top_n=5,
            selected_separate=selected_separate,
        )
        display_column = "dashboard_focus_display"

    chart = build_income_expense_grouped_chart(
        rows,
        display_column,
        title,
        group_title="Item",
    )
    st.plotly_chart(chart, use_container_width=True)
    render_other_items_table(title, rows, display_column)
    render_chart_source_check(title, rows, display_column)


def filter_section_month_view(filtered: pd.DataFrame, key_prefix: str) -> pd.DataFrame:
    month_options = ordered_month_labels(filtered)
    if not month_options:
        return filtered
    latest_month = month_options[-1]
    controls = st.columns([1, 2])
    mode = controls[0].radio(
        "Month view",
        ["Specific Month", "All Months"],
        horizontal=True,
        key=f"{key_prefix}-month-view-mode",
    )
    if mode == "All Months":
        controls[1].caption("Showing all months.")
        return filtered
    selected_month = controls[1].selectbox(
        "Select Month",
        month_options,
        index=month_options.index(latest_month),
        key=f"{key_prefix}-month-view-selected",
    )
    return filtered[filtered["dashboard_month_label"].eq(selected_month)]


def build_income_expense_grouped_chart(
    filtered: pd.DataFrame,
    group_column: str,
    title: str,
    group_title: str,
):
    months = ordered_month_labels(filtered)
    summary = financial_summary(filtered, group_column).sort_values("Total Volume", ascending=False)
    group_order = summary["label"].tolist()
    chart_data = (
        filtered.groupby([group_column, "dashboard_month_label"], dropna=False)[["amount_income", "amount_expense"]]
        .sum()
        .reset_index()
        .rename(columns={"amount_income": "Income", "amount_expense": "Expenses"})
    )

    bar_width = 0.32
    metric_gap = 0.08
    month_gap = 0.26
    group_gap = 1.0
    positioned_rows = []
    group_centers = []
    group_ranges = []
    month_centers = []
    x_cursor = 0.0

    for group in group_order:
        group_positions = []
        for month in months:
            row = chart_data[chart_data[group_column].eq(group) & chart_data["dashboard_month_label"].eq(month)]
            income = float(row["Income"].sum()) if not row.empty else 0.0
            expenses = float(row["Expenses"].sum()) if not row.empty else 0.0
            income_x = x_cursor
            expense_x = x_cursor + bar_width + metric_gap
            positioned_rows.extend(
                [
                    {"x": income_x, "Amount": income, "Group": group, "Month": month, "Metric": "Income", "Color": "#1f7a3d"},
                    {"x": expense_x, "Amount": expenses, "Group": group, "Month": month, "Metric": "Expenses", "Color": "#c84d62"},
                ]
            )
            group_positions.extend([income_x, expense_x])
            month_centers.append({"x": (income_x + expense_x) / 2, "Month": month})
            x_cursor += (bar_width * 2) + metric_gap + month_gap
        if group_positions:
            group_centers.append((sum(group_positions) / len(group_positions), group))
            group_ranges.append({"start": min(group_positions) - (bar_width / 2), "end": max(group_positions) + (bar_width / 2)})
        x_cursor += group_gap

    chart = go.Figure()
    for row in positioned_rows:
        chart.add_bar(
            x=[row["x"]],
            y=[row["Amount"]],
            width=bar_width,
            marker_color=row["Color"],
            showlegend=False,
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Month: %{customdata[1]}<br>"
                "Metric: %{customdata[2]}<br>"
                "Amount: %{y:,.2f}<extra></extra>"
            ),
            customdata=[[row["Group"], row["Month"], row["Metric"]]],
        )
    add_metric_legend(chart)
    chart.update_layout(
        title=title,
        barmode="overlay",
        bargap=0,
        height=max(520, 420 + (len(group_order) * 18)),
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        font={"family": "Inter, Arial, sans-serif", "color": "#20232a"},
        title_font={"size": 20, "color": "#151515"},
        margin={"l": 34, "r": 24, "t": 86, "b": 300},
        legend=readable_legend(y=-0.52),
        annotations=[
            {
                "text": f"{group_title} shown above month labels.",
                "xref": "paper",
                "yref": "paper",
                "x": 0,
                "y": 1.08,
                "showarrow": False,
                "align": "left",
                "font": {"size": 12, "color": "#60646c"},
            }
        ],
    )
    add_multilevel_x_axis_labels(chart, group_centers, month_centers, group_ranges, group_label_max_chars=16)
    chart.update_yaxes(
        title="Amount",
        tickformat=",.2f",
        gridcolor="rgba(28, 31, 35, 0.08)",
        zerolinecolor="rgba(28, 31, 35, 0.18)",
    )
    chart.update_xaxes(title="", showticklabels=False, showgrid=False)
    return chart


def add_metric_legend(chart: go.Figure) -> None:
    chart.add_scatter(
        x=[None],
        y=[None],
        mode="markers",
        name="Income",
        marker={"color": "#1f7a3d", "size": 11, "symbol": "square"},
        showlegend=True,
    )
    chart.add_scatter(
        x=[None],
        y=[None],
        mode="markers",
        name="Expenses",
        marker={"color": "#c84d62", "size": 11, "symbol": "square"},
        showlegend=True,
    )


def assign_yf_area(row: pd.Series) -> str:
    text = dimension_text(row)
    if "yf extra" in text:
        return "YF Extra"
    if "yf main" in text:
        return "YF Main"
    extra_keywords = ("erasmus", "project", "travel", "exchange", "x-change", "nva", "esf")
    return "YF Extra" if any(keyword in text for keyword in extra_keywords) else "YF Main"


def map_core_operation(value: str) -> str | None:
    normalized = normalize_label(value)
    if "membership" in normalized:
        return "Membership"
    if "service" in normalized:
        return "Services"
    if "donation" in normalized:
        return "Donations"
    if "operational" in normalized or "salar" in normalized:
        return "Operational Expenses"
    return None


def membership_rows(filtered: pd.DataFrame) -> pd.DataFrame:
    return filtered[filtered.apply(lambda row: "membership" in dimension_text(row) or assign_membership_label(row) is not None, axis=1)]


def services_rows(filtered: pd.DataFrame) -> pd.DataFrame:
    return filtered[filtered.apply(lambda row: assign_service_label(row) is not None, axis=1)]


def erasmus_rows(filtered: pd.DataFrame) -> pd.DataFrame:
    return filtered[filtered.apply(lambda row: "erasmus" in dimension_text(row), axis=1)]


def assign_membership_label(row: pd.Series) -> str | None:
    text = dimension_text(row)
    membership_items = {
        "Forever Young": ("forever young",),
        "YF Kids": ("yf kids", "kids"),
        "YF Teens": ("yf teens", "teens"),
        "YF Youth": ("yf youth", "youth"),
    }
    return first_matching_label(text, membership_items)


def assign_service_label(row: pd.Series) -> str | None:
    text = dimension_text(row)
    service_items = {
        "English": ("english",),
        "German": ("german",),
        "Latvian": ("latvian",),
        "Academic Drawing": ("academic drawing", "drawing"),
        "Workshops": ("workshop", "workshops"),
    }
    explicit_label = first_matching_label(text, service_items)
    if explicit_label:
        return explicit_label
    return "Other services" if "service" in text else None


def assign_erasmus_label(row: pd.Series) -> str | None:
    division = str(row.get("dashboard_division", "")).strip()
    subdivision = str(row.get("dashboard_subdivision", "")).strip()
    if division and normalize_label(division) not in {"unknown", "unknown division", "unclassified"}:
        return division
    if subdivision and normalize_label(subdivision) not in {"unknown", "no subdivision", "unclassified"}:
        return subdivision
    return "Erasmus+"


def first_matching_label(text: str, mapping: dict[str, tuple[str, ...]]) -> str | None:
    for label, keywords in mapping.items():
        if any(keyword in text for keyword in keywords):
            return label
    return None


def dimension_text(row: pd.Series) -> str:
    values = [
        row.get("dashboard_category", ""),
        row.get("dashboard_division", ""),
        row.get("dashboard_subdivision", ""),
        row.get("category", ""),
        row.get("project_name", ""),
        row.get("description", ""),
    ]
    text = " ".join(str(value).lower() for value in values if pd.notna(value))
    return re.sub(r"\s+", " ", text)


def categories_with_amount(filtered: pd.DataFrame, amount_column: str) -> list[str]:
    grouped = filtered.groupby("dashboard_category", dropna=False)[amount_column].sum()
    return sorted(grouped[grouped.gt(0)].index.tolist())


def apply_other_grouping(
    rows: pd.DataFrame,
    item_column: str,
    display_column: str,
    amount_columns: tuple[str, str] = ("amount_income", "amount_expense"),
    top_n: int = 5,
    selected_separate: list[str] | None = None,
) -> pd.DataFrame:
    selected_separate = selected_separate or []
    data = rows.copy()
    volume = (
        data.groupby(item_column, dropna=False)[list(amount_columns)]
        .sum()
        .abs()
        .sum(axis=1)
        .sort_values(ascending=False)
    )
    default_items = set(volume.head(top_n).index.tolist())
    visible_items = default_items | set(selected_separate)
    data[display_column] = data[item_column].where(data[item_column].isin(visible_items), "Other")
    return data


def other_item_options(rows: pd.DataFrame, item_column: str, top_n: int = 5) -> list[str]:
    volume = (
        rows.groupby(item_column, dropna=False)[["amount_income", "amount_expense"]]
        .sum()
        .abs()
        .sum(axis=1)
        .sort_values(ascending=False)
    )
    return volume.iloc[top_n:].index.tolist()


def render_other_items_table(title: str, rows: pd.DataFrame, display_column: str) -> None:
    if display_column not in rows or "Other" not in set(rows[display_column].dropna()):
        return
    other_rows = rows[rows[display_column].eq("Other")].copy()
    if other_rows.empty:
        return
    table = source_trace_table(other_rows)
    with st.expander(f"{title} — Items included in Other", expanded=False):
        st.dataframe(table, use_container_width=True, hide_index=True)


def render_chart_source_check(title: str, rows: pd.DataFrame, group_column: str) -> None:
    if rows.empty:
        with st.expander(f"{title} — Source Check", expanded=False):
            st.info("No source rows for this chart.")
        return

    table = validation_group_table(rows, group_column)
    chart_income = float(table["Income"].sum())
    chart_expenses = float(table["Expenses"].sum())
    source_income = float(rows["amount_income"].sum())
    source_expenses = float(rows["amount_expense"].sum())
    diff_income = chart_income - source_income
    diff_expenses = chart_expenses - source_expenses
    if abs(diff_income) > 0.01 or abs(diff_expenses) > 0.01:
        st.warning(
            "Chart totals do not match source data. "
            f"Income difference: {format_money(diff_income)}; "
            f"Expenses difference: {format_money(diff_expenses)}."
        )

    with st.expander(f"{title} — Source Check", expanded=False):
        summary = pd.DataFrame(
            [
                {
                    "Metric": "Income",
                    "Chart total": chart_income,
                    "Source filtered total": source_income,
                    "Difference": diff_income,
                },
                {
                    "Metric": "Expenses",
                    "Chart total": chart_expenses,
                    "Source filtered total": source_expenses,
                    "Difference": diff_expenses,
                },
            ]
        )
        st.dataframe(summary, use_container_width=True, hide_index=True)
        st.dataframe(table, use_container_width=True, hide_index=True)


def validation_group_table(rows: pd.DataFrame, group_column: str) -> pd.DataFrame:
    data = rows.copy()
    data["Group shown on chart"] = data[group_column] if group_column in data else "All"
    return (
        data.groupby(["dashboard_month_label", "Group shown on chart"], dropna=False)
        .agg(
            Income=("amount_income", "sum"),
            Expenses=("amount_expense", "sum"),
            **{"Number of transactions": ("amount_income", "size"), "Source rows count": ("amount_income", "size")},
        )
        .reset_index()
        .rename(columns={"dashboard_month_label": "Month"})
    )


def source_trace_table(rows: pd.DataFrame) -> pd.DataFrame:
    source_cols = ["original_category", "original_division", "original_sub", "original_subdivision"]
    data = rows.copy()
    for column in source_cols:
        if column not in data:
            data[column] = ""
    return (
        data.groupby(["dashboard_month_label", *source_cols], dropna=False)
        .agg(
            Income=("amount_income", "sum"),
            Expenses=("amount_expense", "sum"),
            **{"Number of transactions": ("amount_income", "size")},
        )
        .reset_index()
        .rename(
            columns={
                "dashboard_month_label": "Month",
                "original_category": "Original Category",
                "original_division": "Original Division",
                "original_sub": "Original Sub",
                "original_subdivision": "Original Subdivision",
            }
        )
    )


def build_category_division_breakdown_by_month_chart(filtered: pd.DataFrame, category: str):
    category_rows = filtered[filtered["dashboard_category"].eq(category)].copy()
    if not has_specific_division_data(category_rows):
        return None

    months = ordered_month_labels(category_rows)
    summary = financial_summary(category_rows, "dashboard_division").sort_values("Total Volume", ascending=False).head(12)
    division_order = summary["label"].tolist()
    chart_data = (
        category_rows.groupby(["dashboard_division", "dashboard_month_label"], dropna=False)[
            ["amount_income", "amount_expense"]
        ]
        .sum()
        .reset_index()
        .rename(columns={"amount_income": "Income", "amount_expense": "Expenses"})
    )

    income_palette = ["#14532d", "#2f7d45", "#68a878", "#a7d0ad", "#d5ead7"]
    expense_palette = ["#9f2f45", "#c84d62", "#e18491", "#efb3bc", "#f7d8dd"]
    bar_width = 0.30
    metric_gap = 0.08
    month_gap = 0.26
    division_gap = 0.95

    positioned_rows = []
    division_centers = []
    division_ranges = []
    month_centers = []
    x_cursor = 0.0

    for division in division_order:
        division_positions = []
        for month_index, month in enumerate(months):
            row = chart_data[
                chart_data["dashboard_division"].eq(division)
                & chart_data["dashboard_month_label"].eq(month)
            ]
            income = float(row["Income"].sum()) if not row.empty else 0.0
            expense = float(row["Expenses"].sum()) if not row.empty else 0.0
            income_x = x_cursor
            expense_x = x_cursor + bar_width + metric_gap
            month_center = (income_x + expense_x) / 2

            positioned_rows.extend(
                [
                    {
                        "x": income_x,
                        "Amount": income,
                        "Division": division,
                        "Month": month,
                        "Metric": "Income",
                        "Color": income_palette[month_index % len(income_palette)],
                    },
                    {
                        "x": expense_x,
                        "Amount": expense,
                        "Division": division,
                        "Month": month,
                        "Metric": "Expenses",
                        "Color": expense_palette[month_index % len(expense_palette)],
                    },
                ]
            )
            division_positions.extend([income_x, expense_x])
            month_centers.append({"x": month_center, "Month": month})
            x_cursor += (bar_width * 2) + metric_gap + month_gap

        if division_positions:
            division_centers.append((sum(division_positions) / len(division_positions), division))
            division_ranges.append(
                {
                    "start": min(division_positions) - (bar_width / 2),
                    "end": max(division_positions) + (bar_width / 2),
                }
            )
        x_cursor += division_gap

    chart = go.Figure()
    for row in positioned_rows:
        chart.add_bar(
            x=[row["x"]],
            y=[row["Amount"]],
            width=bar_width,
            marker_color=row["Color"],
            showlegend=False,
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Month: %{customdata[1]}<br>"
                "Metric: %{customdata[2]}<br>"
                "Amount: %{y:,.2f}<extra></extra>"
            ),
            customdata=[[row["Division"], row["Month"], row["Metric"]]],
        )

    for month_index, month in enumerate(months):
        chart.add_scatter(
            x=[None],
            y=[None],
            mode="markers",
            name=month,
            marker={
                "color": income_palette[month_index % len(income_palette)],
                "size": 10,
                "symbol": "square",
            },
            showlegend=True,
        )

    chart.update_layout(
        title=f"Division Breakdown — {category}",
        barmode="overlay",
        bargap=0,
        height=560,
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        font={"family": "Inter, Arial, sans-serif", "color": "#20232a"},
        title_font={"size": 20, "color": "#151515"},
        margin={"l": 34, "r": 24, "t": 92, "b": 330},
        legend=readable_legend(y=-0.60, title="Month"),
        annotations=[
            {
                "text": "Green bars show income. Red bars show expenses.",
                "xref": "paper",
                "yref": "paper",
                "x": 0,
                "y": 1.09,
                "showarrow": False,
                "align": "left",
                "font": {"size": 12, "color": "#60646c"},
            }
        ],
    )
    add_multilevel_x_axis_labels(
        chart,
        division_centers,
        month_centers,
        division_ranges,
        group_label_max_chars=16,
    )
    chart.update_yaxes(
        title="Amount",
        tickformat=",.2f",
        gridcolor="rgba(28, 31, 35, 0.08)",
        zerolinecolor="rgba(28, 31, 35, 0.18)",
    )
    chart.update_xaxes(
        title="",
        showticklabels=False,
        showgrid=False,
    )
    return chart


def render_custom_division_breakdown(filtered: pd.DataFrame) -> None:
    st.subheader("Custom Division Breakdown")
    categories = sorted(filtered["dashboard_category"].dropna().unique().tolist())
    selected_category = st.selectbox(
        "Choose category",
        ["Select a category"] + categories,
        key="dashboard-custom-division-category",
    )
    if selected_category == "Select a category":
        st.info("Select a category to view division breakdown.")
        return

    category_rows = filtered[filtered["dashboard_category"].eq(selected_category)]
    division_options = sorted(category_rows["dashboard_division"].dropna().unique().tolist())
    selected_divisions = st.multiselect(
        "Divisions to show in custom breakdown",
        division_options,
        default=division_options,
        key=f"dashboard-custom-division-filter-{normalize_label(selected_category)}",
    )
    if not selected_divisions:
        st.info("Select at least one division to show custom breakdown.")
        return

    chart = build_category_division_breakdown_by_month_chart(
        category_rows[category_rows["dashboard_division"].isin(selected_divisions)],
        selected_category,
    )
    if chart is None:
        st.info(f"No division data available for this category: {selected_category}.")
        return
    st.plotly_chart(chart, use_container_width=True)
    render_chart_source_check("Custom Division Breakdown", category_rows[category_rows["dashboard_division"].isin(selected_divisions)], "dashboard_division")


def has_specific_division_data(rows: pd.DataFrame) -> bool:
    if "Division" not in rows.columns:
        return False
    cleaned = rows["Division"].fillna("").astype(str).str.strip()
    return cleaned.ne("").any()


def build_expense_composition_chart(filtered: pd.DataFrame):
    category_column = "dashboard_category_display" if "dashboard_category_display" in filtered else "dashboard_category"
    expenses = (
        filtered.groupby(["dashboard_month_sort", "dashboard_month_label", category_column], dropna=False)["amount_expense"]
        .sum()
        .reset_index(name="Expenses")
        .rename(columns={category_column: "dashboard_category"})
    )
    if "dashboard_category_display" not in filtered:
        expenses = collapse_to_top_categories(expenses, "dashboard_category", "Expenses", top_n=5)
    if expenses.empty:
        expenses = pd.DataFrame({
            "dashboard_month_sort": ["No date"],
            "dashboard_month_label": ["No date"],
            "dashboard_category": ["No expenses"],
            "Expenses": [0],
        })
    expenses["Month Total"] = expenses.groupby("dashboard_month_label")["Expenses"].transform("sum")
    expenses["Share"] = expenses.apply(
        lambda row: row["Expenses"] / row["Month Total"] if row["Month Total"] else 0,
        axis=1,
    )
    return build_stacked_percentage_chart(
        expenses,
        title="Expense Composition by Month",
        category_title="Expense Category",
        amount_column="Expenses",
        amount_title="Expenses",
        y_axis_title="Share of monthly expenses",
        month_order=ordered_month_labels(filtered),
    )


def build_revenue_dependency_chart(filtered: pd.DataFrame):
    category_column = "dashboard_category_display" if "dashboard_category_display" in filtered else "dashboard_category"
    income = (
        filtered.groupby(["dashboard_month_sort", "dashboard_month_label", category_column], dropna=False)["amount_income"]
        .sum()
        .reset_index(name="Income")
        .rename(columns={category_column: "dashboard_category"})
    )
    if "dashboard_category_display" not in filtered:
        income = collapse_to_top_categories(income, "dashboard_category", "Income", top_n=5)
    if income.empty:
        income = pd.DataFrame({
            "dashboard_month_sort": ["No date"],
            "dashboard_month_label": ["No date"],
            "dashboard_category": ["No income"],
            "Income": [0],
        })
    income["Month Total"] = income.groupby("dashboard_month_label")["Income"].transform("sum")
    income["Share"] = income.apply(
        lambda row: row["Income"] / row["Month Total"] if row["Month Total"] else 0,
        axis=1,
    )
    return build_stacked_percentage_chart(
        income,
        title="Revenue Dependency by Month",
        category_title="Income Category",
        amount_column="Income",
        amount_title="Income",
        y_axis_title="Share of monthly income",
        month_order=ordered_month_labels(filtered),
    )


def build_stacked_percentage_chart(
    data: pd.DataFrame,
    title: str,
    category_title: str,
    amount_column: str,
    amount_title: str,
    y_axis_title: str,
    month_order: list[str],
):
    chart_data = add_short_label_column(data, "dashboard_category", "dashboard_category_short")
    chart_data["Segment Label"] = chart_data.apply(
        lambda row: (
            f"{row['dashboard_category_short']}<br>{row['Share']:.0%}"
            if row["Share"] >= 0.075 and row[amount_column] > 0
            else ""
        ),
        axis=1,
    )
    category_count = max(1, chart_data["dashboard_category_short"].nunique())
    month_count = max(1, chart_data["dashboard_month_label"].nunique())
    chart_height = max(500, 360 + (category_count * 34) + (month_count * 22))

    chart = px.bar(
        chart_data,
        x="dashboard_month_label",
        y="Share",
        color="dashboard_category_short",
        text="Segment Label",
        title=title,
        labels={"dashboard_month_label": "Month", "dashboard_category_short": category_title},
        category_orders={"dashboard_month_label": month_order},
        custom_data=["dashboard_category", "dashboard_month_label", amount_column, "Share"],
        color_discrete_map=composition_color_map(chart_data["dashboard_category_short"].unique()),
    )
    chart.update_traces(
        texttemplate="%{text}",
        textposition="inside",
        insidetextanchor="middle",
        textfont={"color": "#ffffff", "size": 12},
        cliponaxis=False,
        hovertemplate=(
            "<b>%{customdata[0]}</b><br>"
            "Month: %{customdata[1]}<br>"
            "Share: %{customdata[3]:.1%}<br>"
            f"{amount_title}: " + "%{customdata[2]:,.2f}<extra></extra>"
        ),
    )
    chart.update_layout(
        barmode="stack",
        title=title,
        height=chart_height,
        bargap=0.34 if month_count <= 3 else 0.24,
        uniformtext={"minsize": 10, "mode": "hide"},
        legend=readable_legend(y=-0.18, title=category_title),
        margin={"l": 24, "r": 24, "t": 72, "b": 130},
    )
    chart.update_yaxes(tickformat=".0%", range=[0, 1], title=y_axis_title)
    style_chart(chart)
    chart.update_layout(
        height=chart_height,
        bargap=0.34 if month_count <= 3 else 0.24,
        uniformtext={"minsize": 10, "mode": "hide"},
        legend=readable_legend(y=-0.18, title=category_title),
        margin={"l": 24, "r": 24, "t": 72, "b": 130},
    )
    return chart


def collapse_to_top_categories(df: pd.DataFrame, category_column: str, amount_column: str, top_n: int) -> pd.DataFrame:
    if df.empty:
        return df

    top_categories = (
        df.groupby(category_column)[amount_column]
        .sum()
        .sort_values(ascending=False)
        .head(top_n)
        .index
    )
    collapsed = df.copy()
    collapsed[category_column] = collapsed[category_column].where(collapsed[category_column].isin(top_categories), "Other")
    return (
        collapsed.groupby(["dashboard_month_sort", "dashboard_month_label", category_column], dropna=False)[amount_column]
        .sum()
        .reset_index()
        .sort_values(["dashboard_month_sort", amount_column], ascending=[True, False])
    )


def composition_color_map(labels: Any) -> dict[str, str]:
    palette = ["#146c43", "#c84d62", "#2563a8", "#6d4aa8", "#374151", "#0f766e"]
    return {
        label: ("#9ca3af" if label == "Other" else palette[index % len(palette)])
        for index, label in enumerate(labels)
    }


def render_drilldown_summary(drilldown: pd.DataFrame) -> None:
    kpis = st.columns(4)
    kpis[0].metric("Total income", format_money(drilldown["amount_income"].sum()))
    kpis[1].metric("Total expenses", format_money(drilldown["amount_expense"].sum()))
    kpis[2].metric("Net result", format_money(drilldown["signed_amount"].sum()))
    kpis[3].metric("Transactions", f"{len(drilldown):,}")


def financial_summary(df: pd.DataFrame, label_column: str) -> pd.DataFrame:
    summary = (
        df.assign(label=df[label_column].fillna("").replace("", "(blank)"))
        .groupby("label", dropna=False)[["amount_income", "amount_expense"]]
        .sum()
        .reset_index()
    )
    summary = summary.rename(columns={"amount_income": "Income", "amount_expense": "Expenses"})
    summary["Net Result"] = summary["Income"] - summary["Expenses"]
    summary["Total Volume"] = summary["Income"].abs() + summary["Expenses"].abs()
    return summary


def render_active_rules_overview() -> None:
    st.subheader("Active Classification Rules")
    rules = build_rules_overview()
    filtered = filter_rules_overview(rules)

    summary = st.columns(6)
    summary[0].metric("Total active rules", f"{int(rules['is_active'].sum()) if not rules.empty else 0:,}")
    summary[1].metric("Historical rules", f"{(rules['rule_source'] == 'historical').sum() if not rules.empty else 0:,}")
    summary[2].metric("Manual rules", f"{(rules['rule_source'] == 'manual').sum() if not rules.empty else 0:,}")
    summary[3].metric("Default rules", f"{(rules['rule_source'] == 'default').sum() if not rules.empty else 0:,}")
    summary[4].metric("Disabled rules", f"{(rules['is_active'] == False).sum() if not rules.empty else 0:,}")
    summary[5].metric("Low-confidence", f"{(rules['confidence_score'] < 0.9).sum() if not rules.empty else 0:,}")

    category_rules = filtered[filtered["rule_type"] == "category"].copy()
    project_rules = filtered[filtered["rule_type"] == "project"].copy()

    st.subheader("Category Rules")
    edited_category = render_rules_editor(category_rules, "category-rules-editor", "category")
    st.subheader("Project Name Rules")
    edited_project = render_rules_editor(project_rules, "project-rules-editor", "project_name")

    delete_requested = bool(
        edited_category.get("delete_rule", pd.Series(dtype=bool)).any()
        or edited_project.get("delete_rule", pd.Series(dtype=bool)).any()
    )
    delete_confirmed = True
    if delete_requested:
        st.warning("Deleting this rule may affect future categorization consistency.")
        delete_confirmed = st.checkbox("I understand the consistency risk.", key="confirm-rule-delete")

    if st.button("Save Rule Changes", type="primary", disabled=delete_requested and not delete_confirmed):
        saved = save_rule_overview_changes(pd.concat([edited_category, edited_project], ignore_index=True))
        st.success(f"Saved {saved:,} rule change(s).")
        st.rerun()


def render_rules_editor(rules: pd.DataFrame, key: str, label_column: str) -> pd.DataFrame:
    if rules.empty:
        st.info("No rules match the current filters.")
        return rules

    visible = [
        "rule_id",
        "keyword_pattern",
        "match_field",
        "transaction_type",
        label_column,
        "occurrence_count",
        "confidence_score",
        "priority",
        "rule_source",
        "created_at",
        "is_active",
        "reviewed",
        "status",
        "delete_rule",
    ]
    disabled = [
        "rule_id",
        "keyword_pattern",
        "match_field",
        "transaction_type",
        label_column,
        "occurrence_count",
        "confidence_score",
        "rule_source",
        "created_at",
        "status",
    ]
    return st.data_editor(
        rules[visible],
        use_container_width=True,
        hide_index=True,
        height=420,
        key=key,
        disabled=disabled,
        column_config={
            "is_active": st.column_config.CheckboxColumn("Active"),
            "reviewed": st.column_config.CheckboxColumn("Reviewed"),
            "delete_rule": st.column_config.CheckboxColumn("Delete"),
            "priority": st.column_config.NumberColumn("Priority", step=1),
        },
    )


def filter_transactions(
    transactions: pd.DataFrame,
    key_prefix: str,
    include_category: bool = False,
    include_project: bool = False,
    include_text: bool = True,
) -> pd.DataFrame:
    filtered = transactions.copy()
    valid_dates = pd.to_datetime(filtered["date"], errors="coerce").dropna()

    filter_cols = st.columns(4)
    if valid_dates.empty:
        date_range = None
        filter_cols[0].info("No valid dates found.")
    else:
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()
        date_range = filter_cols[0].date_input(
            "Date range",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            key=f"{key_prefix}-date-range",
        )

    transaction_types = sorted(value for value in filtered["transaction_type"].dropna().unique() if value)
    selected_types = filter_cols[1].multiselect(
        "Transaction type",
        transaction_types,
        default=transaction_types,
        key=f"{key_prefix}-transaction-type",
    )

    selected_categories = None
    selected_projects = None
    if include_category:
        categories = sorted(value for value in filtered["category"].dropna().unique() if value)
        selected_categories = filter_cols[2].multiselect(
            "Category",
            categories,
            default=categories,
            key=f"{key_prefix}-category",
        )
    if include_project:
        projects = sorted(value for value in filtered["project_name"].dropna().unique() if value)
        selected_projects = filter_cols[3].multiselect(
            "Project Name",
            projects,
            default=projects,
            key=f"{key_prefix}-project",
        )

    if include_text:
        text_cols = st.columns(2)
        description_search = text_cols[0].text_input("Search by description", key=f"{key_prefix}-description-search")
        counterparty_search = text_cols[1].text_input("Search by counterparty", key=f"{key_prefix}-counterparty-search")
    else:
        description_search = ""
        counterparty_search = ""

    if date_range and len(date_range) == 2:
        start_date = pd.to_datetime(date_range[0])
        end_date = pd.to_datetime(date_range[1])
        filtered = filtered[filtered["date"].between(start_date, end_date, inclusive="both") | filtered["date"].isna()]
    if selected_types:
        filtered = filtered[filtered["transaction_type"].isin(selected_types)]
    if selected_categories is not None and selected_categories:
        filtered = filtered[filtered["category"].isin(selected_categories)]
    if selected_projects is not None and selected_projects:
        filtered = filtered[filtered["project_name"].isin(selected_projects)]
    if description_search:
        filtered = filtered[filtered["description"].fillna("").astype(str).str.contains(description_search, case=False, na=False)]
    if counterparty_search:
        filtered = filtered[filtered["counterparty_raw"].fillna("").astype(str).str.contains(counterparty_search, case=False, na=False)]
    return filtered


def filter_rules_overview(rules: pd.DataFrame) -> pd.DataFrame:
    if rules.empty:
        return rules

    cols = st.columns(4)
    search = cols[0].text_input("Search keyword", key="rules-search")
    labels = sorted(set(rules["category"].dropna().tolist() + rules["project_name"].dropna().tolist()))
    selected_labels = cols[1].multiselect("Category / Project Name", labels, default=labels, key="rules-label-filter")
    sources = sorted(rules["rule_source"].dropna().unique())
    selected_sources = cols[2].multiselect("Rule source", sources, default=sources, key="rules-source-filter")
    min_confidence = cols[3].slider("Confidence threshold", 0.0, 1.0, 0.0, 0.05, key="rules-confidence-filter")

    sort_by = st.selectbox("Sort rules by", ["occurrence_count", "confidence_score", "priority"], key="rules-sort")

    filtered = rules.copy()
    if search:
        filtered = filtered[filtered["keyword_pattern"].fillna("").str.contains(search, case=False, na=False)]
    if selected_labels:
        filtered = filtered[filtered["category"].isin(selected_labels) | filtered["project_name"].isin(selected_labels)]
    if selected_sources:
        filtered = filtered[filtered["rule_source"].isin(selected_sources)]
    filtered = filtered[filtered["confidence_score"] >= min_confidence]
    return filtered.sort_values(sort_by, ascending=sort_by == "priority")


def format_transactions(transactions: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    display = transactions[columns].copy()
    if "date" in display:
        display["date"] = pd.to_datetime(display["date"], errors="coerce").dt.strftime("%Y-%m-%d")
    if "signed_amount" in display:
        display["signed_amount"] = pd.to_numeric(display["signed_amount"], errors="coerce").map(format_amount)
    return display.fillna("").rename(columns=DISPLAY_NAMES)


def format_editor_rows(transactions: pd.DataFrame) -> pd.DataFrame:
    editor = transactions.copy()
    editor.insert(0, "row_id", transactions.index)
    for column in ["date"]:
        editor[column] = pd.to_datetime(editor[column], errors="coerce").dt.strftime("%Y-%m-%d")
    editor["signed_amount"] = pd.to_numeric(editor["signed_amount"], errors="coerce").map(format_amount)
    return editor[EDITOR_COLUMNS].fillna("")


def extract_editor_corrections(edited: pd.DataFrame, transactions: pd.DataFrame) -> pd.DataFrame:
    if edited.empty:
        return pd.DataFrame()

    row_ids = edited["row_id"].astype(int).tolist()
    source = transactions.loc[row_ids].copy()
    edited_by_row = edited.set_index("row_id")
    source["old_category"] = source["category"].fillna("").astype(str)
    source["new_category"] = edited_by_row.loc[source.index, "category"].fillna("").astype(str).values
    source["old_project_name"] = source["project_name"].fillna("").astype(str)
    source["new_project_name"] = edited_by_row.loc[source.index, "project_name"].fillna("").astype(str).values
    source["old_commentary"] = source["commentary"].fillna("").astype(str)
    source["new_commentary"] = edited_by_row.loc[source.index, "commentary"].fillna("").astype(str).values

    changed = source[
        (source["old_category"] != source["new_category"])
        | (source["old_project_name"] != source["new_project_name"])
        | (source["old_commentary"] != source["new_commentary"])
    ].copy()
    return changed


def apply_commentary_overrides(transactions: pd.DataFrame) -> pd.DataFrame:
    ensure_rules_database(RULES_DATABASE_FILE)
    result = transactions.copy()
    result["commentary"] = ""
    with connect(RULES_DATABASE_FILE) as connection:
        rows = connection.execute(
            """
            SELECT c.transaction_id, c.commentary
            FROM manual_corrections c
            JOIN (
                SELECT transaction_id, MAX(edited_at) AS latest_edit
                FROM manual_corrections
                GROUP BY transaction_id
            ) latest
                ON c.transaction_id = latest.transaction_id
                AND c.edited_at = latest.latest_edit
            """
        ).fetchall()
    comments = {transaction_id: commentary or "" for transaction_id, commentary in rows}
    result["commentary"] = result["transaction_id"].map(comments).fillna("")
    return result


def build_duplicate_report(corrections: pd.DataFrame, transactions: pd.DataFrame) -> dict:
    categories = sorted(set(build_category_options(transactions)))
    projects = sorted(set(build_project_options(transactions)))
    new_categories = sorted(
        value
        for value in corrections.get("new_category", pd.Series(dtype=str)).dropna().unique()
        if value.strip() and is_new_label(value, categories)
    )
    new_projects = sorted(
        value
        for value in corrections.get("new_project_name", pd.Series(dtype=str)).dropna().unique()
        if value.strip() and is_new_label(value, projects)
    )
    return {
        "new_categories": new_categories,
        "new_projects": new_projects,
        "category_duplicates": find_near_duplicates(new_categories, categories),
        "project_duplicates": find_near_duplicates(new_projects, projects),
    }


def render_duplicate_warnings(report: dict) -> None:
    for label, existing, score in report["category_duplicates"]:
        st.warning(f"A similar category already exists: {existing}. Similarity: {score:.0%}.")
    for label, existing, score in report["project_duplicates"]:
        st.warning(f"A similar project already exists: {existing}. Similarity: {score:.0%}.")
    if report["new_categories"]:
        st.info("New categories: " + ", ".join(report["new_categories"]))
    if report["new_projects"]:
        st.info("New project names: " + ", ".join(report["new_projects"]))


def build_rules_overview() -> pd.DataFrame:
    ensure_rules_database(RULES_DATABASE_FILE)
    rules = []
    rules.extend(rule_rows(load_historical_category_rules(HISTORICAL_RULE_SOURCE_FILE), "category"))
    rules.extend(rule_rows(load_historical_project_rules(HISTORICAL_RULE_SOURCE_FILE), "project"))
    rules.extend(rule_rows(load_category_rules(CATEGORY_RULES_FILE), "category"))
    rules.extend(rule_rows(load_project_rules(PROJECT_RULES_FILE), "project"))
    rules.extend(load_manual_rule_rows())

    df = pd.DataFrame(rules)
    if df.empty:
        return pd.DataFrame(columns=rule_overview_columns())

    overrides = load_rule_overrides()
    for index, row in df.iterrows():
        override = overrides.get(row["rule_id"], {})
        if override.get("priority_override") is not None:
            df.at[index, "priority"] = int(override["priority_override"])
        if override.get("is_active") is not None:
            df.at[index, "is_active"] = bool(override["is_active"])
        if override.get("deleted_at"):
            df.at[index, "is_active"] = False
            df.at[index, "deleted"] = True
        if override.get("reviewed_at"):
            df.at[index, "reviewed"] = True

    df["duplicate_rule"] = df.duplicated(["keyword_pattern", "match_field", "transaction_type", "category", "project_name"], keep=False)
    df["status"] = df.apply(rule_status, axis=1)
    df["delete_rule"] = False
    return df[rule_overview_columns()]


def rule_rows(rules: list[dict], rule_type: str) -> list[dict]:
    rows = []
    for index, rule in enumerate(rules, start=1):
        label = rule.get("label", "")
        rows.append(
            {
                "rule_id": rule.get("rule_id") or f"{rule.get('origin', 'default')}-{rule_type}-{index}",
                "rule_type": rule_type,
                "keyword_pattern": rule.get("keyword") or rule.get("category", ""),
                "match_field": rule.get("source", "any"),
                "transaction_type": rule.get("transaction_type", "any"),
                "category": label if rule_type == "category" else rule.get("category", ""),
                "project_name": label if rule_type == "project" else "",
                "occurrence_count": int(rule.get("occurrence_count", 0) or 0),
                "confidence_score": float(rule.get("confidence", 0.0) or 0.0),
                "priority": int(rule.get("priority", 999)),
                "rule_source": rule.get("rule_source") or rule.get("origin", "default"),
                "created_at": rule.get("created_at", ""),
                "is_active": True,
                "reviewed": False,
                "deleted": False,
            }
        )
    return rows


def load_manual_rule_rows() -> list[dict]:
    ensure_rules_database(RULES_DATABASE_FILE)
    with connect(RULES_DATABASE_FILE) as connection:
        rows = connection.execute(
            """
            SELECT id, rule_type, transaction_type, source, keyword, label, priority, is_active, created_at, reviewed_at
            FROM manual_rules_v2
            """
        ).fetchall()
    output = []
    for row_id, rule_type, transaction_type, source, keyword, label, priority, is_active, created_at, reviewed_at in rows:
        output.append(
            {
                "rule_id": f"manual-{row_id}",
                "rule_type": rule_type,
                "keyword_pattern": keyword,
                "match_field": source,
                "transaction_type": transaction_type,
                "category": label if rule_type == "category" else "",
                "project_name": label if rule_type == "project" else "",
                "occurrence_count": 1,
                "confidence_score": 1.0,
                "priority": int(priority),
                "rule_source": "manual",
                "created_at": created_at,
                "is_active": bool(is_active),
                "reviewed": bool(reviewed_at),
                "deleted": False,
            }
        )
    return output


def rule_overview_columns() -> list[str]:
    return [
        "rule_id",
        "rule_type",
        "keyword_pattern",
        "match_field",
        "transaction_type",
        "category",
        "project_name",
        "occurrence_count",
        "confidence_score",
        "priority",
        "rule_source",
        "created_at",
        "is_active",
        "reviewed",
        "status",
        "delete_rule",
    ]


def rule_status(row: pd.Series) -> str:
    statuses = []
    if not row["is_active"]:
        statuses.append("DISABLED")
    if row["confidence_score"] < 0.9:
        statuses.append("LOW CONFIDENCE")
    if row.get("duplicate_rule"):
        statuses.append("SIMILAR")
    if row.get("deleted"):
        statuses.append("DELETED")
    return ", ".join(statuses)


def apply_rule_overrides(rules: list[dict]) -> list[dict]:
    overrides = load_rule_overrides()
    active = []
    for rule in rules:
        if (rule.get("rule_source") or rule.get("origin")) == "historical":
            active.append(rule)
            continue

        rule_id = rule.get("rule_id", "")
        override = overrides.get(rule_id, {})
        if override.get("deleted_at") or override.get("is_active") == 0:
            continue
        if override.get("priority_override") is not None:
            rule = rule.copy()
            rule["priority"] = int(override["priority_override"])
        active.append(rule)
    return sorted(active, key=rule_precedence_key)


def rule_precedence_key(rule: dict) -> tuple[int, int]:
    source = rule.get("rule_source") or rule.get("origin", "default")
    source_rank = {"manual": 0, "historical": 1, "default": 2}.get(str(source), 3)
    return source_rank, int(rule.get("priority", 999))


def load_rule_overrides() -> dict:
    ensure_rules_database(RULES_DATABASE_FILE)
    with connect(RULES_DATABASE_FILE) as connection:
        rows = connection.execute(
            """
            SELECT rule_id, is_active, priority_override, reviewed_at, deleted_at
            FROM classification_rule_overrides
            """
        ).fetchall()
    return {
        rule_id: {
            "is_active": is_active,
            "priority_override": priority_override,
            "reviewed_at": reviewed_at,
            "deleted_at": deleted_at,
        }
        for rule_id, is_active, priority_override, reviewed_at, deleted_at in rows
    }


def save_rule_overview_changes(edited: pd.DataFrame) -> int:
    ensure_rules_database(RULES_DATABASE_FILE)
    saved = 0
    with connect(RULES_DATABASE_FILE) as connection:
        for _, row in edited.iterrows():
            if row.get("rule_source") == "historical":
                continue

            rule_id = row["rule_id"]
            connection.execute(
                """
                INSERT INTO classification_rule_overrides (rule_id, is_active, priority_override, reviewed_at, deleted_at, updated_at)
                VALUES (?, ?, ?, CASE WHEN ? THEN CURRENT_TIMESTAMP ELSE NULL END, CASE WHEN ? THEN CURRENT_TIMESTAMP ELSE NULL END, CURRENT_TIMESTAMP)
                ON CONFLICT(rule_id)
                DO UPDATE SET
                    is_active = excluded.is_active,
                    priority_override = excluded.priority_override,
                    reviewed_at = excluded.reviewed_at,
                    deleted_at = COALESCE(excluded.deleted_at, classification_rule_overrides.deleted_at),
                    updated_at = CURRENT_TIMESTAMP
                """,
                (
                    rule_id,
                    1 if bool(row.get("is_active")) else 0,
                    int(row.get("priority", 999)),
                    bool(row.get("reviewed")),
                    bool(row.get("delete_rule")),
                ),
            )
            if bool(row.get("delete_rule")) and row.get("rule_source") == "manual" and str(rule_id).startswith("manual-"):
                manual_id = int(str(rule_id).split("-", 1)[1])
                connection.execute("UPDATE manual_rules_v2 SET is_active = 0 WHERE id = ?", (manual_id,))
            saved += 1
    return saved


def build_category_options(transactions: pd.DataFrame) -> list[str]:
    analysis = get_historical_analysis()
    historical = analysis.get("taxonomy_categories", pd.DataFrame())
    values = historical.get("Category", pd.Series(dtype=str)).dropna().tolist()
    values += transactions["category"].dropna().tolist()
    return sorted({value for value in values if str(value).strip() and value != DEFAULT_CATEGORY})


def build_project_options(transactions: pd.DataFrame) -> list[str]:
    analysis = get_historical_analysis()
    historical = analysis.get("taxonomy_projects", pd.DataFrame())
    values = historical.get("Project Name", pd.Series(dtype=str)).dropna().tolist()
    values += transactions["project_name"].dropna().tolist()
    return sorted({value for value in values if str(value).strip() and value != DEFAULT_PROJECT_NAME})


def is_new_label(value: str, existing_values: list[str]) -> bool:
    normalized = normalize_label(value)
    return bool(normalized) and normalized not in {normalize_label(existing) for existing in existing_values}


def find_near_duplicates(new_values: list[str], existing_values: list[str]) -> list[tuple[str, str, float]]:
    duplicates = []
    for new_value in new_values:
        normalized_new = normalize_label(new_value)
        for existing in existing_values:
            normalized_existing = normalize_label(existing)
            if not normalized_new or normalized_new == normalized_existing:
                continue
            score = SequenceMatcher(None, normalized_new, normalized_existing).ratio()
            if score >= 0.82:
                duplicates.append((new_value, existing, score))
                break
    return duplicates


def normalize_label(value: str) -> str:
    normalized = str(value).strip().lower()
    normalized = re.sub(r"[^\w\s]", "", normalized)
    normalized = re.sub(r"\s+", " ", normalized)
    if normalized.endswith("s"):
        normalized = normalized[:-1]
    return normalized


def wrap_axis_label(value: str, max_chars: int = 18) -> str:
    words = short_chart_label(value).split()
    if not words:
        return ""

    lines = []
    current = words[0]
    for word in words[1:]:
        if len(current) + len(word) + 1 <= max_chars:
            current = f"{current} {word}"
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return "<br>".join(lines)


def short_chart_label(value: str, max_words: int = 2) -> str:
    label = str(value).strip()
    if not label:
        return ""
    if len(label) <= 14 and len(label.split()) <= max_words:
        return label

    cleaned = re.sub(r"\([^)]*\)", "", label)
    cleaned = re.sub(r"\b(19|20)\d{2}\b", "", cleaned)
    cleaned = re.sub(r"\bjan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec\b", "", cleaned, flags=re.IGNORECASE)
    words = [word.strip(" -_/|,:;") for word in cleaned.split()]
    words = [word for word in words if word]
    stop_words = {
        "in",
        "of",
        "the",
        "and",
        "for",
        "with",
        "to",
        "as",
        "a",
        "an",
        "project",
        "projects",
        "program",
        "programme",
        "division",
        "category",
    }
    keyword_words = [word for word in words if word.lower() not in stop_words]
    chosen = keyword_words[:max_words] or words[:max_words]
    return " ".join(chosen) if chosen else label


def add_short_label_column(df: pd.DataFrame, source_column: str, label_column: str) -> pd.DataFrame:
    data = df.copy()
    data[label_column] = data[source_column].apply(short_chart_label)
    duplicates = data.groupby(label_column)[source_column].transform("nunique").gt(1)
    data.loc[duplicates, label_column] = data.loc[duplicates, source_column]
    return data


def add_multilevel_x_axis_labels(
    chart: go.Figure,
    group_centers: list[tuple[float, str]],
    month_centers: list[dict[str, object]],
    group_ranges: list[dict[str, object]],
    group_label_max_chars: int = 18,
) -> None:
    for center, label in group_centers:
        chart.add_annotation(
            text=f"<b>{wrap_axis_label(label, group_label_max_chars)}</b>",
            x=center,
            y=-0.18,
            xref="x",
            yref="paper",
            showarrow=False,
            align="center",
            font={"size": 11, "color": "#5f6675"},
        )

    for month_marker in month_centers:
        chart.add_annotation(
            text=str(month_marker["Month"]),
            x=float(month_marker["x"]),
            y=-0.36,
            xref="x",
            yref="paper",
            showarrow=False,
            textangle=0,
            font={"size": 10, "color": "#7b8191"},
        )

    for left, right in zip(group_ranges, group_ranges[1:]):
        boundary = (float(left["end"]) + float(right["start"])) / 2
        chart.add_shape(
            type="line",
            x0=boundary,
            x1=boundary,
            y0=0,
            y1=1,
            xref="x",
            yref="paper",
            line={"color": "rgba(28, 31, 35, 0.12)", "width": 1},
        )


def aggregate_amount(df: pd.DataFrame, label_column: str, amount_column: str, transaction_type: str | None) -> pd.DataFrame:
    data = df.copy()
    if transaction_type:
        data = data[data["transaction_type"] == transaction_type]
    grouped = (
        data.assign(label=data[label_column].fillna("").replace("", "(blank)"))
        .groupby("label", dropna=False)[amount_column]
        .sum()
        .abs()
        .sort_values(ascending=False)
        .reset_index(name="amount")
        .head(15)
    )
    return grouped


def build_income_expense_comparison(df: pd.DataFrame, label_column: str) -> pd.DataFrame:
    base = df.copy()
    base[label_column] = base[label_column].fillna("").replace("", "(blank)")
    income = (
        base.groupby(label_column, dropna=False)["amount_income"]
        .sum()
        .abs()
        .reset_index(name="amount")
        .assign(Type="Income")
    )
    expenses = (
        base.groupby(label_column, dropna=False)["amount_expense"]
        .sum()
        .abs()
        .reset_index(name="amount")
        .assign(Type="Expenses")
    )
    combined = pd.concat([income, expenses], ignore_index=True)
    combined = combined.rename(columns={label_column: "label"})
    order = (
        combined.groupby("label")["amount"]
        .sum()
        .sort_values(ascending=True)
        .index
        .tolist()
    )
    combined["label"] = pd.Categorical(combined["label"], categories=order, ordered=True)
    return combined.sort_values("label")


def render_drilldown_selector(
    transactions: pd.DataFrame,
    label_column: str,
    default_label: str,
    selector_label: str,
    key: str,
) -> None:
    values = sorted(value for value in transactions[label_column].fillna("").unique() if str(value).strip())
    selected = st.selectbox(selector_label, [default_label] + values, key=key)
    if selected == default_label:
        return

    drilldown = transactions[transactions[label_column] == selected].copy()
    kpis = st.columns(4)
    kpis[0].metric("Total income", format_money(drilldown["amount_income"].sum()))
    kpis[1].metric("Total expenses", format_money(drilldown["amount_expense"].sum()))
    kpis[2].metric("Net amount", format_money(drilldown["signed_amount"].sum()))
    kpis[3].metric("Transactions", f"{len(drilldown):,}")
    st.dataframe(
        format_transactions(drilldown, TRANSACTION_COLUMNS),
        use_container_width=True,
        hide_index=True,
        height=420,
        key=f"{key}-table",
    )


def format_money(value: float) -> str:
    return f"{value:,.2f}"


def format_amount(value: float) -> str:
    if pd.isna(value):
        return ""
    return f"{value:,.2f}"


def format_date_value(value: pd.Timestamp) -> str:
    if pd.isna(value):
        return ""
    return value.date().isoformat()


if __name__ == "__main__":
    main()
