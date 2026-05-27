from pathlib import Path
import base64
import html
import re
from difflib import SequenceMatcher
from sqlite3 import connect

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

    section_gap()
    chart_filtered = filter_by_month_view(filtered)
    if chart_filtered.empty:
        st.info("No transactions match the selected month view.")
        return

    section_gap()
    sort_by = st.selectbox(
        "Sort category performance by",
        ["Income", "Expenses"],
        key="dashboard-category-sort",
    )
    category_chart = build_category_performance_chart(chart_filtered, sort_by)
    st.plotly_chart(category_chart, use_container_width=True)

    section_gap()
    st.subheader("Division Breakdown for Top 5 Categories")
    st.caption("Top five categories ranked by total financial volume: income + expenses.")
    for category in top_categories_by_volume(chart_filtered, limit=5):
        division_chart = build_category_division_breakdown_by_month_chart(chart_filtered, category)
        if division_chart is None:
            st.info(f"No division data available for this category: {category}.")
        else:
            st.plotly_chart(division_chart, use_container_width=True)

    section_gap()
    render_custom_division_breakdown(chart_filtered)

    section_gap()
    expense_chart = build_expense_composition_chart(chart_filtered)
    st.plotly_chart(expense_chart, use_container_width=True)

    section_gap()
    revenue_dependency_chart = build_revenue_dependency_chart(chart_filtered)
    st.plotly_chart(revenue_dependency_chart, use_container_width=True)

    section_gap()
    with st.expander("Detailed Category / Transaction Drilldown", expanded=False):
        selected_category = st.selectbox(
            "Select category for detailed transactions",
            ["All categories"] + sorted(filtered["dashboard_category"].dropna().unique().tolist()),
            key="dashboard-category-drilldown",
        )
        drilldown = chart_filtered if selected_category == "All categories" else chart_filtered[chart_filtered["dashboard_category"] == selected_category]
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
        font={"family": "Inter, Arial, sans-serif", "color": "#171717"},
        title={"font": {"size": 22, "color": "#171717"}, "x": 0.02},
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        legend={
            "orientation": "h",
            "yanchor": "bottom",
            "y": 1.02,
            "xanchor": "right",
            "x": 1,
        },
        margin={"l": 24, "r": 24, "t": 72, "b": 36},
    )
    chart.update_xaxes(
        gridcolor="rgba(23,23,23,0.08)",
        zerolinecolor="rgba(23,23,23,0.12)",
        title_font={"color": "#6b6b5f"},
    )
    chart.update_yaxes(
        gridcolor="rgba(23,23,23,0.08)",
        zerolinecolor="rgba(23,23,23,0.12)",
        title_font={"color": "#6b6b5f"},
    )


def prepare_dashboard_data(transactions: pd.DataFrame) -> pd.DataFrame:
    data = transactions.copy()
    data["date"] = pd.to_datetime(data["date"], errors="coerce")
    data["dashboard_month_sort"] = data["date"].dt.strftime("%Y-%m").fillna("No date")
    data["dashboard_month_label"] = data["date"].dt.strftime("%b %Y").fillna("No date")
    data["amount_income"] = pd.to_numeric(data["amount_income"], errors="coerce").fillna(0)
    data["amount_expense"] = pd.to_numeric(data["amount_expense"], errors="coerce").fillna(0)
    data["signed_amount"] = pd.to_numeric(data["signed_amount"], errors="coerce").fillna(0)
    data["dashboard_category"] = clean_dimension_series(data.get("Category", data.get("category")), "Unclassified")
    if "category" in data:
        data["dashboard_category"] = data["dashboard_category"].where(
            data["dashboard_category"].ne("Unclassified"),
            clean_dimension_series(data["category"], "Unclassified"),
        )
    data["dashboard_division"] = clean_dimension_series(data.get("Division", data.get("project_name")), "Unknown division")
    if "project_name" in data:
        data["dashboard_division"] = data["dashboard_division"].where(
            data["dashboard_division"].ne("Unknown division"),
            clean_dimension_series(data["project_name"], "Unknown division"),
        )
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


def filter_by_month_view(filtered: pd.DataFrame) -> pd.DataFrame:
    st.subheader("Month View")
    month_options = ordered_month_labels(filtered)
    mode_cols = st.columns([1, 2])
    month_mode = mode_cols[0].radio(
        "Month view",
        ["All Months", "Specific Month"],
        horizontal=True,
        key="dashboard-month-view-mode",
    )
    if month_mode == "All Months":
        mode_cols[1].caption("All charts below compare months inside the same chart.")
        return filtered

    selected_month = mode_cols[1].selectbox(
        "Select Month",
        month_options,
        key="dashboard-month-view-selected",
    )
    return filtered[filtered["dashboard_month_label"].eq(selected_month)]


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
    dependency_value, dependency_hint = grant_donation_dependency(filtered)

    kpis = st.columns(6)
    kpis[0].markdown(kpi_card("Total Income", format_money(total_income), "positive"), unsafe_allow_html=True)
    kpis[1].markdown(kpi_card("Total Expenses", format_money(total_expenses), "negative"), unsafe_allow_html=True)
    kpis[2].markdown(
        kpi_card("Net Result", format_money(net_result), "positive" if net_result >= 0 else "negative"),
        unsafe_allow_html=True,
    )
    kpis[3].markdown(kpi_card("Largest Revenue Category", largest_revenue, "neutral"), unsafe_allow_html=True)
    kpis[4].markdown(kpi_card("Largest Expense Category", largest_expense, "neutral"), unsafe_allow_html=True)
    kpis[5].markdown(kpi_card("Grant / Donation Dependency", dependency_value, "neutral", dependency_hint), unsafe_allow_html=True)


def kpi_card(label: str, value: str, tone: str, hint: str = "") -> str:
    hint_markup = f"<em>{html.escape(hint)}</em>" if hint else ""
    return f"""
    <div class="yf-kpi-card yf-kpi-{tone}">
        <span>{html.escape(label)}</span>
        <strong>{html.escape(value)}</strong>
        {hint_markup}
    </div>
    """


def top_label(summary: pd.DataFrame, column: str, fallback: str) -> str:
    rows = summary[summary[column].gt(0)].sort_values(column, ascending=False)
    if rows.empty:
        return fallback
    row = rows.iloc[0]
    return f"{row['label']} ({format_money(row[column])})"


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
    monthly["Net Result"] = monthly["amount_income"] - monthly["amount_expense"]
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
    chart.add_scatter(
        name="Net Result",
        x=x_values,
        y=monthly["Net Result"],
        mode="lines+markers",
        line={"color": "#171717", "width": 4},
    )
    chart.update_layout(
        barmode="group",
        title="Monthly Income, Expenses, and Net Result",
        xaxis_title="Month",
        yaxis_title="Amount",
    )
    chart.update_yaxes(tickformat=",.2f")
    style_chart(chart)
    return chart


def build_category_performance_chart(filtered: pd.DataFrame, sort_by: str):
    summary = financial_summary(filtered, "dashboard_category")
    sort_map = {
        "Income": "Income",
        "Expenses": "Expenses",
    }
    sort_column = sort_map.get(sort_by, "Income")
    summary = summary.sort_values(sort_column, ascending=False)
    category_order = summary["label"].tolist()
    months = ordered_month_labels(filtered)

    chart_data = (
        filtered.groupby(["dashboard_category", "dashboard_month_sort", "dashboard_month_label"], dropna=False)[
            ["amount_income", "amount_expense"]
        ]
        .sum()
        .reset_index()
    )
    chart_data = chart_data.rename(columns={"amount_income": "Income", "amount_expense": "Expenses"})

    income_palette = ["#14532d", "#2f7d45", "#68a878", "#a7d0ad", "#d5ead7"]
    expense_palette = ["#9f2f45", "#c84d62", "#e18491", "#efb3bc", "#f7d8dd"]
    bar_width = 0.32
    metric_gap = 0.08
    month_gap = 0.28
    category_gap = 1.0

    positioned_rows = []
    category_centers = []
    month_centers = []
    x_cursor = 0.0

    for category in category_order:
        category_positions = []
        for month_index, month in enumerate(months):
            row = chart_data[
                chart_data["dashboard_category"].eq(category)
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
                        "Category": category,
                        "Month": month,
                        "Metric": "Income",
                        "Color": income_palette[month_index % len(income_palette)],
                    },
                    {
                        "x": expense_x,
                        "Amount": expense,
                        "Category": category,
                        "Month": month,
                        "Metric": "Expenses",
                        "Color": expense_palette[month_index % len(expense_palette)],
                    },
                ]
            )
            category_positions.extend([income_x, expense_x])
            month_centers.append({"x": month_center, "Month": month})
            x_cursor += (bar_width * 2) + metric_gap + month_gap

        if category_positions:
            category_centers.append((sum(category_positions) / len(category_positions), category))
        x_cursor += category_gap

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
            customdata=[[row["Category"], row["Month"], row["Metric"]]],
        )

    for month_index, month in enumerate(months):
        chart.add_scatter(
            x=[None],
            y=[None],
            mode="markers",
            name=month,
            marker={
                "color": income_palette[month_index % len(income_palette)],
                "size": 11,
                "symbol": "square",
            },
            showlegend=True,
        )

    chart.update_layout(
        title="Category Performance — Income vs Expenses by Month",
        barmode="overlay",
        bargap=0,
        height=620,
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        font={"family": "Inter, Arial, sans-serif", "color": "#20232a"},
        title_font={"size": 22, "color": "#151515"},
        margin={"l": 34, "r": 24, "t": 96, "b": 148},
        legend={
            "title": {"text": "Month"},
            "orientation": "h",
            "yanchor": "bottom",
            "y": 1.04,
            "xanchor": "right",
            "x": 1,
            "font": {"size": 12},
        },
        annotations=[
            {
                "text": "Green bars show income. Red bars show expenses. Darker/lighter tones distinguish months.",
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
    for month_marker in month_centers:
        chart.add_annotation(
            text=month_marker["Month"],
            x=month_marker["x"],
            y=-0.16,
            xref="x",
            yref="paper",
            showarrow=False,
            textangle=-90 if len(category_order) > 5 else 0,
            font={"size": 10, "color": "#6b7280"},
        )
    chart.update_yaxes(
        title="Amount",
        tickformat=",.2f",
        gridcolor="rgba(28, 31, 35, 0.08)",
        zerolinecolor="rgba(28, 31, 35, 0.18)",
    )
    chart.update_xaxes(
        title="Category",
        tickmode="array",
        tickvals=[center for center, _ in category_centers],
        ticktext=[category for _, category in category_centers],
        tickangle=-25 if len(category_order) > 5 else 0,
        showgrid=False,
        tickfont={"size": 12},
    )
    return chart


def top_categories_by_volume(filtered: pd.DataFrame, limit: int) -> list[str]:
    summary = financial_summary(filtered, "dashboard_category")
    return summary.sort_values("Total Volume", ascending=False)["label"].head(limit).tolist()


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
        margin={"l": 34, "r": 24, "t": 92, "b": 148},
        legend={
            "title": {"text": "Month"},
            "orientation": "h",
            "yanchor": "bottom",
            "y": 1.04,
            "xanchor": "right",
            "x": 1,
            "font": {"size": 12},
        },
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
    for month_marker in month_centers:
        chart.add_annotation(
            text=month_marker["Month"],
            x=month_marker["x"],
            y=-0.16,
            xref="x",
            yref="paper",
            showarrow=False,
            textangle=-90 if len(division_order) > 5 else 0,
            font={"size": 10, "color": "#6b7280"},
        )
    chart.update_yaxes(
        title="Amount",
        tickformat=",.2f",
        gridcolor="rgba(28, 31, 35, 0.08)",
        zerolinecolor="rgba(28, 31, 35, 0.18)",
    )
    chart.update_xaxes(
        title="Division",
        tickmode="array",
        tickvals=[center for center, _ in division_centers],
        ticktext=[division for _, division in division_centers],
        tickangle=-25 if len(division_order) > 5 else 0,
        showgrid=False,
        tickfont={"size": 12},
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

    chart = build_category_division_breakdown_by_month_chart(filtered, selected_category)
    if chart is None:
        st.info(f"No division data available for this category: {selected_category}.")
        return
    st.plotly_chart(chart, use_container_width=True)


def has_specific_division_data(rows: pd.DataFrame) -> bool:
    if "Division" not in rows.columns:
        return False
    cleaned = rows["Division"].fillna("").astype(str).str.strip()
    return cleaned.ne("").any()


def build_expense_composition_chart(filtered: pd.DataFrame):
    expenses = (
        filtered.groupby(["dashboard_month_sort", "dashboard_month_label", "dashboard_category"], dropna=False)["amount_expense"]
        .sum()
        .reset_index(name="Expenses")
    )
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
    chart = px.bar(
        expenses,
        x="dashboard_month_label",
        y="Share",
        color="dashboard_category",
        title="Expense Composition by Month",
        labels={"dashboard_month_label": "Month", "dashboard_category": "Expense Category"},
        category_orders={"dashboard_month_label": ordered_month_labels(filtered)},
        hover_data={"Expenses": ":,.2f", "Share": ":.1%", "dashboard_month_sort": False},
        color_discrete_map=composition_color_map(expenses["dashboard_category"].unique()),
    )
    chart.update_layout(barmode="stack", title="Expense Composition by Month")
    chart.update_yaxes(tickformat=".0%", range=[0, 1], title="Share of monthly expenses")
    style_chart(chart)
    return chart


def build_revenue_dependency_chart(filtered: pd.DataFrame):
    income = (
        filtered.groupby(["dashboard_month_sort", "dashboard_month_label", "dashboard_category"], dropna=False)["amount_income"]
        .sum()
        .reset_index(name="Income")
    )
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
    chart = px.bar(
        income,
        x="dashboard_month_label",
        y="Share",
        color="dashboard_category",
        title="Revenue Dependency by Month",
        labels={"dashboard_month_label": "Month", "dashboard_category": "Income Category"},
        category_orders={"dashboard_month_label": ordered_month_labels(filtered)},
        hover_data={"Income": ":,.2f", "Share": ":.1%", "dashboard_month_sort": False},
        color_discrete_map=composition_color_map(income["dashboard_category"].unique()),
    )
    chart.update_layout(barmode="stack")
    chart.update_yaxes(tickformat=".0%", range=[0, 1], title="Share of monthly income")
    style_chart(chart)
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
    palette = ["#0a8527", "#ff6f91", "#55b6ff", "#ffd84d", "#171717", "#8fd18f"]
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
