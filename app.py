from pathlib import Path
import re
from difflib import SequenceMatcher
from sqlite3 import connect

import pandas as pd
import plotly.express as px
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
INPUT_FILE = APP_ROOT / "data" / "input" / "bank_statement_2026_01_raw.xlsx"
HISTORICAL_RULE_SOURCE_FILE = APP_ROOT / "data" / "input" / "bank_statement_2026_01_updated.xlsx"
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


@st.cache_data
def get_historical_analysis() -> dict:
    if not HISTORICAL_RULE_SOURCE_FILE.exists():
        return {}
    return analyze_historical_file(HISTORICAL_RULE_SOURCE_FILE)


def main() -> None:
    st.title("NGO Finance Dashboard")
    st.caption(INPUT_FILE.name)

    if not INPUT_FILE.exists():
        st.error(f"Input file not found: {INPUT_FILE}")
        return

    try:
        transactions, validation_issues, debug_summary = load_classified_transactions()
    except Exception as exc:
        st.error(f"Could not load the input file: {exc}")
        return

    render_sidebar(transactions)

    import_tab, manual_tab, dashboard_tab = st.tabs(["Import Review", "Manual Review", "Dashboard"])
    with import_tab:
        render_import_review(transactions, validation_issues, debug_summary)
    with manual_tab:
        render_manual_review(transactions)
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

    st.subheader("Debug Summary")
    debug_cols = st.columns(4)
    debug_cols[0].metric("Original Excel row count", f"{debug_summary['original_excel_row_count']:,}")
    debug_cols[1].metric("Cleaned dataframe row count", f"{debug_summary['cleaned_dataframe_row_count']:,}")
    debug_cols[2].metric("Removed rows count", f"{debug_summary['removed_rows_count']:,}")
    debug_cols[3].write("")
    st.caption(f"Removal reasons: {debug_summary['removed_rows_reason']}")

    render_historical_source_summary(transactions)
    render_historical_match_check(transactions)

    st.subheader("All Transactions")
    filtered = filter_transactions(transactions, key_prefix="import")
    st.caption(f"Showing {len(filtered):,} of {len(transactions):,} rows.")
    st.dataframe(
        format_transactions(filtered, TRANSACTION_COLUMNS),
        use_container_width=True,
        hide_index=True,
        height=760,
    )

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
    st.subheader("Dashboard Filters")
    filtered = filter_transactions(
        transactions,
        key_prefix="dashboard",
        include_category=True,
        include_project=True,
        include_text=False,
    )

    total = len(filtered)
    categorized_count = int((filtered["category"] != DEFAULT_CATEGORY).sum())
    unclassified_count = int((filtered["category"] == DEFAULT_CATEGORY).sum())
    categorized_pct = categorized_count / total * 100 if total else 0
    unclassified_pct = unclassified_count / total * 100 if total else 0

    kpis = st.columns(5)
    kpis[0].metric("Total income", format_money(filtered["amount_income"].sum()))
    kpis[1].metric("Total expenses", format_money(filtered["amount_expense"].sum()))
    kpis[2].metric("Net cash flow", format_money(filtered["signed_amount"].sum()))
    kpis[3].metric("Categorized %", f"{categorized_pct:.1f}%")
    kpis[4].metric("Unclassified %", f"{unclassified_pct:.1f}%")

    chart_df = filtered.copy()
    chart_df["month_sort"] = pd.to_datetime(chart_df["date"], errors="coerce").dt.strftime("%Y-%m")
    chart_df["month_label"] = pd.to_datetime(chart_df["date"], errors="coerce").dt.strftime("%B")

    monthly = (
        chart_df.groupby(["month_sort", "month_label"], dropna=False)[["amount_income", "amount_expense"]]
        .sum()
        .reset_index()
        .sort_values("month_sort")
    )
    monthly_long = monthly.melt(
        id_vars=["month_sort", "month_label"],
        value_vars=["amount_income", "amount_expense"],
        var_name="Type",
        value_name="Amount",
    )
    monthly_long["Type"] = monthly_long["Type"].replace({"amount_income": "Income", "amount_expense": "Expenses"})
    monthly_long["Amount"] = pd.to_numeric(monthly_long["Amount"], errors="coerce").abs()
    month_order = monthly["month_label"].dropna().tolist()
    monthly_chart = px.bar(
        monthly_long,
        x="month_label",
        y="Amount",
        color="Type",
        barmode="group",
        title="Monthly Income vs Expenses",
        category_orders={"month_label": month_order},
        labels={"month_label": "Month"},
    )
    monthly_chart.update_yaxes(tickformat=",.2f")
    st.plotly_chart(monthly_chart, use_container_width=True)

    category_chart_df = build_income_expense_comparison(chart_df, "category")
    category_chart = px.bar(
        category_chart_df,
        x="amount",
        y="label",
        color="Type",
        barmode="group",
        orientation="h",
        title="Income vs Expenses by Category",
        labels={"amount": "Amount", "label": "Category"},
    )
    category_chart.update_xaxes(tickformat=",.2f")
    st.plotly_chart(category_chart, use_container_width=True)
    render_drilldown_selector(
        filtered,
        label_column="category",
        default_label="All categories",
        selector_label="Select category to inspect",
        key="dashboard-category-drilldown",
    )

    project_chart_df = build_income_expense_comparison(chart_df, "project_name")
    project_chart = px.bar(
        project_chart_df,
        x="amount",
        y="label",
        color="Type",
        barmode="group",
        orientation="h",
        title="Income vs Expenses by Project Name",
        labels={"amount": "Amount", "label": "Project Name"},
    )
    project_chart.update_xaxes(tickformat=",.2f")
    st.plotly_chart(project_chart, use_container_width=True)
    render_drilldown_selector(
        filtered,
        label_column="project_name",
        default_label="All projects",
        selector_label="Select project to inspect",
        key="dashboard-project-drilldown",
    )

    top_counterparties = aggregate_amount(chart_df, "counterparty_raw", "signed_amount", None).head(15)
    counterparty_chart = px.bar(
        top_counterparties,
        x="amount",
        y="label",
        orientation="h",
        title="Top Counterparties",
        labels={"amount": "Amount", "label": "Counterparty"},
    )
    counterparty_chart.update_xaxes(tickformat=",.2f")
    st.plotly_chart(counterparty_chart, use_container_width=True)


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
