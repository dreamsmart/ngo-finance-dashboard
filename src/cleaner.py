from __future__ import annotations

from typing import Any

import pandas as pd


REQUIRED_COLUMNS = [
    "Date",
    "Name Surname",
    "Personal Code",
    "Konta numurs",
    "Bankas SWIFT",
    "Purpose",
    "K (KREDITS)",
    "D (DEBETS)",
]


def clean_bank_statement(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    """Preserve original columns and append normalized bank statement fields."""
    missing_columns = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")

    original_row_count = len(df)
    row_is_fully_empty = df.apply(lambda row: all(not is_filled(value) for value in row), axis=1)
    raw_date_filled = df["Date"].apply(is_filled)
    raw_income_filled = df["K (KREDITS)"].apply(is_filled)
    raw_expense_filled = df["D (DEBETS)"].apply(is_filled)
    removable_rows = row_is_fully_empty & ~raw_date_filled & ~raw_income_filled & ~raw_expense_filled
    removed_rows = df.loc[removable_rows].copy()

    cleaned = df.loc[~removable_rows].copy()

    parsed_dates = pd.to_datetime(cleaned["Date"], errors="coerce", dayfirst=True)
    income_amounts = cleaned["K (KREDITS)"].apply(parse_amount)
    expense_amounts = cleaned["D (DEBETS)"].apply(parse_amount)

    income_filled = income_amounts.apply(has_numeric_amount)
    expense_filled = expense_amounts.apply(has_numeric_amount)

    transaction_type = pd.Series("Invalid", index=cleaned.index, dtype=object)
    transaction_type.loc[income_filled & ~expense_filled] = "Income"
    transaction_type.loc[expense_filled & ~income_filled] = "Expense"

    amount_income = pd.Series(0.0, index=cleaned.index, dtype="Float64")
    amount_expense = pd.Series(0.0, index=cleaned.index, dtype="Float64")
    amount_income.loc[income_filled] = income_amounts.loc[income_filled]
    amount_expense.loc[expense_filled] = expense_amounts.loc[expense_filled]

    signed_amount = pd.Series(pd.NA, index=cleaned.index, dtype="Float64")
    signed_amount.loc[transaction_type == "Income"] = amount_income.loc[transaction_type == "Income"]
    signed_amount.loc[transaction_type == "Expense"] = -amount_expense.loc[transaction_type == "Expense"]

    missing_description = ~cleaned["Purpose"].apply(is_filled)
    missing_counterparty = ~cleaned["Name Surname"].apply(is_filled)
    missing_date = parsed_dates.isna()
    raw_income_filled = cleaned["K (KREDITS)"].apply(is_filled)
    raw_expense_filled = cleaned["D (DEBETS)"].apply(is_filled)
    invalid_amount = (raw_income_filled & income_amounts.isna()) | (raw_expense_filled & expense_amounts.isna())

    cleaned["date"] = parsed_dates
    cleaned["excel_row"] = cleaned.index + 2
    cleaned["transaction_id"] = cleaned["excel_row"].apply(lambda row_number: f"bank_statement_2026_01_raw:{row_number}")
    cleaned["month"] = parsed_dates.dt.to_period("M").astype("string")
    cleaned["counterparty_raw"] = cleaned["Name Surname"]
    cleaned["personal_code"] = cleaned["Personal Code"]
    cleaned["account_number"] = cleaned["Konta numurs"]
    cleaned["bank_swift"] = cleaned["Bankas SWIFT"]
    cleaned["description"] = cleaned["Purpose"]
    cleaned["amount_income"] = amount_income
    cleaned["amount_expense"] = amount_expense
    cleaned["transaction_type"] = transaction_type
    cleaned["signed_amount"] = signed_amount
    cleaned["description_clean"] = cleaned["Purpose"].apply(clean_text)
    cleaned["counterparty_clean"] = cleaned["Name Surname"].apply(clean_text)
    cleaned["category"] = "Unclassified"
    cleaned["missing_description"] = missing_description
    cleaned["missing_counterparty"] = missing_counterparty
    cleaned["invalid_amount"] = invalid_amount
    cleaned["missing_date"] = missing_date

    validation_issues = build_validation_issues(
        cleaned=cleaned,
        parsed_dates=parsed_dates,
        income_filled=income_filled,
        expense_filled=expense_filled,
        income_amounts=income_amounts,
        expense_amounts=expense_amounts,
    )

    debug_summary = {
        "original_excel_row_count": original_row_count,
        "cleaned_dataframe_row_count": len(cleaned),
        "removed_rows_count": len(removed_rows),
        "removed_rows_reason": (
            "Fully empty Excel rows with no date and no K (KREDITS) or D (DEBETS) amount."
            if len(removed_rows) > 0
            else "No rows removed."
        ),
    }

    return cleaned.reset_index(drop=True), validation_issues, debug_summary


def build_validation_issues(
    cleaned: pd.DataFrame,
    parsed_dates: pd.Series,
    income_filled: pd.Series,
    expense_filled: pd.Series,
    income_amounts: pd.Series,
    expense_amounts: pd.Series,
) -> pd.DataFrame:
    issues: list[dict[str, Any]] = []

    for index, row in cleaned.iterrows():
        row_number = index + 2

        if pd.isna(parsed_dates.loc[index]):
            issues.append(issue(row_number, "missing date", "Date is empty or could not be parsed."))

        if not is_filled(row["Purpose"]):
            issues.append(issue(row_number, "missing description", "Purpose is empty."))

        if income_filled.loc[index] and expense_filled.loc[index]:
            issues.append(issue(row_number, "both K and D filled", "Income and expense columns are both populated."))

        if not income_filled.loc[index] and not expense_filled.loc[index]:
            issues.append(issue(row_number, "both K and D empty", "Income and expense columns are both empty."))

        if income_filled.loc[index] and pd.isna(income_amounts.loc[index]):
            issues.append(issue(row_number, "invalid amount", "K (KREDITS) is populated but not numeric."))

        if expense_filled.loc[index] and pd.isna(expense_amounts.loc[index]):
            issues.append(issue(row_number, "invalid amount", "D (DEBETS) is populated but not numeric."))

    return pd.DataFrame(issues, columns=["row_number", "issue", "details"])


def issue(row_number: int, issue_name: str, details: str) -> dict[str, Any]:
    return {
        "row_number": row_number,
        "issue": issue_name,
        "details": details,
    }


def is_filled(value: Any) -> bool:
    if pd.isna(value):
        return False

    return str(value).strip() != ""


def has_numeric_amount(value: Any) -> bool:
    if pd.isna(value):
        return False

    return float(value) != 0


def parse_amount(value: Any) -> float | pd.NA:
    if not is_filled(value):
        return pd.NA

    if isinstance(value, int | float):
        return float(value)

    normalized = str(value).strip()
    normalized = normalized.replace("\u00a0", "")
    normalized = normalized.replace(" ", "")

    if "," in normalized and "." in normalized:
        normalized = normalized.replace(".", "").replace(",", ".")
    else:
        normalized = normalized.replace(",", ".")

    try:
        return float(normalized)
    except ValueError:
        return pd.NA


def clean_text(value: Any) -> str:
    if not is_filled(value):
        return ""

    return " ".join(str(value).strip().split())
