from __future__ import annotations

import re
from collections import Counter, defaultdict
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import pandas as pd
import yaml

from src.cleaner import clean_bank_statement
from src.categorizer import DEFAULT_CATEGORY, DEFAULT_PROJECT_NAME
from src.loader import load_excel_file


DANGEROUS_KEYWORDS = {
    "a",
    "as",
    "bank",
    "card",
    "eur",
    "https",
    "jan",
    "komisija",
    "maksājumu",
    "par",
    "payment",
    "pirkums",
    "riga",
    "sia",
    "the",
    "transfer",
    "uzdevuma",
    "yf",
}
MIN_OCCURRENCES = 2
MIN_CONFIDENCE = 0.8
LOW_CONFIDENCE_THRESHOLD = 0.9
DUPLICATE_THRESHOLD = 0.82


def analyze_historical_file(path: Path) -> dict[str, Any]:
    raw = load_excel_file(path)
    normalized, _, _ = clean_bank_statement(raw)
    category_column = historical_category_column(raw)
    project_column = historical_project_column(raw)
    missing_columns = []
    if not category_column:
        missing_columns.append("Category")
    if not project_column:
        missing_columns.append("Project Name or Division")

    normalized["category"] = historical_labels_for_cleaned_rows(raw, normalized, category_column)
    normalized["project_name"] = historical_labels_for_cleaned_rows(raw, normalized, project_column)

    proposals = pd.DataFrame(
        mine_category_rules(normalized) + mine_project_rules(normalized)
    )
    if not proposals.empty:
        proposals = proposals.sort_values(
            ["rule_type", "confidence", "occurrence_count", "priority"],
            ascending=[True, False, False, True],
        ).reset_index(drop=True)
        proposals.insert(0, "proposal_id", proposals.index.map(lambda index: f"R{index + 1:04d}"))

    ambiguous = pd.DataFrame(find_ambiguous_patterns(normalized))
    low_confidence = proposals[proposals["confidence"] < LOW_CONFIDENCE_THRESHOLD].copy() if not proposals.empty else pd.DataFrame()

    return {
        "transactions": normalized,
        "proposals": proposals,
        "ambiguous": ambiguous,
        "low_confidence": low_confidence,
        "duplicate_categories": duplicate_labels(manual_labels(normalized["category"])),
        "duplicate_projects": duplicate_labels(manual_labels(normalized["project_name"])),
        "taxonomy_categories": taxonomy_counts(normalized, "category", "Category"),
        "taxonomy_projects": taxonomy_counts(normalized, "project_name", "Project Name"),
        "source_file": str(path),
        "source_file_name": path.name,
        "category_source_column": category_column,
        "project_source_column": project_column,
        "missing_columns": missing_columns,
        "summary": historical_source_summary(normalized, proposals, source_row_count=len(raw)),
    }


def apply_exact_historical_labels(transactions: pd.DataFrame, path: Path) -> pd.DataFrame:
    matched = transactions.copy()
    ensure_historical_match_columns(matched)

    if not path.exists():
        return matched

    analysis = analyze_historical_file(path)
    historical_transactions = analysis.get("transactions", pd.DataFrame())
    if historical_transactions.empty:
        return matched

    historical_lookup = exact_historical_lookup(historical_transactions)
    if not historical_lookup:
        return matched

    for index, row in matched.iterrows():
        historical = None
        matched_key = ""
        for match_key in exact_match_keys(row):
            historical = historical_lookup.get(match_key)
            if historical:
                matched_key = match_key
                break
        matched.at[index, "historical_match_key"] = matched_key
        if not historical:
            continue

        matched.at[index, "historical_exact_match"] = True
        matched.at[index, "historical_category"] = historical.get("category", "")
        matched.at[index, "historical_project_name"] = historical.get("project_name", "")
        matched.at[index, "rule_source"] = "historical_exact"

        if has_manual_label(historical.get("category", "")):
            matched.at[index, "category"] = historical["category"]
            matched.at[index, "matched_rule"] = "historical_exact"
            matched.at[index, "category_confidence"] = 1.0
            matched.at[index, "category_rule_source"] = "historical_exact"
            matched.at[index, "historical_category_applied"] = True

        if has_manual_label(historical.get("project_name", "")):
            matched.at[index, "project_name"] = historical["project_name"]
            matched.at[index, "project_matched_rule"] = "historical_exact"
            matched.at[index, "project_confidence"] = 1.0
            matched.at[index, "project_rule_source"] = "historical_exact"
            matched.at[index, "historical_project_applied"] = True

    return matched


def load_historical_category_rules(path: Path) -> list[dict[str, Any]]:
    return proposal_rules_for_categorizer(path, "category")


def load_historical_project_rules(path: Path) -> list[dict[str, Any]]:
    return proposal_rules_for_categorizer(path, "project")


def proposal_rules_for_categorizer(path: Path, rule_type: str) -> list[dict[str, Any]]:
    if not path.exists():
        return []

    proposals = analyze_historical_file(path)["proposals"]
    if proposals.empty:
        return []

    rows = proposals[proposals["rule_type"].eq(rule_type)]
    rules: list[dict[str, Any]] = []
    for _, row in rows.iterrows():
        label = row.get("category") if rule_type == "category" else row.get("project_name")
        if not has_manual_label(label):
            continue
        keyword = str(row.get("keyword", "")).lower()
        source = row.get("match_field", "any")
        transaction_type = row.get("transaction_type", "any")

        rule = {
            "rule_id": "-".join(
                [
                    "historical",
                    rule_type,
                    normalize_label(transaction_type),
                    normalize_label(source),
                    normalize_label(keyword),
                    normalize_label(label),
                ]
            ),
            "rule_type": rule_type,
            "group": transaction_type,
            "transaction_type": transaction_type,
            "priority": int(row.get("priority", 999)),
            "keyword": keyword,
            "label": label,
            "source": source,
            "confidence": float(row.get("confidence", 0.0) or 0.0),
            "origin": "historical",
            "rule_source": "historical",
            "occurrence_count": int(row.get("occurrence_count", 0) or 0),
            "created_at": "",
        }
        if rule_type == "project" and has_manual_label(row.get("category")):
            rule["category"] = row.get("category")
        rules.append(rule)

    return sorted(rules, key=lambda rule: rule["priority"])


def mine_category_rules(df: pd.DataFrame) -> list[dict[str, Any]]:
    return mine_label_rules(
        df=df[df["category"].map(has_manual_label)],
        rule_type="category",
        label_column="category",
        output_field="category",
    )


def mine_project_rules(df: pd.DataFrame) -> list[dict[str, Any]]:
    rules = mine_label_rules(
        df=df[df["project_name"].map(has_manual_label)],
        rule_type="project",
        label_column="project_name",
        output_field="project_name",
    )
    rules.extend(mine_project_category_rules(df))
    return rules


def mine_label_rules(
    df: pd.DataFrame,
    rule_type: str,
    label_column: str,
    output_field: str,
) -> list[dict[str, Any]]:
    rules: list[dict[str, Any]] = []
    for match_field in ["description_clean", "counterparty_clean"]:
        buckets: dict[tuple[str, str], list[str]] = defaultdict(list)

        for _, row in df.iterrows():
            transaction_type = clean_label(row.get("transaction_type", ""))
            label = preserve_historical_label(row.get(label_column, ""))
            for keyword in candidate_keywords(row.get(match_field, "")):
                buckets[(transaction_type, keyword)].append(label)

        for (transaction_type, keyword), labels in buckets.items():
            rule = build_rule_from_labels(
                labels=labels,
                keyword=keyword,
                match_field=match_field,
                transaction_type=transaction_type,
                rule_type=rule_type,
                output_field=output_field,
            )
            if rule:
                rules.append(rule)

    return dedupe_rules(rules)


def mine_project_category_rules(df: pd.DataFrame) -> list[dict[str, Any]]:
    buckets: dict[tuple[str, str], list[str]] = defaultdict(list)
    for _, row in df[df["project_name"].map(has_manual_label)].iterrows():
        category = preserve_historical_label(row.get("category", ""))
        if not has_manual_label(category):
            continue
        transaction_type = clean_label(row.get("transaction_type", ""))
        buckets[(transaction_type, category)].append(preserve_historical_label(row.get("project_name", "")))

    rules = []
    for (transaction_type, category), project_names in buckets.items():
        if len(project_names) < 3:
            continue
        top_label, confidence = top_label_confidence(project_names)
        if confidence < 0.95:
            continue
        rules.append(
            {
                "rule_type": "project",
                "keyword": "",
                "match_field": "category",
                "transaction_type": transaction_type,
                "category": category,
                "project_name": top_label,
                "priority": 500,
                "confidence": round(confidence, 2),
                "occurrence_count": len(project_names),
            }
        )
    return rules


def build_rule_from_labels(
    labels: list[str],
    keyword: str,
    match_field: str,
    transaction_type: str,
    rule_type: str,
    output_field: str,
) -> dict[str, Any] | None:
    if len(labels) < MIN_OCCURRENCES or is_dangerous_keyword(keyword):
        return None

    label, confidence = top_label_confidence(labels)
    if confidence < MIN_CONFIDENCE:
        return None

    priority = suggested_priority(match_field, confidence, len(labels), keyword)
    return {
        "rule_type": rule_type,
        "keyword": keyword,
        "match_field": match_field,
        "transaction_type": transaction_type,
        output_field: label,
        "priority": priority,
        "confidence": round(confidence, 2),
        "occurrence_count": len(labels),
    }


def candidate_keywords(value: Any) -> set[str]:
    text = clean_label(value).lower()
    if not text:
        return set()

    tokens = [
        token
        for token in re.findall(r"[a-zA-ZāčēģīķļņšūžĀČĒĢĪĶĻŅŠŪŽ]+", text)
        if len(token) >= 4 and not is_dangerous_keyword(token)
    ]
    candidates = set()

    for size in range(1, min(4, len(tokens)) + 1):
        for index in range(0, len(tokens) - size + 1):
            phrase_tokens = tokens[index : index + size]
            if any(token in {"eur", "https", "pirkums"} for token in phrase_tokens):
                continue
            phrase = " ".join(phrase_tokens)
            if not is_dangerous_keyword(phrase):
                candidates.add(phrase)

    raw_tokens = re.findall(r"[a-zA-ZāčēģīķļņšūžĀČĒĢĪĶĻŅŠŪŽ]+", text)
    if len(text) <= 60 and not is_dangerous_keyword(text) and not any(
        is_dangerous_keyword(token) for token in raw_tokens
    ):
        candidates.add(text)

    return candidates


def find_ambiguous_patterns(df: pd.DataFrame) -> list[dict[str, Any]]:
    rows = []
    for label_column, rule_type in [("category", "category"), ("project_name", "project")]:
        for match_field in ["description_clean", "counterparty_clean"]:
            buckets: dict[str, list[str]] = defaultdict(list)
            for _, row in df.iterrows():
                label = preserve_historical_label(row.get(label_column, ""))
                if not has_manual_label(label):
                    continue
                for keyword in candidate_keywords(row.get(match_field, "")):
                    buckets[keyword].append(label)

            for keyword, labels in buckets.items():
                if len(labels) < MIN_OCCURRENCES:
                    continue
                top_label, confidence = top_label_confidence(labels)
                if confidence < MIN_CONFIDENCE and len(set(labels)) > 1:
                    rows.append(
                        {
                            "rule_type": rule_type,
                            "keyword": keyword,
                            "match_field": match_field,
                            "top_label": top_label,
                            "confidence": round(confidence, 2),
                            "occurrence_count": len(labels),
                            "labels_seen": ", ".join(repr(label) for label in sorted(set(labels))),
                        }
                    )
    return rows


def top_label_confidence(labels: list[str]) -> tuple[str, float]:
    counts = Counter(label for label in labels if has_manual_label(label))
    if not counts:
        return "", 0.0
    label, count = counts.most_common(1)[0]
    return label, count / len(labels)


def suggested_priority(match_field: str, confidence: float, occurrence_count: int, keyword: str) -> int:
    base = 200 if match_field == "description_clean" else 300
    if confidence >= 0.95:
        base -= 60
    if occurrence_count >= 5:
        base -= 30
    if len(keyword.split()) >= 2:
        base -= 20
    return max(20, base)


def dedupe_rules(rules: list[dict[str, Any]]) -> list[dict[str, Any]]:
    best: dict[tuple[str, str, str, str, str], dict[str, Any]] = {}
    for rule in rules:
        label = rule.get("category") or rule.get("project_name")
        key = (
            rule["rule_type"],
            rule["transaction_type"],
            rule["match_field"],
            rule["keyword"],
            label,
        )
        current = best.get(key)
        if current is None or (rule["confidence"], rule["occurrence_count"]) > (
            current["confidence"],
            current["occurrence_count"],
        ):
            best[key] = rule
    return list(best.values())


def append_approved_rules(
    category_rules_path: Path,
    project_rules_path: Path,
    approved_rules: pd.DataFrame,
) -> tuple[int, int]:
    category_rules = load_yaml(category_rules_path, default={})
    project_rules = load_yaml(project_rules_path, default=[])
    category_count = 0
    project_count = 0

    for _, row in approved_rules.iterrows():
        rule_type = row["rule_type"]
        transaction_type = row["transaction_type"] or "any"
        if rule_type == "category":
            group = transaction_type if transaction_type in {"Income", "Expense"} else "Any"
            category_rules.setdefault(group, [])
            new_rule = {
                "keyword": row["keyword"],
                "match_field": row["match_field"],
                "transaction_type": transaction_type,
                "category": row["category"],
                "priority": int(row["priority"]),
                "confidence": float(row["confidence"]),
            }
            if new_rule not in category_rules[group]:
                category_rules[group].append(new_rule)
                category_count += 1
        elif rule_type == "project":
            new_rule = {
                "match_field": row["match_field"],
                "transaction_type": transaction_type,
                "project_name": row["project_name"],
                "priority": int(row["priority"]),
                "confidence": float(row["confidence"]),
            }
            if clean_label(row.get("keyword", "")):
                new_rule["keyword"] = row["keyword"]
            if has_manual_label(row.get("category", "")):
                new_rule["category"] = row["category"]
            if new_rule not in project_rules:
                project_rules.append(new_rule)
                project_count += 1

    write_yaml(category_rules_path, category_rules)
    write_yaml(project_rules_path, project_rules)
    return category_count, project_count


def load_yaml(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    with path.open("r", encoding="utf-8") as file:
        return yaml.safe_load(file) or default


def write_yaml(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as file:
        yaml.safe_dump(data, file, allow_unicode=True, sort_keys=False)


def duplicate_labels(values: list[str] | Any) -> pd.DataFrame:
    labels = sorted({preserve_historical_label(value) for value in values if has_manual_label(value)})
    rows = []
    for index, first in enumerate(labels):
        for second in labels[index + 1 :]:
            score = SequenceMatcher(None, normalize_label(first), normalize_label(second)).ratio()
            if score >= DUPLICATE_THRESHOLD:
                rows.append({"value": first, "similar_value": second, "similarity": round(score, 2)})
    return pd.DataFrame(rows, columns=["value", "similar_value", "similarity"])


def is_dangerous_keyword(keyword: str) -> bool:
    normalized = normalize_label(keyword)
    if not normalized or normalized in DANGEROUS_KEYWORDS:
        return True
    if len(normalized) < 4:
        return True
    if re.fullmatch(r"\d+", normalized):
        return True
    return False


def clean_label(value: Any) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().split())


def historical_project_column(df: pd.DataFrame) -> str:
    candidates = {
        "project name": "Project Name",
        "division": "Division",
        "project_name": "project_name",
    }
    for column in df.columns:
        normalized = str(column).strip().lower()
        if normalized in candidates:
            return column
    return ""


def historical_category_column(df: pd.DataFrame) -> str:
    for column in df.columns:
        if str(column).strip().lower() == "category":
            return column
    return ""


def historical_labels_for_cleaned_rows(
    raw: pd.DataFrame,
    normalized: pd.DataFrame,
    source_column: str,
) -> pd.Series:
    if not source_column:
        return pd.Series([""] * len(normalized), index=normalized.index, dtype=object)

    raw_indexes = (normalized["excel_row"] - 2).astype(int)
    labels = raw.loc[raw_indexes, source_column].reset_index(drop=True)
    return labels.map(preserve_historical_label)


def exact_historical_lookup(historical_transactions: pd.DataFrame) -> dict[str, dict[str, str]]:
    buckets: dict[str, list[dict[str, str]]] = defaultdict(list)
    for _, row in historical_transactions.iterrows():
        category = preserve_historical_label(row.get("category", ""))
        project_name = preserve_historical_label(row.get("project_name", ""))
        labels = {"category": category, "project_name": project_name}
        for key in exact_match_keys(row):
            if key:
                buckets[key].append(labels)

    lookup: dict[str, dict[str, str]] = {}
    for key, labels_for_key in buckets.items():
        category_labels = {item["category"] for item in labels_for_key if has_manual_label(item["category"])}
        project_labels = {item["project_name"] for item in labels_for_key if has_manual_label(item["project_name"])}
        if len(category_labels) > 1 or len(project_labels) > 1:
            continue

        lookup[key] = {
            "category": next(iter(category_labels), ""),
            "project_name": next(iter(project_labels), ""),
        }

    return lookup


def exact_match_keys(row: pd.Series) -> list[str]:
    keys = [
        exact_match_key(row, include_date=True, include_counterparty=True),
        exact_match_key(row, include_date=False, include_counterparty=True),
        exact_match_key(row, include_date=True, include_counterparty=False),
        exact_match_key(row, include_date=False, include_counterparty=False),
    ]
    return list(dict.fromkeys(key for key in keys if key))


def exact_match_key(
    row: pd.Series,
    include_counterparty: bool = True,
    include_date: bool = True,
) -> str:
    date_value = pd.to_datetime(row.get("date"), errors="coerce")
    date_part = "" if pd.isna(date_value) else date_value.date().isoformat()
    amount_value = signed_amount_for_match(row)
    amount_part = "" if pd.isna(amount_value) else f"{float(amount_value):.2f}"
    description_part = normalize_key_text(row.get("description", row.get("Purpose", "")))
    counterparty_part = normalize_key_text(row.get("counterparty_raw", row.get("Name Surname", "")))
    parts = []
    if include_date:
        parts.append(date_part)
    parts.extend([description_part, amount_part])
    if include_counterparty:
        parts.append(counterparty_part)
    return "|".join(parts)


def signed_amount_for_match(row: pd.Series) -> float | pd.NA:
    signed_amount = pd.to_numeric(row.get("signed_amount"), errors="coerce")
    if not pd.isna(signed_amount):
        return float(signed_amount)

    income = pd.to_numeric(row.get("amount_income", row.get("K (KREDITS)")), errors="coerce")
    expense = pd.to_numeric(row.get("amount_expense", row.get("D (DEBETS)")), errors="coerce")
    income = 0.0 if pd.isna(income) else float(income)
    expense = 0.0 if pd.isna(expense) else float(expense)
    if income != 0 and expense == 0:
        return income
    if expense != 0 and income == 0:
        return -expense
    return pd.NA


def normalize_key_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().split())


def ensure_historical_match_columns(df: pd.DataFrame) -> None:
    defaults = {
        "historical_exact_match": False,
        "historical_category": "",
        "historical_project_name": "",
        "historical_match_key": "",
        "historical_category_applied": False,
        "historical_project_applied": False,
        "category_rule_source": "",
        "project_rule_source": "",
        "rule_source": "",
    }
    for column, default in defaults.items():
        if column not in df.columns:
            df[column] = default


def preserve_historical_label(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value)


def has_manual_label(value: Any) -> bool:
    label = preserve_historical_label(value)
    return bool(label.strip()) and label not in {DEFAULT_CATEGORY, DEFAULT_PROJECT_NAME}


def manual_labels(values: pd.Series) -> list[str]:
    return [preserve_historical_label(value) for value in values if has_manual_label(value)]


def taxonomy_counts(df: pd.DataFrame, column: str, display_name: str) -> pd.DataFrame:
    labels = df[df[column].map(has_manual_label)][column]
    if labels.empty:
        return pd.DataFrame(columns=[display_name, "Count"])

    return (
        labels.value_counts(dropna=False)
        .rename_axis(display_name)
        .reset_index(name="Count")
    )


def historical_source_summary(
    df: pd.DataFrame,
    proposals: pd.DataFrame,
    source_row_count: int,
) -> dict[str, Any]:
    category_filled = df["category"].map(has_manual_label)
    project_filled = df["project_name"].map(has_manual_label)
    return {
        "historical_rows": int(source_row_count),
        "category_filled_rows": int(category_filled.sum()),
        "project_filled_rows": int(project_filled.sum()),
        "manually_categorized_rows": int((category_filled | project_filled).sum()),
        "unique_categories": int(df.loc[category_filled, "category"].nunique()),
        "unique_project_names": int(df.loc[project_filled, "project_name"].nunique()),
        "unique_category_values": sorted(df.loc[category_filled, "category"].dropna().unique().tolist()),
        "unique_project_name_values": sorted(df.loc[project_filled, "project_name"].dropna().unique().tolist()),
        "extracted_historical_rules": int(len(proposals)),
    }


def normalize_label(value: str) -> str:
    normalized = clean_label(value).lower()
    normalized = re.sub(r"[^\w\s]", "", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    if normalized.endswith("s"):
        normalized = normalized[:-1]
    return normalized
