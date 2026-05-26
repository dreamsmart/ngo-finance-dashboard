from __future__ import annotations

from pathlib import Path
from sqlite3 import connect
from typing import Any

import pandas as pd
import yaml


DEFAULT_CATEGORY = "Unclassified"
DEFAULT_PROJECT_NAME = "Unknown"
DESCRIPTION_CONFIDENCE = 1.0
COUNTERPARTY_CONFIDENCE = 0.75
CATEGORY_CONFIDENCE = 0.65
MANUAL_CONFIDENCE = 1.0
MANUAL_PRIORITY = -1000


def load_category_rules(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []

    with path.open("r", encoding="utf-8") as rules_file:
        config = yaml.safe_load(rules_file) or {}

    rules: list[dict[str, Any]] = []
    for group_name, group_rules in config.items():
        for raw_rule in group_rules or []:
            rules.append(
                {
                    "rule_id": "-".join(
                        [
                            "default",
                            "category",
                            str(raw_rule.get("transaction_type", group_name)).strip().lower(),
                            str(raw_rule.get("match_field", raw_rule.get("source", "any"))).strip().lower(),
                            str(raw_rule["keyword"]).strip().lower(),
                            str(raw_rule["category"]).strip().lower(),
                        ]
                    ),
                    "rule_type": "category",
                    "group": str(group_name),
                    "transaction_type": str(raw_rule.get("transaction_type", group_name)),
                    "priority": int(raw_rule["priority"]),
                    "keyword": str(raw_rule["keyword"]).lower(),
                    "label": str(raw_rule["category"]),
                    "source": str(raw_rule.get("match_field", raw_rule.get("source", "any"))),
                    "confidence": float(raw_rule.get("confidence", 0.0)),
                    "occurrence_count": int(raw_rule.get("occurrence_count", 0) or 0),
                    "origin": "default",
                    "rule_source": "default",
                    "created_at": "",
                }
            )

    return sorted(rules, key=lambda rule: rule["priority"])


def load_project_rules(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []

    with path.open("r", encoding="utf-8") as rules_file:
        raw_rules = yaml.safe_load(rules_file) or []

    rules: list[dict[str, Any]] = []
    for raw_rule in raw_rules:
        keyword = raw_rule.get("keyword")
        category = raw_rule.get("category")
        rules.append(
            {
                "rule_id": "-".join(
                    [
                        "default",
                        "project",
                        str(raw_rule.get("transaction_type", "any")).strip().lower(),
                        str(raw_rule.get("match_field", raw_rule.get("source", "any"))).strip().lower(),
                        str(keyword or category or "").strip().lower(),
                        str(raw_rule["project_name"]).strip().lower(),
                    ]
                ),
                "rule_type": "project",
                "transaction_type": str(raw_rule.get("transaction_type", "any")),
                "priority": int(raw_rule["priority"]),
                "keyword": str(keyword).lower() if keyword else "",
                "category": str(category) if category else "",
                "label": str(raw_rule["project_name"]),
                "source": str(raw_rule.get("match_field", raw_rule.get("source", "any"))),
                "confidence": float(raw_rule.get("confidence", 0.0)),
                "occurrence_count": int(raw_rule.get("occurrence_count", 0) or 0),
                "origin": "default",
                "rule_source": "default",
                "created_at": "",
            }
        )

    return sorted(rules, key=lambda rule: rule["priority"])


def load_manual_category_rules(database_path: Path) -> list[dict[str, Any]]:
    return load_manual_rules(database_path, "category")


def load_manual_project_rules(database_path: Path) -> list[dict[str, Any]]:
    return load_manual_rules(database_path, "project")


def load_manual_rules(database_path: Path, rule_type: str) -> list[dict[str, Any]]:
    ensure_rules_database(database_path)
    rules = load_manual_v2_rules(database_path, rule_type)

    if rule_type == "category":
        rules.extend(load_legacy_manual_category_rules(database_path))

    return sorted(rules, key=lambda rule: rule["priority"])


def load_manual_v2_rules(database_path: Path, rule_type: str) -> list[dict[str, Any]]:
    with connect(database_path) as connection:
        rows = connection.execute(
            """
            SELECT id, rule_type, transaction_type, source, keyword, label, priority, created_at
            FROM manual_rules_v2
            WHERE is_active = 1 AND rule_type = ?
            ORDER BY priority ASC, id ASC
            """,
            (rule_type,),
        ).fetchall()

    return [
        {
            "rule_id": f"manual-{row_id}",
            "rule_type": row_rule_type,
            "group": transaction_type,
            "transaction_type": transaction_type,
            "priority": int(priority),
            "keyword": str(keyword).lower(),
            "label": label,
            "source": source,
            "origin": "manual",
            "rule_source": "manual",
            "occurrence_count": 1,
            "confidence": MANUAL_CONFIDENCE,
            "created_at": created_at,
        }
        for row_id, row_rule_type, transaction_type, source, keyword, label, priority, created_at in rows
    ]


def load_legacy_manual_category_rules(database_path: Path) -> list[dict[str, Any]]:
    with connect(database_path) as connection:
        rows = connection.execute(
            """
            SELECT id, transaction_type, source, keyword, category, priority, created_at
            FROM manual_rules
            WHERE is_active = 1
            ORDER BY priority ASC, id ASC
            """
        ).fetchall()

    return [
        {
            "rule_id": f"legacy-manual-{row_id}",
            "rule_type": "category",
            "group": transaction_type,
            "transaction_type": transaction_type,
            "priority": int(priority),
            "keyword": str(keyword).lower(),
            "label": category,
            "source": source,
            "origin": "manual",
            "rule_source": "manual",
            "occurrence_count": 1,
            "confidence": MANUAL_CONFIDENCE,
            "created_at": created_at,
        }
        for row_id, transaction_type, source, keyword, category, priority, created_at in rows
    ]


def categorize_transactions(df: pd.DataFrame, rules: list[dict[str, Any]]) -> pd.DataFrame:
    categorized = df.copy()
    ensure_category_columns(categorized)

    for index, row in categorized.iterrows():
        if is_manual_label(row.get("category"), DEFAULT_CATEGORY):
            continue

        match = match_category(row, rules)
        categorized.at[index, "category"] = match["category"]
        categorized.at[index, "matched_rule"] = match["matched_rule"]
        categorized.at[index, "category_confidence"] = match["category_confidence"]

    return categorized


def assign_project_names(df: pd.DataFrame, rules: list[dict[str, Any]]) -> pd.DataFrame:
    projected = df.copy()
    ensure_project_columns(projected)

    for index, row in projected.iterrows():
        if is_manual_label(row.get("project_name"), DEFAULT_PROJECT_NAME):
            continue

        match = match_project(row, rules)
        projected.at[index, "project_name"] = match["project_name"]
        projected.at[index, "project_matched_rule"] = match["project_matched_rule"]
        projected.at[index, "project_confidence"] = match["project_confidence"]

    return projected


def apply_manual_correction_overrides(
    df: pd.DataFrame,
    database_path: Path,
    include_category: bool = True,
    include_project: bool = True,
) -> pd.DataFrame:
    ensure_rules_database(database_path)
    corrected = df.copy()

    if "transaction_id" not in corrected.columns:
        return corrected

    with connect(database_path) as connection:
        rows = connection.execute(
            """
            SELECT c.transaction_id, c.new_category, c.new_project_name
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

    overrides = {
        transaction_id: {
            "category": new_category,
            "project_name": new_project_name,
        }
        for transaction_id, new_category, new_project_name in rows
    }

    for index, row in corrected.iterrows():
        override = overrides.get(row.get("transaction_id"))
        if not override:
            continue

        if include_category and override.get("category"):
            corrected.at[index, "category"] = override["category"]
            corrected.at[index, "matched_rule"] = "manual correction"
            corrected.at[index, "category_confidence"] = MANUAL_CONFIDENCE

        if include_project and override.get("project_name"):
            corrected.at[index, "project_name"] = override["project_name"]
            corrected.at[index, "project_matched_rule"] = "manual correction"
            corrected.at[index, "project_confidence"] = MANUAL_CONFIDENCE

    return corrected


def save_manual_corrections(database_path: Path, corrections: pd.DataFrame) -> int:
    ensure_rules_database(database_path)
    saved_count = 0

    with connect(database_path) as connection:
        for _, correction in corrections.iterrows():
            transaction_id = str(correction.get("transaction_id", "")).strip()
            if not transaction_id:
                continue

            old_category = clean_label(correction.get("old_category", ""))
            new_category = clean_label(correction.get("new_category", ""))
            old_project_name = clean_label(correction.get("old_project_name", ""))
            new_project_name = clean_label(correction.get("new_project_name", ""))
            commentary = clean_label(correction.get("new_commentary", correction.get("commentary", "")))

            connection.execute(
                """
                INSERT INTO manual_corrections (
                    transaction_id,
                    old_category,
                    new_category,
                    old_project_name,
                    new_project_name,
                    edit_source,
                    commentary
                )
                VALUES (?, ?, ?, ?, ?, 'manual', ?)
                """,
                (
                    transaction_id,
                    old_category,
                    new_category,
                    old_project_name,
                    new_project_name,
                    commentary,
                ),
            )

            if new_category and new_category != old_category:
                upsert_manual_rule(connection, "category", correction, new_category)

            if new_project_name and new_project_name != old_project_name:
                upsert_manual_rule(connection, "project", correction, new_project_name)

            saved_count += 1

    return saved_count


def save_manual_category_rules(database_path: Path, corrections: pd.DataFrame) -> int:
    mapped = corrections.copy()
    mapped["old_category"] = DEFAULT_CATEGORY
    mapped["new_category"] = mapped["category"]
    mapped["old_project_name"] = ""
    mapped["new_project_name"] = ""
    return save_manual_corrections(database_path, mapped)


def upsert_manual_rule(connection: Any, rule_type: str, correction: pd.Series, label: str) -> None:
    keyword, source = manual_rule_source(correction)
    if not keyword:
        return

    transaction_type = str(correction.get("transaction_type", "")).strip()
    connection.execute(
        """
        INSERT INTO manual_rules_v2 (
            rule_type,
            transaction_type,
            source,
            keyword,
            label,
            priority,
            is_active
        )
        VALUES (?, ?, ?, ?, ?, ?, 1)
        ON CONFLICT(rule_type, transaction_type, source, keyword)
        DO UPDATE SET
            label = excluded.label,
            priority = excluded.priority,
            is_active = 1,
            updated_at = CURRENT_TIMESTAMP
        """,
        (
            rule_type,
            transaction_type,
            source,
            keyword,
            label,
            MANUAL_PRIORITY,
        ),
    )


def ensure_rules_database(database_path: Path) -> None:
    database_path.parent.mkdir(parents=True, exist_ok=True)

    with connect(database_path) as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS manual_rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                transaction_type TEXT NOT NULL,
                source TEXT NOT NULL,
                keyword TEXT NOT NULL,
                category TEXT NOT NULL,
                priority INTEGER NOT NULL DEFAULT -1000,
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(transaction_type, source, keyword)
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS manual_rules_v2 (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rule_type TEXT NOT NULL,
                transaction_type TEXT NOT NULL,
                source TEXT NOT NULL,
                keyword TEXT NOT NULL,
                label TEXT NOT NULL,
                priority INTEGER NOT NULL DEFAULT -1000,
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                reviewed_at TEXT,
                UNIQUE(rule_type, transaction_type, source, keyword)
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS manual_corrections (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                transaction_id TEXT NOT NULL,
                old_category TEXT,
                new_category TEXT,
                old_project_name TEXT,
                new_project_name TEXT,
                edited_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                edit_source TEXT NOT NULL DEFAULT 'manual',
                commentary TEXT
            )
            """
        )
        add_column_if_missing(connection, "manual_rules_v2", "reviewed_at", "TEXT")
        add_column_if_missing(connection, "manual_corrections", "commentary", "TEXT")
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS classification_rule_overrides (
                rule_id TEXT PRIMARY KEY,
                is_active INTEGER NOT NULL DEFAULT 1,
                priority_override INTEGER,
                reviewed_at TEXT,
                deleted_at TEXT,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )


def add_column_if_missing(connection: Any, table_name: str, column_name: str, column_type: str) -> None:
    columns = [row[1] for row in connection.execute(f"PRAGMA table_info({table_name})").fetchall()]
    if column_name not in columns:
        connection.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}")


def match_category(row: pd.Series, rules: list[dict[str, Any]]) -> dict[str, Any]:
    eligible_rules = filter_rules_for_transaction_type(rules, row.get("transaction_type", ""))

    description_match = find_keyword_match(row, eligible_rules, "description_clean")
    if description_match:
        return {
            "category": description_match["label"],
            "matched_rule": format_matched_rule(description_match, "description_clean"),
            "category_confidence": confidence_for_rule(description_match, DESCRIPTION_CONFIDENCE),
        }

    counterparty_match = find_keyword_match(row, eligible_rules, "counterparty_clean")
    if counterparty_match:
        return {
            "category": counterparty_match["label"],
            "matched_rule": format_matched_rule(counterparty_match, "counterparty_clean"),
            "category_confidence": confidence_for_rule(counterparty_match, COUNTERPARTY_CONFIDENCE),
        }

    return {
        "category": DEFAULT_CATEGORY,
        "matched_rule": "",
        "category_confidence": 0.0,
    }


def match_project(row: pd.Series, rules: list[dict[str, Any]]) -> dict[str, Any]:
    eligible_rules = filter_rules_for_transaction_type(rules, row.get("transaction_type", ""))

    for source in ["description_clean", "counterparty_clean", "category"]:
        match = find_project_match(row, eligible_rules, source)
        if match:
            return {
                "project_name": match["label"],
                "project_matched_rule": format_matched_rule(match, source),
                "project_confidence": project_confidence_for_source(match, source),
            }

    return {
        "project_name": DEFAULT_PROJECT_NAME,
        "project_matched_rule": "",
        "project_confidence": 0.0,
    }


def filter_rules_for_transaction_type(rules: list[dict[str, Any]], transaction_type: Any) -> list[dict[str, Any]]:
    normalized_type = normalize_match_text(transaction_type)
    matching_rules = [
        rule
        for rule in rules
        if normalize_match_text(rule.get("transaction_type", rule.get("group", "any"))) in {"any", normalized_type}
    ]
    return sorted(matching_rules or rules, key=rule_precedence_key)


def rule_precedence_key(rule: dict[str, Any]) -> tuple[int, int]:
    source = rule.get("rule_source") or rule.get("origin", "default")
    source_rank = {"manual": 0, "historical": 1, "default": 2}.get(str(source), 3)
    return source_rank, int(rule.get("priority", 999))


def find_keyword_match(row: pd.Series, rules: list[dict[str, Any]], source: str) -> dict[str, Any] | None:
    text = normalize_match_text(row.get(source, ""))
    if not text:
        return None

    for rule in rules:
        rule_source = normalize_match_text(rule.get("source", "any"))
        if rule_source not in {"any", source}:
            continue

        keyword = normalize_match_text(rule.get("keyword", ""))
        if keyword and keyword in text:
            return rule

    return None


def find_project_match(row: pd.Series, rules: list[dict[str, Any]], source: str) -> dict[str, Any] | None:
    if source == "category":
        category = normalize_match_text(row.get("category", ""))
        for rule in rules:
            rule_category = normalize_match_text(rule.get("category", ""))
            if rule_category and rule_category == category:
                return rule
        return None

    return find_keyword_match(row, rules, source)


def format_matched_rule(rule: dict[str, Any], source: str) -> str:
    origin = rule.get("origin", "yaml")
    keyword = rule.get("keyword") or rule.get("category", "")
    return f"{origin} {source}: {keyword}"


def confidence_for_rule(rule: dict[str, Any], fallback: float) -> float:
    if rule.get("origin") == "manual":
        return MANUAL_CONFIDENCE
    if float(rule.get("confidence", 0.0) or 0.0) > 0:
        return float(rule["confidence"])

    return fallback


def project_confidence_for_source(rule: dict[str, Any], source: str) -> float:
    if rule.get("origin") == "manual":
        return MANUAL_CONFIDENCE
    if float(rule.get("confidence", 0.0) or 0.0) > 0:
        return float(rule["confidence"])
    if source == "category":
        return CATEGORY_CONFIDENCE
    if source == "counterparty_clean":
        return COUNTERPARTY_CONFIDENCE
    return DESCRIPTION_CONFIDENCE


def ensure_category_columns(df: pd.DataFrame) -> None:
    if "category" not in df.columns:
        df["category"] = DEFAULT_CATEGORY
    if "matched_rule" not in df.columns:
        df["matched_rule"] = ""
    if "category_confidence" not in df.columns:
        df["category_confidence"] = 0.0


def ensure_project_columns(df: pd.DataFrame) -> None:
    if "project_name" not in df.columns:
        df["project_name"] = DEFAULT_PROJECT_NAME
    if "project_matched_rule" not in df.columns:
        df["project_matched_rule"] = ""
    if "project_confidence" not in df.columns:
        df["project_confidence"] = 0.0


def is_manual_label(value: Any, default_value: str) -> bool:
    if pd.isna(value):
        return False

    label = str(value).strip()
    return label != "" and label != default_value


def clean_label(value: Any) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().split())


def normalize_match_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().lower()


def manual_rule_source(row: pd.Series) -> tuple[str, str]:
    description = normalize_match_text(row.get("description_clean", ""))
    if description:
        return description, "description_clean"

    counterparty = normalize_match_text(row.get("counterparty_clean", ""))
    if counterparty:
        return counterparty, "counterparty_clean"

    category = normalize_match_text(row.get("category", ""))
    if category:
        return category, "category"

    return "", ""
