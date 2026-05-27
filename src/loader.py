from pathlib import Path

import pandas as pd


EXCEL_COLUMN_MAPPING = {
    "Date": "Date",
    "Name Surname": "Name Surname",
    "Personal Code": "Personal Code",
    "Column4": "Personal Code",
    "Konta numurs": "Konta numurs",
    "Column3": "Konta numurs",
    "Bankas SWIFT": "Bankas SWIFT",
    "Column2": "Bankas SWIFT",
    "Unnamed: 5": "Purpose",
    "Purpose": "Purpose",
    "K (KREDITS)": "K (KREDITS)",
    "D (DEBETS)": "D (DEBETS)",
}


def load_excel_file(path: Path) -> pd.DataFrame:
    """Load a local Excel file while preserving the workbook's original columns."""
    if not path.exists():
        raise FileNotFoundError(path)

    if path.suffix.lower() not in {".xlsx", ".xls"}:
        raise ValueError("Expected an Excel file with .xlsx or .xls extension.")

    df = pd.read_excel(path, dtype=object)
    return df.rename(columns={column: EXCEL_COLUMN_MAPPING[column] for column in df.columns if column in EXCEL_COLUMN_MAPPING})
