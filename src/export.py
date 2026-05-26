from pathlib import Path

import pandas as pd


def export_to_excel(df: pd.DataFrame, output_path: Path) -> Path:
    """Export cleaned transactions to an Excel workbook."""
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Transactions", index=False)

    return output_path
