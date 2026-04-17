from pathlib import Path
from typing import List, Dict
import pandas as pd


def write_excel(output_path: Path, folders_data: List[Dict]) -> None:
    """Write summary and detail tables to a single Excel sheet.

    Layout of sheet ``QR_Results``:
        - **Table 1 – Summary**   : one row per folder
        - Blank separator row
        - **Table 2 – Details**   : one row per image file

    Column widths are auto-adjusted based on cell content.
    """
    summary_rows = [
        {
            "Folder": fd["folder_name"],
            "Date": fd["reading_date"],
            "Total Files": fd["total_files"],
            "Success": fd["success_count"],
            "Fail": fd["fail_count"],
        }
        for fd in folders_data
    ]
    summary_df = pd.DataFrame(summary_rows)

    detail_rows: List[Dict] = []
    for fd in folders_data:
        detail_rows.extend(fd["details"])

    detail_df = pd.DataFrame(
        detail_rows,
        columns=[
            "folder",
            "file_name",
            "qr_content",
            "col1",
            "TEN DU AN",
            "TEN COT",
            "KIEN SO",
            "SO CHI TIET",
            "KL TINH",
            "status",
        ],
    )

    with pd.ExcelWriter(str(output_path), engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="QR_Results", index=False, startrow=0)
        detail_start_row = len(summary_df) + 2
        detail_df.to_excel(
            writer, sheet_name="QR_Results", index=False, startrow=detail_start_row
        )

        worksheet = writer.sheets["QR_Results"]
        for column_cells in worksheet.columns:
            max_length = 0
            col_letter = column_cells[0].column_letter
            for cell in column_cells:
                try:
                    cell_len = len(str(cell.value)) if cell.value is not None else 0
                    if cell_len > max_length:
                        max_length = cell_len
                except Exception:
                    pass
            adjusted_width = min(max_length + 4, 80)
            worksheet.column_dimensions[col_letter].width = adjusted_width
