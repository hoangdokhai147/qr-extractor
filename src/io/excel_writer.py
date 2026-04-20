from pathlib import Path
from typing import List, Dict
import pandas as pd


def write_excel(output_path: Path, folders_data: List[Dict]) -> None:
    """Write extracted QR data to a single Excel sheet.

    Column widths are auto-adjusted based on cell content.
    """
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
            "ten_du_an",
            "ten_cot",
            "kien_so",
            "so_chi_tiet",
            "kl_tinh",
            "status",
        ],
    )

    with pd.ExcelWriter(str(output_path), engine="openpyxl") as writer:
        detail_df.to_excel(
            writer, sheet_name="QR_Results", index=False, startrow=0
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
