"""
excel_utils.py
--------------
Shared helpers for loading and saving Excel files.
"""

import io
from typing import Optional, Set

import pandas as pd


BIZNET_SHEET_NAME = "הזמנות BIZNET"
COORD_DATE_COL = "תאריך מתואם"


def load_sheets(uploaded_file, sheet_names: list[str]) -> dict[str, pd.DataFrame]:
    """
    Load specific sheets from an uploaded Streamlit file object.

    Parameters
    ----------
    uploaded_file : UploadedFile
        The file object from st.file_uploader.
    sheet_names : list[str]
        List of sheet names to load.

    Returns
    -------
    dict[str, pd.DataFrame]
        Mapping of sheet name → DataFrame (columns stripped of whitespace).
    """
    result = {}
    for name in sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=name, dtype=str, header=1)
        # Strip surrounding whitespace from column names and string values
        df.columns = [str(c).strip() for c in df.columns]
        df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)
        # Drop repeated header rows (the file sometimes has a header row at the bottom)
        order_col = "מספר הזמנה"
        if order_col in df.columns:
            df = df[df[order_col] != order_col].reset_index(drop=True)
        result[name] = df
    return result


def dfs_to_excel_bytes(
    sheets: dict[str, pd.DataFrame],
    text_columns: Optional[Set[str]] = None,
) -> bytes:
    """
    Serialize one or more DataFrames into an in-memory Excel file.

    Parameters
    ----------
    sheets : dict[str, pd.DataFrame]
        Mapping of sheet name → DataFrame to write.
    text_columns : set[str] | None
        Column names to force as Excel text cells.

    Returns
    -------
    bytes
        Raw bytes of the .xlsx file, ready for st.download_button.
    """
    text_columns = text_columns or set()
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            sheet_text_columns = set(text_columns)
            if sheet_name == BIZNET_SHEET_NAME:
                sheet_text_columns.add(COORD_DATE_COL)
            for col_idx, col_name in enumerate(df.columns, start=1):
                if col_name not in sheet_text_columns:
                    continue
                for row_idx in range(2, len(df) + 2):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    cell.number_format = "@"
                    if cell.value is not None:
                        cell.value = str(cell.value)
    return buffer.getvalue()
