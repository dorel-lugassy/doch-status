"""
excel_utils.py
--------------
Shared helpers for loading and saving Excel files.
"""

import io
import pandas as pd


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


def dfs_to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    """
    Serialize one or more DataFrames into an in-memory Excel file.

    Parameters
    ----------
    sheets : dict[str, pd.DataFrame]
        Mapping of sheet name → DataFrame to write.

    Returns
    -------
    bytes
        Raw bytes of the .xlsx file, ready for st.download_button.
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buffer.getvalue()
