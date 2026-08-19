"""
excel_utils.py
--------------
Shared helpers for loading and saving Excel files.
"""

import io
import posixpath
import re
import zipfile
from typing import Optional, Set
from xml.etree import ElementTree as ET

import pandas as pd


BIZNET_SHEET_NAME = "הזמנות BIZNET"
COORD_DATE_COL = "תאריך מתואם"
FILTERED_ORIGINAL_SELLERS = ("אליאור ביטון", "זהבה בלאי", "עומר בר מוחה")
SELLER_FILTER_COLUMNS = {
    "סיבים": 7,      # Column G: שם מוכרן
    "נחושת": 7,      # Column G: שם מוכרן
    "כל השאר": 9,   # Column I
}


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


def _cell_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _uploaded_file_bytes(uploaded_file) -> bytes:
    if isinstance(uploaded_file, bytes):
        return uploaded_file
    return uploaded_file.getvalue()


def _filter_worksheet_by_seller(worksheet, seller_col_idx: int, allowed_sellers: set[str]) -> None:
    rows_to_delete = []
    for row_idx in range(worksheet.max_row, 2, -1):
        seller_name = _cell_text(worksheet.cell(row=row_idx, column=seller_col_idx).value)
        if seller_name not in allowed_sellers:
            rows_to_delete.append(row_idx)

    run_start = None
    previous_row = None
    for row_idx in rows_to_delete:
        if run_start is None:
            run_start = previous_row = row_idx
            continue

        if row_idx == previous_row - 1:
            previous_row = row_idx
            continue

        worksheet.delete_rows(previous_row, run_start - previous_row + 1)
        run_start = previous_row = row_idx

    if run_start is not None:
        worksheet.delete_rows(previous_row, run_start - previous_row + 1)


_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS = {"main": _MAIN_NS, "rel": _REL_NS}
_CELL_REF_RE = re.compile(r"^([A-Z]+)([0-9]+)$")

ET.register_namespace("", _MAIN_NS)
ET.register_namespace("r", _OFFICE_REL_NS)


def _column_letters(col_idx: int) -> str:
    letters = ""
    while col_idx:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    result = 0
    for char in letters:
        result = result * 26 + (ord(char.upper()) - 64)
    return result


def _renumber_cell_ref(cell_ref: str, row_idx: int) -> str:
    match = _CELL_REF_RE.match(cell_ref)
    if not match:
        return cell_ref
    return f"{match.group(1)}{row_idx}"


def _read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []

    values = []
    for item in root.findall("main:si", _NS):
        values.append("".join(item.itertext()))
    return values


def _shared_cell_text(cell, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "s":
        value_node = cell.find("main:v", _NS)
        if value_node is None or value_node.text is None:
            return ""
        try:
            return _cell_text(shared_strings[int(value_node.text)])
        except (ValueError, IndexError):
            return ""
    if cell_type == "inlineStr":
        inline = cell.find("main:is", _NS)
        return _cell_text("".join(inline.itertext()) if inline is not None else "")

    value_node = cell.find("main:v", _NS)
    return _cell_text(value_node.text if value_node is not None else "")


def _workbook_sheet_paths(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels_root.findall("rel:Relationship", _NS)
    }

    sheet_paths = {}
    for sheet in workbook_root.find("main:sheets", _NS):
        rel_id = sheet.attrib[f"{{{_OFFICE_REL_NS}}}id"]
        target = rel_targets[rel_id]
        if target.startswith("/"):
            sheet_path = target.lstrip("/")
        else:
            sheet_path = posixpath.normpath(posixpath.join("xl", target))
        sheet_paths[sheet.attrib["name"]] = sheet_path
    return sheet_paths


def _filter_sheet_xml(sheet_xml: bytes, seller_col_idx: int, allowed_sellers: set[str], shared_strings: list[str]) -> bytes:
    root = ET.fromstring(sheet_xml)
    sheet_data = root.find("main:sheetData", _NS)
    if sheet_data is None:
        return sheet_xml

    kept_rows = []
    for row in list(sheet_data.findall("main:row", _NS)):
        original_row_idx = int(row.attrib.get("r", "0") or 0)
        if original_row_idx <= 2:
            kept_rows.append(row)
            continue

        seller_value = ""
        for cell in row.findall("main:c", _NS):
            if _column_index(cell.attrib.get("r", "")) == seller_col_idx:
                seller_value = _shared_cell_text(cell, shared_strings)
                break
        if seller_value in allowed_sellers:
            kept_rows.append(row)

    sheet_data[:] = kept_rows
    for new_row_idx, row in enumerate(kept_rows, start=1):
        row.attrib["r"] = str(new_row_idx)
        for cell in row.findall("main:c", _NS):
            cell_ref = cell.attrib.get("r")
            if cell_ref:
                cell.attrib["r"] = _renumber_cell_ref(cell_ref, new_row_idx)

    dimension = root.find("main:dimension", _NS)
    if dimension is not None:
        max_col = 1
        for row in kept_rows:
            for cell in row.findall("main:c", _NS):
                max_col = max(max_col, _column_index(cell.attrib.get("r", "A1")))
        dimension.attrib["ref"] = f"A1:{_column_letters(max_col)}{len(kept_rows)}"

    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def filtered_original_workbook_bytes(uploaded_file) -> bytes:
    """
    Return a copy of the original workbook with only selected seller rows.

    The workbook package is copied as-is and only the relevant worksheet XML
    files are filtered, so sharedStrings, styles, relationships, and metadata
    stay as close as possible to the original source file.
    Rows 1-2 are kept; filtering starts from row 3.
    """
    source_bytes = _uploaded_file_bytes(uploaded_file)
    allowed_sellers = {_cell_text(seller) for seller in FILTERED_ORIGINAL_SELLERS}

    source_buffer = io.BytesIO(source_bytes)
    output_buffer = io.BytesIO()
    with zipfile.ZipFile(source_buffer, "r") as source_archive:
        sheet_paths = _workbook_sheet_paths(source_archive)
        shared_strings = _read_shared_strings(source_archive)
        filtered_paths = {}

        for sheet_name, seller_col_idx in SELLER_FILTER_COLUMNS.items():
            if sheet_name not in sheet_paths:
                raise ValueError(f"חסר גיליון נדרש עבור דוח מוכרנים: {sheet_name}")
            sheet_path = sheet_paths[sheet_name]
            filtered_paths[sheet_path] = _filter_sheet_xml(
                source_archive.read(sheet_path),
                seller_col_idx,
                allowed_sellers,
                shared_strings,
            )

        with zipfile.ZipFile(output_buffer, "w") as output_archive:
            for info in source_archive.infolist():
                data = filtered_paths.get(info.filename)
                if data is None:
                    data = source_archive.read(info.filename)
                output_archive.writestr(info, data)

    return output_buffer.getvalue()
