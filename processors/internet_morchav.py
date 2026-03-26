"""
internet_morchav.py
--------------------
Processor for "ניתוח סטטוס אינטרנט מורכב להרצה".

Input sheets
------------
- סיבים  : work-order rows for fiber orders
- נחושת  : work-order rows for copper orders  (same structure as סיבים)
- כל השאר: supplemental services per order

Output
------
Tuple of two DataFrames:
  1. result_df    – the main processed table
  2. exceptions_df – rows whose "סטטוס שירות" didn't match any known value
"""

import pandas as pd

# ── Column names in the source file ──────────────────────────────────────────
COL_ORDER_NUM   = "מספר הזמנה"
COL_CARD_NUM    = "מספר הכר"         # renamed to "מספר היכר" in output
COL_SVC_STATUS  = "סטטוס שירות"
COL_WO_STATUS   = "תיאור סטטוס הוראת עבודה"
COL_COORD_START = "תאום לקוח התחלה"
COL_CLOSE_DATE  = "תאריך סגירת הזמנה"
COL_SPEED       = "מהירות"

# כל השאר columns
COL_SERVICE      = "שירות"
COL_FULL_SVC     = "שירות מלא"
COL_REST_STATUS  = "סטטוס שירות"
COL_COORD_TASK   = "פרטי תאום לקוח במשימה אחרונה בהוראת עבודה"
COL_MINUTES      = "חבילת דקות"

# ── Known status values ───────────────────────────────────────────────────────
STATUS_CLOSED    = "סגור"
STATUS_OPEN      = "פתוח"
STATUS_CANCELLED = "מבוטל"

WO_STATUS_OPEN      = "פתוח"
WO_STATUS_CLOSED    = "סגור"
WO_STATUS_CANCELLED = "בוטל"

WIFI_MESH_VALUE  = "נקודות WIFI"
MESH_MSG         = "היה בהזמנה שלי ולא הותקן"

# ── Output column names ───────────────────────────────────────────────────────
OUT_ORDER_NUM    = "מספר הזמנה"
OUT_CARD_NUM     = "מספר היכר"
OUT_ORDER_STATUS = "סטטוס הזמנה"
OUT_INSTALL_DT   = "תאריך ושעת התקנה מעודכנים"
OUT_COORD_DATE   = "תאריך מתואם"
OUT_MINUTES      = "דקות קו"


def _get_last_workorder_per_order(df: pd.DataFrame) -> pd.DataFrame:
    """
    Keep only the LAST row for each unique order number.
    Preserves the original file order when grouping.
    """
    return df.groupby(COL_ORDER_NUM, sort=False).last().reset_index()


def _is_empty(val) -> bool:
    """
    Return True if a cell value is considered empty:
    NaN (float), None, or any string form of those ("nan", "none", "").
    Pandas reads empty Excel cells as NaN even with dtype=str,
    so we must handle all these cases.
    """
    if val is None:
        return True
    try:
        import math
        if math.isnan(float(val)):
            return True
    except (TypeError, ValueError):
        pass
    return str(val).strip().lower() in ("", "nan", "none", "nat")


def _classify_order(row: pd.Series) -> dict:
    """
    Apply the business rules to a single row.
    - OUT_ORDER_STATUS is determined solely by COL_SVC_STATUS (סטטוס שירות).
    - Dates are taken from the last work-order row as usual.
    Returns a dict with the output fields.
    """
    svc_status = str(row.get(COL_SVC_STATUS, "")).strip()
    coord_date = row.get(COL_COORD_START, "")
    close_date = row.get(COL_CLOSE_DATE,  "")

    result = {
        OUT_ORDER_STATUS: None,
        OUT_INSTALL_DT:   "",
        OUT_COORD_DATE:   "",
        "_is_exception":  False,
    }

    if svc_status == STATUS_CLOSED:
        result[OUT_ORDER_STATUS] = STATUS_CLOSED
        result[OUT_COORD_DATE]   = close_date

    elif svc_status == STATUS_OPEN:
        result[OUT_ORDER_STATUS] = STATUS_OPEN
        result[OUT_COORD_DATE]   = coord_date
        result[OUT_INSTALL_DT]   = coord_date

    elif svc_status == STATUS_CANCELLED:
        result[OUT_ORDER_STATUS] = STATUS_CANCELLED

    else:
        # Unknown סטטוס שירות → goes to exceptions sheet
        result[OUT_ORDER_STATUS] = svc_status
        result["_is_exception"]  = True

    return result


def _build_mesh_lookup(rest_df: pd.DataFrame) -> set[str]:
    """
    Return the set of order numbers where:
      שירות מלא == "נקודות WIFI"  AND  סטטוס שירות == "מבוטל"
    """
    mask = (
        (rest_df[COL_FULL_SVC].str.strip()   == WIFI_MESH_VALUE) &
        (rest_df[COL_REST_STATUS].str.strip() == STATUS_CANCELLED)
    )
    return set(rest_df.loc[mask, COL_ORDER_NUM].str.strip())


BIZNET_SERVICE_VALUE = "BIZNET"
PHONE_SERVICE_VALUE  = "PHONE"


def _build_biznet_rows(rest_df: pd.DataFrame) -> list[dict]:
    """
    Extract rows from 'כל השאר' where שירות == BIZNET and map them
    to the standard output column structure.
    All matching rows are included (no deduplication).
    """
    biznet_mask = rest_df[COL_SERVICE].str.strip() == BIZNET_SERVICE_VALUE
    biznet_df   = rest_df[biznet_mask].copy()

    rows = []
    for _, row in biznet_df.iterrows():
        rows.append({
            OUT_ORDER_NUM:    str(row.get(COL_ORDER_NUM,   "")).strip(),
            OUT_CARD_NUM:     str(row.get(COL_CARD_NUM,    "")).strip(),
            OUT_ORDER_STATUS: str(row.get(COL_REST_STATUS, "")).strip(),
            OUT_INSTALL_DT:   "",
            OUT_COORD_DATE:   str(row.get(COL_COORD_TASK,  "")).strip()
                              if not _is_empty(row.get(COL_COORD_TASK)) else "",
        })
    return rows


def _build_phone_rows(rest_df: pd.DataFrame) -> list[dict]:
    """
    Extract rows from 'כל השאר' where שירות == PHONE.
    Returns rows for the separate phone-line report.
    """
    phone_mask = (
        (rest_df[COL_SERVICE].str.strip() == PHONE_SERVICE_VALUE) &
        (rest_df[COL_MINUTES].notna()) &
        (rest_df[COL_MINUTES].str.strip() != "")
    )
    phone_df   = rest_df[phone_mask].copy()

    rows = []
    for _, row in phone_df.iterrows():
        rows.append({
            OUT_ORDER_NUM:    str(row.get(COL_ORDER_NUM,   "")).strip(),
            OUT_CARD_NUM:     str(row.get(COL_CARD_NUM,    "")).strip(),
            OUT_ORDER_STATUS: str(row.get(COL_REST_STATUS, "")).strip(),
            OUT_COORD_DATE:   str(row.get(COL_COORD_TASK,  "")).strip()
                              if not _is_empty(row.get(COL_COORD_TASK)) else "",
            OUT_MINUTES:      str(row.get(COL_MINUTES,     "")).strip()
                              if not _is_empty(row.get(COL_MINUTES)) else "",
        })
    return rows


def run(
    fiber_df:  pd.DataFrame,
    copper_df: pd.DataFrame,
    rest_df:   pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Main entry point.  Processes fiber + copper sheets and cross-references
    with the "כל השאר" sheet to produce the final report.

    Returns
    -------
    result_df : pd.DataFrame
        The cleaned, classified orders table.
    exceptions_df : pd.DataFrame
        Orders whose service status was unrecognised.
    """
    # 1. Combine fiber and copper – they share the same structure
    combined_df = pd.concat([fiber_df, copper_df], ignore_index=True)

    # 2. Keep only the last work-order row per order number
    last_rows = _get_last_workorder_per_order(combined_df)

    # 3. Build Mesh lookup from "כל השאר"
    mesh_order_nums = _build_mesh_lookup(rest_df)

    # 4. Classify each order
    output_rows     = []
    exception_rows  = []

    for _, row in last_rows.iterrows():
        order_num = str(row.get(COL_ORDER_NUM, "")).strip()
        card_num  = str(row.get(COL_CARD_NUM,  "")).strip()

        classification = _classify_order(row)

        flat = {
            OUT_ORDER_NUM:    order_num,
            OUT_CARD_NUM:     card_num,
            OUT_ORDER_STATUS: classification[OUT_ORDER_STATUS],
            OUT_INSTALL_DT:   classification[OUT_INSTALL_DT],
            OUT_COORD_DATE:   classification[OUT_COORD_DATE],
        }

        if classification["_is_exception"]:
            exception_rows.append(flat)
        elif classification[OUT_ORDER_STATUS] is None:
            # No status was assigned (unexpected combination) → treat as exception
            exception_rows.append(flat)
        else:
            output_rows.append(flat)

    _COLS = [OUT_ORDER_NUM, OUT_CARD_NUM, OUT_ORDER_STATUS, OUT_INSTALL_DT, OUT_COORD_DATE]
    result_df     = pd.DataFrame(output_rows,    columns=list(flat.keys()) if output_rows    else _COLS)
    exceptions_df = pd.DataFrame(exception_rows, columns=list(flat.keys()) if exception_rows else _COLS)

    # 5. Append BIZNET rows from "כל השאר" to the bottom of the result
    biznet_rows = _build_biznet_rows(rest_df)
    if biznet_rows:
        biznet_df = pd.DataFrame(biznet_rows, columns=_COLS)
        result_df = pd.concat([result_df, biznet_df], ignore_index=True)

    # 6. Build the separate phone-line report
    _PHONE_COLS = [OUT_ORDER_NUM, OUT_CARD_NUM, OUT_ORDER_STATUS, OUT_COORD_DATE, OUT_MINUTES]
    phone_rows  = _build_phone_rows(rest_df)
    phone_df    = pd.DataFrame(phone_rows, columns=_PHONE_COLS) if phone_rows else pd.DataFrame(columns=_PHONE_COLS)

    return result_df, exceptions_df, phone_df
