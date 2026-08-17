"""
main.py
-------
Streamlit entry point for the Excel processing web application.

Structure:
  - Sidebar: action buttons for each available report
  - Main area: dynamic content based on the selected action
"""

import datetime
import io
from copy import copy

import streamlit as st
from openpyxl import load_workbook

from utils.excel_utils import load_sheets, dfs_to_excel_bytes
from processors import internet_morchav

APP_VERSION_UPDATED_AT = "17.08.2026 11:24"

FILTERED_ORIGINAL_SELLERS = ("אליאור ביטון", "זהבה בלאי", "עומר בר מוחה")
SELLER_FILTER_COLUMNS = {
    "סיבים": 7,      # Column G: שם מוכרן
    "נחושת": 7,      # Column G: שם מוכרן
    "כל השאר": 9,   # Column I
}


def _cell_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _copy_row(worksheet, source_row_idx: int, target_row_idx: int) -> None:
    if source_row_idx == target_row_idx:
        return

    source_dim = worksheet.row_dimensions[source_row_idx]
    target_dim = worksheet.row_dimensions[target_row_idx]
    target_dim.height = source_dim.height
    target_dim.hidden = source_dim.hidden
    target_dim.outlineLevel = source_dim.outlineLevel
    target_dim.collapsed = source_dim.collapsed

    for col_idx in range(1, worksheet.max_column + 1):
        source_cell = worksheet.cell(row=source_row_idx, column=col_idx)
        target_cell = worksheet.cell(row=target_row_idx, column=col_idx)

        target_cell.value = source_cell.value
        if source_cell.has_style:
            target_cell._style = copy(source_cell._style)
        target_cell.number_format = source_cell.number_format
        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.protection = copy(source_cell.protection)
        target_cell.comment = copy(source_cell.comment)
        target_cell.hyperlink = copy(source_cell.hyperlink)


def _filter_worksheet_by_seller(worksheet, seller_col_idx: int, allowed_sellers: set[str]) -> None:
    keep_rows = [1, 2]
    for row_idx in range(3, worksheet.max_row + 1):
        seller_name = _cell_text(worksheet.cell(row=row_idx, column=seller_col_idx).value)
        if seller_name in allowed_sellers:
            keep_rows.append(row_idx)

    original_max_row = worksheet.max_row
    for target_row_idx, source_row_idx in enumerate(keep_rows, start=1):
        _copy_row(worksheet, source_row_idx, target_row_idx)

    first_delete_row = len(keep_rows) + 1
    if first_delete_row <= original_max_row:
        worksheet.delete_rows(first_delete_row, original_max_row - first_delete_row + 1)


def filtered_original_workbook_bytes(uploaded_file) -> bytes:
    """
    Return a copy of the original workbook with only selected seller rows.
    Rows 1-2 are kept; filtering starts from row 3.
    """
    source_bytes = uploaded_file.getvalue()
    workbook = load_workbook(io.BytesIO(source_bytes))
    allowed_sellers = {_cell_text(seller) for seller in FILTERED_ORIGINAL_SELLERS}

    for sheet_name, seller_col_idx in SELLER_FILTER_COLUMNS.items():
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"חסר גיליון נדרש עבור דוח מוכרנים: {sheet_name}")

        worksheet = workbook[sheet_name]
        _filter_worksheet_by_seller(worksheet, seller_col_idx, allowed_sellers)

    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="מערכת דוחות אקסל",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── RTL styling ───────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
        body, .stApp { direction: rtl; text-align: right; }
        .stButton button { width: 100%; }
        .stDownloadButton button { background-color: #0a7c59; color: white; width: 100%; }
        .block-container { padding-top: 2rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Sidebar – action selection ────────────────────────────────────────────────
st.sidebar.title("📂 פעולות")
st.sidebar.markdown("---")

ACTIONS = {
    "internet_morchav": "🌐 ניתוח סטטוס אינטרנט מחודש להרצה",
}

# Keep the selected action in session state so clicking doesn't reset the page
if "selected_action" not in st.session_state:
    st.session_state.selected_action = None

for key, label in ACTIONS.items():
    if st.sidebar.button(label, key=f"btn_{key}"):
        st.session_state.selected_action = key
        # Clear any previous results when switching actions
        st.session_state.pop("analysis_result", None)

# ── Main area ─────────────────────────────────────────────────────────────────
st.title("📊 מערכת דוחות אקסל")
st.caption(f"גרסא: {APP_VERSION_UPDATED_AT}")

if st.session_state.selected_action is None:
    st.info("בחר פעולה מהתפריט משמאל כדי להתחיל.")
    st.stop()

# ── Action: ניתוח סטטוס אינטרנט מורכב להרצה ──────────────────────────────────
if st.session_state.selected_action == "internet_morchav":
    st.header("🌐 ניתוח סטטוס אינטרנט מחודש להרצה")
    st.markdown(
        "העלה את קובץ ה-Excel המכיל את הגיליונות **סיבים**, **נחושת** ו-**כל השאר**."
    )

    uploaded = st.file_uploader(
        "בחר קובץ Excel",
        type=["xlsx"],
        key="upload_internet_morchav",
    )

    if uploaded:
        # Run analysis button
        if st.button("▶️ הפעל ניתוח", key="run_internet_morchav"):
            with st.spinner("מנתח את הקובץ..."):
                try:
                    sheets = load_sheets(
                        uploaded,
                        sheet_names=["סיבים", "נחושת", "כל השאר"],
                    )

                    analysis_output = internet_morchav.run(
                        fiber_df  = sheets["סיבים"],
                        copper_df = sheets["נחושת"],
                        rest_df   = sheets["כל השאר"],
                    )
                    if len(analysis_output) == 4:
                        result_df, exceptions_df, phone_df, biznet_df = analysis_output
                    else:
                        result_df, exceptions_df, phone_df = analysis_output
                        biznet_df = result_df.iloc[0:0].copy()

                    filtered_original_bytes = filtered_original_workbook_bytes(uploaded)

                    st.session_state["analysis_result"] = {
                        "result":     result_df,
                        "exceptions": exceptions_df,
                        "phone":      phone_df,
                        "biznet":     biznet_df,
                        "filtered_original": filtered_original_bytes,
                    }
                except Exception as e:
                    import traceback
                    st.error("❌ שגיאה בעיבוד הקובץ")
                    st.markdown(
                        f"""
**סוג השגיאה:** `{type(e).__name__}`

**פירוט:** `{e}`

**מה לבדוק:**
- האם שמות הגיליונות בקובץ הם בדיוק: `סיבים`, `נחושת`, `כל השאר`?
- האם שורת הכותרות נמצאת בשורה **2** של הגיליון?
- האם כל העמודות הנדרשות קיימות בגיליון?
"""
                    )
                    with st.expander("🔍 פרטי שגיאה מלאים (Traceback)"):
                        st.code(traceback.format_exc(), language="python")
                    st.stop()

    # Display results if available
    if "analysis_result" in st.session_state:
        data = st.session_state["analysis_result"]
        result_df     = data["result"]
        exceptions_df = data["exceptions"]
        phone_df      = data["phone"]
        biznet_df     = data.get("biznet", result_df.iloc[0:0].copy())
        biznet_df     = biznet_df.drop(columns=["תאריך ושעת התקנה מעודכנים"], errors="ignore")
        filtered_original_bytes = data.get("filtered_original")

        # ── Split result by "תאריך מתואם" ─────────────────────────────────
        coord_col = "תאריך מתואם"
        has_date_mask = (
            result_df[coord_col].notna()
            & (result_df[coord_col].astype(str).str.strip() != "")
            & (result_df[coord_col].astype(str).str.strip().str.lower() != "nan")
        )
        result_with_date    = result_df[has_date_mask].reset_index(drop=True)
        result_without_date = result_df[~has_date_mask].drop(columns=[coord_col]).reset_index(drop=True)
        filtered_original_msg = (
            "ונוצר קובץ מקור מסונן לפי מוכרן."
            if filtered_original_bytes
            else "קובץ מקור מסונן לפי מוכרן ייווצר בהרצה מחדש."
        )

        st.success(
            f"✅ הניתוח הושלם! "
            f"נמצאו {len(result_with_date)} הזמנות עם תאריך מתואם, "
            f"{len(result_without_date)} הזמנות ללא תאריך מתואם, "
            f"{len(biznet_df)} הזמנות BIZNET, "
            f"{len(phone_df)} הזמנות קו טלפון, "
            f"{filtered_original_msg}"
        )

        # ── Preview: with date ─────────────────────────────────────────────
        st.subheader(f"📋 סטטוס אינטרנט – עם תאריך מתואם ({len(result_with_date)} שורות)")
        st.dataframe(result_with_date, use_container_width=True)

        # ── Preview: without date ──────────────────────────────────────────
        st.subheader(f"📋 סטטוס אינטרנט – ללא תאריך מתואם ({len(result_without_date)} שורות)")
        st.dataframe(result_without_date, use_container_width=True)

        if not exceptions_df.empty:
            st.warning(f"⚠️ נמצאו {len(exceptions_df)} שורות חריגות (סטטוס שירות לא מוכר).")
            with st.expander("הצג חריגים"):
                st.dataframe(exceptions_df, use_container_width=True)

        # ── Phone result preview ───────────────────────────────────────────
        if not phone_df.empty:
            st.subheader("📞 תצוגה מקדימה – הזמנות קו טלפון")
            st.dataframe(phone_df, use_container_width=True)

        # ── BIZNET result preview ─────────────────────────────────────────
        if not biznet_df.empty:
            st.subheader("🌐 תצוגה מקדימה – הזמנות BIZNET")
            st.dataframe(biznet_df, use_container_width=True)

        st.markdown("---")
        today_str = datetime.date.today().strftime("%d.%m.%Y")

        # ── Download 1: Internet – with coordinated date ───────────────────
        sheets_with_date = {"סטטוס הזמנות": result_with_date}
        if not exceptions_df.empty:
            sheets_with_date["חריגים"] = exceptions_df

        st.download_button(
            label="⬇️ הורד קובץ אינטרנט – עם תאריך מתואם",
            data=dfs_to_excel_bytes(sheets_with_date),
            file_name=f"סטטוס אינטרנט מורכב להרצה - עם תאריך - {today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_internet_with_date",
        )

        # ── Download 2: Internet – without coordinated date ────────────────
        sheets_without_date = {"סטטוס הזמנות": result_without_date}
        if not exceptions_df.empty:
            sheets_without_date["חריגים"] = exceptions_df

        st.download_button(
            label="⬇️ הורד קובץ אינטרנט – ללא תאריך מתואם",
            data=dfs_to_excel_bytes(sheets_without_date),
            file_name=f"סטטוס אינטרנט מורכב להרצה - ללא תאריך - {today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_internet_without_date",
        )

        # ── Download 3: BIZNET ─────────────────────────────────────────────
        if not biznet_df.empty:
            st.download_button(
                label="⬇️ הורד קובץ BIZNET",
                data=dfs_to_excel_bytes({"הזמנות BIZNET": biznet_df}),
                file_name=f"סטטוס BIZNET - {today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_biznet",
            )

        # ── Download 4: Filtered original workbook ────────────────────────
        if filtered_original_bytes:
            st.download_button(
                label="⬇️ הורד קובץ מקור מסונן לפי מוכרן",
                data=filtered_original_bytes,
                file_name=f"דוח מקור מסונן לפי מוכרן - {today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_filtered_original",
            )

        # ── Download 5: Phone lines ────────────────────────────────────────
        if not phone_df.empty:
            st.download_button(
                label="⬇️ הורד קובץ קו טלפון",
                data=dfs_to_excel_bytes({"הזמנות קו טלפון": phone_df}),
                file_name=f"סטטוס קו טלפון - {today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_phone",
            )
