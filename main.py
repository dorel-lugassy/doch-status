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

import streamlit as st

from utils.excel_utils import load_sheets, dfs_to_excel_bytes
from processors import internet_morchav

APP_VERSION_UPDATED_AT = "17.08.2026 11:38"

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


def filtered_original_workbook_bytes(uploaded_file) -> bytes:
    """
    Return a copy of the original workbook with only selected seller rows.
    Rows 1-2 are kept; filtering starts from row 3.
    """
    from openpyxl import load_workbook

    source_bytes = _uploaded_file_bytes(uploaded_file)
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

                    st.session_state["analysis_result"] = {
                        "result":     result_df,
                        "exceptions": exceptions_df,
                        "phone":      phone_df,
                        "biznet":     biznet_df,
                        "source_workbook": uploaded.getvalue(),
                        "filtered_original": None,
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
        source_workbook_bytes = data.get("source_workbook")
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
        st.success(
            f"✅ הניתוח הושלם! "
            f"נמצאו {len(result_with_date)} הזמנות עם תאריך מתואם, "
            f"{len(result_without_date)} הזמנות ללא תאריך מתואם, "
            f"{len(biznet_df)} הזמנות BIZNET, "
            f"{len(phone_df)} הזמנות קו טלפון."
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
        if not filtered_original_bytes and source_workbook_bytes:
            if st.button("⚙️ הכן קובץ מקור מסונן לפי מוכרן", key="prepare_filtered_original"):
                with st.spinner("מכין קובץ מקור מסונן לפי מוכרן..."):
                    filtered_original_bytes = filtered_original_workbook_bytes(source_workbook_bytes)
                    data["filtered_original"] = filtered_original_bytes

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
