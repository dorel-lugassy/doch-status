"""
main.py
-------
Streamlit entry point for the Excel processing web application.

Structure:
  - Sidebar: action buttons for each available report
  - Main area: dynamic content based on the selected action
"""

import datetime
import streamlit as st

from utils.excel_utils import load_sheets, dfs_to_excel_bytes
from processors import internet_morchav

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

                    result_df, exceptions_df, phone_df = internet_morchav.run(
                        fiber_df  = sheets["סיבים"],
                        copper_df = sheets["נחושת"],
                        rest_df   = sheets["כל השאר"],
                    )

                    st.session_state["analysis_result"] = {
                        "result":     result_df,
                        "exceptions": exceptions_df,
                        "phone":      phone_df,
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

        # ── Download 3: Phone lines ────────────────────────────────────────
        if not phone_df.empty:
            st.download_button(
                label="⬇️ הורד קובץ קו טלפון",
                data=dfs_to_excel_bytes({"הזמנות קו טלפון": phone_df}),
                file_name=f"סטטוס קו טלפון - {today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_phone",
            )
