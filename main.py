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
    "internet_morchav": "🌐 ניתוח סטטוס אינטרנט מורכב להרצה",
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
    st.header("🌐 ניתוח סטטוס אינטרנט מורכב להרצה")
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

        st.success(f"✅ הניתוח הושלם! נמצאו {len(result_df)} הזמנות אינטרנט, {len(phone_df)} הזמנות קו טלפון.")

        # ── Internet result preview ────────────────────────────────────────
        st.subheader("📋 תצוגה מקדימה – סטטוס אינטרנט מורכב")
        st.dataframe(result_df, use_container_width=True)

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

        # ── Download 1: Internet Morchav ───────────────────────────────────
        internet_sheets = {"סטטוס הזמנות": result_df}
        if not exceptions_df.empty:
            internet_sheets["חריגים"] = exceptions_df

        st.download_button(
            label="⬇️ הורד קובץ אינטרנט מורכב",
            data=dfs_to_excel_bytes(internet_sheets),
            file_name=f"סטטוס אינטרנט מורכב להרצה - {today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_internet",
        )

        # ── Download 2: Phone lines ────────────────────────────────────────
        if not phone_df.empty:
            st.download_button(
                label="⬇️ הורד קובץ קו טלפון",
                data=dfs_to_excel_bytes({"הזמנות קו טלפון": phone_df}),
                file_name=f"סטטוס קו טלפון - {today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_phone",
            )
