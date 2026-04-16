"""
ui_pmax.py — PMAX Conversion Channel Distribution UI
Upload Google PMX Channel distribution report → output summary table
"""

import io
import streamlit as st
import pandas as pd

from secom_processor import (
    load_pmax_channel,
    build_pmax_channel_report,
    PMAX_CHANNEL_OUTPUT_COLUMNS,
)


def _step(n, title):
    st.markdown(
        f'<div class="step-row"><div class="step-num">{n}</div>'
        f'<div class="step-title">{title}</div></div>',
        unsafe_allow_html=True,
    )


def render_pmax():

    # ── STEP 1: Upload ────────────────────────────────────────────────────────
    _step("01", "Upload PMX Channel Distribution File")
    st.caption("Upload the Google PMX Channel distribution report (.csv)")

    uploaded = st.file_uploader(
        "Drag & drop your PMX Channel distribution CSV",
        type=["csv"],
        accept_multiple_files=False,
        key="pmax_uploader",
    )

    if not uploaded:
        st.info("👆 Upload your `Google_PMX_Channel_distribution_report.csv` to begin.")
        with st.expander("ℹ️ How to export this file", expanded=False):
            st.markdown("""
1. Go to **Google Ads** → **Campaigns** → select your PMAX campaigns
2. Click **Reports** → **Channel distribution**
3. Export as **CSV** (UTF-16 format)
4. Include "pmx" and "channel" in the filename for easy identification
            """)
        st.stop()

    # ── Load ──────────────────────────────────────────────────────────────────
    try:
        df_raw = load_pmax_channel(uploaded)
    except Exception as e:
        st.error(f"❌ Failed to load file: {e}")
        st.stop()

    # ── STEP 2: Preview ───────────────────────────────────────────────────────
    _step("02", "Preview Raw Data")
    c1, c2 = st.columns(2)
    c1.metric("Rows", f"{len(df_raw):,}")
    c2.metric("Columns", len(df_raw.columns))
    with st.expander("👁 Raw data preview", expanded=False):
        st.dataframe(df_raw.head(10), use_container_width=True)

    # ── STEP 3: Settings ──────────────────────────────────────────────────────
    _step("03", "Settings")
    c1, c2 = st.columns([2, 1])
    with c1:
        strip_suffix = st.text_input(
            "Remove suffix from Campaign Name",
            value="_WE/SEC26018",
            help="Stripped before detecting CONSIDERATION / CONVERSION type.",
        )

    # ── RUN ───────────────────────────────────────────────────────────────────
    run_col, _ = st.columns([1, 3])
    with run_col:
        run_btn = st.button("▶ Build Report", type="primary", use_container_width=True)

    if not run_btn and "pmax_result" not in st.session_state:
        st.info("Click **▶ Build Report** to process.")
        st.stop()

    if run_btn:
        with st.spinner("Processing…"):
            try:
                df_out = build_pmax_channel_report(df_raw, strip_suffix)
                st.session_state["pmax_result"] = df_out
            except Exception as e:
                st.error(f"❌ Processing error: {e}")
                import traceback; st.code(traceback.format_exc())
                st.stop()

    df_out = st.session_state["pmax_result"]

    # ── STEP 4: Summary ───────────────────────────────────────────────────────
    _step("04", "Summary")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Rows", f"{len(df_out):,}")

    consideration = df_out[df_out['Original Name'].str.contains('CONSIDERATION', na=False)]
    conversion    = df_out[df_out['Original Name'].str.contains('CONVERSION', na=False)]
    c2.metric("CONSIDERATION rows", f"{len(consideration):,}")
    c3.metric("CONVERSION rows", f"{len(conversion):,}")

    # ── STEP 5: Preview Output ────────────────────────────────────────────────
    _step("05", "Output Table")

    # Style the table — highlight row by type
    def highlight_row(row):
        if 'CONSIDERATION' in str(row['Original Name']):
            return ['background-color: rgba(29,158,117,0.06)'] * len(row)
        else:
            return ['background-color: rgba(83,74,183,0.06)'] * len(row)

    st.dataframe(
        df_out.style.apply(highlight_row, axis=1),
        use_container_width=True,
        height=400,
    )

    # ── STEP 6: Download ──────────────────────────────────────────────────────
    _step("06", "Download")
    dl1, dl2 = st.columns(2)

    with dl1:
        csv = df_out.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            "⬇️ Download CSV",
            csv,
            "pmax_channel_report.csv",
            "text/csv",
            use_container_width=True,
        )
        st.caption("UTF-8 BOM — Excel-safe")

    with dl2:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df_out.to_excel(w, index=False, sheet_name="PMAX Channel Report")
        st.download_button(
            "⬇️ Download Excel",
            buf.getvalue(),
            "pmax_channel_report.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption("All data in one sheet")
