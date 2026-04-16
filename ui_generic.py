"""
ui_generic.py
=============
Streamlit UI for Generic Report mode.
"""

import io
import streamlit as st
import pandas as pd
from processor import (
    load_file, standardize_columns, merge_files,
    apply_filters, calculate_kpis, build_summary_table,
)


def _step(n, title):
    st.markdown(
        f'<div class="step-row">'
        f'<div class="step-num">{n}</div>'
        f'<div class="step-title">{title}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def render_generic(drop_dupes, drop_empty, strip_ws, agg_func):

    # ── STEP 1: Upload ────────────────────────────────────────────────────────
    _step("01", "Upload Raw Data Files")

    uploaded_files = st.file_uploader(
        "Drag & drop Excel or CSV files",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("👆 Upload one or more Excel / CSV files to get started.")
        st.stop()

    raw_frames = {}
    for f in uploaded_files:
        try:
            raw_frames[f.name] = load_file(f)
        except Exception as e:
            st.error(f"❌ `{f.name}`: {e}")

    if not raw_frames:
        st.stop()

    with st.expander(f"📁 File Preview ({len(raw_frames)} files)", expanded=True):
        tabs = st.tabs(list(raw_frames.keys()))
        for tab, (name, df) in zip(tabs, raw_frames.items()):
            with tab:
                c1, c2, c3 = st.columns(3)
                c1.metric("Rows", f"{len(df):,}")
                c2.metric("Columns", len(df.columns))
                c3.metric("File", name[:20])
                st.dataframe(df.head(8), use_container_width=True)

    # ── STEP 2: Clean ─────────────────────────────────────────────────────────
    _step("02", "Clean & Standardize")

    all_cols = sorted(set(c for df in raw_frames.values() for c in df.columns))

    with st.expander("🗂 Column Rename (optional)", expanded=False):
        col_map = {}
        ui_cols = st.columns(3)
        for i, col in enumerate(all_cols):
            with ui_cols[i % 3]:
                new = st.text_input(col, value=col, key=f"cmap_{col}")
                if new != col:
                    col_map[col] = new

    cleaned = {
        name: standardize_columns(df, col_map, drop_dupes, drop_empty, strip_ws)
        for name, df in raw_frames.items()
    }
    st.success(f"✅ Cleaned {len(cleaned)} file(s)")

    # ── STEP 3: Merge ─────────────────────────────────────────────────────────
    _step("03", "Merge Files")

    merge_mode = st.radio(
        "Strategy",
        ["Stack rows (vertical append)", "Join on common key (horizontal)"],
        horizontal=True,
    )

    join_key = None
    join_how = "outer"
    if "Join" in merge_mode:
        common = sorted(set.intersection(*[set(df.columns) for df in cleaned.values()]))
        if not common:
            st.warning("No common columns — falling back to vertical stack.")
            merge_mode = "Stack rows (vertical append)"
        else:
            join_key = st.selectbox("Join key column", common)
            join_how = st.selectbox("Join type", ["inner", "left", "outer"])

    merged = merge_files(
        list(cleaned.values()),
        mode="append" if "Stack" in merge_mode else "join",
        key=join_key,
        how=join_how,
    )

    st.metric("Merged rows", f"{len(merged):,}")
    with st.expander("👁 Preview merged"):
        st.dataframe(merged.head(15), use_container_width=True)

    # ── STEP 4: Filter ────────────────────────────────────────────────────────
    _step("04", "Filter & Segment")

    filter_cols = st.multiselect("Filter on columns", merged.columns.tolist())
    active_filters = {}

    if filter_cols:
        fcols = st.columns(min(len(filter_cols), 3))
        for i, col in enumerate(filter_cols):
            with fcols[i % 3]:
                uniq = merged[col].dropna().unique().tolist()
                if len(uniq) <= 50:
                    sel = st.multiselect(f"**{col}**", uniq, default=uniq, key=f"f_{col}")
                    active_filters[col] = sel
                else:
                    txt = st.text_input(f"**{col}** contains", key=f"f_{col}")
                    if txt:
                        active_filters[col] = txt

    filtered = apply_filters(merged, active_filters)
    if filter_cols:
        st.caption(f"→ {len(filtered):,} rows after filtering (from {len(merged):,})")

    # ── STEP 5: KPIs ──────────────────────────────────────────────────────────
    _step("05", "KPI Summary")

    numeric_cols = filtered.select_dtypes(include="number").columns.tolist()
    if numeric_cols:
        kpi_cols = st.multiselect("Select numeric columns", numeric_cols, default=numeric_cols[:5])
        if kpi_cols:
            kpis = calculate_kpis(filtered, kpi_cols)
            cards = st.columns(min(len(kpi_cols), 4))
            for i, col in enumerate(kpi_cols):
                val = kpis.get(col, {}).get(agg_func, 0)
                with cards[i % 4]:
                    st.markdown(
                        f'<div class="metric-card">'
                        f'<div class="metric-val">{val:,.2f}</div>'
                        f'<div class="metric-lbl">{col} ({agg_func})</div>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
            st.markdown("")
            with st.expander("📐 Full KPI breakdown"):
                st.dataframe(pd.DataFrame(kpis).T.style.format("{:.2f}"), use_container_width=True)
    else:
        st.info("No numeric columns detected.")

    # ── STEP 6: Summary Table ─────────────────────────────────────────────────
    _step("06", "Summary Table")

    cat_cols = filtered.select_dtypes(include=["object", "category"]).columns.tolist()
    summary_df = filtered

    if cat_cols and numeric_cols:
        group_col = st.selectbox("Group by", cat_cols)
        agg_c = st.multiselect("Aggregate columns", numeric_cols, default=numeric_cols[:3])
        if group_col and agg_c:
            summary_df = build_summary_table(filtered, group_col, agg_c, agg_func)
            st.dataframe(summary_df, use_container_width=True)

    # ── STEP 7: Download ──────────────────────────────────────────────────────
    _step("07", "Download")

    export = filtered if st.radio("Export", ["Full filtered data", "Summary table"], horizontal=True) == "Full filtered data" else summary_df

    c1, c2 = st.columns(2)
    with c1:
        csv = export.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button("⬇️ Download CSV", csv, "report.csv", "text/csv", use_container_width=True)
    with c2:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            export.to_excel(w, index=False, sheet_name="Report")
        st.download_button(
            "⬇️ Download Excel", buf.getvalue(), "report.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    st.caption(f"Exporting {len(export):,} rows × {len(export.columns)} columns")
