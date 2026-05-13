"""
ui_secom.py — SECOM Multi-Source Monthly Report UI
Sources: Meta Ads (.xlsx) + Google Ads YT/DMG/GDN (.csv) + GA4 Session (.csv)
"""

import io
import streamlit as st
import pandas as pd

from secom_processor import (
    load_meta_ads, load_google_ads, load_ga4_session,
    build_meta_report, build_google_report, build_ga4_report,
    combine_ad_sources, merge_reports, finalize_master, qa_report,
    META_ADS_DEFAULT_MAPPING, GOOGLE_ADS_DEFAULT_MAPPING, GA4_DEFAULT_MAPPING,
    OUTPUT_COLUMNS,
)


def _step(n, title):
    st.markdown(
        f'<div class="step-row">'
        f'<div class="step-num">{n}</div>'
        f'<div class="step-title">{title}</div>'
        f'</div>', unsafe_allow_html=True,
    )


def _pill(label, ok=True):
    cls = "pill-ok" if ok else "pill-err"
    return f'<span class="{cls}">{"✓" if ok else "✗"} {label}</span>'


def render_secom():

    # ── STEP 1: Upload ────────────────────────────────────────────────────────
    _step("01", "Upload Raw Data Files")

    uploaded = st.file_uploader(
        "Upload Meta Ads (.xlsx), Google Ads (.csv) and GA4 Session (.csv)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
    )

    if not uploaded:
        st.info("👆 Upload your files to begin. Meta Ads and GA4 are required. Google Ads is optional.")
        _render_help()
        st.stop()

    # ── Auto-detect files ─────────────────────────────────────────────────────
    meta_file   = None
    google_file = None
    pmx_file    = None
    sem_file    = None
    ga4_file    = None

    for f in uploaded:
        name = f.name.lower()
        if any(k in name for k in ["meta", "facebook"]) and name.endswith((".xlsx", ".xls")):
            meta_file = meta_file or f
        elif any(k in name for k in ["pmx", "pmax"]) and name.endswith(".csv"):
            pmx_file = pmx_file or f
        elif any(k in name for k in ["sem"]) and name.endswith(".csv"):
            sem_file = sem_file or f
        elif any(k in name for k in ["youtube", "_yt_", "dmg", "gdn", "google_ads"]) and name.endswith(".csv"):
            google_file = google_file or f
        elif any(k in name for k in ["ga4", "session"]) and name.endswith(".csv"):
            ga4_file = ga4_file or f
        elif name.endswith((".xlsx", ".xls")) and meta_file is None:
            meta_file = f
        elif name.endswith(".csv") and ga4_file is None:
            ga4_file = f

    # Show detection
    st.markdown("**File Detection:**")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.markdown(
            f'<div class="file-row"><span>📗</span>'
            f'<span class="key">Meta Ads</span>'
            f'<span class="ml">{_pill(meta_file.name[:20] if meta_file else "Missing", bool(meta_file))}</span></div>',
            unsafe_allow_html=True,
        )
    with col_b:
        st.markdown(
            f'<div class="file-row"><span>📙</span>'
            f'<span class="key">Google Ads</span>'
            f'<span class="ml">{_pill(google_file.name[:20] if google_file else "Optional", True)}</span></div>',
            unsafe_allow_html=True,
        )
    with col_c:
        st.markdown(
            f'<div class="file-row"><span>📘</span>'
            f'<span class="key">GA4 Session</span>'
            f'<span class="ml">{_pill(ga4_file.name[:20] if ga4_file else "Missing", bool(ga4_file))}</span></div>',
            unsafe_allow_html=True,
        )

    if not meta_file and not google_file:
        st.error("Need at least one ad source file: Meta Ads (.xlsx) or Google Ads (.csv).")
        st.stop()
    if not ga4_file:
        st.error("GA4 Session CSV is required.")
        st.stop()

    # ── Load files ────────────────────────────────────────────────────────────
    df_meta_raw = df_google_raw = None

    if meta_file:
        try:
            df_meta_raw = load_meta_ads(meta_file)
        except Exception as e:
            st.error(f"❌ Meta Ads load failed: {e}"); st.stop()

    if google_file:
        try:
            df_google_raw = load_google_ads(google_file)
        except Exception as e:
            st.error(f"❌ Google Ads load failed: {e}"); st.stop()

    try:
        result = load_ga4_session(ga4_file)
        df_ga4_raw, detected_skip = result
    except Exception as e:
        st.error(f"❌ GA4 load failed: {e}"); st.stop()

    with st.expander("⚙️ GA4 CSV Settings (auto-detected)", expanded=False):
        st.success(f"✅ Header auto-detected at row {detected_skip}")
        manual_skip = st.number_input("Override skip rows (only if data looks wrong)", 0, 20, detected_skip)
        if manual_skip != detected_skip:
            ga4_file.seek(0)
            df_ga4_raw, _ = load_ga4_session(ga4_file, int(manual_skip))


    # ── STEP 2: Preview ───────────────────────────────────────────────────────
    _step("02", "Preview Raw Data")
    preview_tabs = []
    preview_dfs  = []
    if df_meta_raw is not None:
        preview_tabs.append("📗 Meta Ads"); preview_dfs.append(df_meta_raw)
    if df_google_raw is not None:
        preview_tabs.append("📙 Google Ads"); preview_dfs.append(df_google_raw)
    preview_tabs.append("📘 GA4 Session"); preview_dfs.append(df_ga4_raw)

    tabs = st.tabs(preview_tabs)
    for tab, df in zip(tabs, preview_dfs):
        with tab:
            c1, c2 = st.columns(2)
            c1.metric("Rows", f"{len(df):,}")
            c2.metric("Columns", len(df.columns))
            st.dataframe(df.head(6), use_container_width=True)

    # ── STEP 3: Column Mapping ────────────────────────────────────────────────
    _step("03", "Column Mapping")

    mapping_tabs = []
    if df_meta_raw is not None:   mapping_tabs.append("📗 Meta Ads")
    if df_google_raw is not None: mapping_tabs.append("📙 Google Ads")
    mapping_tabs.append("📘 GA4 Session")

    map_tab_objs = st.tabs(mapping_tabs)
    tab_idx = 0

    meta_mapping   = {}
    google_mapping = {}
    ga4_mapping    = {}

    # Meta Ads mapping
    if df_meta_raw is not None:
        with map_tab_objs[tab_idx]:
            tab_idx += 1
            meta_cols = list(df_meta_raw.columns)
            std_names = list(META_ADS_DEFAULT_MAPPING.values())
            ui_cols = st.columns(3)
            for i, (raw_def, std_name) in enumerate(META_ADS_DEFAULT_MAPPING.items()):
                with ui_cols[i % 3]:
                    default_idx = meta_cols.index(raw_def) + 1 if raw_def in meta_cols else 0
                    sel = st.selectbox(f"→ **{std_name}**", ["(skip)"] + meta_cols,
                                       index=default_idx, key=f"meta_{std_name}")
                    if sel != "(skip)":
                        meta_mapping[sel] = std_name

    # Google Ads mapping
    if df_google_raw is not None:
        with map_tab_objs[tab_idx]:
            tab_idx += 1
            google_cols = list(df_google_raw.columns)
            ui_cols = st.columns(3)
            for i, (raw_def, std_name) in enumerate(GOOGLE_ADS_DEFAULT_MAPPING.items()):
                with ui_cols[i % 3]:
                    default_idx = google_cols.index(raw_def) + 1 if raw_def in google_cols else 0
                    sel = st.selectbox(f"→ **{std_name}**", ["(skip)"] + google_cols,
                                       index=default_idx, key=f"google_{std_name}")
                    if sel != "(skip)":
                        google_mapping[sel] = std_name
            st.caption("ℹ️ For GDN campaigns, **Ad** will auto-extract from `utm_term=` in Final URL when Ad name is empty.")

    # GA4 mapping
    with map_tab_objs[tab_idx]:
        ga4_cols = list(df_ga4_raw.columns)
        ui_cols  = st.columns(3)
        for i, (raw_def, std_name) in enumerate(GA4_DEFAULT_MAPPING.items()):
            with ui_cols[i % 3]:
                default_idx = ga4_cols.index(raw_def) + 1 if raw_def in ga4_cols else 0
                sel = st.selectbox(f"→ **{std_name}**", ["(skip)"] + ga4_cols,
                                   index=default_idx, key=f"ga4_{std_name}")
                if sel != "(skip)":
                    ga4_mapping[sel] = std_name

    # ── STEP 4: Campaign Name Cleaner ─────────────────────────────────────────
    _step("04", "Campaign Name Cleaner")
    c1, c2 = st.columns([2, 1])
    with c1:
        strip_suffix = st.text_input(
            "Remove this suffix from Campaign Name (both Meta Ads & Google Ads)",
            value="_WE/SEC26018",
            help="Must be stripped so campaign names match GA4. Also handles truncated variants like _WE/SEC2601.",
        )
        if not strip_suffix.strip():
            st.warning("⚠️ Suffix is empty — Campaign Names won't match GA4. Recommended: `_WE/SEC26018`")
    with c2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("👁 Preview"):
            samples = []
            if df_meta_raw is not None and "Campaign name" in df_meta_raw.columns:
                from secom_processor import _strip_suffix
                s = df_meta_raw["Campaign name"].astype(str).apply(
                    lambda x: _strip_suffix(x, strip_suffix)).drop_duplicates().head(4).tolist()
                samples += [f"[Meta] {x}" for x in s]
            if df_google_raw is not None and "Campaign" in df_google_raw.columns:
                from secom_processor import _strip_suffix, _clean_str
                s = df_google_raw["Campaign"].astype(str).apply(
                    lambda x: _strip_suffix(_clean_str(x), strip_suffix)).drop_duplicates().head(4).tolist()
                samples += [f"[Google] {x}" for x in s]
            for s in samples:
                st.caption(f"• {s}")

    # ── RUN ───────────────────────────────────────────────────────────────────
    run_col, _ = st.columns([1, 3])
    with run_col:
        run_btn = st.button("▶ Run Pipeline", type="primary", use_container_width=True)

    if not run_btn and "secom_result" not in st.session_state:
        st.info("Configure mappings above then click **▶ Run Pipeline**.")
        st.stop()

    if run_btn:
        with st.spinner("Processing…"):
            try:
                # Build individual source reports
                df_meta_report   = build_meta_report(df_meta_raw, meta_mapping, strip_suffix) if df_meta_raw is not None else None
                df_google_report = build_google_report(df_google_raw, google_mapping, strip_suffix) if df_google_raw is not None else None
                df_ga4_report    = build_ga4_report(df_ga4_raw, ga4_mapping)

                # Combine ad sources
                df_combined = combine_ad_sources(df_meta_report, df_google_report)

                # Merge with GA4
                df_merged = merge_reports(df_combined, df_ga4_report)
                qa        = qa_report(df_merged, len(df_combined))
                df_final  = finalize_master(df_merged)

                st.session_state["secom_result"] = {
                    "final": df_final, "qa": qa,
                    "meta_rows":   len(df_meta_report)   if df_meta_report   is not None else 0,
                    "google_rows": len(df_google_report) if df_google_report is not None else 0,
                }
            except Exception as e:
                st.error(f"❌ Pipeline error: {e}")
                import traceback; st.code(traceback.format_exc())
                st.stop()

    if "secom_result" not in st.session_state:
        st.stop()

    result   = st.session_state["secom_result"]
    df_final = result["final"]
    qa       = result["qa"]

    # ── STEP 5: Source Summary ────────────────────────────────────────────────
    _step("05", "Source Summary")
    src_counts = df_final.groupby("Report Source").size().reset_index(name="Rows")
    s_cols = st.columns(len(src_counts) + 1)
    for i, row in src_counts.iterrows():
        s_cols[i].markdown(
            f'<div class="metric-card"><div class="metric-val">{row["Rows"]}</div>'
            f'<div class="metric-lbl">{row["Report Source"]}</div></div>',
            unsafe_allow_html=True,
        )
    s_cols[-1].markdown(
        f'<div class="metric-card"><div class="metric-val">{len(df_final):,}</div>'
        f'<div class="metric-lbl">Total Rows</div></div>',
        unsafe_allow_html=True,
    )

    # ── STEP 6: Merge QA Check ────────────────────────────────────────────────
    _step("06", "Merge Quality Check")
    qa_cols = st.columns(4)
    qa_cols[0].metric("Total Rows",      f"{qa['total']:,}")
    qa_cols[1].metric("Matched with GA4", f"{qa['matched']:,}")
    qa_cols[2].metric("Unmatched",        f"{qa['unmatched']:,}")
    qa_cols[3].metric("Match Rate",       f"{qa['match_pct']}%")

    bar_color = "#16a34a" if qa['match_pct'] >= 70 else "#f59e0b" if qa['match_pct'] >= 40 else "#ef4444"
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:10px;margin:8px 0 12px">'
        f'<div class="qa-bar-outer" style="flex:1"><div class="qa-bar-inner" '
        f'style="width:{int(qa["match_pct"])}%;background:{bar_color}"></div></div>'
        f'<span class="qa-pct" style="color:{bar_color}">{qa["match_pct"]}%</span></div>',
        unsafe_allow_html=True,
    )
    if qa['match_pct'] < 70:
        st.warning("⚠️ Low match rate — check Step 03 GA4 column mapping or Step 04 suffix cleaner.")

    if qa['unmatched'] > 0 and not qa["unmatched_rows"].empty:
        with st.expander(f"🔍 Unmatched rows ({qa['unmatched']})", expanded=False):
            st.dataframe(qa["unmatched_rows"], use_container_width=True)

    # ── STEP 7: Preview ───────────────────────────────────────────────────────
    _step("07", "Preview Master Report")
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Rows",           f"{len(df_final):,}")
    p2.metric("Columns",        len(df_final.columns))
    p3.metric("Total Cost",     f"{pd.to_numeric(df_final.get('Cost', pd.Series(dtype=float)), errors='coerce').sum():,.0f}")
    p4.metric("Total Leads",    f"{pd.to_numeric(df_final.get('Leads', pd.Series(dtype=float)), errors='coerce').sum():,.0f}")
    st.dataframe(df_final, use_container_width=True, height=320)

    # ── STEP 8: Download ──────────────────────────────────────────────────────
    _step("08", "Download Results")
    dl1, dl2, dl3 = st.columns(3)

    with dl1:
        csv = df_final.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button("⬇️ Download CSV", csv,
                           "master_monthly_report.csv", "text/csv",
                           use_container_width=True)
        st.caption("UTF-8 BOM — Excel-safe")

    with dl2:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df_final.to_excel(w, index=False, sheet_name="Master Report")
            qa_df = pd.DataFrame({
                "Metric": ["Total", "Matched", "Unmatched", "Match %"],
                "Value":  [qa["total"], qa["matched"], qa["unmatched"], qa["match_pct"]],
            })
            qa_df.to_excel(w, index=False, sheet_name="QA Summary")
        st.download_button("⬇️ Download Excel", buf.getvalue(),
                           "master_monthly_report.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
        st.caption("Includes QA Summary sheet")

    with dl3:
        if qa["unmatched"] > 0 and not qa["unmatched_rows"].empty:
            st.download_button("⬇️ Unmatched Rows",
                               qa["unmatched_rows"].to_csv(index=False).encode("utf-8-sig"),
                               "unmatched.csv", "text/csv", use_container_width=True)
            st.caption(f"{qa['unmatched']} rows without GA4 data")
        else:
            st.markdown('<br><span class="pill-ok">✓ All rows matched</span>', unsafe_allow_html=True)


def _render_help():
    with st.expander("ℹ️ File naming tips for auto-detection", expanded=False):
        st.markdown("""
| File | Include in filename |
|------|-------------------|
| Meta Ads `.xlsx` | `meta` or `facebook` |
| Google Ads `.csv` | `youtube`, `yt`, `dmg`, `gdn`, or `google_ads` |
| GA4 Session `.csv` | `ga4` or `session` |
        """)
