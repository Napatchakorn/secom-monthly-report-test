"""
ui_secom.py — SECOM Multi-Source Monthly Report UI
Sources: Meta Ads (.xlsx) + Google Ads YT/DMG/GDN (.csv) + GA4 Session (.csv)
"""

import io
import streamlit as st
import pandas as pd

from secom_processor import (
    load_meta_ads, load_google_ads, load_google_pmx, load_google_sem, load_ga4_session,
    build_meta_report, build_google_report, build_google_pmx_report, build_google_sem_report,
    build_ga4_report, combine_ad_sources, merge_reports, finalize_master, qa_report,
    detect_campaign_suffix,
    META_ADS_DEFAULT_MAPPING, GOOGLE_ADS_DEFAULT_MAPPING,
    GOOGLE_PMX_DEFAULT_MAPPING, GOOGLE_SEM_DEFAULT_MAPPING,
    GA4_DEFAULT_MAPPING, OUTPUT_COLUMNS,
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


def _file_row(icon, label, f, optional=False):
    if f:
        badge = '<span class="pill-ok">✓ Detected</span>'
        fname = f'<span style="font-family:monospace;font-size:0.78rem;color:#2563eb;margin-left:8px">{f.name[:25]}</span>'
    elif optional:
        badge = '<span class="pill-warn">– Optional</span>'
        fname = ""
    else:
        badge = '<span class="pill-err">✗ Missing</span>'
        fname = ""
    st.markdown(
        f'<div class="file-row"><span>{icon}</span>'
        f'<span class="key" style="min-width:140px">{label}</span>'
        f'{fname}<span class="ml">{badge}</span></div>',
        unsafe_allow_html=True,
    )


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
    r1c1, r1c2, r1c3 = st.columns(3)
    r2c1, r2c2, r2c3 = st.columns(3)
    with r1c1:
        _file_row("📗", "Meta Ads",            meta_file,   optional=True)
    with r1c2:
        _file_row("📙", "Google YT/DMG/GDN",   google_file, optional=True)
    with r1c3:
        _file_row("📘", "GA4 Session",          ga4_file,    optional=False)
    with r2c1:
        _file_row("📒", "Google PMX",           pmx_file,    optional=True)
    with r2c2:
        _file_row("📓", "Google SEM",           sem_file,    optional=True)

    if not meta_file and not google_file:
        st.error("Need at least one ad source file: Meta Ads (.xlsx) or Google Ads (.csv).")
        st.stop()
    if not ga4_file:
        st.error("GA4 Session CSV is required.")
        st.stop()

    # ── Load files ────────────────────────────────────────────────────────────
    df_meta_raw = df_google_raw = df_pmx_raw = df_sem_raw = None

    if meta_file:
        try: df_meta_raw = load_meta_ads(meta_file)
        except Exception as e: st.error(f"❌ Meta Ads: {e}"); st.stop()

    if google_file:
        try: df_google_raw = load_google_ads(google_file)
        except Exception as e: st.error(f"❌ Google Ads: {e}"); st.stop()

    if pmx_file:
        try: df_pmx_raw = load_google_pmx(pmx_file)
        except Exception as e: st.error(f"❌ Google PMX: {e}"); st.stop()

    if sem_file:
        try: df_sem_raw = load_google_sem(sem_file)
        except Exception as e: st.error(f"❌ Google SEM: {e}"); st.stop()

    try:
        result = load_ga4_session(ga4_file)
        df_ga4_raw, detected_skip = result
    except Exception as e: st.error(f"❌ GA4: {e}"); st.stop()

    with st.expander("⚙️ GA4 CSV Settings (auto-detected)", expanded=False):
        st.success(f"✅ Header auto-detected at row {detected_skip}")
        manual_skip = st.number_input("Override skip rows (only if data looks wrong)", 0, 20, detected_skip)
        if manual_skip != detected_skip:
            ga4_file.seek(0)
            df_ga4_raw, _ = load_ga4_session(ga4_file, int(manual_skip))

    # ── STEP 2: Preview ───────────────────────────────────────────────────────
    _step("02", "Preview Raw Data")
    tab_labels, tab_dfs = [], []
    for label, df in [("📗 Meta Ads", df_meta_raw), ("📙 Google YT/DMG/GDN", df_google_raw),
                       ("📒 Google PMX", df_pmx_raw), ("📓 Google SEM", df_sem_raw),
                       ("📘 GA4 Session", df_ga4_raw)]:
        if df is not None:
            tab_labels.append(label); tab_dfs.append(df)

    for tab, df in zip(st.tabs(tab_labels), tab_dfs):
        with tab:
            c1, c2 = st.columns(2)
            c1.metric("Rows", f"{len(df):,}")
            c2.metric("Columns", len(df.columns))
            st.dataframe(df.head(5), use_container_width=True)

    # ── STEP 3: Column Mapping ────────────────────────────────────────────────
    _step("03", "Column Mapping")

    map_labels = []
    if df_meta_raw   is not None: map_labels.append("📗 Meta Ads")
    if df_google_raw is not None: map_labels.append("📙 Google YT/DMG/GDN")
    if df_pmx_raw    is not None: map_labels.append("📒 Google PMX")
    if df_sem_raw    is not None: map_labels.append("📓 Google SEM")
    map_labels.append("📘 GA4 Session")

    meta_mapping = {}; google_mapping = {}; pmx_mapping = {}; ga4_mapping = {}

    map_tabs = st.tabs(map_labels)
    idx = 0

    if df_meta_raw is not None:
        with map_tabs[idx]:
            meta_cols = list(df_meta_raw.columns)
            ui = st.columns(3)
            for i, (raw_def, std_name) in enumerate(META_ADS_DEFAULT_MAPPING.items()):
                with ui[i % 3]:
                    def_idx = meta_cols.index(raw_def) + 1 if raw_def in meta_cols else 0
                    sel = st.selectbox(f"→ **{std_name}**", ["(skip)"] + meta_cols,
                                       index=def_idx, key=f"meta_{std_name}_{i}")
                    if sel != "(skip)": meta_mapping[sel] = std_name
        idx += 1

    if df_google_raw is not None:
        with map_tabs[idx]:
            google_cols = list(df_google_raw.columns)
            ui = st.columns(3)
            for i, (raw_def, std_name) in enumerate(GOOGLE_ADS_DEFAULT_MAPPING.items()):
                with ui[i % 3]:
                    def_idx = google_cols.index(raw_def) + 1 if raw_def in google_cols else 0
                    sel = st.selectbox(f"→ **{std_name}**", ["(skip)"] + google_cols,
                                       index=def_idx, key=f"google_{std_name}_{i}")
                    if sel != "(skip)": google_mapping[sel] = std_name
            st.caption("ℹ️ GDN campaigns auto-extract Ad from utm_term= in Final URL.")
        idx += 1

    if df_pmx_raw is not None:
        with map_tabs[idx]:
            st.info("ℹ️ PMX is an asset-group level report. Campaign Name auto-extracted.")
            pmx_cols = list(df_pmx_raw.columns)
            ui = st.columns(3)
            for i, (raw_def, std_name) in enumerate(GOOGLE_PMX_DEFAULT_MAPPING.items()):
                with ui[i % 3]:
                    def_idx = pmx_cols.index(raw_def) + 1 if raw_def in pmx_cols else 0
                    sel = st.selectbox(f"→ **{std_name}**", ["(skip)"] + pmx_cols,
                                       index=def_idx, key=f"pmx_{std_name}_{i}")
                    if sel != "(skip)": pmx_mapping[sel] = std_name
        idx += 1

    if df_sem_raw is not None:
        with map_tabs[idx]:
            st.info("ℹ️ SEM is an ad-group level report. Campaign Name auto-extracted from combined column.")
        idx += 1

    with map_tabs[idx]:
        ga4_cols = list(df_ga4_raw.columns)
        ui = st.columns(3)
        for i, (raw_def, std_name) in enumerate(GA4_DEFAULT_MAPPING.items()):
            with ui[i % 3]:
                def_idx = ga4_cols.index(raw_def) + 1 if raw_def in ga4_cols else 0
                sel = st.selectbox(f"→ **{std_name}**", ["(skip)"] + ga4_cols,
                                   index=def_idx, key=f"ga4_{std_name}_{i}")
                if sel != "(skip)": ga4_mapping[sel] = std_name

    # ── STEP 4: Campaign Cleaner ───────────────────────────────────────────────
    _step("04", "Campaign Name Cleaner")

    # Auto-detect suffixes from all loaded raw files
    detected_suffixes = detect_campaign_suffix([df_meta_raw, df_google_raw, df_pmx_raw, df_sem_raw])

    st.caption("Suffixes detected from your files — each will be stripped from Campaign Names.")

    # One input row per detected suffix (plus one blank for manual add)
    suffix_list = list(detected_suffixes.keys()) if detected_suffixes else ["_WE/SEC26018"]

    strip_suffixes = []
    for i, default_val in enumerate(suffix_list):
        count = detected_suffixes.get(default_val, 0)
        c1, c2 = st.columns([3, 1])
        with c1:
            val = st.text_input(
                f"Suffix {i+1}",
                value=default_val,
                key=f"suffix_{i}",
                label_visibility="collapsed",
                placeholder="e.g. _WE/SEC26018",
            )
        with c2:
            st.markdown(
                f'<div style="padding:8px 0;font-size:0.8rem;color:#64748b">found {count}×</div>',
                unsafe_allow_html=True,
            )
        if val.strip():
            strip_suffixes.append(val.strip())

    # Extra row to add a custom suffix manually
    extra = st.text_input("+ Add another suffix (optional)", value="",
                          key="suffix_extra", placeholder="e.g. _WE/SEC26022")
    if extra.strip():
        strip_suffixes.append(extra.strip())

    if not strip_suffixes:
        st.warning("⚠️ No suffixes set — Campaign Names may not match GA4.")
    else:
        st.caption(f"Will strip: {' | '.join(strip_suffixes)}")

    # Use list for strip_suffix throughout pipeline
    strip_suffix = strip_suffixes

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
                df_meta_rep   = build_meta_report(df_meta_raw, meta_mapping, strip_suffix) if df_meta_raw is not None else None
                df_google_rep = build_google_report(df_google_raw, google_mapping, strip_suffix) if df_google_raw is not None else None
                df_pmx_rep    = build_google_pmx_report(df_pmx_raw, pmx_mapping, strip_suffix) if df_pmx_raw is not None else None
                df_sem_rep    = build_google_sem_report(df_sem_raw, {}, strip_suffix) if df_sem_raw is not None else None
                df_ga4_rep    = build_ga4_report(df_ga4_raw, ga4_mapping)

                df_combined = combine_ad_sources(df_meta_rep, df_google_rep, df_pmx_rep, df_sem_rep)
                df_merged   = merge_reports(df_combined, df_ga4_rep)
                qa          = qa_report(df_merged, len(df_combined))
                df_final    = finalize_master(df_merged)

                st.session_state["secom_result"] = {"final": df_final, "qa": qa}
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
    cols = st.columns(len(src_counts) + 1)
    for i, row in src_counts.iterrows():
        cols[i].markdown(
            f'<div class="metric-card"><div class="metric-val">{row["Rows"]}</div>'
            f'<div class="metric-lbl">{row["Report Source"]}</div></div>',
            unsafe_allow_html=True,
        )
    cols[-1].markdown(
        f'<div class="metric-card"><div class="metric-val">{len(df_final):,}</div>'
        f'<div class="metric-lbl">Total Rows</div></div>',
        unsafe_allow_html=True,
    )

    # ── STEP 6: QA Check ─────────────────────────────────────────────────────
    _step("06", "Merge Quality Check")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Rows",       f"{qa['total']:,}")
    c2.metric("Matched with GA4", f"{qa['matched']:,}")
    c3.metric("Unmatched",        f"{qa['unmatched']:,}")
    c4.metric("Match Rate",       f"{qa['match_pct']}%")

    bar_color = "#16a34a" if qa["match_pct"] >= 70 else "#f59e0b" if qa["match_pct"] >= 40 else "#ef4444"
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:10px;margin:8px 0 12px">' 
        f'<div class="qa-bar-outer" style="flex:1"><div class="qa-bar-inner" '
        f'style="width:{int(qa["match_pct"])}%;background:{bar_color}"></div></div>'
        f'<span class="qa-pct" style="color:{bar_color}">{qa["match_pct"]}%</span></div>',
        unsafe_allow_html=True,
    )
    if qa["match_pct"] < 70:
        st.warning("⚠️ Low match rate. Check Step 03 column mapping and Step 04 suffix.")

    if qa["unmatched"] > 0 and not qa["unmatched_rows"].empty:
        with st.expander(f"🔍 Unmatched rows ({qa['unmatched']})", expanded=False):
            st.dataframe(qa["unmatched_rows"], use_container_width=True)

    # ── STEP 7: Preview ───────────────────────────────────────────────────────
    _step("07", "Preview Master Report")
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Rows",        f"{len(df_final):,}")
    p2.metric("Columns",     len(df_final.columns))
    p3.metric("Total Cost",  f"{pd.to_numeric(df_final.get('Cost',  pd.Series(dtype=float)), errors='coerce').sum():,.0f}")
    p4.metric("Total Leads", f"{pd.to_numeric(df_final.get('Leads', pd.Series(dtype=float)), errors='coerce').sum():,.0f}")
    st.dataframe(df_final, use_container_width=True, height=320)

    # ── STEP 8: Download ──────────────────────────────────────────────────────
    _step("08", "Download Results")
    dl1, dl2 = st.columns(2)

    with dl1:
        csv = df_final.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button("⬇️ Download CSV", csv, "master_monthly_report.csv", "text/csv", use_container_width=True)
        st.caption("UTF-8 BOM — Excel-safe")

    with dl2:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df_final.to_excel(w, index=False, sheet_name="Master Report")
            pd.DataFrame({
                "Metric": ["Total", "Matched", "Unmatched", "Match %"],
                "Value":  [qa["total"], qa["matched"], qa["unmatched"], qa["match_pct"]],
            }).to_excel(w, index=False, sheet_name="QA Summary")
        st.download_button("⬇️ Download Excel", buf.getvalue(), "master_monthly_report.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
        st.caption("Includes QA Summary sheet")


def _render_help():
    with st.expander("ℹ️ File naming tips", expanded=False):
        st.markdown("""
| File | Include in filename |
|------|-------------------|
| Meta Ads `.xlsx` | `meta` or `facebook` |
| Google YT/DMG/GDN `.csv` | `youtube`, `dmg`, `gdn` |
| Google PMX `.csv` | `pmx` or `pmax` |
| Google SEM `.csv` | `sem` |
| GA4 Session `.csv` | `ga4` or `session` |
        """)
