"""
Monthly Report Manager v2
=========================
Two modes:
  1. Generic  — upload any CSV/Excel, clean, merge, filter, KPI, download
  2. SECOM     — dedicated Meta Ads + GA4 → master_monthly_report pipeline
Run: streamlit run app.py
"""

import streamlit as st

st.set_page_config(
    page_title="Monthly Report Manager",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Global CSS ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; max-width: 1200px; }

.app-header {
    display:flex; align-items:center; gap:14px;
    padding:1.2rem 1.6rem; margin-bottom:1.4rem;
    background:linear-gradient(135deg,#0f1b35 0%,#1a2f5e 100%);
    border-radius:14px; color:white;
}
.app-header .logo  { font-size:2rem; }
.app-header .title { font-size:1.35rem; font-weight:700; letter-spacing:-0.02em; margin:0; }
.app-header .subtitle { font-size:0.8rem; color:#8da8d8; margin:2px 0 0; }
.mode-badge        { margin-left:auto; background:#2563eb; color:white; padding:4px 14px;
                     border-radius:20px; font-size:0.75rem; font-weight:600; letter-spacing:0.04em; }
.mode-badge.secom  { background:#0ea5e9; }

.step-row  { display:flex; align-items:center; gap:10px; margin:1.4rem 0 0.7rem; }
.step-num  { background:#0f1b35; color:#60a5fa; width:28px; height:28px; border-radius:8px;
             display:flex; align-items:center; justify-content:center;
             font-size:0.78rem; font-weight:700; font-family:'DM Mono',monospace; flex-shrink:0; }
.step-title { font-size:0.95rem; font-weight:600; color:#0f1b35; }

.metric-card  { background:white; border:1px solid #e5e9f2; border-radius:10px;
                padding:1rem 1.2rem; box-shadow:0 1px 3px rgba(0,0,0,.06); }
.metric-val   { font-size:1.7rem; font-weight:700; color:#1a2f5e; font-family:'DM Mono'; }
.metric-lbl   { font-size:0.72rem; color:#7b8ab5; text-transform:uppercase;
                letter-spacing:0.07em; margin-top:2px; }

.pill-ok  { display:inline-block; background:#dcfce7; color:#166534;
            padding:2px 10px; border-radius:20px; font-size:0.75rem; font-weight:600; }
.pill-warn{ display:inline-block; background:#fef9c3; color:#854d0e;
            padding:2px 10px; border-radius:20px; font-size:0.75rem; font-weight:600; }
.pill-err { display:inline-block; background:#fee2e2; color:#991b1b;
            padding:2px 10px; border-radius:20px; font-size:0.75rem; font-weight:600; }

.file-row  { display:flex; align-items:center; gap:10px; padding:8px 14px;
             border-radius:8px; margin-bottom:6px; background:white;
             border:1px solid #e5e9f2; font-size:0.85rem; }
.file-row .fname { font-family:'DM Mono'; font-size:0.78rem; color:#2563eb; }
.file-row .key   { font-weight:600; color:#0f1b35; }
.file-row .ml    { margin-left:auto; }

.qa-bar-outer { flex:1; background:#e5e9f2; border-radius:4px; height:8px; }
.qa-bar-inner { height:8px; border-radius:4px; background:#2563eb; }
.qa-pct { font-family:'DM Mono'; font-size:0.78rem; color:#2563eb; font-weight:600; width:46px; text-align:right; }

section[data-testid="stSidebar"] > div { background:#0f1b35; }
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stMarkdown,
section[data-testid="stSidebar"] p { color:#c8d8f0 !important; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color:white !important; }

div[data-testid="stExpander"] { border:1px solid #e5e9f2 !important; border-radius:10px !important; background:white; }

div.stDownloadButton > button {
    background:#0f1b35 !important; color:white !important; border-radius:8px !important;
    font-weight:600 !important; border:none !important; width:100%; padding:0.6rem 1.2rem !important;
}
div.stDownloadButton > button:hover { background:#1a2f5e !important; }
</style>
""", unsafe_allow_html=True)

# ── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Report Manager")
    st.markdown("---")
    mode = st.radio("Select Mode", ["🔷 Generic Report", "🎯 SECOM Monthly Report", "📊 PMAX Conversion"], index=0)
    st.markdown("---")

    if "Generic" in mode:
        st.markdown("### ⚙️ Clean Options")
        drop_dupes = st.checkbox("Remove duplicate rows", True)
        drop_empty = st.checkbox("Drop fully-empty rows", True)
        strip_ws   = st.checkbox("Strip whitespace", True)
        st.markdown("### 📊 KPI Aggregation")
        agg_func   = st.selectbox("Default function", ["sum", "mean", "count", "max", "min"])
    elif "SECOM" in mode:
        st.markdown("### 🎯 SECOM Pipeline")
        st.caption("All sources → Master Monthly Report")
        st.markdown("**Required:**")
        st.markdown("- 📘 GA4 Session `.csv`")
        st.markdown("**Optional (upload any):**")
        st.markdown("- 📗 Meta Ads `.xlsx`")
        st.markdown("- 📙 Google YT/DMG/GDN `.csv`")
        st.markdown("- 📒 Google PMX `.csv`")
        st.markdown("- 📓 Google SEM `.csv`")
        st.markdown("---")
        st.markdown("**Output columns:**")
        for c in ["Report Source","Campaign Name","Ad Group","Ad","Cost",
                  "Imprs.","Views","Clicks","Engagements","GA4 | Sessions",
                  "GA4 | Avg. Session Duration (Minutes)","Bounce rate%",
                  "Leads","Call Events","LINE Events"]:
            st.caption(f"• {c}")

    else:
        st.markdown("### 📊 PMAX Conversion")
        st.caption("Channel distribution report processor")
        st.markdown("**File needed:**")
        st.markdown("- 📒 PMX Channel distribution `.csv`")
        st.markdown("**Output columns:**")
        for c in ["Report Source","Original Name","Channel",
                  "Impressions","Clicks","Cost","Conversions"]:
            st.caption(f"• {c}")

# ── Header ───────────────────────────────────────────────────────────────────
badge_cls  = "mode-badge secom" if "SECOM" in mode else "mode-badge"
badge_text = "SECOM MODE" if "SECOM" in mode else ("PMAX MODE" if "PMAX" in mode else "GENERIC MODE")
st.markdown(f"""
<div class="app-header">
  <div class="logo">📊</div>
  <div>
    <p class="title">Monthly Report Manager</p>
    <p class="subtitle">Upload · Process · Download</p>
  </div>
  <span class="{badge_cls}">{badge_text}</span>
</div>
""", unsafe_allow_html=True)

# ── Route ────────────────────────────────────────────────────────────────────
if "Generic" in mode:
    from ui_generic import render_generic
    render_generic(drop_dupes, drop_empty, strip_ws, agg_func)
elif "PMAX" in mode:
    from ui_pmax import render_pmax
    render_pmax()
else:
    from ui_secom import render_secom
    render_secom()
