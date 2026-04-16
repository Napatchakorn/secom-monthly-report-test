# 📊 Monthly Report Manager v2

Internal team tool — upload raw data, process, download results.

---

## Quick Start

```bash
pip install -r requirements.txt
streamlit run app.py
```
Opens at `http://localhost:8501`

Share with the team on local network:
```bash
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```
Team accesses at `http://YOUR_IP:8501`

---

## Two Modes

### 🔷 Generic Mode
Upload any Excel/CSV files. Clean, merge, filter, calculate KPIs, build summary tables, download.

### 🎯 SECOM Meta Ads Mode
Dedicated pipeline for SECOM monthly reporting:

| Step | What it does |
|------|-------------|
| 1 | Upload Meta Ads `.xlsx` + GA4 Session `.csv` |
| 2 | Auto-detect files by keyword in filename |
| 3 | Visual column mapping (configurable, no code changes) |
| 4 | Campaign name suffix cleaner (e.g. `_WE/SEC26018`) |
| 5 | Left merge on Campaign Name + Ad Group + Ad |
| 6 | Merge QA Check — shows match rate + unmatched rows |
| 7 | Preview 15-column master report |
| 8 | Download CSV / Excel (with QA sheet) |

---

## File Naming (for auto-detection)

| File | Include in filename |
|------|-------------------|
| Meta Ads Excel | `meta` or `ads` |
| GA4 Session CSV | `ga4` or `session` |

---

## File Structure

```
monthly_report_tool/
├── app.py              ← Entry point + CSS + sidebar
├── ui_secom.py         ← SECOM mode UI
├── ui_generic.py       ← Generic mode UI
├── secom_processor.py  ← SECOM data logic (no UI)
├── processor.py        ← Generic data logic (no UI)
├── requirements.txt
└── README.md
```

---

## Customizing the SECOM Pipeline

### Change column defaults (no UI needed)
Edit `secom_processor.py`:

```python
META_ADS_DEFAULT_MAPPING = {
    "Campaign name":      "Campaign Name",   # ← change raw col name here
    "Amount spent (THB)": "Cost",
    ...
}
```

### Add a new output column
1. Add it to `OUTPUT_COLUMNS` list in `secom_processor.py`
2. Map it in the UI (Step 3)

### Change GA4 skiprows
Adjust in the GA4 CSV Settings expander in the UI (default = 6).

---

## Bug Fixes vs. Original Notebook

| # | Original Notebook Bug | Fix |
|---|----------------------|-----|
| 1 | GA4 used `df.index` as Campaign Name (wrong) | Now uses column position 0 correctly |
| 2 | Steps 3–7 ran twice (duplicate cells) | Removed duplication |
| 3 | Google Drive / Colab dependency | Replaced with Streamlit file upload |
| 4 | Hardcoded `_WE/SEC26018` suffix | Configurable text input |
| 5 | No validation of merge quality | QA check with match rate + unmatched rows |
