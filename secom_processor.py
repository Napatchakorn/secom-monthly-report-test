"""
secom_processor.py
==================
SECOM Multi-Source + GA4 → Master Monthly Report pipeline.
Sources: Meta Ads (.xlsx) + Google Ads YT/DMG/GDN (.csv) + GA4 Session (.csv)
"""

import io
import re
import pandas as pd
from typing import Optional


# ── Column mapping defaults ──────────────────────────────────────────────────

META_ADS_DEFAULT_MAPPING = {
    "Campaign name":          "Campaign Name",
    "Ad set name":            "Ad Group",
    "Ad name":                "Ad",
    "Amount spent (THB)":     "Cost",
    "Impressions":            "Imprs.",
    "3-second video plays":   "Views",
    "Link clicks":            "Clicks",
    "Post engagements":       "Engagements",
    "Leads":                  "Leads",
    "Contacts":               "Call Events",
    "Subscriptions":          "LINE Events",
}

GOOGLE_ADS_DEFAULT_MAPPING = {
    "Campaign":        "Campaign Name",
    "Ad group":        "Ad Group",
    "Ad name":         "Ad",           # overridden for GDN via utm_term
    "Cost":            "Cost",
    "Impr.":           "Imprs.",
    "TrueView views":  "Views",
    "Clicks":          "Clicks",
    "Engagements":     "Engagements",
    "Leads":           "Leads",
    "Contact call":    "Call Events",
    "Contact line":    "LINE Events",
}

GA4_DEFAULT_MAPPING = {
    "Session manual campaign name": "Campaign Name",
    "Session manual ad content":    "Ad Group",
    "Session manual term":          "Ad",
    "Sessions":                     "GA4 | Sessions",
    "Average session duration":     "GA4 | Avg. Session Duration",
    "Bounce rate":                  "Bounce rate%",
}

OUTPUT_COLUMNS = [
    "Report Source",
    "Campaign Name",
    "Ad Group",
    "Ad",
    "Cost",
    "Imprs.",
    "Views",
    "Clicks",
    "Engagements",
    "GA4 | Sessions",
    "GA4 | Avg. Session Duration (Minutes)",
    "Bounce rate%",
    "Leads",
    "Call Events",
    "LINE Events",
]


# ── Helpers ──────────────────────────────────────────────────────────────────

def _clean_str(s):
    """Remove zero-width spaces and strip whitespace."""
    return re.sub(r'[\u200b\u200c\u200d\ufeff]', '', str(s)).strip()

def _clean_number(val):
    """Remove commas and dashes from numeric strings → float."""
    s = str(val).replace(',', '').strip()
    if s in ('--', '-', '', 'nan', 'None'):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0

def _extract_utm_term(url: str) -> str:
    """Extract utm_term value from a Final URL string."""
    url = str(url)
    if 'utm_term=' in url:
        term = url.split('utm_term=')[-1].split('&')[0].strip()
        return term if term else ''
    return ''

def _detect_google_source(campaign: str) -> str:
    """Detect Report Source label from Campaign Name."""
    c = campaign.upper()
    if '_YT_' in c or 'YT_' in c:
        return 'Google YT'
    if '_DMG_' in c or 'DMG_' in c:
        return 'Google DMG'
    if '_GDN_' in c or 'GDN_' in c:
        return 'Google GDN'
    return 'Google Ads'

def _strip_suffix(value: str, suffix: str) -> str:
    """Strip suffix pattern; also handles zero-width spaces and regex variant."""
    if not suffix:
        return value
    # Also strip zero-width spaces around the suffix
    cleaned = re.sub(r'[\u200b\u200c\u200d]', '', value)
    # Use regex to handle truncated suffix variants (e.g. _WE/SEC2601 vs _WE/SEC26018)
    escaped = re.escape(suffix)
    cleaned = re.sub(escaped, '', cleaned)
    # Fallback: strip _WE/SEC followed by any digits
    cleaned = re.sub(r'_WE/SEC\d+', '', cleaned)
    return cleaned.strip()


# ── File loaders ─────────────────────────────────────────────────────────────

def load_meta_ads(uploaded_file) -> pd.DataFrame:
    """Load Meta Ads Excel export (header row 0)."""
    raw = uploaded_file.read()
    uploaded_file.seek(0)
    df = pd.read_excel(io.BytesIO(raw), header=0)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_google_ads(uploaded_file) -> pd.DataFrame:
    """
    Load Google Ads YT/DMG/GDN CSV export.
    Format: UTF-16 LE, tab-delimited, 2 metadata rows at top.
    Strips summary/totals rows (NaN or '--' in Campaign column).
    """
    raw = uploaded_file.read()
    uploaded_file.seek(0)

    df = None
    for enc in ['utf-16', 'utf-16-le', 'utf-16-be']:
        try:
            df = pd.read_csv(
                io.BytesIO(raw),
                encoding=enc,
                sep='\t',
                skiprows=2,
                on_bad_lines='skip',
                dtype=str,        # keep everything as string for safe cleaning
            )
            if len(df.columns) > 5:
                break
            df = None
        except Exception:
            continue

    if df is None:
        raise ValueError(
            "Cannot load Google Ads CSV. "
            "Expected UTF-16 tab-delimited format from Google Ads export."
        )

    df.columns = [str(c).strip() for c in df.columns]

    # Drop totals / summary rows — Campaign column is NaN or ' --'
    if 'Campaign' in df.columns:
        df = df[df['Campaign'].notna()]
        df = df[~df['Campaign'].astype(str).str.strip().isin(['--', '-', '', 'nan'])]

    return df.reset_index(drop=True)


def load_ga4_session(uploaded_file, skiprows: int = 6) -> pd.DataFrame:
    """
    Load GA4 Session CSV.
    Pre-strips 'Grand total' row to prevent pandas column-shift bug.
    """
    raw = uploaded_file.read()
    uploaded_file.seek(0)

    text = None
    for enc in ["utf-8-sig", "utf-8", "latin-1", "cp874"]:
        try:
            text = raw.decode(enc)
            break
        except (UnicodeDecodeError, UnicodeError):
            continue

    if text is None:
        raise ValueError(
            "Cannot decode GA4_Session.csv.\n"
            "Fix: open the file in Excel → Save As → CSV UTF-8 (Comma delimited) → re-upload."
        )

    clean_lines = [line for line in text.split("\n") if "Grand total" not in line]
    clean_bytes  = "\n".join(clean_lines).encode("utf-8")

    try:
        df = pd.read_csv(io.BytesIO(clean_bytes), encoding="utf-8",
                         skiprows=skiprows, index_col=None)
    except Exception as e:
        raise ValueError(f"Could not parse GA4 CSV: {e}")

    if df.empty or len(df.columns) < 2:
        raise ValueError(
            f"GA4 CSV loaded with only {len(df.columns)} column(s). "
            f"Try adjusting 'Skip rows' (currently {skiprows})."
        )

    df = df[df.iloc[:, 0].notna()].copy()
    df = df[df.iloc[:, 0].astype(str).str.strip() != ""].reset_index(drop=True)
    return df


# ── Transformers ─────────────────────────────────────────────────────────────

def build_meta_report(df, col_mapping, strip_suffix=""):
    """Map raw Meta Ads columns to standard names."""
    result = {"Report Source": "Meta Ads"}
    for raw_col, std_name in col_mapping.items():
        result[std_name] = df[raw_col] if raw_col in df.columns else None

    out = pd.DataFrame(result)

    if strip_suffix and "Campaign Name" in out.columns:
        out["Campaign Name"] = out["Campaign Name"].astype(str).apply(
            lambda x: _strip_suffix(x, strip_suffix)
        )

    if "Ad" in out.columns:
        out["Ad"] = out["Ad"].astype(str).replace("nan", "")

    return out.reset_index(drop=True)


def build_google_report(df, col_mapping=None, strip_suffix=""):
    """
    Map Google Ads columns to standard names.
    - Detects Report Source (Google YT / DMG / GDN) from Campaign name
    - For GDN campaigns: extracts Ad from utm_term= in Final URL
    - Cleans comma-formatted numbers
    - Strips campaign suffix
    """
    if col_mapping is None:
        col_mapping = GOOGLE_ADS_DEFAULT_MAPPING

    result = {}
    for raw_col, std_name in col_mapping.items():
        if raw_col in df.columns:
            result[std_name] = df[raw_col].values
        else:
            result[std_name] = None

    out = pd.DataFrame(result)

    # Strip suffix from Campaign Name (handles zero-width spaces + truncated variants)
    if "Campaign Name" in out.columns:
        out["Campaign Name"] = out["Campaign Name"].astype(str).apply(
            lambda x: _strip_suffix(_clean_str(x), strip_suffix)
        )

    # Detect Report Source from Campaign Name
    out["Report Source"] = out["Campaign Name"].apply(_detect_google_source)

    # Fix Ad column for GDN: use utm_term from Final URL when Ad name is '--' or empty
    if "Final URL" in df.columns and "Ad" in out.columns:
        final_urls = df["Final URL"].values
        ads = out["Ad"].astype(str).values
        fixed_ads = []
        campaigns = out["Campaign Name"].astype(str).values
        for i, (ad, url, camp) in enumerate(zip(ads, final_urls, campaigns)):
            is_gdn = 'GDN' in camp.upper()
            ad_empty = ad.strip() in ('', '--', '-', 'nan', 'None')
            if is_gdn or ad_empty:
                utm = _extract_utm_term(str(url))
                fixed_ads.append(utm if utm else ad)
            else:
                fixed_ads.append(ad)
        out["Ad"] = fixed_ads

    # Clean numeric columns (remove commas, replace '--' with 0)
    numeric_std = ["Cost", "Imprs.", "Views", "Clicks", "Engagements",
                   "Leads", "Call Events", "LINE Events"]
    for col in numeric_std:
        if col in out.columns:
            out[col] = out[col].apply(_clean_number)

    return out.reset_index(drop=True)


def build_ga4_report(df, col_mapping=None):
    """Map GA4 Session columns to standard names."""
    if col_mapping is None:
        col_mapping = GA4_DEFAULT_MAPPING

    result = {}
    for raw_col, std_name in col_mapping.items():
        if raw_col in df.columns:
            result[std_name] = df[raw_col].values
        else:
            result[std_name] = None

    out = pd.DataFrame(result, index=df.index)

    if "Ad" in out.columns:
        out["Ad"] = out["Ad"].astype(str).replace("nan", "")

    return out.reset_index(drop=True)


# ── Combine & Merge ───────────────────────────────────────────────────────────

def combine_ad_sources(*report_dfs) -> pd.DataFrame:
    """Stack multiple ad source DataFrames (Meta, Google, etc.) vertically."""
    frames = [df for df in report_dfs if df is not None and not df.empty]
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)


def merge_reports(combined_df, ga4_df):
    """Left join combined ad sources with GA4 on [Campaign Name, Ad Group, Ad]."""
    return pd.merge(
        combined_df, ga4_df,
        on=["Campaign Name", "Ad Group", "Ad"],
        how="left",
        suffixes=("", "_ga4"),
    )


# ── Finalize ─────────────────────────────────────────────────────────────────

def finalize_master(df):
    """Select output columns, convert types, fill NaN."""
    if "GA4 | Avg. Session Duration" in df.columns and \
       "GA4 | Avg. Session Duration (Minutes)" not in df.columns:
        df = df.rename(columns={
            "GA4 | Avg. Session Duration": "GA4 | Avg. Session Duration (Minutes)"
        })

    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = 0

    out = df[OUTPUT_COLUMNS].copy()

    if "Bounce rate%" in out.columns:
        br = pd.to_numeric(out["Bounce rate%"], errors="coerce")
        if br.dropna().max() <= 1.0:
            br = br * 100
        out["Bounce rate%"] = br

    if "GA4 | Sessions" in out.columns:
        out["GA4 | Sessions"] = pd.to_numeric(out["GA4 | Sessions"], errors="coerce")

    dur = "GA4 | Avg. Session Duration (Minutes)"
    if dur in out.columns:
        out[dur] = pd.to_numeric(out[dur], errors="coerce")

    return out.fillna(0).reset_index(drop=True)


# ── QA ───────────────────────────────────────────────────────────────────────

def qa_report(merged_df, total_ad_rows):
    """Compute match statistics after merge."""
    total     = len(merged_df)
    ga4_col   = "GA4 | Sessions" if "GA4 | Sessions" in merged_df.columns else None
    matched   = int(merged_df[ga4_col].notna().sum()) if ga4_col else 0
    unmatched = total - matched
    match_pct = round(matched / total * 100, 1) if total > 0 else 0.0

    unmatched_rows = pd.DataFrame()
    if ga4_col:
        unmatched_rows = (
            merged_df[merged_df[ga4_col].isna()][
                ["Report Source", "Campaign Name", "Ad Group", "Ad"]
            ].drop_duplicates().head(20)
        )

    return {
        "total":          total,
        "matched":        matched,
        "unmatched":      unmatched,
        "match_pct":      match_pct,
        "unmatched_rows": unmatched_rows,
    }


# ── Google PMX mapping ────────────────────────────────────────────────────────
# NOTE: PMX Asset group asset details report does NOT contain Campaign or
# Asset group columns. Only asset-level metrics are available.
# Campaign Name and Ad Group will be empty (no GA4 merge expected).
GOOGLE_PMX_DEFAULT_MAPPING = {
    "Campaign":     "Campaign Name",
    "Asset group":  "Ad Group",    # Asset group used for both Ad Group and Ad
    "Cost":         "Cost",
    "Impr.":        "Imprs.",
    "Clicks":       "Clicks",
    "Engagements":  "Engagements",
    # Views / Leads / Call Events / LINE Events not in this report type
}

# ── Google SEM mapping ────────────────────────────────────────────────────────
# Campaign column format: {CampaignName}-{AdGroup}_WE/SEC26018
# → Campaign Name extracted by stripping ad group suffix
# → Ad = Ad group (no ad-level data in this report)
GOOGLE_SEM_DEFAULT_MAPPING = {
    "Ad group":       "Ad Group",
    "Ad group":       "Ad",        # same column used for both
    "Cost":           "Cost",
    "Impr.":          "Imprs.",
    "Leads":          "Leads",
    "Contact call":   "Call Events",
    "Contact line":   "LINE Events",
    # TrueView views and Engagements not in SEM reports (search only)
}


def load_google_pmx(uploaded_file) -> pd.DataFrame:
    """Load Google PMX Asset group asset details CSV. UTF-16 tab, skiprows=2."""
    raw = uploaded_file.read()
    uploaded_file.seek(0)

    for enc in ['utf-16', 'utf-16-le']:
        try:
            df = pd.read_csv(io.BytesIO(raw), encoding=enc, sep='\t',
                             skiprows=2, on_bad_lines='skip', dtype=str)
            if len(df.columns) > 5:
                df.columns = [str(c).strip() for c in df.columns]
                # Drop totals rows: 'Total: Account', 'Total: Filtered...'
                if 'Asset group status' in df.columns:
                    df = df[~df['Asset group status'].astype(str).str.startswith('Total')]
                # Drop rows where Campaign is NaN (new format) or Asset name is NaN (old format)
                if 'Campaign' in df.columns:
                    df = df[df['Campaign'].notna()]
                    df = df[~df['Campaign'].astype(str).str.strip().isin(['--', '-', '', 'nan'])]
                elif 'Asset name' in df.columns:
                    df = df[df['Asset name'].notna()]
                    df = df[~df['Asset name'].astype(str).str.strip().isin(['--', '-', '', 'nan'])]
                return df.reset_index(drop=True)
        except Exception:
            continue
    raise ValueError("Cannot load Google PMX CSV. Expected UTF-16 tab-delimited format.")


def load_google_sem(uploaded_file) -> pd.DataFrame:
    """Load Google SEM Ad group report CSV. UTF-16 tab, skiprows=2."""
    raw = uploaded_file.read()
    uploaded_file.seek(0)

    for enc in ['utf-16', 'utf-16-le']:
        try:
            df = pd.read_csv(io.BytesIO(raw), encoding=enc, sep='\t',
                             skiprows=2, on_bad_lines='skip', dtype=str)
            if len(df.columns) > 5:
                df.columns = [str(c).strip() for c in df.columns]
                # Drop totals rows: "Total: Display", "Total: Search", etc.
                if 'Ad group status' in df.columns:
                    df = df[~df['Ad group status'].astype(str).str.startswith('Total')]
                # Drop rows where Campaign is NaN
                if 'Campaign' in df.columns:
                    df = df[df['Campaign'].notna()]
                    df = df[~df['Campaign'].astype(str).str.strip().isin(['--', '-', '', 'nan'])]
                return df.reset_index(drop=True)
        except Exception:
            continue
    raise ValueError("Cannot load Google SEM CSV. Expected UTF-16 tab-delimited format.")


def _extract_sem_campaign_name(campaign_raw: str, ad_group: str) -> str:
    """
    Extract clean campaign name from SEM Campaign column.
    Format: {CampaignName}​​-{AdGroup}_WE/SEC26018
    """
    camp = re.sub(r'[\u200b\u200c\u200d]', '', str(campaign_raw)).strip()
    camp = re.sub(r'_WE/SEC\d+', '', camp).strip()
    if ad_group and ad_group not in ('nan', ''):
        camp = re.sub(r'-?' + re.escape(str(ad_group)) + r'$', '', camp).strip('-').strip()
    return camp


def build_google_pmx_report(df, col_mapping=None, strip_suffix=""):
    """
    Build Google PMX report.
    Uses Campaign and Asset group columns (asset-group level report).
    Ad = Asset group (no ad-level data in this report type).
    """
    if col_mapping is None:
        col_mapping = GOOGLE_PMX_DEFAULT_MAPPING

    result = {"Report Source": "Google PMX"}

    for raw_col, std_name in col_mapping.items():
        if raw_col in df.columns:
            result[std_name] = df[raw_col].values
        else:
            result[std_name] = 0

    out = pd.DataFrame(result)

    # Strip suffix from Campaign Name
    if "Campaign Name" in out.columns:
        out["Campaign Name"] = out["Campaign Name"].astype(str).apply(
            lambda x: _strip_suffix(_clean_str(x), strip_suffix)
        )

    # Ad = Asset group (same as Ad Group — no ad-level data)
    if "Ad Group" in out.columns:
        out["Ad Group"] = out["Ad Group"].apply(lambda x: re.sub(r'[​‌‍]', '', str(x) if x is not None else '').strip())
        out["Ad"] = out["Ad Group"]

    # Views / Leads / Call Events / LINE Events not in PMX asset group report
    for col in ["Views", "Leads", "Call Events", "LINE Events"]:
        if col not in out.columns:
            out[col] = 0

    # Clean numeric columns
    for col in ["Cost", "Imprs.", "Views", "Clicks", "Engagements", "Leads", "Call Events", "LINE Events"]:
        if col in out.columns:
            out[col] = out[col].apply(_clean_number)

    return out.reset_index(drop=True)


def build_google_sem_report(df, col_mapping=None, strip_suffix=""):
    """
    Build Google SEM report.
    Extracts clean Campaign Name from the combined Campaign column.
    Ad = Ad Group (no ad-level data available in this report type).
    """
    if col_mapping is None:
        col_mapping = GOOGLE_SEM_DEFAULT_MAPPING

    out = pd.DataFrame()

    # Extract Campaign Name from combined Campaign column
    if 'Campaign' in df.columns and 'Ad group' in df.columns:
        out["Campaign Name"] = [
            _extract_sem_campaign_name(camp, adg)
            for camp, adg in zip(df['Campaign'], df['Ad group'])
        ]
    else:
        out["Campaign Name"] = ""

    # Ad Group and Ad both come from Ad group column
    if 'Ad group' in df.columns:
        ad_group_vals = df['Ad group'].astype(str).values
        out["Ad Group"] = [re.sub(r'[\u200b\u200c\u200d]', '', v).strip() for v in ad_group_vals]
        out["Ad"]       = out["Ad Group"]   # SEM is ad-group level only
    else:
        out["Ad Group"] = ""
        out["Ad"]       = ""

    # Map remaining metric columns
    metric_map = {
        "Cost":          "Cost",
        "Impr.":         "Imprs.",
        "Clicks":        "Clicks",
        "Leads":         "Leads",
        "Contact call":  "Call Events",
        "Contact line":  "LINE Events",
    }
    for raw_col, std_name in metric_map.items():
        if raw_col in df.columns:
            out[std_name] = df[raw_col].apply(_clean_number).values
        else:
            out[std_name] = 0

    # Views and Engagements not in SEM
    out["Views"]         = 0
    out["Engagements"]   = 0
    out["Report Source"] = "Google SEM"

    return out.reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════════════════
# PMAX Channel Distribution Report
# ══════════════════════════════════════════════════════════════════════════════

PMAX_CHANNEL_OUTPUT_COLUMNS = [
    "Report Source",
    "Original Name",
    "Channel",
    "Impressions",
    "Clicks",
    "Cost",
    "Conversions",
]


def load_pmax_channel(uploaded_file) -> pd.DataFrame:
    """Load Google PMX Channel distribution report. UTF-16 tab, skiprows=2."""
    raw = uploaded_file.read()
    uploaded_file.seek(0)

    for enc in ['utf-16', 'utf-16-le']:
        try:
            df = pd.read_csv(io.BytesIO(raw), encoding=enc, sep='\t',
                             skiprows=2, on_bad_lines='skip', dtype=str)
            if len(df.columns) > 5:
                df.columns = [str(c).strip() for c in df.columns]
                # Drop totals rows
                if 'Status' in df.columns:
                    df = df[df['Status'] != 'Total: Campaigns']
                # Drop rows where Campaigns is NaN
                if 'Campaigns' in df.columns:
                    df = df[df['Campaigns'].notna()]
                    df = df[~df['Campaigns'].astype(str).str.strip().isin(['--','-','','nan'])]
                return df.reset_index(drop=True)
        except Exception:
            continue
    raise ValueError("Cannot load PMX Channel distribution CSV. Expected UTF-16 tab-delimited format.")


def _extract_result_value(results_str: str, key: str) -> str:
    """Extract a specific conversion value from the Results column string."""
    m = re.search(re.escape(key) + r':\s*([\d,\.]+)', str(results_str))
    return m.group(1) if m else '0'


def build_pmax_channel_report(df: pd.DataFrame, strip_suffix: str = "") -> pd.DataFrame:
    """
    Build PMAX Channel Distribution report.
    - Original Name: 'Campaigns (CONSIDERATION)' or 'Campaigns (CONVERSION)'
    - Conversions: Contact value for CONSIDERATION, Submit lead form for CONVERSION
    """
    df = df.copy()

    # Derive Original Name from Campaign
    def _campaign_type(c):
        c = re.sub(r'[\u200b\u200c\u200d]', '', str(c)).strip()
        if strip_suffix:
            c = re.sub(re.escape(strip_suffix), '', c)
            c = re.sub(r'_WE/SEC\d+', '', c).strip()
        if 'CONSIDERATION' in c.upper():
            return 'Campaigns (CONSIDERATION)'
        if 'CONVERSION' in c.upper():
            return 'Campaigns (CONVERSION)'
        return c

    df['Original Name'] = df['Campaigns'].apply(_campaign_type)

    # Extract Conversions based on campaign type
    def _get_conversion(row):
        if 'CONSIDERATION' in str(row['Original Name']):
            val = _extract_result_value(row.get('Results', ''), 'Contact')
        else:
            val = _extract_result_value(row.get('Results', ''), 'Submit lead form')
        try:
            return float(str(val).replace(',', ''))
        except ValueError:
            return 0.0

    df['Conversions']   = df.apply(_get_conversion, axis=1)
    df['Report Source'] = 'Google PMX'

    # Map columns
    out = pd.DataFrame({
        'Report Source': df['Report Source'],
        'Original Name': df['Original Name'],
        'Channel':       df['Channels'] if 'Channels' in df.columns else '',
        'Impressions':   df['Impr.'] if 'Impr.' in df.columns else 0,
        'Clicks':        df['Clicks'] if 'Clicks' in df.columns else 0,
        'Cost':          df['Cost'] if 'Cost' in df.columns else 0,
        'Conversions':   df['Conversions'],
    })

    return out.reset_index(drop=True)
