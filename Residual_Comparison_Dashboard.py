
# Residual Comparison Dashboard (Streamlit)
# Upload two Excel files (last month & this month) and get a summary + Excel diff output
# Author: M365 Copilot for Kevin Blamer

import io
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Residual Comparison Dashboard", layout="wide")

st.title("Residual Comparison Dashboard")
st.caption(
    "Upload last month and this month Basic_Information workbooks (US or Canada) to compare by "
    "Residual Master Title: Terms & Optional Miles, plus Normal Miles changes."
)

# --- Helpers ---
REQ_COLUMNS = [
    "Residual Master Title", "Normal Miles", "Optional Miles", "Terms"
]
# Optional (used for filtering, if present in either file)
OPTIONAL_COLUMNS = ["Finance Company"]

@st.cache_data
def parse_list(val):
    """Parse comma-separated list of numbers/strings into a tuple preserving order and uniqueness."""
    if pd.isna(val):
        return tuple()
    if isinstance(val, (int, float)):
        try:
            return (int(val),)
        except Exception:
            return (val,)
    parts = [p.strip() for p in str(val).split(',') if p.strip() != '']
    values = []
    for p in parts:
        try:
            values.append(int(p))
        except ValueError:
            try:
                f = float(p)
                if f.is_integer():
                    values.append(int(f))
                else:
                    values.append(p)
            except Exception:
                values.append(p)
    # unique preserving order
    uniq = []
    for v in values:
        if v not in uniq:
            uniq.append(v)
    return tuple(uniq)

@st.cache_data
def load_sheet(_file_bytes: bytes):
    """
    Load the first sheet from an uploaded Excel file.
    NOTE: Leading underscore in '_file_bytes' tells Streamlit not to hash this arg for caching.
    Pass raw bytes from UploadedFile.read().
    """
    df = pd.read_excel(io.BytesIO(_file_bytes), sheet_name=0, engine='openpyxl')
    df.columns = [str(c).strip() for c in df.columns]
    # Minimal check (required columns)
    missing = [c for c in REQ_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")
    return df

def safe_get(series, key, default=""):
    """Get series[key] if present; otherwise default."""
    try:
        return series.get(key, default)
    except Exception:
        return default

@st.cache_data
def compare_frames(last_df: pd.DataFrame, this_df: pd.DataFrame):
    """Return (summary_df, result_df, changes_df) comparing last vs this month by Residual Master Title."""

    # Build keyed dicts by Residual Master Title
    last_keyed = {}
    for _, row in last_df.iterrows():
        title = str(row['Residual Master Title']).strip()
        last_keyed[title] = {
            'Normal Miles': row['Normal Miles'],
            'Optional Miles': parse_list(row['Optional Miles']),
            'Terms': parse_list(row['Terms']),
            # Include optional fields (e.g., Finance Company)
            'Finance Company': safe_get(row, 'Finance Company')
        }

    this_keyed = {}
    for _, row in this_df.iterrows():
        title = str(row['Residual Master Title']).strip()
        this_keyed[title] = {
            'Normal Miles': row['Normal Miles'],
            'Optional Miles': parse_list(row['Optional Miles']),
            'Terms': parse_list(row['Terms']),
            'Finance Company': safe_get(row, 'Finance Company')
        }

    all_titles = sorted(set(last_keyed.keys()) | set(this_keyed.keys()))

    def build_change_summary(prev, curr, status, normal_prev, normal_curr,
                             terms_prev, terms_curr, opt_prev, opt_curr):
        """Return a concise human-readable description of what changed."""
        notes = []
        if status.startswith("Removed"):
            notes.append("Removed program")
            # Optionally list prior terms/miles removed
            if terms_prev:
                notes.append(f"Terms removed: {', '.join(map(str, terms_prev))}")
            if opt_prev:
                notes.append(f"Optional Miles removed: {', '.join(map(str, opt_prev))}")
            return "; ".join(notes)

        if status.startswith("New"):
            notes.append("New program")
            if terms_curr:
                notes.append(f"Terms: {', '.join(map(str, terms_curr))}")
            if opt_curr:
                notes.append(f"Optional Miles: {', '.join(map(str, opt_curr))}")
            return "; ".join(notes)

        # Changed: enumerate specifics
        if normal_prev != normal_curr:
            notes.append(f"Normal Miles: {normal_prev} → {normal_curr}")

        terms_added = [t for t in terms_curr if t not in terms_prev]
        terms_removed = [t for t in terms_prev if t not in terms_curr]
        if terms_added:
            notes.append(f"Terms added: {', '.join(map(str, terms_added))}")
        if terms_removed:
            notes.append(f"Terms removed: {', '.join(map(str, terms_removed))}")

        opt_added = [m for m in opt_curr if m not in opt_prev]
        opt_removed = [m for m in opt_prev if m not in opt_curr]
        if opt_added:
            notes.append(f"Optional Miles added: {', '.join(map(str, opt_added))}")
        if opt_removed:
            notes.append(f"Optional Miles removed: {', '.join(map(str, opt_removed))}")

        return "; ".join(notes) if notes else "—"

    records = []
    for title in all_titles:
        prev = last_keyed.get(title)
        curr = this_keyed.get(title)

        # Compose comparison values
        normal_prev = prev['Normal Miles'] if prev else np.nan
        normal_curr = curr['Normal Miles'] if curr else np.nan
        opt_prev = prev['Optional Miles'] if prev else tuple()
        opt_curr = curr['Optional Miles'] if curr else tuple()
        terms_prev = prev['Terms'] if prev else tuple()
        terms_curr = curr['Terms'] if curr else tuple()
        finance_company = (curr and curr.get('Finance Company')) or (prev and prev.get('Finance Company')) or ""

        if prev and not curr:
            status = 'Removed (only in last month)'
        elif curr and not prev:
            status = 'New (only in this month)'
        else:
            normal_changed = normal_prev != normal_curr
            terms_added = [t for t in terms_curr if t not in terms_prev]
            terms_removed = [t for t in terms_prev if t not in terms_curr]
            opt_added = [m for m in opt_curr if m not in opt_prev]
            opt_removed = [m for m in opt_prev if m not in opt_curr]
            status = 'Changed' if (normal_changed or terms_added or terms_removed or opt_added or opt_removed) else 'Unchanged'

        what_changed = build_change_summary(prev, curr, status, normal_prev, normal_curr,
                                            terms_prev, terms_curr, opt_prev, opt_curr)

        records.append({
            'Residual Master Title': title,
            'Finance Company': finance_company,
            'Status': status,
            'What Changed': what_changed,  # <-- NEW COLUMN directly after Status
            'Normal Miles (prev)': normal_prev,
            'Normal Miles (curr)': normal_curr,
            'Normal Miles Changed': (normal_prev != normal_curr),
            'Terms (prev)': ', '.join(map(str, terms_prev)) if terms_prev else '',
            'Terms (curr)': ', '.join(map(str, terms_curr)) if terms_curr else '',
            'Terms Added': ', '.join(map(str, [t for t in terms_curr if t not in terms_prev])),
            'Terms Removed': ', '.join(map(str, [t for t in terms_prev if t not in terms_curr])),
            'Optional Miles (prev)': ', '.join(map(str, opt_prev)) if opt_prev else '',
            'Optional Miles (curr)': ', '.join(map(str, opt_curr)) if opt_curr else '',
            'Optional Miles Added': ', '.join(map(str, [m for m in opt_curr if m not in opt_prev])),
            'Optional Miles Removed': ', '.join(map(str, [m for m in opt_prev if m not in opt_curr])),
        })

    result_df = pd.DataFrame(records)
    summary = {
        'Total programs (last month)': len(last_keyed),
        'Total programs (this month)': len(this_keyed),
        'Unchanged': int((result_df['Status'] == 'Unchanged').sum()),
        'Changed': int((result_df['Status'] == 'Changed').sum()),
        'New': int((result_df['Status'].str.startswith('New')).sum()),
        'Removed': int((result_df['Status'].str.startswith('Removed')).sum()),
    }
    summary_df = pd.DataFrame(list(summary.items()), columns=['Metric','Value'])

    # Filtered changes view
    changes_df = result_df[result_df['Status'] == 'Changed'].copy()

    # Order columns (place 'What Changed' right after Status)
    ordered_cols = [
        'Residual Master Title', 'Finance Company', 'Status', 'What Changed',
        'Normal Miles (prev)', 'Normal Miles (curr)', 'Normal Miles Changed',
        'Terms (prev)', 'Terms (curr)', 'Terms Added', 'Terms Removed',
        'Optional Miles (prev)', 'Optional Miles (curr)', 'Optional Miles Added', 'Optional Miles Removed'
    ]
    # Keep columns that exist (Finance Company may be missing if not present in inputs)
    ordered_cols = [c for c in ordered_cols if c in result_df.columns]

    result_df = result_df[ordered_cols]
    changes_df = changes_df[[c for c in ordered_cols if c in changes_df.columns]]

    return summary_df, result_df, changes_df

# --- UI ---
col_l, col_r = st.columns(2)
with col_l:
    last_upl = st.file_uploader("Upload LAST month Basic_Information workbook", type=["xlsx"], key="last")
with col_r:
    this_upl = st.file_uploader("Upload THIS month Basic_Information workbook", type=["xlsx"], key="this")

run_btn = st.button("Run Comparison")

if run_btn:
    if not last_upl or not this_upl:
        st.error("Please upload both files before running the comparison.")
        st.stop()
    try:
        # Use .read() to get bytes; pass to loader with underscore arg (excluded from caching)
        last_df = load_sheet(last_upl.read())
        this_df = load_sheet(this_upl.read())
    except Exception as e:
        st.error(f"Failed to read files: {e}")
        st.stop()

    summary_df, result_df, changes_df = compare_frames(last_df, this_df)

    # --- Summary KPI ---
    st.subheader("Summary")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Last Month", int(summary_df.loc[summary_df['Metric']=='Total programs (last month)','Value']))
    m2.metric("This Month", int(summary_df.loc[summary_df['Metric']=='Total programs (this month)','Value']))
    m3.metric("Unchanged", int(summary_df.loc[summary_df['Metric']=='Unchanged','Value']))
    m4.metric("Changed", int(summary_df.loc[summary_df['Metric']=='Changed','Value']))
    m5.metric("New / Removed",
              f"{int(summary_df.loc[summary_df['Metric']=='New','Value'])} / {int(summary_df.loc[summary_df['Metric']=='Removed','Value'])}")

    st.divider()

    # --- Filters (default Status='Changed') ---
    with st.expander("Filters", expanded=True):
        status_options = sorted(result_df['Status'].unique().tolist())
        default_status = 'Changed' if 'Changed' in status_options else status_options[0]
        status_sel = st.selectbox("Status", options=status_options, index=status_options.index(default_status))

        # Finance Company filter (if present)
        finance_companies = result_df['Finance Company'].dropna().unique().tolist() if 'Finance Company' in result_df.columns else []
        fin_sel = st.multiselect("Finance Company", options=sorted(finance_companies)) if finance_companies else []

        title_query = st.text_input("Title contains", value="")

    # Apply filters
    filtered_df = result_df.copy()
    if status_sel:
        filtered_df = filtered_df[filtered_df['Status'] == status_sel]
    if fin_sel and 'Finance Company' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['Finance Company'].isin(fin_sel)]
    if title_query:
        filtered_df = filtered_df[filtered_df['Residual Master Title'].str.contains(title_query, case=False, na=False)]

    # --- Views ---
    st.subheader("Filtered View")
    st.dataframe(filtered_df, use_container_width=True)

    st.subheader("Only Changes (raw)")
    st.dataframe(changes_df, use_container_width=True)

    # --- Downloads ---
    out_name_full = f"Residual_Changes_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
    out_name_filtered = f"Residual_Changes_Filtered_{datetime.today().strftime('%Y-%m-%d')}.xlsx"

    def build_workbook_bytes(df_all, df_changes, summary):
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine='openpyxl') as writer:
            # Summary
            summary.to_excel(writer, sheet_name='Summary', index=False)
            # All Programs
            df_all.to_excel(writer, sheet_name='All_Programs', index=False)
            # Only Changes
            df_changes.to_excel(writer, sheet_name='Only_Changes', index=False)
        bio.seek(0)
        return bio

    # Full workbook (unfiltered)
    full_bio = build_workbook_bytes(result_df, changes_df, summary_df)
    st.download_button(
        label=f"Download Excel (Full) — {out_name_full}",
        data=full_bio,
        file_name=out_name_full,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    # Filtered workbook (based on current filters)
    # Recompute changes_df for the filtered subset (if Status != Changed, this tab may be empty)
    filtered_changes = filtered_df[filtered_df['Status'] == 'Changed'] if 'Status' in filtered_df.columns else pd.DataFrame()
    filtered_summary = summary_df.copy()  # Keep original summary; or compute filtered metrics if desired
    filtered_bio = build_workbook_bytes(filtered_df, filtered_changes, filtered_summary)
    st.download_button(
        label=f"Download Excel (Filtered) — {out_name_filtered}",
        data=filtered_bio,
        file_name=out_name_filtered,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# --- EOF ---
