
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
    # Minimal check
    missing = [c for c in REQ_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")
    return df

@st.cache_data
def compare_frames(last_df: pd.DataFrame, this_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    # Build keyed dicts by Residual Master Title
    last_keyed = {}
    for _, row in last_df.iterrows():
        title = str(row['Residual Master Title']).strip()
        last_keyed[title] = {
            'Normal Miles': row['Normal Miles'],
            'Optional Miles': parse_list(row['Optional Miles']),
            'Terms': parse_list(row['Terms'])
        }

    this_keyed = {}
    for _, row in this_df.iterrows():
        title = str(row['Residual Master Title']).strip()
        this_keyed[title] = {
            'Normal Miles': row['Normal Miles'],
            'Optional Miles': parse_list(row['Optional Miles']),
            'Terms': parse_list(row['Terms'])
        }

    all_titles = sorted(set(last_keyed.keys()) | set(this_keyed.keys()))

    records = []
    for title in all_titles:
        prev = last_keyed.get(title)
        curr = this_keyed.get(title)
        status = 'Unchanged'
        normal_prev = prev['Normal Miles'] if prev else np.nan
        normal_curr = curr['Normal Miles'] if curr else np.nan
        opt_prev = prev['Optional Miles'] if prev else tuple()
        opt_curr = curr['Optional Miles'] if curr else tuple()
        terms_prev = prev['Terms'] if prev else tuple()
        terms_curr = curr['Terms'] if curr else tuple()

        if prev and not curr:
            status = 'Removed (only in last month)'
            terms_added, terms_removed, opt_added, opt_removed = [], list(terms_prev), [], list(opt_prev)
        elif curr and not prev:
            status = 'New (only in this month)'
            terms_added, terms_removed, opt_added, opt_removed = list(terms_curr), [], list(opt_curr), []
        else:
            normal_changed = normal_prev != normal_curr
            terms_added = [t for t in terms_curr if t not in terms_prev]
            terms_removed = [t for t in terms_prev if t not in terms_curr]
            opt_added = [m for m in opt_curr if m not in opt_prev]
            opt_removed = [m for m in opt_prev if m not in opt_curr]
            if normal_changed or terms_added or terms_removed or opt_added or opt_removed:
                status = 'Changed'
            else:
                status = 'Unchanged'

        records.append({
            'Residual Master Title': title,
            'Status': status,
            'Normal Miles (prev)': normal_prev,
            'Normal Miles (curr)': normal_curr,
            'Normal Miles Changed': (normal_prev != normal_curr),
            'Terms (prev)': ', '.join(map(str, terms_prev)) if terms_prev else '',
            'Terms (curr)': ', '.join(map(str, terms_curr)) if terms_curr else '',
            'Terms Added': ', '.join(map(str, terms_added)),
            'Terms Removed': ', '.join(map(str, terms_removed)),
            'Optional Miles (prev)': ', '.join(map(str, opt_prev)) if opt_prev else '',
            'Optional Miles (curr)': ', '.join(map(str, opt_curr)) if opt_curr else '',
            'Optional Miles Added': ', '.join(map(str, opt_added)),
            'Optional Miles Removed': ', '.join(map(str, opt_removed)),
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

    # Order columns
    ordered_cols = [
        'Residual Master Title','Status',
        'Normal Miles (prev)','Normal Miles (curr)','Normal Miles Changed',
        'Terms (prev)','Terms (curr)','Terms Added','Terms Removed',
        'Optional Miles (prev)','Optional Miles (curr)','Optional Miles Added','Optional Miles Removed'
    ]
    result_df = result_df[ordered_cols]
    changes_df = changes_df[ordered_cols]

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

    st.subheader("Summary")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Last Month", int(summary_df.loc[summary_df['Metric']=='Total programs (last month)','Value']))
    m2.metric("This Month", int(summary_df.loc[summary_df['Metric']=='Total programs (this month)','Value']))
    m3.metric("Unchanged", int(summary_df.loc[summary_df['Metric']=='Unchanged','Value']))
    m4.metric("Changed", int(summary_df.loc[summary_df['Metric']=='Changed','Value']))
    m5.metric("New / Removed", f"{int(summary_df.loc[summary_df['Metric']=='New','Value'])} / {int(summary_df.loc[summary_df['Metric']=='Removed','Value'])}")

    st.divider()
    st.subheader("Only Changes")
    st.dataframe(changes_df, use_container_width=True)

    st.subheader("All Programs")
    st.dataframe(result_df, use_container_width=True)

    # Build downloadable Excel in-memory
    out_name = f"Residual_Changes_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        result_df.to_excel(writer, sheet_name='All_Programs', index=False)
        changes_df.to_excel(writer, sheet_name='Only_Changes', index=False)
    bio.seek(0)

    st.download_button(
        label=f"Download Excel ({out_name})",
        data=bio,
        file_name=out_name,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# --- EOF ---
