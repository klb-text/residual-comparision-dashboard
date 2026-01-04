
# Residual Comparison Dashboard (Streamlit)
# Upload two Excel files (last month & this month), filter to "Changed" by default,
# show "What Changed", download workbooks, and push files to Dropbox.
# Fix: cache loader per-file by hashing file bytes; add "Clear cache & rerun".
# Author: M365 Copilot for Kevin Blamer

import io
import hashlib
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime

# Dropbox SDK
import dropbox
from dropbox.files import WriteMode, UploadSessionCursor, CommitInfo

st.set_page_config(page_title="Residual Comparison Dashboard", layout="wide")

st.title("Residual Comparison Dashboard")
st.caption(
    "Upload last month and this month Basic_Information workbooks (US or Canada) to compare by "
    "Residual Master Title: Terms & Optional Miles, plus Normal Miles changes."
)

# --- Top utility: clear cache & rerun ---
with st.sidebar:
    if st.button("🔄 Clear cache & rerun"):
        st.cache_data.clear()
        st.rerun()

# --- Helpers ---
REQ_COLUMNS = ["Residual Master Title", "Normal Miles", "Optional Miles", "Terms"]
OPTIONAL_COLUMNS = ["Finance Company"]  # used if present

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

@st.cache_data(show_spinner=True)
def load_sheet(_file_bytes: bytes, file_key: str):
    """
    Load the first sheet from an uploaded Excel file.
    NOTE:
      • Leading underscore excludes raw bytes from caching hash (per Streamlit docs).
      • 'file_key' IS hashed, so a different file produces a different cache entry.
    """
    df = pd.read_excel(io.BytesIO(_file_bytes), sheet_name=0, engine='openpyxl')
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQ_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")
    return df

def safe_get(series, key, default=""):
    try:
        return series.get(key, default)
    except Exception:
        return default

@st.cache_data(show_spinner=True)
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
        """Concise description of deltas."""
        notes = []
        if status.startswith("Removed"):
            notes.append("Removed program")
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
        # Changed specifics
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
            'What Changed': what_changed,
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
    changes_df = result_df[result_df['Status'] == 'Changed'].copy()

    ordered_cols = [
        'Residual Master Title', 'Finance Company', 'Status', 'What Changed',
        'Normal Miles (prev)', 'Normal Miles (curr)', 'Normal Miles Changed',
        'Terms (prev)', 'Terms (curr)', 'Terms Added', 'Terms Removed',
        'Optional Miles (prev)', 'Optional Miles (curr)', 'Optional Miles Added', 'Optional Miles Removed'
    ]
    ordered_cols = [c for c in ordered_cols if c in result_df.columns]
    result_df = result_df[ordered_cols]
    changes_df = changes_df[[c for c in ordered_cols if c in changes_df.columns]]

    return summary_df, result_df, changes_df

# --- Dropbox helpers ---
def upload_bytes_to_dropbox(dbx: dropbox.Dropbox, data: bytes, dest_path: str, overwrite: bool = True):
    """Upload bytes to Dropbox (chunked for >4MB), default overwrite mode."""
    mode = WriteMode.overwrite if overwrite else WriteMode.add
    CHUNK = 4 * 1024 * 1024  # 4MB
    size = len(data)
    if size <= CHUNK:
        dbx.files_upload(data, dest_path, mode=mode)
        return
    start = dbx.files_upload_session_start(data[:CHUNK])
    cursor = UploadSessionCursor(session_id=start.session_id, offset=CHUNK)
    while cursor.offset < size:
        next_offset = min(size, cursor.offset + CHUNK)
        chunk = data[cursor.offset:next_offset]
        if next_offset >= size:
            commit = CommitInfo(path=dest_path, mode=mode)
            dbx.files_upload_session_finish(chunk, cursor, commit)
        else:
            dbx.files_upload_session_append_v2(chunk, cursor)
            cursor.offset = next_offset

# --- UI: Uploads ---
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
        # Read bytes once and compute digest keys so the cache differentiates different files
        last_bytes = last_upl.read()
        this_bytes = this_upl.read()

        last_key = hashlib.sha256(last_bytes).hexdigest()
        this_key = hashlib.sha256(this_bytes).hexdigest()

        last_df = load_sheet(last_bytes, last_key)
        this_df = load_sheet(this_bytes, this_key)
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

    # --- Build workbooks ---
    out_name_full = f"Residual_Changes_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
    out_name_filtered = f"Residual_Changes_Filtered_{datetime.today().strftime('%Y-%m-%d')}.xlsx"

    def build_workbook_bytes(df_all, df_changes, summary):
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine='openpyxl') as writer:
            summary.to_excel(writer, sheet_name='Summary', index=False)
            df_all.to_excel(writer, sheet_name='All_Programs', index=False)
            df_changes.to_excel(writer, sheet_name='Only_Changes', index=False)
        bio.seek(0)
        return bio

    full_bio = build_workbook_bytes(result_df, changes_df, summary_df)
    filtered_changes = filtered_df[filtered_df['Status'] == 'Changed'] if 'Status' in filtered_df.columns else pd.DataFrame()
    filtered_summary = summary_df.copy()
    filtered_bio = build_workbook_bytes(filtered_df, filtered_changes, filtered_summary)

    # --- Downloads ---
    st.download_button(
        label=f"Download Excel (Full) — {out_name_full}",
        data=full_bio,
        file_name=out_name_full,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    st.download_button(
        label=f"Download Excel (Filtered) — {out_name_filtered}",
        data=filtered_bio,
        file_name=out_name_filtered,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    # --- Dropbox drop-off (optional) ---
    with st.expander("Dropbox (optional)", expanded=False):
        token = st.secrets.get("DROPBOX_ACCESS_TOKEN", None)
        folder = st.text_input("Dropbox folder path (starts with /)", value="/ResidualComparison/inbox")
        overwrite = st.checkbox("Overwrite existing files", value=True)
        st.markdown(
            "Set your `DROPBOX_ACCESS_TOKEN` in **Advanced settings → Secrets** when deploying on Streamlit Cloud.",
            help="Configure secrets safely in Cloud; never commit tokens to Git."
        )
        if not token:
            st.info("No Dropbox token found. Add `DROPBOX_ACCESS_TOKEN` via Streamlit Cloud Secrets.", icon="⚠️")
        else:
            dbx = dropbox.Dropbox(token)
            if st.button("Upload ORIGINAL files to Dropbox"):
                try:
                    name_last = last_upl.name or "last_month.xlsx"
                    name_this = this_upl.name or "this_month.xlsx"
                    dest_last = f"{folder.rstrip('/')}/{name_last}"
                    dest_this = f"{folder.rstrip('/')}/{name_this}"
                    upload_bytes_to_dropbox(dbx, last_bytes, dest_last, overwrite=overwrite)
                    upload_bytes_to_dropbox(dbx, this_bytes, dest_this, overwrite=overwrite)
                    st.success(f"Uploaded: {dest_last} and {dest_this}")
                except Exception as e:
                    st.error(f"Dropbox upload failed: {e}")

            if st.button("Upload GENERATED workbooks to Dropbox"):
                try:
                    dest_full = f"{folder.rstrip('/')}/{out_name_full}"
                    dest_filtered = f"{folder.rstrip('/')}/{out_name_filtered}"
                    upload_bytes_to_dropbox(dbx, full_bio.getvalue(), dest_full, overwrite=overwrite)
                    upload_bytes_to_dropbox(dbx, filtered_bio.getvalue(), dest_filtered, overwrite=overwrite)
                    st.success(f"Uploaded: {dest_full} and {dest_filtered}")
                except Exception as e:
                    st.error(f"Dropbox upload failed: {e}")

# --- EOF ---
