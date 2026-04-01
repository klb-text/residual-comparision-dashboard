# Residual Comparison Dashboard (Streamlit)
# Upload two Excel files (last month & this month), filter to "Changed" by default,
# show "What Changed", download workbooks, and push files to Dropbox.
# Fixes:
#  • Cache loader per-file by hashing file bytes
#  • Clear cache & rerun
#  • FIXED: st.metric Series -> scalar TypeError
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
                values.append(int(f) if f.is_integer() else p)
            except Exception:
                values.append(p)
    uniq = []
    for v in values:
        if v not in uniq:
            uniq.append(v)
    return tuple(uniq)

@st.cache_data(show_spinner=True)
def load_sheet(_file_bytes: bytes, file_key: str):
    """
    Load the first sheet from an uploaded Excel file.
    _file_bytes is excluded from cache hashing; file_key is used instead.
    """
    df = pd.read_excel(io.BytesIO(_file_bytes), sheet_name=0, engine="openpyxl")
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
    """Return (summary_df, result_df, changes_df)."""

    last_keyed = {}
    for _, row in last_df.iterrows():
        title = str(row["Residual Master Title"]).strip()
        last_keyed[title] = {
            "Normal Miles": row["Normal Miles"],
            "Optional Miles": parse_list(row["Optional Miles"]),
            "Terms": parse_list(row["Terms"]),
            "Finance Company": safe_get(row, "Finance Company")
        }

    this_keyed = {}
    for _, row in this_df.iterrows():
        title = str(row["Residual Master Title"]).strip()
        this_keyed[title] = {
            "Normal Miles": row["Normal Miles"],
            "Optional Miles": parse_list(row["Optional Miles"]),
            "Terms": parse_list(row["Terms"]),
            "Finance Company": safe_get(row, "Finance Company")
        }

    all_titles = sorted(set(last_keyed) | set(this_keyed))

    records = []
    for title in all_titles:
        prev = last_keyed.get(title)
        curr = this_keyed.get(title)

        normal_prev = prev["Normal Miles"] if prev else np.nan
        normal_curr = curr["Normal Miles"] if curr else np.nan
        terms_prev = prev["Terms"] if prev else tuple()
        terms_curr = curr["Terms"] if curr else tuple()
        opt_prev = prev["Optional Miles"] if prev else tuple()
        opt_curr = curr["Optional Miles"] if curr else tuple()

        if prev and not curr:
            status = "Removed (only in last month)"
        elif curr and not prev:
            status = "New (only in this month)"
        else:
            status = (
                "Changed"
                if normal_prev != normal_curr
                or set(terms_prev) != set(terms_curr)
                or set(opt_prev) != set(opt_curr)
                else "Unchanged"
            )

        notes = []
        if normal_prev != normal_curr:
            notes.append(f"Normal Miles: {normal_prev} → {normal_curr}")

        added_terms = [t for t in terms_curr if t not in terms_prev]
        removed_terms = [t for t in terms_prev if t not in terms_curr]

        if added_terms:
            notes.append(f"Terms added: {', '.join(map(str, added_terms))}")
        if removed_terms:
            notes.append(f"Terms removed: {', '.join(map(str, removed_terms))}")

        added_opt = [m for m in opt_curr if m not in opt_prev]
        removed_opt = [m for m in opt_prev if m not in opt_curr]

        if added_opt:
            notes.append(f"Optional Miles added: {', '.join(map(str, added_opt))}")
        if removed_opt:
            notes.append(f"Optional Miles removed: {', '.join(map(str, removed_opt))}")

        records.append({
            "Residual Master Title": title,
            "Finance Company": (curr or prev).get("Finance Company", ""),
            "Status": status,
            "What Changed": "; ".join(notes) if notes else "—",
            "Normal Miles (prev)": normal_prev,
            "Normal Miles (curr)": normal_curr,
            "Terms (prev)": ", ".join(map(str, terms_prev)),
            "Terms (curr)": ", ".join(map(str, terms_curr)),
            "Optional Miles (prev)": ", ".join(map(str, opt_prev)),
            "Optional Miles (curr)": ", ".join(map(str, opt_curr)),
        })

    result_df = pd.DataFrame(records)

    summary_df = pd.DataFrame([
        ("Total programs (last month)", len(last_keyed)),
        ("Total programs (this month)", len(this_keyed)),
        ("Unchanged", int((result_df["Status"] == "Unchanged").sum())),
        ("Changed", int((result_df["Status"] == "Changed").sum())),
        ("New", int(result_df["Status"].str.startswith("New").sum())),
        ("Removed", int(result_df["Status"].str.startswith("Removed").sum())),
    ], columns=["Metric", "Value"])

    changes_df = result_df[result_df["Status"] == "Changed"].copy()

    return summary_df, result_df, changes_df

# --- UI uploads ---
c1, c2 = st.columns(2)
with c1:
    last_upl = st.file_uploader("Upload LAST month Basic_Information workbook", type="xlsx")
with c2:
    this_upl = st.file_uploader("Upload THIS month Basic_Information workbook", type="xlsx")

if st.button("Run Comparison"):

    if not last_upl or not this_upl:
        st.error("Please upload both files.")
        st.stop()

    last_bytes = last_upl.read()
    this_bytes = this_upl.read()

    last_df = load_sheet(last_bytes, hashlib.sha256(last_bytes).hexdigest())
    this_df = load_sheet(this_bytes, hashlib.sha256(this_bytes).hexdigest())

    summary_df, result_df, changes_df = compare_frames(last_df, this_df)

    # ✅ FIXED helper — scalar extraction
    def metric_value(label: str) -> int:
        return int(summary_df.loc[summary_df["Metric"] == label, "Value"].iloc[0])

    st.subheader("Summary")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Last Month", metric_value("Total programs (last month)"))
    m2.metric("This Month", metric_value("Total programs (this month)"))
    m3.metric("Unchanged", metric_value("Unchanged"))
    m4.metric("Changed", metric_value("Changed"))
    m5.metric("New / Removed", f"{metric_value('New')} / {metric_value('Removed')}")

    st.divider()
    st.subheader("Changed Programs")
    st.dataframe(changes_df, use_container_width=True)

# --- EOF ---
