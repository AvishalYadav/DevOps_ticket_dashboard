"""
Ticket Dashboard (Streamlit)
============================
Reads the latest tickets export (CSV or Excel), derives three columns on the
fly, and renders pivot-style views equivalent to the Excel pivot tables.

Run locally:
    pip install streamlit pandas openpyxl
    streamlit run ticket_dashboard.py

All transformation logic lives in pure functions at the top so it can be
unit-tested and so the rules are easy to edit in one place.
"""

import os
import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# CONFIG — edit these three blocks to tune behaviour. Nothing else needs changing.
# ----------------------------------------------------------------------------

# Column headers in your export. Referenced by NAME, so column order can change
# in the source file without breaking anything.
# Default local file to auto-load on startup. If this file exists it loads
# automatically (just hit Refresh after a new export). Leave the uploader as an
# override — and when deployed to a server where this path doesn't exist, the
# app simply falls back to asking for an upload.
DEFAULT_DATA_PATH = r"C:\Cursor\DevOps\eAppSys_Provider_DM.csv"

COL_ID        = "ID"
COL_TITLE     = "Title"
COL_ASSIGNEE  = "Assigned To"
COL_STATE     = "State"
COL_ITERATION = "Iteration Path"
COL_PRIORITY  = "Priority"

# Object extraction rules — mirrors your Excel formula exactly.
# (label shown in the dashboard, substring searched case-insensitively in Title)
# Order here = the order objects appear in the combined "Objects" string.
OBJECT_RULES = [
    ("Locations",       "location"),
    ("Suppliers",       "supplier"),
    ("Customer",        "customer"),
    ("AP Invoices",     "ap invoice"),
    ("AR Invoices",     "ar invoice"),
    ("AR Transactions", "ar transaction"),
    ("AR Receipts",     "ar receipt"),
    ("AR Debtor Notes", "ar debtor note"),
    ("GL Balances",     "gl balance"),
    ("GL Budgets",      "gl budget"),
]
NO_OBJECT_LABEL = "Object not assigned"

# Actual Priority — rank order of objects (rank 1 = most important).
# This matches your Image 5 (Suppliers=2, Customer=3, AP Invoices=4,
# AR Transactions=5, AR Invoices=6). If your true order differs, just
# reorder this list — that is the ONLY place priority is defined.
PRIORITY_ORDER = [
    "Locations",        # 1
    "Suppliers",        # 2
    "Customer",         # 3
    "AP Invoices",      # 4
    "AR Transactions",  # 5
    "AR Invoices",      # 6
    "AR Receipts",      # 7
    "AR Debtor Notes",  # 8
    "GL Balances",      # 9
    "GL Budgets",       # 10
]
UNASSIGNED_PRIORITY = 99  # rows with no recognised object


# ----------------------------------------------------------------------------
# TRANSFORMATIONS (pure functions — testable without Streamlit)
# ----------------------------------------------------------------------------

def clean_assignee(value):
    """'Cosmin Calin <cosmin@x.com>'  ->  'Cosmin Calin'"""
    if pd.isna(value):
        return "Unassigned"
    name = str(value).split("<")[0].strip()
    return name if name else "Unassigned"


def extract_objects(title):
    """Replicates the Excel SEARCH formula: build a comma-separated list of
    every object whose keyword appears in the Title (case-insensitive)."""
    text = "" if pd.isna(title) else str(title).lower()
    found = [label for label, keyword in OBJECT_RULES if keyword in text]
    return ", ".join(found) if found else NO_OBJECT_LABEL


def actual_priority(objects_str):
    """Lowest (best) rank among the objects present on the row."""
    if objects_str == NO_OBJECT_LABEL:
        return UNASSIGNED_PRIORITY
    objs = [o.strip() for o in objects_str.split(",")]
    ranks = [PRIORITY_ORDER.index(o) + 1 for o in objs if o in PRIORITY_ORDER]
    return min(ranks) if ranks else UNASSIGNED_PRIORITY


def enrich(df):
    """Add the three derived columns to a raw tickets dataframe."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]  # tidy stray spaces in headers

    if COL_ASSIGNEE in df.columns:
        # overwrite the full 'name <email>' with just the name, in place,
        # so the enriched export/raw view shows the name only
        df[COL_ASSIGNEE] = df[COL_ASSIGNEE].apply(clean_assignee)
        df["Assignee"] = df[COL_ASSIGNEE]
    else:
        df["Assignee"] = "Unassigned"

    title_col = COL_TITLE if COL_TITLE in df.columns else None
    df["Objects"] = (df[title_col] if title_col else "").apply(extract_objects) \
        if title_col else NO_OBJECT_LABEL
    df["Actual Priority"] = df["Objects"].apply(actual_priority)
    return df


# ----------------------------------------------------------------------------
# DATA LOADING
# ----------------------------------------------------------------------------

def load_data(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)


def count_table(df, column, label):
    """A 'value -> Count' table sorted like the Excel pivots, plus Grand Total."""
    counts = df[column].fillna("(blank)").value_counts().sort_index()
    out = counts.rename_axis(label).reset_index(name="Count")
    return out


def with_grand_total(table, label_col):
    """Append a 'Grand Total' summary row to a two-column count table."""
    total = pd.DataFrame({label_col: ["Grand Total"], "Count": [table["Count"].sum()]})
    return pd.concat([table, total], ignore_index=True)


def bold_totals(df, label_cols):
    """Return a Styler that makes the Grand Total row and column stand out
    (bold + a light fill, like the Excel total rows)."""
    label_cols = label_cols if isinstance(label_cols, list) else [label_cols]
    total_css = "font-weight: bold; background-color: #eef2ff"

    def _row_style(row):
        is_total = any(str(row[c]) == "Grand Total" for c in label_cols)
        return [total_css if is_total else "" for _ in row]

    styler = df.style.apply(_row_style, axis=1)
    if "Grand Total" in df.columns:
        styler = styler.set_properties(subset=["Grand Total"], **{
            "font-weight": "bold", "background-color": "#eef2ff",
        })
    return styler


def build_pivot(df, index, columns):
    """Count-based pivot with Grand Total margins, returned flat and Arrow-safe.

    The Grand Total margin injects the string 'Grand Total' into otherwise
    numeric index levels (e.g. Actual Priority), which crashes Streamlit's
    Arrow conversion. Flattening to columns and stringifying the label columns
    fixes that while keeping the count cells numeric.
    """
    work = df.copy()
    if COL_ID in work.columns:
        values, aggfunc = COL_ID, "count"
    else:
        work["_one"] = 1
        values, aggfunc = "_one", "sum"

    pivot = pd.pivot_table(
        work,
        index=index,
        columns=columns,
        values=values,
        aggfunc=aggfunc,
        fill_value=0,
        margins=True,
        margins_name="Grand Total",
    )
    flat = pivot.reset_index()
    # the index column(s) now mix labels with 'Grand Total' -> force to str
    index_cols = index if isinstance(index, list) else [index]
    for col in index_cols:
        flat[col] = flat[col].astype(str)
    return flat


# ----------------------------------------------------------------------------
# STREAMLIT UI
# ----------------------------------------------------------------------------

def main():
    st.set_page_config(page_title="Ticket Dashboard", layout="wide")
    st.title("🎫 Ticket Dashboard")

    with st.sidebar:
        st.header("Data source")
        uploaded = st.file_uploader(
            "Upload a tickets export (overrides the default file)",
            type=["csv", "xlsx", "xls"],
        )
        st.caption("The three derived columns are computed live — no manual prep.")

    # Priority: an uploaded file wins; otherwise auto-load the default path.
    if uploaded is not None:
        raw = load_data(uploaded)
        st.sidebar.success(f"Loaded upload: {uploaded.name}")
    elif os.path.exists(DEFAULT_DATA_PATH):
        raw = pd.read_csv(DEFAULT_DATA_PATH)
        st.sidebar.success(f"Loaded: {os.path.basename(DEFAULT_DATA_PATH)}")
        st.sidebar.caption("Re-run a fresh export to this path, then hit ↻ Rerun.")
    else:
        st.info(
            "⬅️ No default file found at the configured path. "
            "Upload a tickets CSV/Excel in the sidebar to begin."
        )
        st.stop()

    df = enrich(raw)

    # ---- Filters ----------------------------------------------------------
    with st.sidebar:
        st.header("Filters")

        def multi(label, col):
            if col not in df.columns:
                return None
            opts = sorted(df[col].dropna().astype(str).unique())
            return st.multiselect(label, opts, default=opts)

        sel_assignee  = multi("Assignee", "Assignee")
        sel_state     = multi("State", COL_STATE)
        sel_iteration = multi("Iteration Path", COL_ITERATION)
        sel_objects   = multi("Objects", "Objects")

    f = df.copy()
    if sel_assignee  is not None: f = f[f["Assignee"].astype(str).isin(sel_assignee)]
    if sel_state     is not None: f = f[f[COL_STATE].astype(str).isin(sel_state)]
    if sel_iteration is not None: f = f[f[COL_ITERATION].astype(str).isin(sel_iteration)]
    if sel_objects   is not None: f = f[f["Objects"].astype(str).isin(sel_objects)]

    # ---- KPIs -------------------------------------------------------------
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total tickets", len(f))
    c2.metric("Total assignees", f["Assignee"].nunique())
    c3.metric("States", f[COL_STATE].nunique() if COL_STATE in f.columns else 0)
    if COL_PRIORITY in f.columns:
        p1_count = int((pd.to_numeric(f[COL_PRIORITY], errors="coerce") == 1).sum())
    else:
        p1_count = 0
    c4.metric("Priority 1 tickets", p1_count, help="Tickets where Priority = 1 (highest)")

    st.divider()

    # ---- Tabbed views (one per Excel pivot) -------------------------------
    tab_state, tab_iter, tab_assignee, tab_pivot, tab_data = st.tabs(
        ["By State", "By Iteration Path", "Assignee × State", "Priority × Assignee", "Raw + Download"]
    )

    with tab_state:
        if COL_STATE in f.columns:
            t = count_table(f, COL_STATE, "State")
            disp = with_grand_total(t, "State")
            left, right = st.columns([1, 1.4])
            left.dataframe(bold_totals(disp, "State"), hide_index=True, use_container_width=True)
            right.bar_chart(t.set_index("State"))  # chart excludes the total row
        else:
            st.warning(f"No '{COL_STATE}' column found.")

    with tab_iter:
        if COL_ITERATION in f.columns:
            t = count_table(f, COL_ITERATION, "Iteration Path")
            disp = with_grand_total(t, "Iteration Path")
            st.dataframe(bold_totals(disp, "Iteration Path"), hide_index=True, use_container_width=True)
            st.bar_chart(t.set_index("Iteration Path"))  # chart excludes the total row
        else:
            st.warning(f"No '{COL_ITERATION}' column found.")

    with tab_assignee:
        st.subheader("Assignee × State")
        st.caption("Each assignee's tickets broken down by state, with Grand Total.")
        if COL_STATE in f.columns:
            piv = build_pivot(f, index="Assignee", columns=COL_STATE)
            st.dataframe(bold_totals(piv, "Assignee"), hide_index=True, use_container_width=True)
        else:
            st.warning(f"No '{COL_STATE}' column found.")

    with tab_pivot:
        st.subheader("Actual Priority × Object, counted per Assignee")
        st.caption("Equivalent of your Image 5 pivot.")
        piv = build_pivot(f, index=["Actual Priority", "Objects"], columns="Assignee")
        st.dataframe(bold_totals(piv, ["Actual Priority", "Objects"]),
                     hide_index=True, use_container_width=True)

    with tab_data:
        st.subheader("Enriched data")
        # 'Assigned To' now holds the clean name, so drop the duplicate helper column
        disp = f.drop(columns=["Assignee"]) if "Assignee" in f.columns else f
        st.dataframe(disp, use_container_width=True, hide_index=True)
        csv_bytes = disp.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download enriched CSV",
            data=csv_bytes,
            file_name="tickets_enriched.csv",
            mime="text/csv",
        )


if __name__ == "__main__":
    main()