"""
Ticket Dashboard (Streamlit)
============================
Reads the latest tickets export (CSV or Excel), derives extra columns on the
fly, and renders pivot-style views equivalent to the Excel pivot tables.

Run locally:
    pip install streamlit pandas openpyxl
    streamlit run ticket_dashboard.py

All transformation logic lives in pure functions at the top so it can be
unit-tested and so the rules are easy to edit in one place.
"""

import os
from datetime import date

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# CONFIG — edit these blocks to tune behaviour. Nothing else needs changing.
# ----------------------------------------------------------------------------

# Default local file to auto-load on startup. If this file exists it loads
# automatically (just hit Refresh after a new export). Leave the uploader as an
# override — and when deployed to a server where this path doesn't exist, the
# app simply falls back to asking for an upload.
DEFAULT_DATA_PATH = r"C:\Cursor\DevOps\eAppSys_Provider_DM.csv"

# Column headers in your export. Referenced by NAME, so column order can change
# in the source file without breaking anything.
COL_ID        = "ID"
COL_TITLE     = "Title"
COL_ASSIGNEE  = "Assigned To"
COL_STATE     = "State"
COL_ITERATION = "Iteration Path"
COL_PRIORITY  = "Priority"
COL_CREATED   = "Created Date"   # used for "Created today"
COL_CHANGED   = "Changed Date"   # used for "Set to Monitor today" / "Closed today"

# Extra header spellings tolerated for the date columns. Matching is
# case-insensitive and ignores stray spaces, so "Changed date", "ChangedDate",
# etc. all resolve. The first existing match (in order) wins.
CREATED_DATE_CANDIDATES = [COL_CREATED, "CreatedDate", "Created", "Created On", "Create Date"]
CHANGED_DATE_CANDIDATES = [COL_CHANGED, "ChangedDate", "Changed", "Changed On",
                           "Last Changed Date", "State Change Date", "State Changed Date"]

# State values (compared case-insensitively, stray spaces ignored).
STATE_DONE     = "Done"      # the "closed" bucket
RETEST_KEYWORD = "retest"    # ANY State containing this counts ("Retest in progress",
                             # "Retest In Test", ...) — substring, case-insensitive

# eAppsys team. A ticket counts as "eAppsys" if its assignee CONTAINS any of
# these names, so trailing identifiers (e.g. a name with a trailing "1")
# still match. Edit this list to add/remove people.
EAPPSYS_TEAM = [
    "Allen Kalathur",
    "Raghavendra Pai",
    "Guna Kurmadasu",
    "Prathyusha Challuri",
    "Sathwika Thallapelli",
    "Vishal Yadav",
    "Alexander Hoff",
    "Varun Jukanti",
    "Anand Satti",
    "Akshith Gouravelli",
    "Bharath Bekkam",
]

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


def is_eappsys(value):
    """True if the assignee name contains any eAppsys team member.
    Substring match handles trailing identifiers (e.g. a trailing '1')."""
    if pd.isna(value):
        return False
    name = str(value).casefold()
    return any(member.casefold() in name for member in EAPPSYS_TEAM)


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


def find_col(df, candidates):
    """Resolve the real column name in `df` matching any candidate, ignoring
    case and stray spaces. Returns the actual column name, or None."""
    norm = {str(c).strip().casefold(): c for c in df.columns}
    for cand in candidates:
        hit = norm.get(str(cand).strip().casefold())
        if hit is not None:
            return hit
    return None


def date_series(frame, col):
    """Parsed DATE-ONLY series (time dropped) for the resolved column `col`;
    all-NaT if `col` is None/missing. Robust to mixed string/Excel formats."""
    if not col or col not in frame.columns:
        return pd.Series([pd.NaT] * len(frame), index=frame.index)
    return pd.to_datetime(frame[col], errors="coerce", dayfirst=True).dt.date


def state_norm(frame):
    """Normalised (stripped, casefolded) State series for safe comparisons."""
    if COL_STATE not in frame.columns:
        return pd.Series([""] * len(frame), index=frame.index)
    return frame[COL_STATE].astype(str).str.strip().str.casefold()


def done_mask(frame):
    """Boolean mask of rows whose State is 'Done'."""
    return state_norm(frame) == STATE_DONE.casefold()


def enrich(df):
    """Add the derived columns to a raw tickets dataframe."""
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
    """Count-based pivot with Grand Total margins, returned flat and Arrow-safe."""
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
    # Trim Streamlit's large default top padding and use a compact heading so
    # more of the dashboard is visible without scrolling.
    st.markdown(
        """
        <style>
          .block-container {padding-top: 0.6rem; padding-bottom: 1rem;}
          [data-testid="stVerticalBlock"] {gap: 0.55rem;}
          hr {margin: 0.4rem 0;}
          h3 {margin: 0.1rem 0 0.3rem 0;}
          [data-testid="stMetric"] {padding: 0;}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("### 🎫 Ticket Dashboard")

    with st.sidebar:
        st.header("Data source")
        uploaded = st.file_uploader(
            "Upload a tickets export (overrides the default file)",
            type=["csv", "xlsx", "xls"],
        )
        st.caption("The derived columns are computed live — no manual prep.")

        st.header("Reporting date")
        as_of = st.date_input(
            "Treat this date as “today”",
            value=date.today(),
            help="Auto-set to the system date. Override it for any as-of date "
                 "or if the server clock is in a different timezone.",
        )

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

    today = as_of

    # Full enriched data — keeps 'Done' rows, which are needed ONLY for the
    # "Closed today" metric. Everything else works off `df` (Done removed).
    df_full = enrich(raw)
    is_done_full = done_mask(df_full)
    df = df_full[~is_done_full].copy()   # active tickets — Done excluded

    # ---- Overview metrics (daily snapshot — NOT affected by the filters) ---
    # Resolve the date columns tolerantly (handles 'Changed date', 'ChangedDate', ...)
    created_col = find_col(df_full, CREATED_DATE_CANDIDATES)
    changed_col = find_col(df_full, CHANGED_DATE_CANDIDATES)

    created_full = date_series(df_full, created_col)   # date-only (time dropped)
    changed_full = date_series(df_full, changed_col)   # date-only (time dropped)
    state_full   = state_norm(df_full)

    retest_mask  = state_full.str.contains(RETEST_KEYWORD, case=False, na=False)

    new_today    = int((created_full == today).sum())
    eappsys_cnt  = int((df_full["Assignee"].apply(is_eappsys) & (~is_done_full)).sum())
    retest_today = int((retest_mask & (changed_full == today)).sum())
    closed_today = int((is_done_full & (changed_full == today)).sum())

    has_created = created_col is not None
    has_changed = changed_col is not None

    # Quick self-check so a 0 is explainable (column not found vs genuinely none).
    with st.sidebar.expander("🔍 Data check"):
        st.write(f"Reporting date: **{today}**")
        st.write(f"Created-date column: **{created_col or 'NOT FOUND'}**")
        st.write(f"Changed-date column: **{changed_col or 'NOT FOUND'}**")
        st.write(f"Rows changed today: **{int((changed_full == today).sum())}**")
        st.write(f"Done rows in file: **{int(is_done_full.sum())}**")
        st.write(f"Retest-state rows: **{int(retest_mask.sum())}**")

    st.caption(f"Daily snapshot for **{today:%a %d %b %Y}** "
               f"(Done excluded everywhere except *Closed today*).")
    o1, o2, o3, o4 = st.columns(4)
    o1.metric(
        "🆕 Created today", new_today,
        help=(f"Tickets where {created_col} (date only) = {today}."
              if has_created else "No Created-date column found in this export."),
    )
    o2.metric(
        "🏢 eAppsys tickets", eappsys_cnt,
        help="Active tickets assigned to the eAppsys team (substring match, so "
             "trailing identifiers still count). Done excluded.",
    )
    o3.metric(
        "🔁 Set to Monitor today", retest_today,
        help=(f"State contains '{RETEST_KEYWORD}' (e.g. Retest in progress, Retest In "
              f"Test) and {changed_col} (date only) = {today}."
              if has_changed else "No Changed-date column found in this export."),
    )
    o4.metric(
        "✅ Closed today", closed_today,
        help=(f"State = '{STATE_DONE}' and {changed_col} (date only) = {today}. "
              "The only metric that includes Done."
              if has_changed else "No Changed-date column found in this export."),
    )

    st.divider()

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

    # ---- KPIs (filtered view — Done already excluded) ---------------------
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
            left, right = st.columns([1, 1.4])
            left.dataframe(bold_totals(disp, "Iteration Path"), hide_index=True, use_container_width=True)
            right.bar_chart(t.set_index("Iteration Path"))  # chart excludes the total row
        else:
            st.warning(f"No '{COL_ITERATION}' column found.")

    with tab_assignee:
        st.subheader("Assignee × State")
        st.caption("Each assignee's tickets broken down by state, with Grand Total.")
        if COL_STATE in f.columns:
            piv = build_pivot(f, index="Assignee", columns=COL_STATE)
            st.dataframe(bold_totals(piv, "Assignee"), hide_index=True, use_container_width=False)
        else:
            st.warning(f"No '{COL_STATE}' column found.")

    with tab_pivot:
        st.subheader("Actual Priority × Object, counted per Assignee")
        st.caption("")
        piv = build_pivot(f, index=["Actual Priority", "Objects"], columns="Assignee")
        st.dataframe(bold_totals(piv, ["Actual Priority", "Objects"]),
                     hide_index=True, use_container_width=False)

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
