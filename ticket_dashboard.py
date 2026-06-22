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

from datetime import date, timedelta

import altair as alt
import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# CONFIG — edit these blocks to tune behaviour. Nothing else needs changing.
# ----------------------------------------------------------------------------

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


def retest_mask_of(frame):
    """Boolean mask of rows whose State contains the retest/monitor keyword
    (e.g. 'Retest in progress', 'Retest In Test')."""
    return state_norm(frame).str.contains(RETEST_KEYWORD, case=False, na=False)


def dt_norm(frame, col):
    """Midnight-normalised datetime64 series for `col` (all-NaT if missing).
    Uses the same dayfirst convention as the rest of the app so day-first
    exports (DD/MM/YYYY) parse correctly."""
    if col and col in frame.columns:
        return pd.to_datetime(frame[col], errors="coerce", dayfirst=True).dt.normalize()
    return pd.Series(pd.NaT, index=frame.index, dtype="datetime64[ns]")


def _daily_counts(dates_dt, mask, start_ts, end_ts):
    """Per-day counts (Series indexed by python date) for rows where `mask` is
    True and the date falls within [start_ts, end_ts] inclusive. NaT-safe."""
    sel = mask & dates_dt.notna() & (dates_dt >= start_ts) & (dates_dt <= end_ts)
    return dates_dt[sel].dt.date.value_counts()


def build_range_trend(frame, created_col, changed_col, start, end):
    """Daily time-series with a continuous, zero-filled day axis over
    [start, end]. Columns: 'Created', 'Set to Monitor', 'Closed'.

    Attribution (current-state snapshot — the most a single-row-per-ticket
    export supports):
      * Created        -> Created Date (immutable; accurate).
      * Set to Monitor -> Changed Date, for rows whose CURRENT State contains
                          the retest keyword.
      * Closed         -> Changed Date, for rows whose CURRENT State is Done.

    Caveat: Changed Date reflects the LAST activity on a ticket (a comment or
    any edit also bumps it), so the two Changed-Date series mean "tickets
    currently in that state whose last activity fell on that day", not exact
    state-transition events. The cards in the UI sum these columns, so the
    chart and the metric totals are guaranteed to agree.
    """
    start_ts, end_ts = pd.Timestamp(start), pd.Timestamp(end)

    created_dt = dt_norm(frame, created_col)
    changed_dt = dt_norm(frame, changed_col)
    all_rows   = pd.Series(True, index=frame.index)

    trend = pd.DataFrame({
        "Created":        _daily_counts(created_dt, all_rows,             start_ts, end_ts),
        "Set to Monitor": _daily_counts(changed_dt, retest_mask_of(frame), start_ts, end_ts),
        "Closed":         _daily_counts(changed_dt, done_mask(frame),     start_ts, end_ts),
    })

    full_axis = pd.date_range(start_ts, end_ts, freq="D").date  # no skipped days
    trend = trend.reindex(full_axis).fillna(0).astype(int)
    trend.index.name = "Day"
    return trend


def data_date_span(frame, created_col, changed_col, fallback):
    """Min/max observed date across Created+Changed, used as the default range.
    Falls back to (fallback, fallback) when no dates are parseable."""
    span = pd.concat([dt_norm(frame, created_col), dt_norm(frame, changed_col)]).dropna()
    if span.empty:
        return fallback, fallback
    return span.min().date(), span.max().date()


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


def show_table(data, label_cols=(), fill_width=True, max_height=560, min_rows=3):
    """Render a dataframe/Styler consistently:

    * auto-height so typical tables fit without clicking fullscreen;
    * wider columns for label/index/'Grand Total' so headers aren't clipped
      (no more manual drag-to-resize, no truncated 'Grand Tota');
    * `fill_width=False` for wide pivots makes the grid exactly as wide as its
      content, which removes the odd blank/'extra empty column' on the right.

    Accepts either a pandas DataFrame or a Styler (from bold_totals)."""
    frame = data.data if hasattr(data, "data") else data   # underlying df if Styler
    rows  = max(len(frame), min_rows)
    height = int(min(38 + (rows + 1) * 35, max_height))     # header + rows, capped

    wide = {str(c) for c in label_cols} | {"Grand Total", "Objects",
                                           "Iteration Path", "Actual Priority"}
    cfg = {c: st.column_config.Column(width="medium")
           for c in frame.columns if str(c) in wide}

    st.dataframe(
        data,
        hide_index=True,
        use_container_width=fill_width,
        height=height,
        column_config=cfg or None,
    )


# Header used for the cosmetic blank spacer column appended to summary tables.
# A couple of spaces keeps the header visually empty while staying unique.
SPACER_COL = "  "


def _to_display_str(v):
    """Render a cell as a clean display string. Integers lose the '.0', NaNs
    become blank. Stringifying is what makes st.dataframe LEFT-align a column:
    the grid right-aligns numeric dtypes by default, which is what clipped the
    right-hand columns and the Grand Total."""
    if pd.isna(v):
        return ""
    if isinstance(v, float) and float(v).is_integer():
        return str(int(v))
    return str(v)


def prep_summary(df, label_cols):
    """Make a dense summary/pivot display-ready:

    * LABEL columns are stringified so they LEFT-align (keeps the readable look);
    * VALUE/count columns (everything else, incl. 'Grand Total') are kept as a
      nullable-integer dtype so the grid SORTS THEM NUMERICALLY. Descending now
      gives 151, 74, 22, 17, … instead of the old text order 9, 74, 5, 3, ….
      As a result numeric columns right-align (Streamlit's default for numbers).
    * append one blank spacer column on the far right and one blank spacer row
      at the bottom — *after* the Grand Total — so the last real column/row is
      never flush against the grid edge. The spacer row uses <NA> in the value
      columns, which renders blank and sorts to the end.

    Returns a plain DataFrame; pass it through bold_totals() afterwards for the
    Grand Total styling (the all-blank spacer row/col are never bolded)."""
    label_cols = label_cols if isinstance(label_cols, list) else [label_cols]

    out = df.copy()
    out[SPACER_COL] = ""                                   # spacer column (far right)
    blank_row = {c: "" for c in out.columns}
    out = pd.concat([out, pd.DataFrame([blank_row])], ignore_index=True)  # spacer row

    # Coerce per column AFTER the spacer row is in place, so the concat can't
    # upcast the numeric columns back to text — that upcast is exactly what made
    # the grid sort 'Grand Total' alphabetically instead of by value.
    for c in out.columns:
        if c in label_cols or c == SPACER_COL:
            out[c] = out[c].map(_to_display_str)           # text  -> left-aligned
        else:
            out[c] = pd.to_numeric(out[c], errors="coerce").astype("Int64")  # numeric -> sortable
    return out


def count_bar(frame, col, title, color):
    """Horizontal bar chart of ticket count per value in `col`, biggest first.
    Horizontal bars keep long Assignee/State labels readable. Driven by the
    same filtered frame as the table above it, so the two always agree."""
    if col not in frame.columns or frame.empty:
        st.info(f"No data for {title}.")
        return
    counts = (
        frame[col].fillna("(blank)").astype(str)
        .value_counts().rename_axis(col).reset_index(name="Tickets")
    )
    chart = (
        alt.Chart(counts)
        .mark_bar(color=color)
        .encode(
            x=alt.X("Tickets:Q", title="Tickets"),
            y=alt.Y(f"{col}:N", sort="-x", title=col),
            tooltip=[alt.Tooltip(f"{col}:N", title=col),
                     alt.Tooltip("Tickets:Q", title="Tickets")],
        )
        .properties(height=max(220, 26 * len(counts)), title=title)
    )
    st.altair_chart(chart, use_container_width=True)


# ----------------------------------------------------------------------------
# STREAMLIT UI
# ----------------------------------------------------------------------------

# Colours shared by the range chart (Created=blue, Monitor=amber, Closed=green —
# matching the 🆕 / 🔁 / ✅ snapshot metrics above).
TREND_COLOURS = {"Created": "#2563eb", "Set to Monitor": "#f59e0b", "Closed": "#16a34a"}


def render_range_tab(df_full, created_col, changed_col,
                     default_start, default_end, range_min, range_max,
                     sel_assignee, sel_iteration, sel_objects):
    """Date-range view: pick a start/end, see Created / Set-to-Monitor / Closed
    totals for the window, and a daily point-line trend of all three."""
    st.subheader("Created · Set to Monitor · Closed over a date range")
    st.caption(
        "Created uses **Created Date**; Set to Monitor and Closed use **Changed Date** "
        "with the ticket's current State. Because any edit (even a comment) updates "
        "Changed Date, read the two Changed-Date metrics as *“tickets currently in that "
        "state whose last activity fell in the window”*, not exact transition days."
    )

    pick, opts = st.columns([1.5, 1])
    with pick:
        rng = st.date_input(
            "Date range (inclusive)",
            value=(default_start, default_end),
            min_value=range_min, max_value=range_max,
            key="range_dates",
            help="Defaults to the last 30 days up to the reporting date. "
                 "Scroll the picker back to reach older data in the file.",
        )
    with opts:
        apply_filters = st.checkbox(
            "Apply sidebar filters",
            value=False,
            help="OFF = the whole file (every ticket). ON = restrict this tab to the "
                 "Assignee / Iteration Path / Objects you've narrowed in the sidebar. "
                 "The sidebar's State filter is ignored here on purpose, so Done "
                 "tickets stay visible for the Closed count.",
        )

    # date_input returns a single date mid-selection; wait for both ends.
    if not (isinstance(rng, (tuple, list)) and len(rng) == 2):
        st.info("Pick both a start and an end date to see the range view.")
        return
    start, end = rng
    if start > end:
        st.warning("Start date is after end date — adjust the range.")
        return

    base = df_full
    narrowed = []   # human-readable note of what was actually narrowed
    if apply_filters:
        if sel_assignee is not None:
            base = base[base["Assignee"].astype(str).isin(sel_assignee)]
            if sel_assignee: narrowed.append(f"{len(sel_assignee)} assignee(s)")
        if sel_iteration is not None and COL_ITERATION in base.columns:
            base = base[base[COL_ITERATION].astype(str).isin(sel_iteration)]
            if sel_iteration: narrowed.append(f"{len(sel_iteration)} iteration(s)")
        if sel_objects is not None and "Objects" in base.columns:
            base = base[base["Objects"].astype(str).isin(sel_objects)]
            if sel_objects: narrowed.append(f"{len(sel_objects)} object(s)")

    trend  = build_range_trend(base, created_col, changed_col, start, end)
    totals = trend.sum()
    span   = (end - start).days + 1

    scope = ("filtered → " + ", ".join(narrowed)) if apply_filters else "all tickets in the file"
    st.caption(f"**{start:%d %b %Y} → {end:%d %b %Y}**  ·  {span} day(s)  ·  {scope}")

    m1, m2, m3 = st.columns(3)
    m1.metric("🆕 Created in range",        int(totals["Created"]))
    m2.metric("🔁 Set to Monitor in range", int(totals["Set to Monitor"]))
    m3.metric("✅ Closed in range",         int(totals["Closed"]))

    # ---- daily point-line chart (3 series) --------------------------------
    long = trend.reset_index().melt("Day", var_name="Metric", value_name="Count")
    order = list(TREND_COLOURS.keys())
    chart = (
        alt.Chart(long)
        .mark_line(point=True, strokeWidth=2)
        .encode(
            x=alt.X("Day:T", title="Day"),
            y=alt.Y("Count:Q", title="Tickets", scale=alt.Scale(nice=True, zero=True)),
            color=alt.Color(
                "Metric:N",
                scale=alt.Scale(domain=order, range=[TREND_COLOURS[m] for m in order]),
                sort=order, title="Metric",
            ),
            tooltip=[alt.Tooltip("Day:T", title="Day"),
                     alt.Tooltip("Metric:N", title="Metric"),
                     alt.Tooltip("Count:Q", title="Count")],
        )
        .properties(height=380)
    )
    st.altair_chart(chart, use_container_width=True)

    # ---- supporting daily table + download --------------------------------
    with st.expander("📋 Daily breakdown table / download"):
        table = trend.reset_index()
        table["Day"] = pd.to_datetime(table["Day"]).dt.strftime("%Y-%m-%d")
        show_table(table, label_cols=["Day"], fill_width=False, max_height=420)
        st.download_button(
            "⬇️ Download daily breakdown (CSV)",
            data=table.to_csv(index=False).encode("utf-8"),
            file_name=f"ticket_trend_{start:%Y%m%d}_{end:%Y%m%d}.csv",
            mime="text/csv",
        )


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
            "Upload a tickets export (CSV or Excel)",
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

    # The dashboard is upload-driven: nothing renders until a file is provided.
    # An uploaded file is the only data source — no auto-loaded local path.
    if uploaded is None:
        st.info(
            "👋 **Welcome to the Ticket Dashboard.**\n\n"
            "⬅️ **Upload a tickets export (CSV or Excel) in the sidebar** to see "
            "the details, metrics and visuals."
        )
        st.stop()

    raw = load_data(uploaded)
    st.sidebar.success(f"Loaded upload: {uploaded.name}")

    today = as_of

    # Full enriched data — keeps 'Done' rows. The overview cards below read this
    # directly (so 'Closed today' works), and the sidebar filters now operate on
    # this full frame so 'Done' is a selectable State (just un-ticked by default).
    df_full = enrich(raw)
    is_done_full = done_mask(df_full)

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
    # Option lists (LOVs) are built from the FULL file, so every value — including
    # 'Done' — is selectable. Only the State filter pre-excludes 'Done' from its
    # DEFAULT selection: the active-ticket views stay Done-free out of the box,
    # but you can tick 'Done' any time to pull those tickets back in.
    with st.sidebar:
        st.header("Filters")

        def multi(label, col, exclude_done=False):
            if col not in df_full.columns:
                return None
            opts = sorted(df_full[col].dropna().astype(str).unique())
            if exclude_done:
                default = [o for o in opts
                           if o.strip().casefold() != STATE_DONE.casefold()]
            else:
                default = opts
            return st.multiselect(label, opts, default=default)

        sel_assignee  = multi("Assignee", "Assignee")
        sel_state     = multi("State", COL_STATE, exclude_done=True)
        sel_iteration = multi("Iteration Path", COL_ITERATION)
        sel_objects   = multi("Objects", "Objects")
        st.caption("‘Done’ is available in the State filter but un-ticked by "
                   "default, so active views stay Done-free unless you add it.")

    # Filter the FULL file (not a pre-stripped copy) so a ticked 'Done' flows
    # straight through. With the default selections this equals 'file minus Done'.
    f = df_full.copy()
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

    # ---- Tabbed views (one per Excel pivot, plus the date-range trend) -----
    # Default range = the last 30 days up to the reporting date ("today").
    # Bounds are widened to also cover the file's own date span, so the user can
    # still scroll the picker back to older data.
    default_end   = today
    default_start = today - timedelta(days=29)        # 30 days inclusive
    data_min, data_max = data_date_span(df_full, created_col, changed_col, today)
    range_min = min(data_min, default_start)
    range_max = max(data_max, default_end)

    tab_range, tab_state, tab_iter, tab_assignee, tab_pivot, tab_data = st.tabs(
        ["📈 Date Range", "By State", "By Iteration Path", "Assignee × State",
         "Priority × Assignee", "Raw + Download"]
    )

    with tab_range:
        render_range_tab(
            df_full, created_col, changed_col,
            default_start, default_end, range_min, range_max,
            sel_assignee, sel_iteration, sel_objects,
        )

    with tab_state:
        if COL_STATE in f.columns:
            t = count_table(f, COL_STATE, "State")
            disp = with_grand_total(t, "State")
            disp = prep_summary(disp, "State")          # left-align + spacer row/col
            left, right = st.columns([1, 1.4])
            with left:
                show_table(bold_totals(disp, "State"), label_cols=["State"])
            right.bar_chart(t.set_index("State"))  # chart excludes the total row
        else:
            st.warning(f"No '{COL_STATE}' column found.")

    with tab_iter:
        if COL_ITERATION in f.columns:
            t = count_table(f, COL_ITERATION, "Iteration Path")
            disp = with_grand_total(t, "Iteration Path")
            disp = prep_summary(disp, "Iteration Path")   # left-align + spacer row/col
            left, right = st.columns([1, 1.4])
            with left:
                show_table(bold_totals(disp, "Iteration Path"), label_cols=["Iteration Path"])
            right.bar_chart(t.set_index("Iteration Path"))  # chart excludes the total row
        else:
            st.warning(f"No '{COL_ITERATION}' column found.")

    with tab_assignee:
        st.subheader("Assignee × State")
        st.caption("Each assignee's tickets broken down by state, with Grand Total.")
        if COL_STATE in f.columns:
            piv = build_pivot(f, index="Assignee", columns=COL_STATE)
            piv = prep_summary(piv, "Assignee")          # left-align + spacer row/col
            show_table(bold_totals(piv, "Assignee"), label_cols=["Assignee"], fill_width=False)

            # ---- Two supporting charts (same filtered data as the table) -----
            st.markdown("##### Visual summary")
            st.caption("Tickets per assignee and per state for the current filters.")
            g1, g2 = st.columns(2)
            with g1:
                count_bar(f, "Assignee", "Tickets per Assignee", "#2563eb")
            with g2:
                count_bar(f, COL_STATE, "Tickets per State", "#16a34a")
        else:
            st.warning(f"No '{COL_STATE}' column found.")

    with tab_pivot:
        st.subheader("Actual Priority × Object, counted per Assignee")
        st.caption("")
        piv = build_pivot(f, index=["Actual Priority", "Objects"], columns="Assignee")
        piv = prep_summary(piv, ["Actual Priority", "Objects"])   # left-align + spacer row/col
        show_table(bold_totals(piv, ["Actual Priority", "Objects"]),
                   label_cols=["Actual Priority", "Objects"], fill_width=False)

    with tab_data:
        st.subheader("Enriched data")
        # 'Assigned To' now holds the clean name, so drop the duplicate helper column
        disp = f.drop(columns=["Assignee"]) if "Assignee" in f.columns else f
        show_table(disp, fill_width=False, max_height=600)
        csv_bytes = disp.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download enriched CSV",
            data=csv_bytes,
            file_name="tickets_enriched.csv",
            mime="text/csv",
        )


if __name__ == "__main__":
    main()
