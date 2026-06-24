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
from html import escape as html_escape
import json

import altair as alt
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

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
# STATIC SUMMARY-TABLE RENDERER
# ----------------------------------------------------------------------------
# The interactive grid (st.dataframe) can't do two things we now need: keep the
# Grand Total frozen at the bottom while a user sorts, and colour the header row.
# So the four summary/pivot views render as STATIC styled HTML tables instead —
# think of an Excel sheet with a coloured header and a totals row that stays put
# no matter how you sort the rows above it. Because nothing is interactive here,
# no spacer rows/columns are needed (that hack only existed to coax the grid),
# which also removes the trailing blank line(s).

TOTAL_LABEL = "Grand Total"
HEADER_BG   = "#4338ca"   # indigo-700 — prominent, clearly distinct from rows
HEADER_FG   = "#ffffff"
TOTAL_FILL  = "#e0e7ff"   # indigo-100 — Grand Total row/column tint
STRIPE_FILL = "#f8fafc"   # subtle zebra striping for readability
GRID_LINE   = "#e5e7eb"


def _is_total_row(row, label_cols):
    """A row is the Grand Total row if any label cell equals 'Grand Total'."""
    return any(str(row[c]).strip() == TOTAL_LABEL for c in label_cols)


_TABLE_TEMPLATE = r"""<!DOCTYPE html><html><head><meta charset="utf-8">
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<style>
  *{box-sizing:border-box;}
  html,body{margin:0;padding:0;font-family:"Source Sans Pro",system-ui,-apple-system,Segoe UI,sans-serif;}
  .bar{display:flex;gap:8px;align-items:center;margin:0 0 8px;}
  .bar button{font:inherit;font-size:0.82rem;font-weight:600;cursor:pointer;
              border:1px solid var(--grid);background:#fff;color:#312e81;
              padding:5px 11px;border-radius:7px;}
  .bar button:hover{background:#eef2ff;border-color:#c7d2fe;}
  .bar #msg{font-size:0.8rem;color:#16a34a;font-weight:600;}
  .wrap{overflow:auto;border:1px solid var(--grid);border-radius:8px;}
  table{border-collapse:separate;border-spacing:0;width:100%;font-size:0.9rem;}
  th,td{padding:7px 14px;white-space:nowrap;border-bottom:1px solid var(--grid);border-right:1px solid var(--grid);}
  thead th{position:sticky;top:0;z-index:3;background:var(--hbg);color:var(--hfg);
           font-weight:700;text-align:right;cursor:pointer;user-select:none;}
  thead th.label{text-align:left;}
  thead th.spacer{cursor:default;}
  thead th:hover:not(.spacer){filter:brightness(1.12);}
  td.num{text-align:right;color:#1f2937;}
  td.label{text-align:left;color:#1f2937;font-weight:600;}
  tbody tr.data td{background:#fff;}
  tbody tr.data:nth-child(even) td{background:var(--stripe);}
  th.stick,td.stick{position:sticky;left:0;z-index:2;}
  thead th.stick{z-index:4;}
  tbody tr.data td.stick{background:#fff;}
  tbody tr.data:nth-child(even) td.stick{background:var(--stripe);}
  tr.total td{background:var(--tfill)!important;font-weight:700;color:#312e81;}
  td.totalcol{background:var(--tfill)!important;font-weight:700;}
  tr.spacer td{height:30px;background:#fff;border-bottom:none;}
  td.spacercol{min-width:42px;background:#fff;}
  th.spacer{min-width:42px;}
  .arr{font-size:0.7rem;margin-left:5px;opacity:0.95;}
  /* During capture: expand fully and drop sticky so the PNG shows every row/column */
  .cap .wrap{overflow:visible!important;max-height:none!important;}
  .cap th,.cap td{position:static!important;}
</style></head><body>
<div class="bar">
  <button id="copyBtn">📋 Copy image</button>
  <button id="pngBtn">⬇ Download PNG</button>
  <button id="maxBtn">🔳 Max view</button>
  <span id="msg"></span>
</div>
<div class="wrap"><table id="t"><thead><tr id="hr"></tr></thead><tbody id="b"></tbody></table></div>
<script>
const P=__PAYLOAD__;
const R=document.documentElement.style;
R.setProperty('--hbg',P.headerBg);R.setProperty('--hfg',P.headerFg);
R.setProperty('--tfill',P.totalFill);R.setProperty('--stripe',P.stripe);R.setProperty('--grid',P.grid);
const totalColIdx=P.cols.findIndex(c=>c.name===P.totalColName);
let sortIdx=P.defaultIdx, sortDir='desc';
function colClass(c,i){let k=c.spacer?'spacercol':(c.label?'label':'num');if(i===0&&c.label)k+=' stick';if(i===totalColIdx)k+=' totalcol';return k;}
function sortData(){if(sortIdx<0)return;const num=P.cols[sortIdx].sortNum,dir=sortDir==='asc'?1:-1;
  P.data.sort((ra,rb)=>{let a=ra[sortIdx].k,b=rb[sortIdx].k;
    if(num){a=(a===null?-Infinity:a);b=(b===null?-Infinity:b);return (a-b)*dir;}
    a=String(a);b=String(b);return a<b?-dir:a>b?dir:0;});}
function header(){const hr=document.getElementById('hr');hr.innerHTML='';
  P.cols.forEach((c,i)=>{const th=document.createElement('th');let cls=c.spacer?'spacer':(c.label?'label':'num');if(i===0&&c.label)cls+=' stick';th.className=cls;
    th.textContent=c.name;
    if(i===sortIdx){const s=document.createElement('span');s.className='arr';s.textContent=sortDir==='asc'?'\u25B2':'\u25BC';th.appendChild(s);}
    if(!c.spacer)th.onclick=()=>{if(sortIdx===i)sortDir=(sortDir==='asc'?'desc':'asc');else{sortIdx=i;sortDir=(c.sortNum?'desc':'asc');}sortData();render();};
    hr.appendChild(th);});}
function mkRow(cells,rowCls){const tr=document.createElement('tr');tr.className=rowCls;
  cells.forEach((cell,i)=>{const td=document.createElement('td');td.className=colClass(P.cols[i],i);td.textContent=cell.t;tr.appendChild(td);});return tr;}
function render(){header();const b=document.getElementById('b');b.innerHTML='';
  P.data.forEach(r=>b.appendChild(mkRow(r,'data')));
  P.totals.forEach(r=>b.appendChild(mkRow(r,'total')));}
function flash(t,col){const m=document.getElementById('msg');m.style.color=col||'#16a34a';m.textContent=t;setTimeout(()=>{if(m.textContent===t)m.textContent='';},4000);}
function snap(cb){
  if(typeof html2canvas==='undefined'){flash('Image tool blocked by network — try the CSV on Raw tab','#b91c1c');return;}
  document.body.classList.add('cap');
  const node=document.querySelector('.wrap');
  html2canvas(node,{backgroundColor:'#ffffff',scale:2,scrollX:0,scrollY:0,
                    windowWidth:node.scrollWidth,windowHeight:node.scrollHeight})
    .then(c=>{document.body.classList.remove('cap');cb(c);})
    .catch(()=>{document.body.classList.remove('cap');flash('Could not render image','#b91c1c');});}
document.getElementById('copyBtn').onclick=()=>snap(c=>c.toBlob(async b=>{
  try{await navigator.clipboard.write([new ClipboardItem({'image/png':b})]);
      flash('Copied! Paste with Ctrl+V');}
  catch(e){flash('Clipboard blocked — use Download PNG instead','#b91c1c');}},'image/png'));
document.getElementById('pngBtn').onclick=()=>snap(c=>c.toBlob(b=>{
  const a=document.createElement('a');a.href=URL.createObjectURL(b);
  a.download=(P.title||'table')+'.png';a.click();
  flash('PNG downloaded');},'image/png'));
document.getElementById('maxBtn').onclick=()=>{
  const w=window.open('','_blank');
  if(!w){flash('Pop-up blocked — allow pop-ups, then retry','#b91c1c');return;}
  // Clone the WHOLE interactive widget so sorting/copy work in the new tab too.
  let doc='<!DOCTYPE html>'+document.documentElement.outerHTML;
  doc=doc.replace('</head>','<style>body{padding:18px;}#maxBtn{display:none;}'+
      '.wrap{max-height:none!important;}</style></head>');
  w.document.open();w.document.write(doc);w.document.close();
};
sortData();render();
</script></body></html>"""


def render_summary_table(df, label_cols, sort_by=None, max_height=720, title="table"):
    """Render a count/pivot table as an INTERACTIVE, self-contained widget.

    Built as HTML + a little vanilla JS (rendered via st.components.v1.html, so
    no pip install is needed). It restores the grid features that the static
    table lost, while keeping the styling the interactive st.dataframe can't do:

      * Coloured, prominent, sticky header (requirement: coloured header).
      * Click any header to SORT (asc/desc toggle, arrow shown).
      * First label column and the header are FROZEN/sticky while scrolling.
      * Grand Total ROW is PINNED at the bottom — excluded from sorting, so it
        never floats to the top; a Grand Total COLUMN (pivots) is tinted + bold.
      * Exactly ONE blank spacer row and ONE blank spacer column for clean
        screenshot framing (nothing clipped at the edges).
      * Auto-sized height so the whole table shows in one view for screenshots.
      * `sort_by=False` keeps the given order (Actual Priority 1..n) initially.
    """
    label_cols = label_cols if isinstance(label_cols, list) else [label_cols]
    work = df.copy()

    total_mask = work.apply(lambda r: _is_total_row(r, label_cols), axis=1)
    totals, body = work[total_mask], work[~total_mask]

    # Initial order: descending by sort column (unless caller keeps order).
    if sort_by is not False:
        if sort_by is None:
            sort_by = (TOTAL_LABEL if TOTAL_LABEL in work.columns
                       else next((c for c in work.columns if c not in label_cols), None))
        if sort_by and sort_by in body.columns:
            body = (body.assign(_s=pd.to_numeric(body[sort_by], errors="coerce"))
                        .sort_values("_s", ascending=False, kind="stable")
                        .drop(columns="_s"))

    cols = list(work.columns)
    value_cols = [c for c in cols if c not in label_cols]

    # A column sorts numerically if every data value parses as a number — this
    # makes 'Actual Priority' sort 1,2,…,10 instead of as text ("10" before "2").
    def _numeric_col(col):
        if col not in label_cols:
            return True
        s = pd.to_numeric(body[col], errors="coerce")
        return len(s) > 0 and bool(s.notna().all())
    numeric_cols = {c: _numeric_col(c) for c in cols}

    def disp(col, v):
        if col in label_cols:
            return "" if pd.isna(v) else str(v)
        num = pd.to_numeric(pd.Series([v]), errors="coerce").iloc[0]
        return "" if pd.isna(num) else f"{num:,.0f}"

    def sort_key(col, v):
        if numeric_cols[col]:
            num = pd.to_numeric(pd.Series([v]), errors="coerce").iloc[0]
            return None if pd.isna(num) else float(num)
        return "" if pd.isna(v) else str(v).lower()

    col_meta = [{"name": str(c), "num": (c not in label_cols),
                 "label": (c in label_cols), "sortNum": numeric_cols[c]}
                for c in cols]

    def row_payload(series):
        return [{"t": disp(c, series[c]), "k": sort_key(c, series[c])} for c in cols]

    data_rows = [row_payload(r) for _, r in body.iterrows()]
    total_rows = [row_payload(r) for _, r in totals.iterrows()]

    if sort_by is False:
        default_idx = -1                        # keep priority order, no active sort
    else:
        default_idx = cols.index(sort_by) if (sort_by in cols) else -1

    payload = {
        "cols": col_meta, "data": data_rows, "totals": total_rows,
        "totalColName": TOTAL_LABEL, "defaultIdx": default_idx, "title": title,
        "headerBg": HEADER_BG, "headerFg": HEADER_FG,
        "totalFill": TOTAL_FILL, "stripe": STRIPE_FILL, "grid": GRID_LINE,
    }

    # Auto height so the full table + toolbar shows in one view.
    visible_rows = len(data_rows) + len(total_rows)
    height = min(max_height, 52 + 40 + visible_rows * 34 + 22)   # +40 = screenshot toolbar
    height = max(height, 190)

    html = _TABLE_TEMPLATE.replace("__PAYLOAD__", json.dumps(payload))
    components.html(html, height=int(height), scrolling=True)


# ----------------------------------------------------------------------------
# KPI OVERVIEW WIDGET  (the snapshot + filtered KPIs as one copyable image)
# ----------------------------------------------------------------------------
_KPI_TEMPLATE = r"""<!DOCTYPE html><html><head><meta charset="utf-8">
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<style>
 *{box-sizing:border-box;}
 html,body{margin:0;padding:0;font-family:"Source Sans Pro",system-ui,-apple-system,Segoe UI,sans-serif;}
 .bar{display:flex;gap:8px;align-items:center;margin:0 0 10px;}
 .bar button{font:inherit;font-size:0.82rem;font-weight:600;cursor:pointer;border:1px solid #e5e7eb;
             background:#fff;color:#312e81;padding:5px 11px;border-radius:7px;}
 .bar button:hover{background:#eef2ff;border-color:#c7d2fe;}
 .bar #msg{font-size:0.8rem;color:#16a34a;font-weight:600;}
 .grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;}
 .card{border:1px solid #e5e7eb;border-radius:12px;padding:14px 16px;background:#fff;
       border-top:4px solid var(--a,#4338ca);box-shadow:0 1px 2px rgba(15,23,42,.04);}
 .card .lbl{font-size:0.82rem;color:#475569;font-weight:600;margin-bottom:6px;
            white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
 .card .val{font-size:1.9rem;font-weight:800;color:#0f172a;line-height:1;}
 #cap{background:#fff;}
 #hd{font-size:0.92rem;color:#334155;font-weight:600;margin:0 0 12px;}
 #hd b{color:#0f172a;}
 @media(max-width:760px){.grid{grid-template-columns:repeat(2,1fr);}}
</style></head><body>
<div class="bar"><button id="copyBtn">📋 Copy image</button><button id="pngBtn">⬇ Download PNG</button><span id="msg"></span></div>
<div id="cap"><div id="hd"></div><div class="grid" id="g"></div></div>
<script>
const C=__PAYLOAD__;
if(C.heading){document.getElementById('hd').innerHTML=C.heading;}else{document.getElementById('hd').style.display='none';}
const g=document.getElementById('g');
C.cards.forEach(c=>{const d=document.createElement('div');d.className='card';d.style.setProperty('--a',c.accent||'#4338ca');
 const l=document.createElement('div');l.className='lbl';l.textContent=c.label;if(c.help)l.title=c.help;
 const v=document.createElement('div');v.className='val';v.textContent=c.value;
 d.appendChild(l);d.appendChild(v);g.appendChild(d);});
function flash(t,col){const m=document.getElementById('msg');m.style.color=col||'#16a34a';m.textContent=t;setTimeout(()=>{if(m.textContent===t)m.textContent='';},4000);}
function snap(cb){if(typeof html2canvas==='undefined'){flash('Image tool blocked by network','#b91c1c');return;}
 html2canvas(document.getElementById('cap'),{backgroundColor:'#ffffff',scale:2}).then(cb).catch(()=>flash('Could not render image','#b91c1c'));}
document.getElementById('copyBtn').onclick=()=>snap(c=>c.toBlob(async b=>{try{await navigator.clipboard.write([new ClipboardItem({'image/png':b})]);flash('Copied! Paste with Ctrl+V');}catch(e){flash('Clipboard blocked — use Download PNG','#b91c1c');}},'image/png'));
document.getElementById('pngBtn').onclick=()=>snap(c=>c.toBlob(b=>{const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=(C.title||'overview')+'.png';a.click();flash('PNG downloaded');},'image/png'));
</script></body></html>"""


def render_kpi_cards(cards, title="overview", heading=""):
    """Render KPI cards as a self-contained widget with a Copy-image button. The
    optional heading (with the snapshot date) sits INSIDE the captured area, so
    the copied image includes it."""
    import math
    payload = {"cards": cards, "title": title, "heading": heading}
    rows = math.ceil(len(cards) / 4)
    height = 50 + (28 if heading else 0) + rows * 104 + 14
    html = _KPI_TEMPLATE.replace("__PAYLOAD__", json.dumps(payload))
    components.html(html, height=int(height), scrolling=False)


# ----------------------------------------------------------------------------
# STREAMLIT UI
# ----------------------------------------------------------------------------

def render_welcome():
    """A visual landing page shown before any file is uploaded."""
    st.markdown(
        """
<div style="margin-top:6px;padding:34px 30px;border-radius:18px;
     background:linear-gradient(135deg,#4338ca 0%,#6366f1 55%,#0ea5e9 100%);color:#fff;">
  <div style="font-size:2.3rem;font-weight:800;line-height:1.1;">🎫 Ticket Dashboard</div>
  <div style="font-size:1.08rem;opacity:.95;margin-top:10px;max-width:760px;">
    Turn a raw tickets export into an interactive, screenshot-ready report in
    seconds — trends, pivots, and one-click images.
  </div>
  <div style="display:inline-block;margin-top:18px;padding:11px 18px;border-radius:10px;
       background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.45);
       font-weight:700;font-size:1rem;">
    ⬅️ Upload a CSV or Excel export in the sidebar to begin
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    cards = [
        ("📈", "Ticket Trends", "Created · Set to Monitor · Closed over any date range, with toggleable series."),
        ("📊", "Smart pivots", "Break tickets down by State, Iteration Path, Assignee, and Priority."),
        ("🎨", "Clear tables", "Coloured headers, click-to-sort, frozen first column, pinned Grand Total."),
        ("📋", "One-click copy", "Copy any table or the KPI overview as an image — then paste it wherever you need."),
        ("🔍", "Max view", "Open any table full-screen for easy reading of wide pivots."),
        ("⬇️", "Clean export", "Download the enriched data as CSV, with derived columns computed for you."),
    ]
    cells = "".join(
        f"""<div style="border:1px solid #e5e7eb;border-radius:14px;padding:16px 18px;background:#fff;">
              <div style="font-size:1.5rem;">{icon}</div>
              <div style="font-weight:700;color:#0f172a;margin:6px 0 4px;">{title}</div>
              <div style="font-size:0.9rem;color:#475569;line-height:1.35;">{desc}</div>
            </div>"""
        for icon, title, desc in cards
    )
    st.markdown(
        f"""
<div style="margin-top:22px;font-weight:700;color:#334155;font-size:1.05rem;">What you'll get</div>
<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-top:12px;">{cells}</div>
<div style="margin-top:20px;color:#64748b;font-size:0.88rem;">
  Accepted files: <b>.csv</b>, <b>.xlsx</b>, <b>.xls</b> &nbsp;·&nbsp;
  Your file is processed in this session only — nothing is uploaded to a server by the dashboard itself.
</div>
""",
        unsafe_allow_html=True,
    )

# Colours shared by the range chart (Created=blue, Monitor=amber, Closed=green —
# matching the 🆕 / 🔁 / ✅ snapshot metrics above).
TREND_COLOURS = {"Created": "#2563eb", "Set to Monitor": "#f59e0b", "Closed": "#16a34a"}


def render_range_tab(df_full, created_col, changed_col,
                     default_start, default_end, range_min, range_max,
                     sel_assignee, sel_iteration, sel_objects):
    """Date-range view: pick a start/end, see Created / Set-to-Monitor / Closed
    totals for the window, and a daily point-line trend of all three."""
    st.subheader("Ticket Trends — Created · Set to Monitor · Closed over a date range")

    pick, opts = st.columns([1.5, 1])
    with pick:
        rng = st.date_input(
            "Trend date range (inclusive)",
            value=(default_start, default_end),
            min_value=range_min, max_value=range_max,
            key="range_dates",
            help="Defaults to the last 7 days up to the reporting date. "
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

    # ---- Series toggles: pick any combination; all three on by default --------
    st.caption("Show series:")
    s1, s2, s3 = st.columns(3)
    show_created = s1.checkbox("🆕 Created",        value=True, key="trend_show_created")
    show_monitor = s2.checkbox("🔁 Set to Monitor", value=True, key="trend_show_monitor")
    show_closed  = s3.checkbox("✅ Closed",          value=True, key="trend_show_closed")
    selected = [name for name, on in (
        ("Created", show_created),
        ("Set to Monitor", show_monitor),
        ("Closed", show_closed),
    ) if on]

    # date_input returns a single date mid-selection; wait for both ends.
    if not (isinstance(rng, (tuple, list)) and len(rng) == 2):
        st.info("Pick both a start and an end date to see the range view.")
        return
    start, end = rng
    if start > end:
        st.warning("Start date is after end date — adjust the range.")
        return
    if not selected:
        st.info("Tick at least one series above (Created / Set to Monitor / Closed) "
                "to see the trend.")
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
    trend  = trend[selected]                       # honour the series checkboxes
    totals = trend.sum()
    span   = (end - start).days + 1

    scope = ("filtered → " + ", ".join(narrowed)) if apply_filters else "all tickets in the file"
    st.caption(f"**{start:%d %b %Y} → {end:%d %b %Y}**  ·  {span} day(s)  ·  {scope}")

    metric_labels = {
        "Created":        "🆕 Created in range",
        "Set to Monitor": "🔁 Set to Monitor in range",
        "Closed":         "✅ Closed in range",
    }
    cols = st.columns(len(selected))
    for col, name in zip(cols, selected):
        col.metric(metric_labels[name], int(totals[name]))

    # ---- daily point-line chart (selected series only) --------------------
    long = trend.reset_index().melt("Day", var_name="Metric", value_name="Count")
    order = [m for m in TREND_COLOURS if m in selected]
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
    st.set_page_config(page_title="Ticket Dashboard · v9", layout="wide")
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
        render_welcome()
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

    # Snapshot KPI values (new_today, eappsys_cnt, retest_today, closed_today) are
    # computed above. The full overview (snapshot + filtered KPIs) is rendered as
    # ONE copyable widget after the filters are applied — see below.

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

    # ---- Overview KPIs (snapshot + filtered) as ONE copyable image -------
    if COL_PRIORITY in f.columns:
        p1_count = int((pd.to_numeric(f[COL_PRIORITY], errors="coerce") == 1).sum())
    else:
        p1_count = 0
    states_cnt = int(f[COL_STATE].nunique()) if COL_STATE in f.columns else 0

    snapshot_heading = (f"Daily snapshot for <b>{today:%a %d %b %Y}</b> "
                        f"(Done excluded everywhere except Closed today).")
    render_kpi_cards([
        {"label": "🆕 Created today",        "value": int(new_today),    "accent": "#2563eb",
         "help": "Tickets created on the reporting date."},
        {"label": "🏢 eAppsys tickets",      "value": int(eappsys_cnt),  "accent": "#7c3aed",
         "help": "Active tickets assigned to the eAppsys team. Done excluded."},
        {"label": "🔁 Set to Monitor today", "value": int(retest_today), "accent": "#f59e0b",
         "help": "Retest-state tickets last changed on the reporting date."},
        {"label": "✅ Closed today",          "value": int(closed_today), "accent": "#16a34a",
         "help": "Tickets moved to Done on the reporting date."},
        {"label": "Total tickets",           "value": int(len(f)),                  "accent": "#0ea5e9"},
        {"label": "Total assignees",         "value": int(f['Assignee'].nunique()), "accent": "#0891b2"},
        {"label": "States",                  "value": states_cnt,                   "accent": "#6366f1"},
        {"label": "Priority 1 tickets",      "value": int(p1_count),                "accent": "#dc2626",
         "help": "Tickets where Priority = 1 (highest)."},
    ], title="overview", heading=snapshot_heading)

    st.divider()

    # ---- Tabbed views (one per Excel pivot, plus the ticket-trends view) ----
    # Default range = the last 7 days up to the reporting date ("today").
    # Bounds are widened to also cover the file's own date span, so the user can
    # still scroll the picker back to older data.
    default_end   = today
    default_start = today - timedelta(days=6)         # 7 days inclusive
    data_min, data_max = data_date_span(df_full, created_col, changed_col, today)
    range_min = min(data_min, default_start)
    range_max = max(data_max, default_end)

    tab_range, tab_state, tab_iter, tab_assignee, tab_pivot, tab_data = st.tabs(
        ["📈 Ticket Trends", "By State", "By Iteration Path", "Assignee × State",
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
            disp = with_grand_total(t, "State")          # Grand Total appended
            left, right = st.columns([1, 1.4])
            with left:
                render_summary_table(disp, label_cols=["State"], title="By_State")  # desc by Count, total pinned
            right.bar_chart(t.set_index("State"))  # chart excludes the total row
        else:
            st.warning(f"No '{COL_STATE}' column found.")

    with tab_iter:
        if COL_ITERATION in f.columns:
            t = count_table(f, COL_ITERATION, "Iteration Path")
            disp = with_grand_total(t, "Iteration Path")  # Grand Total appended
            left, right = st.columns([1, 1.4])
            with left:
                render_summary_table(disp, label_cols=["Iteration Path"], title="By_Iteration_Path")  # desc by Count, total pinned
            right.bar_chart(t.set_index("Iteration Path"))  # chart excludes the total row
        else:
            st.warning(f"No '{COL_ITERATION}' column found.")

    with tab_assignee:
        st.subheader("Assignee × State")
        if COL_STATE in f.columns:
            piv = build_pivot(f, index="Assignee", columns=COL_STATE)
            # Busiest assignee first (desc by their Grand Total); total row pinned.
            render_summary_table(piv, label_cols=["Assignee"], title="Assignee_x_State")

            # ---- Two supporting charts (same filtered data as the table) -----
            st.markdown("##### Visual summary")
            g1, g2 = st.columns(2)
            with g1:
                count_bar(f, "Assignee", "Tickets per Assignee", "#2563eb")
            with g2:
                count_bar(f, COL_STATE, "Tickets per State", "#16a34a")
        else:
            st.warning(f"No '{COL_STATE}' column found.")

    with tab_pivot:
        st.subheader("Actual Priority × Object, counted per Assignee")
        piv = build_pivot(f, index=["Actual Priority", "Objects"], columns="Assignee")
        # Keep the meaningful priority order (1..10, then unassigned); pin total.
        render_summary_table(piv, label_cols=["Actual Priority", "Objects"],
                             sort_by=False, title="Priority_x_Assignee")

    with tab_data:
        st.subheader("Enriched data")
        # 'Assigned To' now holds the clean name, so drop the duplicate helper column
        disp = f.drop(columns=["Assignee"]) if "Assignee" in f.columns else f
        csv_bytes = disp.to_csv(index=False).encode("utf-8")   # clean export (no blanks)

        # On-screen only: one blank spacer column + one blank trailing row (framing).
        view = disp.copy()
        view[" "] = ""
        view = pd.concat([view, pd.DataFrame([{c: "" for c in view.columns}])],
                         ignore_index=True)
        show_table(view, fill_width=False, max_height=600)
        st.download_button(
            "⬇️ Download enriched CSV",
            data=csv_bytes,
            file_name="tickets_enriched.csv",
            mime="text/csv",
        )


if __name__ == "__main__":
    main()
