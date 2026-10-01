import io
import os
import re
import sys
from datetime import datetime

# Vercel's Python runtime imports this file directly via importlib without
# adding its own directory to sys.path, so sibling modules (db.py,
# data_utils.py) can't be found by a bare `import db` unless we add it
# ourselves first.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from dotenv import load_dotenv
from flask import Flask, render_template, request, jsonify, send_from_directory, send_file, g, session
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from werkzeug.exceptions import RequestEntityTooLarge

load_dotenv()
load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".env.local"), override=True)

import db
import mysupport_sync
from data_utils import (
    COLORS, PRIORITY_COLORS, AGEING_COLORS,
    parse_ticket_sheet, parse_project_sheet, parse_client_sheet, parse_milestone_sheet,
    parse_task_detail_sheet, detect_ticket_sheets,
)

app = Flask(
    __name__,
    template_folder=os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "templates"),
    static_folder=None,  # we serve /static/<file> ourselves below, from the project root
)

MAX_UPLOAD_MB = 25
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

# Needed to sign the session cookie. A fixed fallback (rather than a
# randomly generated one) matters here because Vercel's Python runtime can
# spin up a fresh process per request, and a random secret would silently
# invalidate every logged-in session on the next cold start.
app.secret_key = os.environ.get("SECRET_KEY", "sw-dashboard-session-signing-key-change-me")


@app.errorhandler(RequestEntityTooLarge)
def handle_upload_too_large(e):
    """Return JSON (not Werkzeug's default HTML error page) when a request
    body exceeds MAX_CONTENT_LENGTH, so the upload UI can show a readable
    message instead of a fetch failure or a JSON parse error."""
    return jsonify({"success": False,
                    "error": f"File(s) exceed the {MAX_UPLOAD_MB} MB upload limit"}), 413

PROJECT_ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")

# The Client sheet's row for KUIPS's Warranty project is filed under
# "UNISIRAJ" (its Projek ID/Name both spell out KUIPS, e.g.
# "KUIPSSAGA007"/"SAGA KUIPS" -- same entity, inconsistent naming between
# sheets), while its tickets use Company/Client="KUIPS". Used wherever a
# Home page client name needs to be translated into the name its ticket
# data actually uses.
CLIENT_DISPLAY_ALIASES = {"UNISIRAJ": "KUIPS"}

# Vercel sets this automatically on every deploy -- using it as the
# service worker's cache-name/version means sw.js's bytes (and therefore
# its cache) change on every deploy without anyone having to remember to
# bump a version number by hand. Locally (no Vercel env) it falls back to
# "dev", which is fine since local restarts don't need cache-busting.
SW_VERSION = os.environ.get("VERCEL_GIT_COMMIT_SHA", "dev")[:12]


@app.route("/sw.svg")
def brand_watermark():
    return send_from_directory(PROJECT_ROOT, "sw.svg", mimetype="image/svg+xml")


@app.route("/manifest.json")
def pwa_manifest():
    return send_from_directory(PROJECT_ROOT, "manifest.json", mimetype="application/manifest+json")


@app.route("/sw.js")
def pwa_service_worker():
    # Served from the root path (not /static/sw.js) so its default scope
    # covers the whole origin instead of just /static/. Templated (not
    # send_from_directory) so __SW_VERSION__ can be swapped for the
    # current deploy's commit SHA -- see SW_VERSION above.
    with open(os.path.join(PROJECT_ROOT, "sw.js"), "r", encoding="utf-8") as f:
        content = f.read().replace("__SW_VERSION__", SW_VERSION)
    return content, 200, {"Content-Type": "application/javascript", "Cache-Control": "no-cache"}


@app.route("/static/<path:filename>")
def static_files(filename):
    return send_from_directory(os.path.join(PROJECT_ROOT, "static"), filename)


def log(msg, level="INFO"):
    print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {level} {msg}", flush=True)


log("=" * 50)
log("Dashboard starting (Flask + Neon Postgres)")
log(f"Python: {sys.version}")

pio.templates.default = "plotly_white"
# Every chart in this app uses template="plotly_white" explicitly, so
# pinning the hover box style here fixes it everywhere at once. Without
# this, some browsers' "force dark mode for websites" auto-darkening
# (this site never declares a color-scheme, so it's a candidate for
# that heuristic) can darken the hover box background while Plotly's
# own inline SVG text fill stays dark too, leaving dark text on a dark
# box. Pinning both explicitly guarantees readable contrast regardless.
pio.templates["plotly_white"].layout.hoverlabel = dict(
    bgcolor="white", bordercolor="#d0d5dd", font=dict(color="#1f2937", size=12),
)

TICKET_DB_COL_BY_DISPLAY = {display: col for display, col in db.TICKET_DB_COLUMNS}
CLIENT_DB_COL_BY_DISPLAY = {display: col for display, col in db.CLIENT_DB_COLUMNS}
PROJECT_DB_COL_BY_DISPLAY = {display: col for display, col in db.PROJECT_DB_COLUMNS}
MILESTONE_DB_COL_BY_DISPLAY = {display: col for display, col in db.PROJECT_MILESTONE_DB_COLUMNS}
TASK_DETAIL_DB_COL_BY_DISPLAY = {display: col for display, col in db.TASK_DETAIL_DB_COLUMNS}

# Source workbooks live next to this module (api/*.xlsx), while templates and
# static assets live one level up (PROJECT_ROOT, defined below). Resolving a
# workbook name against both keeps source_file -- which the upload stores as a
# bare filename -- usable no matter which directory the app was started from.
SOURCE_WORKBOOK_DIRS = (
    os.path.dirname(os.path.abspath(__file__)),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."),
)

_schema_ready = False


def request_conn():
    """One psycopg2 connection per Flask request, reused by every db.*
    call in that request instead of each opening its own. A fresh Neon
    connect costs real round-trip time, and a single page load needs
    4+ separate queries, so this is what actually made pages fast --
    the indexes only help once the connection overhead isn't dominating.
    """
    if "db_conn" not in g:
        g.db_conn = db.get_conn()
    return g.db_conn


@app.teardown_appcontext
def close_request_conn(exception):
    conn = g.pop("db_conn", None)
    if conn is None:
        return
    try:
        if exception is None:
            conn.commit()
        else:
            conn.rollback()
    except Exception:
        pass
    finally:
        conn.close()


def ensure_schema():
    global _schema_ready
    if _schema_ready:
        return
    db.init_schema(conn=request_conn())
    _schema_ready = True


def lookup_client_for_project(projek_name, parsed_client_df=None):
    """Which client owns this project, for tagging a '<project> TASK DETAIL'
    sheet (which carries no Client column of its own).

    Prefers the Client sheet of the workbook being uploaded -- it is already
    parsed at that point and is the newest statement of the mapping -- and
    falls back to the clients table so the sheet still lands when the
    workbook omits a Client tab or the project was registered earlier.
    """
    if parsed_client_df is not None and not parsed_client_df.empty and "Projek Name" in parsed_client_df.columns:
        names = parsed_client_df["Projek Name"].astype(str).str.strip()
        hit = parsed_client_df.loc[names == projek_name, "Client"]
        if len(hit):
            value = str(hit.iloc[0]).strip()
            if value and value.lower() not in ("nan", "none"):
                return value
    try:
        value = db.find_client_for_project(projek_name, conn=request_conn())
        if value:
            return value
    except Exception as e:
        log(f"Client lookup failed for {projek_name}: {e}", "ERROR")
    return None


def resolve_source_workbook(source_file):
    """Absolute path of the workbook a source_file value refers to, or None.

    source_file is stored as whatever the upload was named, so a rename
    between uploads would make the path unresolvable -- the Excel buttons
    report that rather than guessing a different file to overwrite.
    """
    if not source_file:
        return None
    candidate = os.path.normpath(source_file)
    if os.path.isabs(candidate):
        return candidate if os.path.isfile(candidate) else None
    for base in SOURCE_WORKBOOK_DIRS:
        path = os.path.normpath(os.path.join(base, candidate))
        if os.path.isfile(path):
            return path
    return None


def parse_filters(args):
    filters = {
        "clients": args.getlist("client"),
        "priorities": args.getlist("priority"),
        "statuses": args.getlist("status"),
        "task_types": args.getlist("task_type"),
        "search": args.get("search") or None,
    }
    for key, param in (("date_start", "date_start"), ("date_end", "date_end")):
        raw = args.get(param)
        if raw:
            try:
                datetime.strptime(raw, "%Y-%m-%d")
                filters[key] = raw
            except ValueError:
                pass
    filters["projek_name"] = args.get("projek_name") or None
    return filters


def _narrow_by_projek_name(df, filters):
    """Best-effort narrowing of tickets down to the specific Projek Name
    clicked in Overall Client (e.g. "MYOT MARA" vs "MYCLAIM MARA" under the
    same client) via the tickets' own Project column. Real data shows this
    only lines up for some clients (Project holds a short code like
    "MYCLAIM" that's a substring of the Projek Name once the client's own
    name is stripped off) and not others (e.g. a Projek Name that doesn't
    even share a client-name convention with its tickets) -- so on no
    match this deliberately falls back to the unnarrowed df rather than
    showing an empty page for a client whose data just doesn't follow the
    pattern.
    """
    projek_name = (filters or {}).get("projek_name")
    if not projek_name or df.empty or "Project" not in df.columns:
        return df
    core = projek_name
    clients = (filters or {}).get("clients") or []
    if len(clients) == 1:
        client = re.escape(clients[0])
        core = re.sub(rf"(^\s*{client}\s+|\s+{client}\s*$)", "", projek_name, flags=re.IGNORECASE).strip()
    if not core:
        return df
    narrowed = df[df["Project"].astype(str).str.contains(re.escape(core), case=False, na=False)]
    return narrowed if not narrowed.empty else df


def load_data(filters=None):
    ensure_schema()
    df = db.fetch_tickets_df(filters, conn=request_conn())
    df = _narrow_by_projek_name(df, filters)
    return df, []


def load_project_data():
    ensure_schema()
    return db.fetch_projects_df(conn=request_conn())


def load_client_data():
    ensure_schema()
    return db.fetch_clients_df(conn=request_conn())


def load_milestone_data():
    ensure_schema()
    return db.fetch_project_milestone_df(conn=request_conn())


def build_warranty_charts(df):
    charts = {}
    warranty_df = df[df["Client"] == "Client Warranty"].copy() if "Client" in df.columns else pd.DataFrame()
    if warranty_df.empty:
        return charts

    total = len(warranty_df)
    completed = len(warranty_df[warranty_df["Ticket Status"].isin(["Completed", "Closed"])]) if "Ticket Status" in warranty_df.columns else 0
    pending = len(warranty_df[warranty_df["Ticket Status"] == "Pending"]) if "Ticket Status" in warranty_df.columns else 0
    in_progress = len(warranty_df[warranty_df["Ticket Status"] == "In Progress"]) if "Ticket Status" in warranty_df.columns else 0
    sla_breach = warranty_df["SLA Breach"].sum() if "SLA Breach" in warranty_df.columns else 0

    charts["metrics"] = {
        "total": total, "completed": completed, "pending": pending,
        "in_progress": in_progress, "sla_breach": int(sla_breach),
        "completed_pct": f"{completed / total * 100:.1f}%" if total > 0 else "0%",
        "pending_pct": f"{pending / total * 100:.1f}%" if total > 0 else "0%",
        "in_progress_pct": f"{in_progress / total * 100:.1f}%" if total > 0 else "0%",
    }

    if "Ticket Status" in warranty_df.columns:
        sc = warranty_df["Ticket Status"].value_counts().reset_index()
        sc.columns = ["Status", "Count"]
        fig = px.pie(sc, names="Status", values="Count", title="Warranty Ticket Status",
                      color="Status", color_discrete_map=COLORS, hole=0.3)
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), legend=dict(orientation="h", yanchor="bottom", y=-0.2))
        charts["status_pie"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Task Type" in warranty_df.columns:
        tc = warranty_df["Task Type"].value_counts().reset_index()
        tc.columns = ["Task Type", "Count"]
        fig = px.bar(tc, x="Task Type", y="Count", title="Warranty Tickets by Task Type",
                      color="Task Type", text="Count")
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), showlegend=False, xaxis_tickangle=-45)
        charts["task_type_bar"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Project" in warranty_df.columns:
        pc = warranty_df["Project"].value_counts().reset_index()
        pc.columns = ["Project", "Count"]
        fig = px.bar(pc, x="Project", y="Count", title="Warranty Tickets by Project",
                      color="Project", text="Count")
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), showlegend=False, xaxis_tickangle=-45)
        charts["project_bar"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    display_cols = ["Ticket No", "Task Type", "Project", "Company", "Ticket Title", "Priority", "Ticket Status", "Ticket Created Date", "Days", "Ageing"]
    avail = [c for c in display_cols if c in warranty_df.columns]
    meta_cols = [c for c in ["_row_idx", "Source File"] if c in warranty_df.columns]
    detail = warranty_df[avail + meta_cols].copy()
    if "Ticket Created Date" in detail.columns:
        detail["Ticket Created Date"] = detail["Ticket Created Date"].dt.strftime("%d/%m/%Y")
    charts["detail_data"] = detail.to_dict("records")

    return charts


def recompute_overall_progress(df):
    """Overall Progress Task (%) is meant to be each title's own project
    completion (all its numbered subtasks averaged together), but the
    value that comes in from the source spreadsheet is whatever number was
    last typed in by hand -- e.g. an average taken before later subtasks
    were even added to the sheet, or entirely missing. Recompute it fresh
    here instead of trusting that stale figure.

    A title's real subtasks are its rows numbered "1.", "2." etc in
    Description; the "- " checklist bullets underneath a subtask (e.g. a
    breakdown of "6. Pre UAT:") aren't separate subtasks and would double
    count that one subtask's percentage if averaged in directly.
    """
    if df.empty or not {"Client", "Title", "Percentage"}.issubset(df.columns):
        return df

    df = df.copy()
    pct = pd.to_numeric(df["Percentage"], errors="coerce")
    is_subtask = (
        df["Description"].astype(str).str.match(r"^\s*\d+\.")
        if "Description" in df.columns else pd.Series(True, index=df.index)
    )

    def group_overall(idx):
        group_pct = pct.loc[idx]
        subtask_pct = group_pct[is_subtask.loc[idx]]
        basis = subtask_pct if not subtask_pct.empty else group_pct
        basis = basis.dropna()
        return round(basis.mean(), 1) if not basis.empty else None

    # A plain dict (not .groupby().apply()) sidesteps a pandas edge case:
    # when every group's computed value is None -- as happens for a module
    # with no Percentage data at all, e.g. YIPS's "Outstanding Development"
    # rows -- .apply() can't decide whether the combined result is a Series
    # or an empty DataFrame and raises "Data must be 1-dimensional, got
    # ndarray of shape (0, 0)" instead of just returning all-None.
    overall_by_group = {key: group_overall(g.index) for key, g in df.groupby(["Client", "Title"])}
    df["Overall Progress Task (%)"] = df.set_index(["Client", "Title"]).index.map(overall_by_group).to_numpy()
    return df


def recompute_status_from_percentage(df):
    """Status Progress and Progress are meant to track each row's own
    Percentage (0 = Not Started, 1-99 = In Progress, 100 = Completed), but
    rows saved before that rule existed -- or edited directly in the
    source spreadsheet -- can carry a stale label that no longer matches
    the number. Recompute both label columns from Percentage on every
    load so they (and the metrics/pie chart that count them) never drift
    out of sync with it. Rows with a blank/non-numeric Percentage are left
    untouched since there's nothing to derive a label from.
    """
    if df.empty or "Percentage" not in df.columns:
        return df

    df = df.copy()
    pct = pd.to_numeric(df["Percentage"], errors="coerce")
    has_pct = pct.notna()

    def label(n):
        if n <= 0:
            return "Not Started"
        if n >= 100:
            return "Completed"
        return "In Progress"

    status = pct.apply(lambda n: label(n) if pd.notna(n) else None)
    for col in ("Status Progress", "Progress"):
        if col in df.columns:
            df.loc[has_pct, col] = status.loc[has_pct]
    return df


def recompute_duration(df):
    """Duration is derived, not typed -- Plan Start Date through Plan End
    Date, inclusive of both ends, with Saturdays not counted (so a 7-day
    calendar week is 6 days of duration). Recomputed on every load so
    editing either date always keeps Duration in sync, the same way
    Status Progress stays in sync with Percentage.
    """
    if df.empty or not {"Plan Start Date", "Plan End Date"}.issubset(df.columns):
        return df

    df = df.copy()

    def duration_for(row):
        start, end = row["Plan Start Date"], row["Plan End Date"]
        if pd.isna(start) or pd.isna(end) or end < start:
            return None
        days = (end - start).days + 1
        saturdays = sum(1 for i in range(days) if (start + pd.Timedelta(days=i)).weekday() == 5)
        return f"{days - saturdays} days"

    df["Duration"] = df.apply(duration_for, axis=1)
    return df


def resolve_project_scope(project_df, filters):
    """Narrow the full projects table down to whatever (Client, Projek
    Name) the current filters ask for. Shared by the Project tab and the
    PDF report endpoint so both scope to a project the exact same way.

    The projects table is independent of the tickets table (no shared
    filter query), so a client selected via the sidebar/Overall Client
    table has to be applied here explicitly.
    """
    if filters.get("clients") and "Client" in project_df.columns:
        project_df = project_df[project_df["Client"].isin(filters["clients"])]
    # Unlike tickets (no real Projek Name field, only a best-effort guess
    # against Project), projects rows carry Projek Name directly, so this
    # is an exact match -- still falls back to the unnarrowed set if it
    # matches nothing, same safety rule as the ticket side.
    projek_name = filters.get("projek_name")
    if projek_name and "Projek Name" in project_df.columns and not project_df.empty:
        narrowed = project_df[project_df["Projek Name"] == projek_name]
        if not narrowed.empty:
            project_df = narrowed
        elif project_df["Projek Name"].isna().all():
            # Some clients' Client Project sheet rows never carry their
            # own Projek Name at all (e.g. LKTN) -- there's nothing to
            # disambiguate against, so it's safe to label every row with
            # the project that was actually clicked instead of leaving
            # the column blank. If the client's rows DO carry a
            # (different, non-matching) Projek Name elsewhere, this
            # branch is skipped and the unnarrowed set is shown instead,
            # same safety rule as the ticket side.
            project_df = project_df.copy()
            project_df["Projek Name"] = projek_name
    return project_df


def build_project_charts(df):
    charts = {}
    if df.empty:
        return charts

    total = len(df)
    completed = len(df[df["Status Progress"].str.lower().str.contains("completed", na=False)]) if "Status Progress" in df.columns else 0
    in_progress = len(df[df["Status Progress"].str.lower().str.contains("progress", na=False)]) if "Status Progress" in df.columns else 0
    not_started = len(df[df["Status Progress"].str.lower().str.contains("not started", na=False)]) if "Status Progress" in df.columns else 0

    charts["metrics"] = {
        "total": total, "completed": completed,
        "in_progress": in_progress, "not_started": not_started,
    }

    # "Projects by Client" collapses to a single, useless one-bar chart
    # whenever the sidebar has already scoped the page down to one client
    # (the common case -- it's identical to "all projects" only when every
    # client is being viewed at once). Show progress per module (Title)
    # instead: that's informative both scoped to one client and across all
    # of them, and doubles as an at-a-glance progress readout.
    if "Title" in df.columns and "Overall Progress Task (%)" in df.columns:
        module_df = df.dropna(subset=["Title"])
        module_df = module_df[module_df["Title"].astype(str).str.strip() != ""]
        if not module_df.empty:
            group_cols = ["Client", "Title"] if "Client" in module_df.columns else ["Title"]
            # keep="first" on the table's own row order (not a groupby,
            # which would alphabetize) so the chart's bar order matches
            # the Project Details table below it, exactly like the
            # table's own row order is never re-sorted either.
            modules = module_df.drop_duplicates(subset=group_cols, keep="first")[group_cols + ["Overall Progress Task (%)"]]
            modules = modules.dropna(subset=["Overall Progress Task (%)"])
            if not modules.empty:
                # Color per module (Title), not per client -- with client
                # as the color, long module names on the x-axis and a
                # handful of client swatches in the legend crowded right
                # up against each other and read as clutter, not a legend.
                # Each bar already has its own module name on the x-axis,
                # so a per-module legend would just repeat that -- drop it.
                fig = px.bar(
                    modules, x="Title", y="Overall Progress Task (%)",
                    title="Module Progress", text="Overall Progress Task (%)",
                    color="Title",
                    category_orders={"Title": modules["Title"].tolist()},
                )
                fig.update_traces(texttemplate="%{text:.0f}%", textposition="outside")
                fig.update_layout(
                    template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#374151"), xaxis_tickangle=-45, xaxis_title="Module",
                    yaxis_title="Overall Progress (%)", yaxis_range=[0, 110],
                    showlegend=False,
                )
                charts["client_bar"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Status Progress" in df.columns:
        valid_status = df.dropna(subset=["Status Progress"])
        if not valid_status.empty:
            sc = valid_status["Status Progress"].value_counts().reset_index()
            sc.columns = ["Status", "Count"]
            fig = px.pie(sc, names="Status", values="Count", title="Project Status Progress",
                          hole=0.3)
            fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
            charts["status_pie"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Target Start Date" in df.columns and "Target End Date" in df.columns and "Title" in df.columns:
        valid = df.dropna(subset=["Target Start Date", "Target End Date", "Title"]).copy()
        valid = valid[valid["Title"].astype(str).str.strip() != ""]
        if "Description" in valid.columns:
            valid["Task Label"] = valid["Description"].astype(str)
            valid["Task Label"] = valid["Task Label"].str.replace(r"^\d+\.\s*", "", regex=True)
            valid["Task Label"] = valid["Task Label"].str.split("\n").str[0].str.strip()
        else:
            valid["Task Label"] = valid["Title"].astype(str)
        # A client can have more than one project (e.g. MYCLAIM MARA and
        # MYOT MARA both under MARA), and both commonly reuse the exact
        # same task names (Development/UAT/Go Live/...). Row identity by
        # Task Label alone collapsed those onto the same y-axis row, so
        # whichever bar happened to overlap in time visually covered the
        # other one up entirely -- e.g. MyOT's "Go Live" bar hid MyClaim's
        # own Development/UAT/Go Live bars underneath it. Scope each row to
        # its own project (Projek Name, falling back to Title) plus task
        # name so same-named tasks from different projects always get
        # their own row and are never drawn on top of each other.
        project_key = valid["Projek Name"].where(valid["Projek Name"].astype(str).str.strip().ne(""), valid["Title"]) if "Projek Name" in valid.columns else valid["Title"]
        valid["Row Label"] = project_key.astype(str) + ": " + valid["Task Label"]
        if not valid.empty:
            timeline_charts_html = ""
            if "Client" in valid.columns:
                for client in sorted(valid["Client"].dropna().unique()):
                    cdf_client = valid[valid["Client"] == client]
                    if cdf_client.empty:
                        continue
                    # Broken down one Gantt chart per module (Title) instead
                    # of one giant chart mixing every module's tasks
                    # together -- each module renders as its own collapsed
                    # card (see .gantt-title-card / <details> below) so a
                    # client with several modules (e.g. LKTN's Payroll,
                    # Claim, Asset, ...) isn't one overwhelming wall of
                    # bars; the chart for a module only needs to render
                    # once its card is actually opened. Order preserved
                    # (not sorted) so cards appear in the same order as the
                    # Project Details table's modules.
                    title_cards_html = ""
                    for title in cdf_client["Title"].drop_duplicates():
                        cdf = cdf_client[cdf_client["Title"] == title].copy()
                        if cdf.empty:
                            continue
                        # Row Label (Project: Task) usually gives each task
                        # its own row, but a project can legitimately have
                        # two rows with the exact same description (a
                        # generic recurring checklist item like "Sign-off",
                        # or two "UAT" rounds) -- a *shared string* y-axis
                        # category collapses those onto one line no matter
                        # how the string is built. Give every row its own
                        # guaranteed-unique numeric position instead (one
                        # row of cdf = one position, always, by
                        # construction) and only use Row Label as the tick
                        # text shown at that position -- so two rows with
                        # identical text still each get their own line.
                        #
                        # Preserve cdf's incoming order rather than
                        # re-sorting it (previously by Row Label/Start
                        # date) so the chart's row order matches the
                        # Project Details table's row order -- both
                        # ultimately come from the same project_df, fetched
                        # `ORDER BY id`, so as long as neither re-sorts
                        # they stay in the same sequence.
                        cdf = cdf.reset_index(drop=True)
                        cdf["Y Pos"] = cdf.index
                        # Coloring by Client here was a no-op -- every row
                        # in cdf already shares the same Client, so every
                        # bar came out one uniform color. Color by Category
                        # instead so different kinds of work are visually
                        # distinguishable; but if this module's tasks are
                        # all the same Category too (equally uniform,
                        # equally uninformative), color by the task itself
                        # (Task Label, from Description) so each bar in the
                        # timeline still reads as distinct.
                        categories = cdf["Category"].dropna().unique() if "Category" in cdf.columns else []
                        color_col = "Category" if len(categories) > 1 else "Task Label"
                        color_values = sorted(cdf[color_col].dropna().astype(str).unique().tolist())
                        palette = px.colors.qualitative.Plotly
                        color_map = {val: palette[i % len(palette)] for i, val in enumerate(color_values)}

                        # A milestone (Start date == Due date, e.g. "Go
                        # Live") and a real 1-day task (Due date = Start
                        # date + 1) are both, in plain terms, "this
                        # happened on one day" -- they used to render
                        # completely differently (a diamond marker vs. a
                        # bar), which is the inconsistency being fixed
                        # here. Both now render as the exact same bar: a
                        # milestone's Due date is treated as Start date + 1
                        # day purely for the chart (real Due date/Duration
                        # elsewhere untouched), so every ~1-day task gets
                        # identical, standardized sizing regardless of
                        # which way it happened to be recorded.
                        chart_due = cdf["Target End Date"].where(cdf["Target End Date"] != cdf["Target Start Date"], cdf["Target Start Date"] + pd.Timedelta(days=1))
                        cdf = cdf.assign(**{"Chart Due": chart_due})

                        # A short bar does have an actual duration, so it's
                        # fair to give it a minimum visible length on the
                        # date axis, scaled to this module's own chart span
                        # -- this doesn't invent a date range that never
                        # existed, it just guarantees a short-but-real one
                        # doesn't round down to invisible. Capped at 2
                        # days: uncapped, this scaled with the *whole
                        # chart's* span (which can be dominated by an
                        # unrelated multi-year task elsewhere in the same
                        # module), so on a wide enough chart a boosted
                        # 1-day task could stretch past a genuine,
                        # unboosted 3-4 day task and visually look longer
                        # than something that actually took more real time
                        # -- capping at 2 days keeps it strictly below the
                        # >2-day tier that never gets boosted, so relative
                        # ordering is never inverted by the correction
                        # meant to just aid visibility.
                        span_days = max((cdf["Chart Due"].max() - cdf["Target Start Date"].min()).days, 1)
                        min_bar_ms = min(max(1, round(span_days * 0.015)), 2) * 86400000

                        fig = go.Figure()
                        for val in color_values:
                            rdf = cdf[cdf[color_col].astype(str) == val]
                            if rdf.empty:
                                continue
                            bar_days = (rdf["Chart Due"] - rdf["Target Start Date"]).dt.days
                            # Thicker vertically (row height) AND given a
                            # visible minimum horizontal length -- short
                            # tasks need to stand out in both directions,
                            # not just one, and every ~1-day task
                            # (bar_days <= 1) gets the exact same
                            # standardized thickness/width.
                            bar_widths = bar_days.apply(lambda d: 0.9 if d <= 1 else (0.85 if d <= 2 else 0.7))
                            bar_ms = (rdf["Chart Due"] - rdf["Target Start Date"]).dt.total_seconds() * 1000
                            bar_ms = bar_ms.where(bar_days > 2, bar_ms.clip(lower=min_bar_ms))
                            fig.add_trace(go.Bar(
                                base=rdf["Target Start Date"],
                                # A raw pandas Timedelta isn't
                                # JSON-serializable in every Plotly version
                                # -- milliseconds (a plain float) is how
                                # Plotly represents a bar's width on a date
                                # axis internally either way.
                                x=bar_ms,
                                y=rdf["Y Pos"], orientation="h", width=bar_widths.tolist(),
                                name=val, legendgroup=val, marker_color=color_map[val],
                                customdata=rdf[["Row Label", "Target Start Date", "Target End Date"]].astype(str),
                                hovertemplate="Row=%{customdata[0]}<br>Start=%{customdata[1]}<br>Due=%{customdata[2]}<extra></extra>",
                            ))

                        # tickvals/ticktext (not a categorical axis) is
                        # what lets the same description repeat as text on
                        # two different rows without Plotly merging them
                        # back down to one category. margin (row spacing)
                        # comes from the bar width values above (0.7/0.9),
                        # leaving 10-30% of each row's slot empty.
                        fig.update_yaxes(
                            autorange="reversed", title=None,
                            tickmode="array", tickvals=cdf["Y Pos"], ticktext=cdf["Task Label"],
                            tickfont=dict(size=14),
                        )
                        fig.update_xaxes(title="Tarikh", type="date")
                        fig.update_layout(
                            template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                            font=dict(color="#374151"),
                            # Y-axis already shows the plain task name and
                            # hovering shows the full project+task+dates,
                            # so the color-key legend is redundant screen
                            # space.
                            showlegend=False,
                            height=max(420, 95*len(cdf)),
                            margin=dict(l=180),
                            font_size=15,
                        )
                        chart_html = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False, "responsive": True})
                        title_label = title if str(title).strip() else "(Untitled)"
                        title_cards_html += (
                            f'<details class="gantt-title-card"><summary>{title_label} ({len(cdf)} tasks)</summary>'
                            f'<div class="gantt-title-card-body">{chart_html}</div></details>'
                        )
                    if title_cards_html:
                        timeline_charts_html += f'<div class="client-section"><h4>{client}</h4>{title_cards_html}</div>'
            charts["timeline_chart"] = timeline_charts_html

    display_cols_p = [
        "Client", "Title", "Projek Name", "Description", "Category", "Progress", "Priority",
        "Actual Start Date", "Actual End Date", "Plan Start Date", "Plan End Date",
        "Target Start Date", "Target End Date", "Duration", "Assigned to",
        "Status Progress", "Percentage", "Overall Progress Task (%)",
    ]
    avail_p = [c for c in display_cols_p if c in df.columns]
    meta_p = [c for c in ["_row_idx", "_source_file"] if c in df.columns]
    detail = df[avail_p + meta_p].copy()
    for c in ["Plan Start Date", "Plan End Date", "Target Start Date", "Target End Date", "Actual Start Date", "Actual End Date"]:
        if c in detail.columns:
            detail[c] = detail[c].dt.strftime("%d/%m/%Y") if detail[c].notna().any() else ""
    detail = detail.fillna("")
    detail_records = detail.to_dict("records")

    # Overall Progress Task (%) is one number per Title, not one per row --
    # mark, for each contiguous run of rows sharing the same (Client,
    # Title), how many rows the first one's cell should visually span, so
    # the template can render it as a single merged cell (like the source
    # spreadsheet) instead of repeating the same figure down every row.
    if "Overall Progress Task (%)" in detail.columns and "Title" in detail.columns:
        i, n = 0, len(detail_records)
        while i < n:
            key = (detail_records[i].get("Client"), detail_records[i].get("Title"))
            j = i + 1
            while j < n and (detail_records[j].get("Client"), detail_records[j].get("Title")) == key:
                j += 1
            detail_records[i]["_op_rowspan"] = j - i
            for k in range(i + 1, j):
                detail_records[k]["_op_rowspan"] = 0
            i = j

    charts["detail_data"] = detail_records

    return charts


def build_project_report_data(client, projek_name):
    """Aggregate one project's data into the shape the client-side PDF
    report (pdfmake, see dashboard.html) renders. Everything here is
    derived straight from existing fields -- module/task rows, dates,
    Assigned to -- rather than any hand-written narrative, since the
    database has no field for that. Returns None if nothing matches.
    """
    project_df = load_project_data()
    project_df = resolve_project_scope(project_df, {"clients": [client] if client else [], "projek_name": projek_name})
    project_df = recompute_status_from_percentage(project_df)
    project_df = recompute_overall_progress(project_df)
    project_df = recompute_duration(project_df)
    if project_df.empty:
        return None

    resolved_projek_name = projek_name
    if not resolved_projek_name and "Projek Name" in project_df.columns:
        non_blank = project_df["Projek Name"].dropna()
        if not non_blank.empty:
            resolved_projek_name = non_blank.iloc[0]

    def fmt_date(v):
        if v is None or pd.isna(v):
            return None
        return v.strftime("%d/%m/%Y")

    # Contract-level info (Technology, Projek Status, contract dates) comes
    # from the Client sheet/table, a separate source from the task-level
    # Client Project data above.
    projek_status = technology = contract_start = contract_end = None
    try:
        client_df = load_client_data()
    except Exception:
        client_df = pd.DataFrame()
    if not client_df.empty and "Client" in client_df.columns:
        cdf = client_df[client_df["Client"] == client]
        if resolved_projek_name and "Projek Name" in cdf.columns:
            matched = cdf[cdf["Projek Name"] == resolved_projek_name]
            if not matched.empty:
                cdf = matched
        if not cdf.empty:
            row0 = cdf.iloc[0]
            projek_status = row0.get("Projek Status")
            technology = row0.get("Technology")
            contract_start = fmt_date(row0.get("Start Date"))
            contract_end = fmt_date(row0.get("End Date"))

    date_cols = ["Plan Start Date", "Plan End Date", "Target Start Date", "Target End Date", "Actual Start Date", "Actual End Date"]

    modules = []
    if "Title" in project_df.columns:
        for title in project_df["Title"].dropna().drop_duplicates():
            tdf = project_df[project_df["Title"] == title]
            overall = None
            if "Overall Progress Task (%)" in tdf.columns:
                ov = pd.to_numeric(tdf["Overall Progress Task (%)"], errors="coerce").dropna()
                if not ov.empty:
                    overall = round(float(ov.iloc[0]), 1)
            if overall is None and "Percentage" in tdf.columns:
                pct = pd.to_numeric(tdf["Percentage"], errors="coerce").dropna()
                if not pct.empty:
                    overall = round(float(pct.mean()), 1)
            categories = sorted(set(tdf["Category"].dropna().astype(str))) if "Category" in tdf.columns else []
            assignees = sorted(set(a.strip() for a in tdf["Assigned to"].dropna().astype(str) if a.strip())) if "Assigned to" in tdf.columns else []
            module = {
                "title": str(title),
                "task_count": int(len(tdf)),
                "overall_percent": overall if overall is not None else 0,
                "category": categories[0] if len(categories) == 1 else ("Mixed" if len(categories) > 1 else None),
                "assignees": assignees,
            }
            for c in date_cols:
                key = c.lower().replace(" ", "_")
                if c not in tdf.columns:
                    module[key] = None
                    continue
                is_start = c.endswith("Start Date")
                module[key] = fmt_date(tdf[c].min() if is_start else tdf[c].max())
            modules.append(module)

    overall_percent = round(sum(m["overall_percent"] for m in modules) / len(modules), 1) if modules else 0

    tasks = []
    for _, row in project_df.iterrows():
        desc = row.get("Description")
        desc = str(desc).split("\n")[0].strip()[:160] if pd.notna(desc) else ""
        pct = row.get("Percentage")
        tasks.append({
            "title": str(row.get("Title")) if pd.notna(row.get("Title")) else "",
            "description": desc,
            "category": row.get("Category") if pd.notna(row.get("Category")) else "",
            "priority": row.get("Priority") if pd.notna(row.get("Priority")) else "",
            "assigned_to": row.get("Assigned to") if pd.notna(row.get("Assigned to")) else "",
            "percentage": float(pct) if pd.notna(pct) else None,
            "status_progress": row.get("Status Progress") if pd.notna(row.get("Status Progress")) else "",
            "plan_end": fmt_date(row.get("Plan End Date")),
            "target_start": fmt_date(row.get("Target Start Date")),
            "target_end": fmt_date(row.get("Target End Date")),
            "actual_end": fmt_date(row.get("Actual End Date")),
        })

    team = {}
    if "Assigned to" in project_df.columns:
        for _, row in project_df.iterrows():
            name = row.get("Assigned to")
            name = str(name).strip() if pd.notna(name) and str(name).strip() else "Unassigned"
            entry = team.setdefault(name, {"modules": set(), "task_count": 0, "percentages": []})
            title = row.get("Title")
            if pd.notna(title) and str(title).strip():
                entry["modules"].add(str(title))
            entry["task_count"] += 1
            pct = row.get("Percentage")
            if pd.notna(pct):
                entry["percentages"].append(float(pct))
    team_list = [
        {
            "name": name,
            "modules": sorted(entry["modules"]),
            "task_count": entry["task_count"],
            "avg_percent": round(sum(entry["percentages"]) / len(entry["percentages"]), 1) if entry["percentages"] else None,
        }
        for name, entry in team.items()
    ]
    team_list.sort(key=lambda t: t["name"].lower())

    today = pd.Timestamp.now().normalize()
    overdue, upcoming = [], []
    has_plan_end = "Plan End Date" in project_df.columns
    has_target_end = "Target End Date" in project_df.columns
    if has_plan_end or has_target_end:
        for _, row in project_df.iterrows():
            plan_end = row.get("Plan End Date") if has_plan_end else None
            target_end = row.get("Target End Date") if has_target_end else None
            plan_end = None if pd.isna(plan_end) else plan_end
            target_end = None if pd.isna(target_end) else target_end
            if plan_end is None and target_end is None:
                continue
            pct = row.get("Percentage")
            is_done = pd.notna(pct) and float(pct) >= 100
            if is_done:
                continue
            entry = {
                "title": str(row.get("Title")) if pd.notna(row.get("Title")) else "",
                "description": str(row.get("Description")).split("\n")[0].strip()[:120] if pd.notna(row.get("Description")) else "",
                "plan_end": fmt_date(plan_end),
                "target_end": fmt_date(target_end),
                "assigned_to": row.get("Assigned to") if pd.notna(row.get("Assigned to")) else "",
            }
            # A task is overdue the moment either its Plan End or its Target
            # End has already passed -- Target End slipping past today is
            # just as much a red flag as Plan End slipping, even if Plan
            # End itself is still in the future.
            if (plan_end is not None and plan_end < today) or (target_end is not None and target_end < today):
                overdue.append(entry)
            else:
                soonest = min(d for d in (plan_end, target_end) if d is not None)
                if soonest <= today + pd.Timedelta(days=14):
                    upcoming.append(entry)

    low_progress_modules = [m["title"] for m in modules if m["overall_percent"] < 50]

    # Schedule slippage: how many days a task's Target End Date has moved
    # past its original Plan End Date. This is independent of whether the
    # task is overdue today -- a task can have already slipped from the
    # plan while its (revised) target is still comfortably in the future.
    slippage = []
    if has_plan_end and has_target_end:
        for _, row in project_df.iterrows():
            plan_end = row.get("Plan End Date")
            target_end = row.get("Target End Date")
            if pd.isna(plan_end) or pd.isna(target_end):
                continue
            days = (target_end - plan_end).days
            if days <= 0:
                continue
            slippage.append({
                "title": str(row.get("Title")) if pd.notna(row.get("Title")) else "",
                "description": str(row.get("Description")).split("\n")[0].strip()[:120] if pd.notna(row.get("Description")) else "",
                "plan_end": fmt_date(plan_end),
                "target_end": fmt_date(target_end),
                "slippage_days": int(days),
                "assigned_to": row.get("Assigned to") if pd.notna(row.get("Assigned to")) else "",
            })
    slippage.sort(key=lambda s: s["slippage_days"], reverse=True)

    # At-risk modules: below 50% complete AND already has at least one
    # overdue task -- behind schedule with little buffer left to recover,
    # as distinct from a module that is merely slow but not yet late.
    overdue_module_titles = {o["title"] for o in overdue}
    at_risk_modules = [m["title"] for m in modules if m["overall_percent"] < 50 and m["title"] in overdue_module_titles]

    return {
        "client": client,
        "projek_name": resolved_projek_name,
        "projek_status": projek_status,
        "technology": technology,
        "contract_start": contract_start,
        "contract_end": contract_end,
        "generated_at": pd.Timestamp.now().strftime("%d/%m/%Y %H:%M"),
        "overall_percent": overall_percent,
        "modules": modules,
        "tasks": tasks,
        "team": team_list,
        "attention": {
            "overdue": overdue,
            "upcoming": upcoming,
            "low_progress_modules": low_progress_modules,
            "slippage": slippage,
            "at_risk_modules": at_risk_modules,
        },
    }


def build_ticket_report_data(client, category):
    """Aggregate one client's Warranty or Maintenance tickets into the shape
    the client-side PDF/PPTX report (pdfmake/pptxgenjs, see dashboard.html)
    renders -- the ticket-side counterpart to build_project_report_data.
    Returns None if nothing matches.

    category is "Warranty" or "Maintenance". Warranty tickets are stored
    under a shared "Client Warranty" sentinel Client value with the real
    client only recorded in Company (see build_tab_context's idx == 2
    branch, which this mirrors), so it needs its own lookup path instead of
    the plain clients/task_types filter Maintenance uses.
    """
    lookup_client = CLIENT_DISPLAY_ALIASES.get(client, client)
    if category == "Warranty":
        try:
            df, _ = load_data({"clients": ["Client Warranty"], "priorities": [], "statuses": [], "task_types": [], "search": None})
        except Exception as e:
            log(f"DB error loading warranty tickets: {e}", "ERROR")
            df = pd.DataFrame()
        if "Company" in df.columns:
            df = df[df["Company"] == lookup_client]
    else:
        try:
            df, _ = load_data({"clients": [lookup_client], "priorities": [], "statuses": [], "task_types": [category], "search": None})
        except Exception as e:
            log(f"DB error loading {category} tickets: {e}", "ERROR")
            df = pd.DataFrame()
    if df.empty:
        return None

    def fmt_date(v):
        if v is None or pd.isna(v):
            return None
        return v.strftime("%d/%m/%Y")

    total = len(df)
    statuses = df["Ticket Status"] if "Ticket Status" in df.columns else pd.Series(dtype=object)
    completed = int(statuses.isin(["Completed", "Closed"]).sum())
    pending = int((statuses == "Pending").sum())
    in_progress = int((statuses == "In Progress").sum())
    sla_breach = int(df["SLA Breach"].sum()) if "SLA Breach" in df.columns else 0

    metrics = {
        "total": total, "completed": completed, "pending": pending,
        "in_progress": in_progress, "sla_breach": sla_breach,
        "completed_pct": round(completed / total * 100, 1) if total else 0,
        "pending_pct": round(pending / total * 100, 1) if total else 0,
        "in_progress_pct": round(in_progress / total * 100, 1) if total else 0,
    }

    status_counts = []
    if "Ticket Status" in df.columns:
        sc = df["Ticket Status"].dropna().value_counts()
        status_counts = [{"status": k, "count": int(v)} for k, v in sc.items()]

    priority_counts = []
    if "Priority" in df.columns:
        pc = df["Priority"].dropna().value_counts()
        priority_counts = [{"priority": k, "count": int(v)} for k, v in pc.items()]

    # Per-project breakdown -- a client can run tickets against more than
    # one project (e.g. LKTN's Payroll/Claim/Asset modules each raise their
    # own tickets), so a flat client-wide total hides which project is
    # actually driving the ticket load. Sorted by open-ticket count (same
    # rule build_overall_client_charts uses for its client ranking) so the
    # project needing the most attention sorts to the top.
    project_counts = []
    if "Project" in df.columns:
        for project, pdf_ in df.groupby(df["Project"].fillna("(No Project)")):
            statuses = pdf_["Ticket Status"] if "Ticket Status" in pdf_.columns else pd.Series(dtype=object)
            p_completed = int(statuses.isin(["Completed", "Closed"]).sum())
            p_pending = int((statuses == "Pending").sum())
            p_in_progress = int((statuses == "In Progress").sum())
            project_counts.append({
                "project": str(project),
                "total": int(len(pdf_)),
                "completed": p_completed,
                "pending": p_pending,
                "in_progress": p_in_progress,
                "sla_breach": int(pdf_["SLA Breach"].sum()) if "SLA Breach" in pdf_.columns else 0,
            })
        project_counts.sort(key=lambda p: p["pending"] + p["in_progress"], reverse=True)

    def ticket_entry(row):
        return {
            "ticket_no": row.get("Ticket No") if pd.notna(row.get("Ticket No")) else "",
            "task_type": row.get("Task Type") if pd.notna(row.get("Task Type")) else "",
            "project": row.get("Project") if pd.notna(row.get("Project")) else "",
            "title": row.get("Ticket Title") if pd.notna(row.get("Ticket Title")) else "",
            "priority": row.get("Priority") if pd.notna(row.get("Priority")) else "",
            "status": row.get("Ticket Status") if pd.notna(row.get("Ticket Status")) else "",
            "created": fmt_date(row.get("Ticket Created Date")),
            "completed": fmt_date(row.get("Ticket Completed Date")),
            "closed": fmt_date(row.get("Ticket Closed Date")),
            "ageing": row.get("Ageing") if pd.notna(row.get("Ageing")) else "",
            "sla_breach": bool(row.get("SLA Breach")) if pd.notna(row.get("SLA Breach")) else False,
        }

    # Newest first reads more usefully than source-file order for a report.
    if "Ticket Created Date" in df.columns:
        df = df.sort_values("Ticket Created Date", ascending=False, na_position="last")
    tickets = [ticket_entry(row) for _, row in df.iterrows()]

    open_statuses = {"Pending", "In Progress"}
    open_tickets = [t for t in tickets if t["status"] in open_statuses]
    sla_breaches = [t for t in tickets if t["sla_breach"]]

    # Resolution performance -- Days to Close is only populated once a
    # ticket has actually closed, so this is naturally scoped to resolved
    # tickets rather than needing its own status filter.
    avg_days = median_days = None
    if "Days to Close" in df.columns:
        valid_days = pd.to_numeric(df["Days to Close"], errors="coerce").dropna()
        if not valid_days.empty:
            avg_days = round(float(valid_days.mean()), 1)
            median_days = round(float(valid_days.median()), 1)
    sla_compliance_pct = round((total - sla_breach) / total * 100, 1) if total else None

    # Ageing buckets: Ageing is only populated for still-open tickets (see
    # build_ageing_charts), so this reads as "how old is each open ticket",
    # not a bucketing of every ticket ever raised.
    age_order = ["1-30 Days", "31-60 Days", "> 60 Days"]
    ageing_buckets = []
    if "Ageing" in df.columns:
        ac = df["Ageing"].dropna().value_counts().reindex(age_order, fill_value=0)
        ageing_buckets = [{"bucket": k, "count": int(v)} for k, v in ac.items()]

    # Monthly trend: tickets raised vs. resolved per calendar month, so the
    # report shows whether the workload is growing or the team is keeping
    # pace with it, not just a point-in-time snapshot.
    monthly_trend = []
    if "Ticket Created Date" in df.columns:
        created_by_month = df["Ticket Created Date"].dropna().dt.to_period("M").value_counts()
        completed_col = df["Ticket Completed Date"] if "Ticket Completed Date" in df.columns else pd.Series(dtype="datetime64[ns]")
        completed_by_month = completed_col.dropna().dt.to_period("M").value_counts()
        all_months = sorted(set(created_by_month.index) | set(completed_by_month.index))
        for period in all_months:
            monthly_trend.append({
                "month": str(period),
                "label": period.strftime("%b %Y"),
                "created": int(created_by_month.get(period, 0)),
                "completed": int(completed_by_month.get(period, 0)),
            })

    # Monthly trend broken down by project (tickets raised per month, one
    # series per project) -- capped to the top few projects by ticket
    # volume (project_counts is already sorted, just by open count instead,
    # so re-sort by total here) and the rest folded into "Other", the same
    # "don't let a long tail make the chart unreadable" rule a chart
    # library's own top-N grouping would apply. A project with only a
    # handful of tickets barely shows up on a monthly chart anyway.
    monthly_trend_by_project = []
    if "Project" in df.columns and "Ticket Created Date" in df.columns and all_months:
        MAX_PROJECT_SERIES = 6
        ranked_projects = sorted(project_counts, key=lambda p: p["total"], reverse=True)
        top_projects = {p["project"] for p in ranked_projects[:MAX_PROJECT_SERIES]}
        pdf_ = df.copy()
        pdf_["_report_project"] = pdf_["Project"].fillna("(No Project)").astype(str)
        pdf_["_report_project"] = pdf_["_report_project"].where(pdf_["_report_project"].isin(top_projects), "Other")
        for project, group in pdf_.groupby("_report_project"):
            created_pm = group["Ticket Created Date"].dropna().dt.to_period("M").value_counts()
            points = [{"month": str(p), "label": p.strftime("%b %Y"), "created": int(created_pm.get(p, 0))} for p in all_months]
            monthly_trend_by_project.append({"project": project, "points": points})
        # "Other" (if present) always last regardless of its volume -- it's
        # a catch-all bucket, not a project competing for rank.
        monthly_trend_by_project.sort(key=lambda s: (s["project"] == "Other", -sum(pt["created"] for pt in s["points"])))

    return {
        "client": client,
        "category": category,
        "generated_at": pd.Timestamp.now().strftime("%d/%m/%Y %H:%M"),
        "metrics": metrics,
        "status_counts": status_counts,
        "priority_counts": priority_counts,
        "project_counts": project_counts,
        "tickets": tickets,
        "resolution": {
            "avg_days": avg_days,
            "median_days": median_days,
            "sla_compliance_pct": sla_compliance_pct,
        },
        "ageing_buckets": ageing_buckets,
        "monthly_trend": monthly_trend,
        "monthly_trend_by_project": monthly_trend_by_project,
        "attention": {
            "open_tickets": open_tickets,
            "sla_breaches": sla_breaches,
        },
    }


def build_overall_client_charts(df, tickets_df=None, project_df=None):
    charts = {}
    if df.empty:
        return charts

    # Per-client ticket totals for the Maintenance/Warranty sections below.
    # Tickets aren't reliably linkable to one specific project row (see
    # _narrow_by_projek_name), so this is aggregated per Client rather than
    # per Projek Name -- a client with two rows in the same section will
    # show the same totals on both.
    def ticket_status_stats(statuses):
        return {
            "Pending": int((statuses == "Pending").sum()),
            "In Progress": int((statuses == "In Progress").sum()),
            "Total Tickets": int(len(statuses)),
        }

    ticket_stats_by_client = {}
    # Warranty tickets are all bucketed under the shared "Client Warranty"
    # sentinel Client value, with the real client recorded in Company
    # instead. That's a *separate* dict (not merged into
    # ticket_stats_by_client) because a client like MTIB has both regular
    # and warranty tickets -- its Warranty row needs just the warranty
    # slice, not the same totals its Maintenance row shows.
    warranty_ticket_stats_by_company = {}
    if tickets_df is not None and not tickets_df.empty and "Client" in tickets_df.columns and "Ticket Status" in tickets_df.columns:
        for client, statuses in tickets_df.groupby("Client")["Ticket Status"]:
            ticket_stats_by_client[client] = ticket_status_stats(statuses)
        if "Company" in tickets_df.columns:
            warranty_df = tickets_df[tickets_df["Client"] == "Client Warranty"]
            for company, statuses in warranty_df.groupby("Company")["Ticket Status"]:
                warranty_ticket_stats_by_company[company] = ticket_status_stats(statuses)

    # A project's tasks in the Project Details table each carry their own
    # Actual Start/End Date, which naturally vary from task to task -- there
    # is no single "the" actual date for the project as a whole. Covering
    # the full span means taking the earliest Actual Start Date and the
    # latest Actual End Date across every task under that (Client, Projek
    # Name), the same way a Gantt chart's overall span is read off its
    # first and last bars. Keyed by (Client, Projek Name) rather than just
    # Client so a client with more than one project (e.g. MARA) doesn't
    # blend two unrelated projects' actual dates together.
    show_actual_span_cols = (
        project_df is not None and not project_df.empty
        and {"Client", "Projek Name", "Actual Start Date", "Actual End Date"}.issubset(project_df.columns)
    )
    actual_span_by_project = {}
    # Fallback for a client whose Client Project rows never carry their own
    # Projek Name at all (e.g. LKTN, YIK) -- pandas groupby silently drops
    # NaN keys, so those clients would otherwise never get an entry above.
    # Aggregated per Client instead; only safe to use when nothing needs
    # disambiguating (see the "exactly one Development row" check below),
    # same rule as resolve_project_scope's Projek Name fallback.
    actual_span_by_client_fallback = {}
    if show_actual_span_cols:
        for (client, projek_name), pdf in project_df.groupby(["Client", "Projek Name"]):
            start_min = pdf["Actual Start Date"].min()
            end_max = pdf["Actual End Date"].max()
            if pd.isna(start_min) and pd.isna(end_max):
                continue
            actual_span_by_project[(client, projek_name)] = {
                "Actual Start Date": "" if pd.isna(start_min) else start_min.strftime("%d/%m/%Y"),
                "Actual End Date": "" if pd.isna(end_max) else end_max.strftime("%d/%m/%Y"),
            }
        blank_projek_clients = (
            set(project_df.loc[project_df["Projek Name"].isna(), "Client"])
            - set(project_df.loc[project_df["Projek Name"].notna(), "Client"])
        )
        for client in blank_projek_clients:
            pdf = project_df[project_df["Client"] == client]
            start_min = pdf["Actual Start Date"].min()
            end_max = pdf["Actual End Date"].max()
            if pd.isna(start_min) and pd.isna(end_max):
                continue
            actual_span_by_client_fallback[client] = {
                "Actual Start Date": "" if pd.isna(start_min) else start_min.strftime("%d/%m/%Y"),
                "Actual End Date": "" if pd.isna(end_max) else end_max.strftime("%d/%m/%Y"),
            }

    total = len(df)
    unique_clients = df["Client"].nunique() if "Client" in df.columns else 0
    charts["metrics"] = {"total": total, "clients": unique_clients}

    if "Projek Status" in df.columns:
        valid_status = df.dropna(subset=["Projek Status"])
        if not valid_status.empty:
            sc = valid_status["Projek Status"].value_counts().reset_index()
            sc.columns = ["Projek Status", "Count"]
            charts["status_counts"] = sc.to_dict("records")
            fig = px.pie(
                sc, names="Projek Status", values="Count",
                title="Projects by Status", color="Projek Status",
                hole=0.3,
            )
            fig.update_layout(
                template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#374151"), legend=dict(orientation="h", yanchor="bottom", y=-0.2),
            )
            charts["status_pie"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    display_cols = ["Client", "Projek ID", "Projek Name", "Projek Status", "Progress Status", "Start Date", "End Date", "Technology"]
    avail = [c for c in display_cols if c in df.columns]
    meta_cols = [c for c in ["_row_idx", "Source File"] if c in df.columns]
    detail = df[avail + meta_cols].copy()
    for c in ["Start Date", "End Date"]:
        if c in detail.columns and not detail[c].isna().all():
            detail[c] = detail[c].dt.strftime("%d/%m/%Y")
    detail = detail.fillna("")
    charts["detail_data"] = detail.to_dict("records")

    def section_rows(sdf, status):
        # Actual Start/End Date only makes sense for Development -- a
        # Warranty/Maintenance row is ongoing support work, not a project
        # with a start/end to report on, so those sections don't get the
        # columns at all (not just blank cells).
        if status == "Development" and show_actual_span_cols:
            sdf = sdf.copy()
            end_date_pos = sdf.columns.get_loc("End Date") + 1 if "End Date" in sdf.columns else len(sdf.columns)
            sdf.insert(end_date_pos, "Actual Start Date", "")
            sdf.insert(end_date_pos + 1, "Actual End Date", "")
            # A client-only fallback is only safe when there's nothing to
            # disambiguate -- i.e. this client has exactly one Development
            # row on the Home page. A client with two (e.g. MARA) keeps
            # relying on the exact (Client, Projek Name) match only.
            client_dev_counts = sdf["Client"].value_counts().to_dict()
            for i, row in sdf.iterrows():
                client = row.get("Client")
                span = actual_span_by_project.get((client, row.get("Projek Name")))
                if not span and client_dev_counts.get(client) == 1:
                    span = actual_span_by_client_fallback.get(client)
                if span:
                    sdf.loc[i, "Actual Start Date"] = span["Actual Start Date"]
                    sdf.loc[i, "Actual End Date"] = span["Actual End Date"]

        # Progress Status is a Development-only field (typed on this very
        # page) -- the other sections don't carry the column at all, the
        # same way Actual Start/End Date above are Development-only.
        if status != "Development" and "Progress Status" in sdf.columns:
            sdf = sdf.drop(columns=["Progress Status"])
        rows = sdf.to_dict("records")
        if status in ("Maintenance", "Warranty"):
            stats_source = warranty_ticket_stats_by_company if status == "Warranty" else ticket_stats_by_client
            for row in rows:
                client = row.get("Client")
                lookup_client = CLIENT_DISPLAY_ALIASES.get(client, client)
                stats = stats_source.get(lookup_client, {"Pending": 0, "In Progress": 0, "Total Tickets": 0})
                row.update(stats)
            rows.sort(key=lambda r: r["Pending"] + r["In Progress"], reverse=True)
        return rows

    charts["status_sections"] = {}
    if "Projek Status" in detail.columns:
        status_order = ["Development", "Warranty", "Maintenance"]
        statuses = detail["Projek Status"].dropna().unique()
        for status in status_order:
            if status in statuses:
                sdf = detail[detail["Projek Status"] == status]
                if sdf.empty:
                    continue
                charts["status_sections"][status] = {
                    "count": int(len(sdf)),
                    "rows": section_rows(sdf, status),
                }
        for status in sorted(set(statuses) - set(status_order), key=lambda s: str(s).lower()):
            sdf = detail[detail["Projek Status"] == status]
            if sdf.empty:
                continue
            charts["status_sections"][status] = {
                "count": int(len(sdf)),
                "rows": section_rows(sdf, status),
            }

    return charts


def build_charts(df):
    charts = {}

    total = len(df)
    completed = len(df[df["Ticket Status"].isin(["Completed", "Closed"])]) if "Ticket Status" in df.columns else 0
    pending = len(df[df["Ticket Status"] == "Pending"]) if "Ticket Status" in df.columns else 0
    in_progress = len(df[df["Ticket Status"] == "In Progress"]) if "Ticket Status" in df.columns else 0
    sla_breach = df["SLA Breach"].sum() if "SLA Breach" in df.columns else 0
    avg_days = None
    if "Days to Close" in df.columns:
        valid_days = df["Days to Close"].dropna()
        if len(valid_days) > 0:
            avg_days = round(valid_days.mean(), 1)

    metrics = {
        "total": total,
        "completed": completed,
        "pending": pending,
        "in_progress": in_progress,
        "sla_breach": int(sla_breach),
        "avg_days": avg_days,
        "completed_pct": f"{completed / total * 100:.1f}%" if total > 0 else "0%",
        "pending_pct": f"{pending / total * 100:.1f}%" if total > 0 else "0%",
        "in_progress_pct": f"{in_progress / total * 100:.1f}%" if total > 0 else "0%",
    }

    charts["metrics"] = metrics

    if "Ticket Status" in df.columns:
        status_counts = df["Ticket Status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Bilangan"]
        fig = px.pie(
            status_counts, names="Status", values="Bilangan",
            title="Ticket Status Distribution", color="Status",
            color_discrete_map=COLORS, hole=0.3,
        )
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), legend=dict(orientation="h", yanchor="bottom", y=-0.2))
        charts["status_pie"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Priority" in df.columns:
        priority_counts = df["Priority"].value_counts().reset_index()
        priority_counts.columns = ["Keutamaan", "Bilangan"]
        fig = px.pie(
            priority_counts, names="Keutamaan", values="Bilangan",
            title="Priority Distribution", color="Keutamaan",
            color_discrete_map=PRIORITY_COLORS, hole=0.3,
        )
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), legend=dict(orientation="h", yanchor="bottom", y=-0.2))
        charts["priority_pie"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Client" in df.columns:
        client_dist = df["Client"].value_counts().reset_index()
        client_dist.columns = ["Client", "Bilangan"]
        client_colors = px.colors.qualitative.Plotly[:len(client_dist)]
        fig = px.bar(
            client_dist, x="Client", y="Bilangan",
            title="Tickets by Client", color="Client",
            color_discrete_sequence=client_colors, text="Bilangan",
        )
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), showlegend=False, xaxis_tickangle=-45)
        charts["client_bar"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    return charts


def build_priority_charts(df):
    charts = {}

    if "Priority" not in df.columns:
        return charts

    priority_counts = df["Priority"].value_counts().reset_index()
    priority_counts.columns = ["Keutamaan", "Bilangan"]
    fig = px.pie(
        priority_counts, names="Keutamaan", values="Bilangan",
        title="Priority Distribution", color="Keutamaan",
        color_discrete_map=PRIORITY_COLORS, hole=0.4,
    )
    fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
    charts["priority_pie"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Ticket Status" in df.columns:
        cross = df.groupby(["Priority", "Ticket Status"]).size().reset_index(name="Bilangan")
        fig = px.bar(
            cross, x="Priority", y="Bilangan",
            color="Ticket Status", title="Priority by Status",
            color_discrete_map=COLORS, barmode="group",
        )
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
        charts["priority_status_bar"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Client" in df.columns:
        pivot = df.groupby(["Client", "Priority"]).size().unstack(fill_value=0)
        charts["priority_client_pivot"] = pivot.to_html()

    return charts


def build_ageing_charts(df):
    charts = {}
    age_order = ["1-30 Days", "31-60 Days", "> 60 Days"]
    charts["age_order"] = age_order
    charts["ageing_clients"] = {}

    has_ageing = "Ageing" in df.columns and df["Ageing"].notna().any()
    has_days = "Days" in df.columns and df["Days"].notna().any()

    if not has_ageing and not has_days:
        return charts

    if has_ageing and "Client" in df.columns:
        total_all = df["Ageing"].notna().sum()
        charts["total_ageing"] = int(total_all)

        for client in sorted(df["Client"].unique()):
            dc = df[df["Client"] == client].dropna(subset=["Ageing"])
            if dc.empty:
                continue

            counts = dc["Ageing"].value_counts().reindex(age_order, fill_value=0).reset_index()
            counts.columns = ["Kumpulan Umur", "Bilangan"]

            fig = px.bar(
                counts, x="Kumpulan Umur", y="Bilangan",
                color="Kumpulan Umur", color_discrete_map=AGEING_COLORS,
                text="Bilangan", title=client,
            )
            fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), showlegend=False)
            charts["ageing_clients"][client] = {
                "count": int(len(dc)),
                "chart": fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False}),
                "table": {k: int(counts.set_index("Kumpulan Umur").loc[k, "Bilangan"]) for k in age_order},
            }

    if has_days:
        valid_days = df["Days"].dropna()
        if len(valid_days) > 0:
            fig = px.histogram(
                df.dropna(subset=["Days"]), x="Days", nbins=30,
                title="Days Open Distribution", color_discrete_sequence=["#3498db"],
                marginal="box",
            )
            fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
            charts["days_hist"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "SLA Breach" in df.columns and "Client" in df.columns:
        sla_by_client = df.groupby("Client")["SLA Breach"].sum().reset_index()
        sla_by_client.columns = ["Client", "Pelanggaran SLA"]
        fig = px.bar(
            sla_by_client, x="Client", y="Pelanggaran SLA",
            title="Total SLA Breaches", color_discrete_sequence=["#e74c3c"],
            text="Pelanggaran SLA",
        )
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), xaxis_tickangle=-45)
        charts["sla_breach_bar"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    return charts


def build_client_comparison_charts(df):
    charts = {}

    if "Client" not in df.columns:
        return charts

    client_stats = df.groupby("Client").agg(
        Jumlah=("Ticket No", "count") if "Ticket No" in df.columns else ("Ticket Status", "count"),
    ).reset_index()

    if "Ticket Status" in df.columns:
        status_counts = df.groupby(["Client", "Ticket Status"]).size().unstack(fill_value=0)
        client_stats = client_stats.merge(status_counts, on="Client", how="left")

    if "Days to Close" in df.columns:
        avg_days = df.groupby("Client")["Days to Close"].mean().reset_index()
        avg_days.columns = ["Client", "Purata Hari"]
        client_stats = client_stats.merge(avg_days, on="Client", how="left")

    if "SLA Breach" in df.columns:
        sla = df.groupby("Client")["SLA Breach"].sum().reset_index()
        sla.columns = ["Client", "Pelanggaran SLA"]
        client_stats = client_stats.merge(sla, on="Client", how="left")

    charts["client_stats_table"] = client_stats.to_html(index=False)

    exclude_clients = ["Client Warranty", "KUIPS"]
    chart_clients = client_stats[~client_stats["Client"].isin(exclude_clients)]

    df_filtered = df[~df["Client"].isin(exclude_clients)]
    if "Ticket Status" in df_filtered.columns:
        status_counts = df_filtered["Ticket Status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Bilangan"]
        fig = px.bar(
            status_counts, x="Bilangan", y="Status",
            orientation="h", title="Count by Status",
            color="Status", color_discrete_sequence=px.colors.qualitative.Plotly[:len(status_counts)],
            text="Bilangan",
        )
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
        charts["count_by_status"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Ticket Status" in df_filtered.columns:
        pivot = df_filtered.groupby(["Client", "Ticket Status"]).size().unstack(fill_value=0)
        pivot["Total"] = pivot.sum(axis=1)
        pivot.loc["Total"] = pivot.sum()
        pivot = pivot.astype(int)
        charts["status_pivot"] = pivot.to_html()

    fig = px.bar(
        chart_clients, x="Client", y="Jumlah",
        title="Total Tickets by Client", color="Client", text="Jumlah",
    )
    fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), showlegend=False, xaxis_tickangle=-45)
    charts["client_total_bar"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    status_order = ["Pending", "In Progress", "Completed", "Closed"]
    status_cols = [c for c in status_order if c in chart_clients.columns]
    if status_cols:
        fig = go.Figure()
        for col in status_cols:
            color = COLORS.get(col, "#95a5a6")
            fig.add_trace(go.Bar(name=col, x=chart_clients["Client"], y=chart_clients[col], marker_color=color, text=chart_clients[col], textposition="outside", textfont=dict(color="#374151", size=10)))
        fig.update_layout(
            barmode="group", title="Status by Client",
            template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#374151"), xaxis_tickangle=-45,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        )
        charts["status_by_client"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Priority" in df.columns:
        priority_dummies = pd.get_dummies(df[["Client", "Priority"]], columns=["Priority"])
        radar_data = priority_dummies.groupby("Client").sum().reset_index()
        categories = [c for c in radar_data.columns if c.startswith("Priority_")]
        if categories:
            fig = go.Figure()
            for _, row in radar_data.iterrows():
                values = [row[c] for c in categories]
                values.append(values[0])
                cats = [c.replace("Priority_", "") for c in categories]
                cats.append(cats[0])
                fig.add_trace(go.Scatterpolar(r=values, theta=cats, fill="toself", name=row["Client"]))
            fig.update_layout(polar=dict(radialaxis=dict(visible=True)), title="Priority Profile by Client", template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
            charts["client_radar"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    return charts


def build_timeline_charts(df):
    charts = {}

    if "Ticket Created Date" not in df.columns:
        return charts

    df_dated = df[df["Ticket Created Date"].notna()].copy()
    if len(df_dated) == 0:
        return charts

    df_dated["Bulan"] = df_dated["Ticket Created Date"].dt.to_period("M").astype(str)
    monthly_created = df_dated.groupby("Bulan").size().reset_index(name="Dicipta")

    fig = px.line(
        monthly_created, x="Bulan", y="Dicipta",
        title="Tickets Created by Month", markers=True,
        color_discrete_sequence=["#3498db"],
    )
    fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
    charts["timeline_created"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Ticket Completed Date" in df.columns:
        df_completed = df[df["Ticket Completed Date"].notna()].copy()
        if len(df_completed) > 0:
            df_completed["Bulan"] = df_completed["Ticket Completed Date"].dt.to_period("M").astype(str)
            monthly_completed = df_completed.groupby("Bulan").size().reset_index(name="Selesai")

            merged = monthly_created.merge(monthly_completed, on="Bulan", how="left").fillna(0)
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=merged["Bulan"], y=merged["Dicipta"], mode="lines+markers", name="Dicipta", line=dict(color="#3498db", width=2)))
            fig.add_trace(go.Scatter(x=merged["Bulan"], y=merged["Selesai"], mode="lines+markers", name="Selesai", line=dict(color="#2ecc71", width=2)))
            fig.update_layout(title="Created vs Completed", template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), xaxis_tickangle=-45)
            charts["timeline_created_vs_completed"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Client" in df.columns:
        client_monthly = df_dated.groupby(["Bulan", "Client"]).size().reset_index(name="Bilangan")
        fig = px.area(client_monthly, x="Bulan", y="Bilangan", color="Client", title="Tickets by Client and Month")
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
        charts["timeline_client_area"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Ticket Category" in df.columns:
        cat_monthly = df_dated.groupby(["Bulan", "Ticket Category"]).size().reset_index(name="Bilangan")
        if len(cat_monthly) > 0:
            top_cats = df_dated["Ticket Category"].value_counts().head(8).index.tolist()
            cat_monthly = cat_monthly[cat_monthly["Ticket Category"].isin(top_cats)]
            fig = px.line(cat_monthly, x="Bulan", y="Bilangan", color="Ticket Category", title="Tickets by Category (Top 8)", markers=True)
            fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
            charts["timeline_category"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    return charts


def build_sla_charts(df):
    charts = {}

    if "SLA Breach" not in df.columns:
        return charts

    exclude_clients = ["Client Warranty", "KUIPS"]
    if "Client" in df.columns:
        df = df[~df["Client"].isin(exclude_clients)].copy()

    total = len(df)
    breaches = df["SLA Breach"].sum()
    compliance_rate = round((total - breaches) / total * 100, 1) if total > 0 else 0
    charts["compliance_rate"] = compliance_rate
    charts["total_breaches"] = int(breaches)
    charts["total_compliant"] = int(total - breaches)

    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=compliance_rate,
        title={"text": "SLA Compliance Rate (%)"},
        gauge=dict(
            axis=dict(range=[0, 100]), bar=dict(color="#2ecc71"),
            steps=[
                dict(range=[0, 50], color="#e74c3c"),
                dict(range=[50, 75], color="#f39c12"),
                dict(range=[75, 100], color="#2ecc71"),
            ],
            threshold=dict(line=dict(color="white", width=2), thickness=0.75, value=compliance_rate),
        ),
    ))
    fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), height=350)
    charts["sla_gauge"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Client" in df.columns:
        client_sla = df.groupby("Client").agg(Total=("SLA Breach", "count"), Breaches=("SLA Breach", "sum")).reset_index()
        client_sla["Kadar Pematuhan (%)"] = ((client_sla["Total"] - client_sla["Breaches"]) / client_sla["Total"] * 100).round(1)
        client_sla = client_sla.sort_values("Kadar Pematuhan (%)", ascending=True)

        fig = px.bar(
            client_sla, x="Kadar Pematuhan (%)", y="Client",
            orientation="h", title="SLA Compliance Rate by Client",
            color="Kadar Pematuhan (%)", color_continuous_scale=["#e74c3c", "#f39c12", "#2ecc71"],
            text="Kadar Pematuhan (%)",
        )
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
        charts["sla_client_bar"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

        if "Ticket Status" in df.columns:
            status_map = {
                "Completed": "Closed + Completed",
                "Closed": "Closed + Completed",
                "Pending": "Pending + In Progress",
                "In Progress": "Pending + In Progress",
            }
            status_group = df["Ticket Status"].replace(status_map)
            sla_pivot = df.groupby(["Client", status_group])["SLA Breach"].agg(["sum", "count", "mean"]).reset_index()
            sla_pivot.columns = ["Client", "Status", "Pelanggaran", "Jumlah", "Kadar Pelanggaran"]
            sla_pivot["Kadar Pelanggaran"] = (sla_pivot["Kadar Pelanggaran"] * 100).round(1)

            if "SLA Late" in df.columns and "Ageing" in df.columns:
                sla_valid = pd.to_numeric(df["SLA Late"].astype(str).str.strip().replace({"nan": ""}), errors="coerce").notna()
                age_valid = df["Ageing"].astype(str).str.strip()
                age_valid = age_valid.ne("") & age_valid.str.lower().ne("nan") & age_valid.ne("Not Due")

                open_counts = df[df["Ticket Status"].isin(["Pending", "In Progress"]) & sla_valid & age_valid].groupby("Client").size()
                sla_pivot["Open (SLA+Ageing)"] = sla_pivot.apply(
                    lambda r: int(open_counts.get(r["Client"], 0)) if r["Status"] == "Pending + In Progress" else "",
                    axis=1,
                )

            charts["sla_pivot"] = sla_pivot.to_html(index=False)

    return charts


def default_tab_idx(filters):
    """Mirrors the exact predicate templates/dashboard.html uses (via its
    single_client_mode/client_category {% set %}s) to decide which tab-pane
    gets the "active" class -- kept in one place so index() and
    build_tab_context() can't drift apart on which tab is "the" default."""
    if len(filters["clients"]) == 1:
        category = filters["task_types"][0] if len(filters["task_types"]) == 1 else None
        if category == "Development":
            return 3
        if category == "Warranty":
            return 2
        return 7
    return 9


def build_filter_options(filters):
    try:
        filter_options = db.get_filter_metadata(conn=request_conn())
    except Exception as e:
        log(f"DB error loading filter metadata: {e}", "ERROR")
        filter_options = {}
    filter_options["search"] = request.args.get("search", "")
    filter_options["date_start"] = request.args.get("date_start", "")
    filter_options["date_end"] = request.args.get("date_end", "")
    filter_options["selected_clients"] = filters["clients"]
    filter_options["selected_priorities"] = filters["priorities"]
    filter_options["selected_statuses"] = filters["statuses"]
    filter_options["selected_task_types"] = filters["task_types"]
    filter_options["projek_name"] = filters.get("projek_name") or ""
    return filter_options


def build_tab_context(idx, filters, filter_options, df=None):
    """Returns (template_name, context) for exactly one tab -- the per-tab
    slice of what index() used to compute unconditionally for all 11 tabs
    on every request. `df` lets a caller that already loaded the ticket
    dataframe (index(), for whichever tab is the current default) hand it
    in instead of paying for a second query."""
    single_client_mode = len(filters["clients"]) == 1
    client_category = filters["task_types"][0] if len(filters["task_types"]) == 1 else None
    common = {
        "filter_options": filter_options,
        "single_client_mode": single_client_mode,
        "client_category": client_category,
    }

    if idx == 0:
        if df is None:
            df, _ = load_data(filters)
        overview_charts = build_charts(df) if not df.empty else {}
        return "tabs/tab_0.html", {**common, "overview_charts": overview_charts}

    if idx == 1:
        if df is None:
            df, _ = load_data(filters)
        comparison_charts = build_client_comparison_charts(df) if not df.empty else {}
        return "tabs/tab_1.html", {**common, "comparison_charts": comparison_charts}

    if idx == 2:
        # Warranty tickets are stored under a shared "Client Warranty"
        # sentinel in the tickets table (not the real client name) -- the
        # real client is only recorded in the Company column. So filtering
        # tickets by the real client name (as every other category does)
        # always returns zero rows for Warranty. Detect that case and fetch
        # a Warranty-appropriate dataframe instead, scoped by Company.
        warranty_client = filters["clients"][0] if (single_client_mode and client_category == "Warranty") else None
        if warranty_client:
            # A Home page client name (e.g. UNISIRAJ) can differ from the
            # name its own ticket data uses (KUIPS) -- see
            # CLIENT_DISPLAY_ALIASES.
            warranty_client = CLIENT_DISPLAY_ALIASES.get(warranty_client, warranty_client)
            try:
                warranty_df, _ = load_data({"clients": ["Client Warranty"], "priorities": [], "statuses": [], "task_types": [], "search": None})
            except Exception as e:
                log(f"DB error loading warranty tickets: {e}", "ERROR")
                warranty_df = pd.DataFrame()
            if "Company" in warranty_df.columns:
                warranty_df = warranty_df[warranty_df["Company"] == warranty_client]
        else:
            if df is None:
                df, _ = load_data(filters)
            warranty_df = df
        warranty_charts = build_warranty_charts(warranty_df) if not warranty_df.empty else {}
        return "tabs/tab_2.html", {**common, "warranty_charts": warranty_charts}

    if idx == 3:
        try:
            project_df = load_project_data()
        except Exception as e:
            log(f"DB error loading projects: {e}", "ERROR")
            project_df = pd.DataFrame()
        project_df = resolve_project_scope(project_df, filters)
        project_df = recompute_status_from_percentage(project_df)
        project_df = recompute_overall_progress(project_df)
        project_df = recompute_duration(project_df)
        has_project = not project_df.empty
        project_charts = build_project_charts(project_df) if has_project else {}
        return "tabs/tab_3.html", {**common, "has_project": has_project, "project_charts": project_charts}

    if idx == 4:
        if df is None:
            df, _ = load_data(filters)
        # Matches the shape build_ageing_charts() always returns (even for
        # a genuinely empty df) -- the template unconditionally iterates
        # ageing_charts.ageing_clients.items(), so a bare {} would raise
        # UndefinedError instead of just rendering zero rows.
        ageing_charts = build_ageing_charts(df) if not df.empty else {"age_order": ["1-30 Days", "31-60 Days", "> 60 Days"], "ageing_clients": {}}
        return "tabs/tab_4.html", {**common, "ageing_charts": ageing_charts}

    if idx == 5:
        if df is None:
            df, _ = load_data(filters)
        filtered_has_data = not df.empty
        timeline_charts = build_timeline_charts(df) if filtered_has_data else {}
        return "tabs/tab_5.html", {**common, "filtered_has_data": filtered_has_data, "timeline_charts": timeline_charts}

    if idx == 6:
        if df is None:
            df, _ = load_data(filters)
        sla_charts = build_sla_charts(df) if not df.empty else {}
        return "tabs/tab_6.html", {**common, "sla_charts": sla_charts}

    if idx == 7:
        if df is None:
            df, _ = load_data(filters)
        filtered_has_data = not df.empty
        display_cols = [
            "Client", "Ticket No", "Task Type", "Project", "Company",
            "Ticket Title", "Ticket Category", "Priority", "Ticket Status",
            "Ticket Created Date", "Ticket Completed Date", "Ticket Closed Date",
            "Days to Close", "Ageing", "SLA Breach",
        ]
        avail_cols = [c for c in display_cols if c in df.columns]
        meta_cols = ["_row_idx", "Source File"]
        detail_cols = avail_cols + [c for c in meta_cols if c in df.columns]
        detail_df = df[detail_cols].copy() if filtered_has_data and detail_cols else pd.DataFrame()

        # Contract period per Maintenance client (from the same clients
        # table the Home page renders), used to sort every ticket into
        # the tab's two groups: "Current Contract" (creation date on/after
        # the contract's Start Date) and "Old Contract" (before it). A
        # client with several Maintenance contracts matches each ticket
        # to a contract by its Project name; tickets that don't match any
        # project fall back to the client's earliest contract start.
        contracts = {}
        try:
            clients_df = db.fetch_clients_df(conn=request_conn())
        except Exception as e:
            log(f"DB error loading clients for contract split: {e}", "ERROR")
            clients_df = pd.DataFrame()
        if (
            not clients_df.empty
            and {"Client", "Projek Status", "Projek Name", "Start Date", "End Date"}.issubset(clients_df.columns)
        ):
            maint_rows = clients_df[
                (clients_df["Projek Status"] == "Maintenance") & clients_df["Start Date"].notna()
            ]
            for client, group in maint_rows.groupby("Client"):
                contracts[client] = [
                    (p, s, e)
                    for p, s, e in zip(group["Projek Name"], group["Start Date"], group["End Date"])
                    if p
                ]

        def contract_period_for(project, client):
            """(start, end) for the client's contract matching `project`'s
            name, or the combined earliest-start/latest-end span as a
            fallback, or (None, None) when the client has no contract."""
            rows = contracts.get(client)
            if not rows:
                return None, None
            project_l = (project or "").strip().lower()
            for projek_name, start, end in rows:
                name_l = (projek_name or "").strip().lower()
                if project_l and (project_l in name_l or name_l in project_l):
                    return start, end
            return min(s for _, s, _ in rows), max((e for _, _, e in rows if e is not None), default=None)

        # A maintenance row with a blank Projek Name can't be matched
        # against any ticket's Project -- drop it rather than hand the
        # template a split with no dates to show.
        contracts = {client: rows for client, rows in contracts.items() if rows}

        # The tab renders as two top-level groups -- Current Contract and
        # Old Contract -- each holding one table per client. A ticket is
        # Old when its creation date precedes the contract's Start Date;
        # on/after it (even past the End Date) is Current. Clients with no
        # Maintenance contract can't be classified, so all their tickets
        # sit under Current Contract without any contract dates shown.
        created_col = "Ticket Created Date" if "Ticket Created Date" in detail_df.columns else None
        groups = {"current": [], "old": []}
        by_client = {}
        for rec in detail_df.to_dict("records"):
            client = rec.get("Client", "Unknown")
            start, end = contract_period_for(rec.get("Project"), client)
            if start is None:
                group = "current"
            else:
                created = rec.get(created_col) if created_col else None
                is_old = created is not None and not pd.isna(created) and created < start
                group = "old" if is_old else "current"
            key = (group, client)
            entry = by_client.get(key)
            if entry is None:
                entry = {"client": client, "rows": [], "starts": [], "ends": []}
                by_client[key] = entry
                groups[group].append(entry)
            entry["rows"].append(rec)
            if start is not None:
                entry["starts"].append(start)
                if end is not None:
                    entry["ends"].append(end)

        for entry in by_client.values():
            entry["contract_start"] = min(entry["starts"]) if entry["starts"] else None
            entry["contract_end"] = max(entry["ends"]) if entry["ends"] else None
            del entry["starts"], entry["ends"]

        for col in ("Ticket Created Date", "Ticket Completed Date", "Ticket Closed Date"):
            for group_rows in groups.values():
                for entry in group_rows:
                    for rec in entry["rows"]:
                        val = rec.get(col)
                        if pd.isna(val):
                            rec[col] = "NaT"
                        else:
                            rec[col] = pd.Timestamp(val).strftime("%d/%m/%Y")
        for entry in by_client.values():
            for key in ("contract_start", "contract_end"):
                if entry[key] is not None:
                    entry[key] = pd.Timestamp(entry[key]).strftime("%d/%m/%Y")
        detail_groups = {
            "current": groups["current"],
            "old": groups["old"],
            "current_total": sum(len(e["rows"]) for e in groups["current"]),
            "old_total": sum(len(e["rows"]) for e in groups["old"]),
        }
        # Status roll-up shown next to each group's heading. The four
        # statuses the user expects always appear (even at 0); anything
        # else in the data (e.g. Deleted) is appended. "In Progress" is
        # the rare alternate spelling of "Inprogress" -- merged so it
        # doesn't show up as its own chip.
        for group_key in ("current", "old"):
            counts = {}
            for entry in groups[group_key]:
                for rec in entry["rows"]:
                    raw = (rec.get("Ticket Status") or "").strip() or "Unknown"
                    status = "Inprogress" if raw.lower() in ("inprogress", "in progress") else raw
                    counts[status] = counts.get(status, 0) + 1
            ordered = [(s, counts.pop(s, 0)) for s in ("Pending", "Inprogress", "Completed", "Closed")]
            ordered.extend(sorted(counts.items()))
            detail_groups[group_key + "_status"] = ordered
        return "tabs/tab_7.html", {**common, "detail_groups": detail_groups, "data_info": {"total_filtered": len(df)}}

    if idx == 8:
        if df is None:
            df, _ = load_data(filters)
        ageing_list_data = {}
        if not df.empty and "Ageing" in df.columns and df["Ageing"].notna().sum() > 0:
            age_order = ["1-30 Days", "31-60 Days", "> 60 Days"]
            ageing_cols = [c for c in ["Client", "Ticket No", "Ticket Title", "Ticket Status", "Priority", "Ticket Created Date", "Days", "_row_idx", "Source File"] if c in df.columns]
            ageing_df = df.dropna(subset=["Ageing"]).copy()
            if "Ticket Created Date" in ageing_df.columns:
                ageing_df["Ticket Created Date"] = ageing_df["Ticket Created Date"].dt.strftime("%d/%m/%Y")
            ageing_list_data = {"total": int(ageing_df["Ageing"].notna().sum()), "buckets": {}}
            for bucket in age_order:
                bucket_df = ageing_df[ageing_df["Ageing"] == bucket]
                if bucket_df.empty:
                    continue
                clients_in_bucket = {}
                for client in sorted(bucket_df["Client"].unique()):
                    client_rows_df = bucket_df[bucket_df["Client"] == client]
                    clients_in_bucket[client] = {
                        "count": len(client_rows_df),
                        "rows": client_rows_df[ageing_cols].to_dict("records"),
                    }
                ageing_list_data["buckets"][bucket] = clients_in_bucket
        return "tabs/tab_8.html", {**common, "ageing_list_data": ageing_list_data}

    if idx == 9:
        # Make sure transfer_history (part of SCHEMA_SQL) exists before the
        # auto-transfer below queries it -- load_client_data() would
        # otherwise be the first thing to create the schema.
        ensure_schema()
        # A Development row whose Progress Status is "Completed" and whose
        # End Date has passed is finished work that now belongs under
        # Warranty, so the move happens here, before the data is read --
        # one page load is all it takes and no button press is needed.
        # Never let a failed write take the Home page down with it: the
        # rows simply stay in Development until the next load.
        try:
            moved = db.auto_transfer_development_rows(conn=request_conn())
            if moved:
                names = ", ".join(f"{r['client']}/{r['projek_name'] or r['projek_id']}" for r in moved)
                log(f"Auto-transferred {len(moved)} completed Development row(s) to Warranty: {names}")
        except Exception as e:
            log(f"Auto-transfer to Warranty failed: {e}", "ERROR")
        try:
            client_df = load_client_data()
        except Exception as e:
            log(f"DB error loading clients: {e}", "ERROR")
            client_df = pd.DataFrame()
        has_client = not client_df.empty
        tickets_df = pd.DataFrame()
        project_df = pd.DataFrame()
        if has_client:
            try:
                tickets_df, _ = load_data({})
            except Exception as e:
                log(f"DB error loading tickets for Home totals: {e}", "ERROR")
            try:
                project_df = load_project_data()
            except Exception as e:
                log(f"DB error loading projects for Home actual dates: {e}", "ERROR")
        overall_client_charts = build_overall_client_charts(client_df, tickets_df, project_df) if has_client else {}
        try:
            transfer_history = db.fetch_transfer_history(conn=request_conn())
        except Exception as e:
            log(f"DB error loading transfer history: {e}", "ERROR")
            transfer_history = []
        return "tabs/tab_9.html", {**common, "has_client": has_client, "overall_client_charts": overall_client_charts, "transfer_history": transfer_history}

    if idx == 10:
        return "tabs/tab_10.html", common

    if idx == 11:
        # Dedup Seq is re-upload plumbing (see idx_projectmilestone_dedup_key),
        # not a field anyone edits -- same skip the Project Details tab does.
        cols = [c for c, _ in db.PROJECT_MILESTONE_DB_COLUMNS if c != "Dedup Seq"]
        try:
            milestone_df = load_milestone_data()
        except Exception as e:
            log(f"DB error loading milestones: {e}", "ERROR")
            milestone_df = pd.DataFrame()
        if milestone_df.empty:
            milestone_df = pd.DataFrame(columns=["_row_idx"] + cols)
        # Same narrowing as every other tab: the ?client=... / ?projek_name=...
        # in the URL (set when a Home row is clicked) scopes this table too,
        # so a single-client view doesn't list every other client's tasks.
        if filters["clients"]:
            milestone_df = milestone_df[milestone_df["Client"].isin(filters["clients"])]
        if filters.get("projek_name"):
            milestone_df = milestone_df[
                milestone_df["Projectname"].fillna("").str.contains(
                    filters["projek_name"], case=False, regex=False
                )
            ]

        detail = milestone_df[["_row_idx"] + [c for c in cols if c in milestone_df.columns]].copy()
        for c in ("Startdate", "Enddate"):
            if c in detail.columns and not detail[c].isna().all():
                detail[c] = detail[c].dt.strftime("%d/%m/%Y")
        detail = detail.fillna("")
        return (
            "tabs/tab_11.html",
            {
                **common,
                "milestone_cols": [c for c in cols if c in detail.columns],
                "milestone_rows": detail.to_dict("records"),
            },
        )

    if idx == 12:
        # The Task Detail tab: one row per process line of a
        # "<project> TASK DETAIL" sheet, grouped into modules. Scoped by the
        # same ?client= / ?projek_name= every other tab narrows on, so the
        # sidebar button only ever shows this project's rows.
        cols = db.TASK_DETAIL_COLUMNS
        try:
            task_df = db.fetch_project_task_detail_df(
                clients=filters["clients"] or None,
                projek_name=filters.get("projek_name"),
                conn=request_conn(),
            )
        except Exception as e:
            log(f"DB error loading task detail: {e}", "ERROR")
            task_df = pd.DataFrame()
        if task_df.empty:
            task_df = pd.DataFrame(columns=["_row_idx"] + cols)

        detail = task_df[["_row_idx"] + [c for c in cols if c in task_df.columns]].copy()
        if "Target Date" in detail.columns and detail["Target Date"].notna().any():
            # dd/mm/yyyy, the format buildEditableCellInput's ddmmyyyyToIso
            # expects so the date input picks the value up on edit.
            detail["Target Date"] = detail["Target Date"].dt.strftime("%d/%m/%Y")
        if "Percentage" in detail.columns:
            detail["Percentage"] = detail["Percentage"].map(
                lambda v: f"{v:g}%" if pd.notna(v) else ""
            )
        # astype(object) first: a fillna("") on an Int64/float column would
        # try to cast the string and raise.
        detail = detail.astype(object).where(detail.notna(), "")

        rows = detail.to_dict("records")
        # Consecutive rows sharing a No/Modul form one module block -- the
        # template draws a band per block instead of repeating 55 labels.
        modules, seen = [], set()
        for row in rows:
            key = (row.get("No", ""), row.get("Modul", ""))
            if key not in seen:
                seen.add(key)
                modules.append({"no": row.get("No", ""), "modul": row.get("Modul", ""), "rows": []})
            modules[-1]["rows"].append(row)

        return (
            "tabs/tab_12.html",
            {
                **common,
                "task_detail_cols": [c for c in cols if c in detail.columns],
                "task_detail_rows": rows,
                "task_detail_modules": modules,
                # The dedup-key columns: shown for context, kept read-only
                # so an edit can't collide with idx_ptd_dedup_key.
                "task_detail_readonly_cols": ["No", "Modul", "Proses"],
            },
        )

    raise ValueError(f"Unknown tab index: {idx}")


def render_tab_html(idx, filters, filter_options, df=None):
    template_name, ctx = build_tab_context(idx, filters, filter_options, df=df)
    return render_template(template_name, **ctx)


@app.route("/")
def index():
    user_role = session.get("role")
    filters = parse_filters(request.args)
    try:
        df, load_errors = load_data(filters)
    except Exception as e:
        log(f"DB error loading tickets: {e}", "ERROR")
        df, load_errors = pd.DataFrame(), [str(e)]

    filter_options = build_filter_options(filters)

    try:
        counts = db.get_counts(conn=request_conn())
    except Exception:
        counts = {"tickets": len(df), "projects": 0, "last_updated": None}

    # has_data drives the page shell (sidebar + tabs vs. the "upload data"
    # empty state) and must reflect whether the database has any tickets
    # at all -- not whether the current client/task-type filter happens to
    # match anything. Otherwise clicking into a client + category combo
    # with zero matching tickets (e.g. a client with no "Development"
    # tickets) would incorrectly collapse the whole dashboard back to the
    # "no data yet" prompt instead of showing that pane empty.
    has_data = counts.get("tickets", 0) > 0

    data_info = {
        "total_raw": counts.get("tickets", len(df)),
        "total_filtered": len(df),
        "load_errors": load_errors,
        "columns": list(df.columns) if not df.empty else [],
        "counts": counts,
    }

    # The sidebar only advertises Task Detail when the current scope has
    # rows -- a bare EXISTS, so no project other than the one with a
    # "<project> TASK DETAIL" sheet gets a button that opens an empty tab.
    # load_data() above has already run ensure_schema().
    try:
        has_task_detail = db.has_task_detail(
            clients=filters["clients"] or None,
            projek_name=filters.get("projek_name"),
            conn=request_conn(),
        )
    except Exception as e:
        log(f"DB error checking task detail: {e}", "ERROR")
        has_task_detail = False

    # Only the tab that would be "active" by default is computed/rendered
    # here -- every other tab is fetched lazily by the browser (see
    # /api/tab/<idx> below and switchTab() in dashboard.html) the first
    # time the user actually clicks into it, instead of every tab's charts
    # being built on every single page view.
    idx = default_tab_idx(filters)
    default_tab_html = render_tab_html(idx, filters, filter_options, df=df)

    return render_template(
        "dashboard.html",
        user_role=user_role,
        has_data=has_data,
        data_info=data_info,
        filter_options=filter_options,
        default_tab_idx=idx,
        default_tab_html=default_tab_html,
        has_task_detail=has_task_detail,
        now=datetime.now().strftime("%d-%m-%Y %H:%M"),
        max_upload_mb=MAX_UPLOAD_MB,
    )


@app.route("/api/tab/<int:idx>")
def api_tab(idx):
    if idx < 0 or idx > 12:
        return "Not found", 404
    filters = parse_filters(request.args)
    filter_options = build_filter_options(filters)
    try:
        return render_tab_html(idx, filters, filter_options)
    except Exception as e:
        log(f"Tab {idx} render error: {e}", "ERROR")
        return f"<div class='tab-loading'>Failed to load: {e}</div>", 500


@app.route("/api/project_report_data")
def api_project_report_data():
    """Feeds the client-side PDF report (pdfmake, built in dashboard.html)
    -- generation happens entirely in the browser, this just hands back
    the one project's data as JSON, scoped the same way the Project tab
    itself is (see resolve_project_scope)."""
    ensure_schema()
    client = request.args.get("client") or ""
    projek_name = request.args.get("projek_name") or None
    if not client:
        return jsonify({"success": False, "error": "client is required"}), 400
    try:
        data = build_project_report_data(client, projek_name)
    except Exception as e:
        log(f"Project report data error: {e}", "ERROR")
        return jsonify({"success": False, "error": str(e)}), 500
    if data is None:
        label = f"{client} / {projek_name}" if projek_name else client
        return jsonify({"success": False, "error": f"No project data found for {label}"}), 404
    return jsonify({"success": True, "data": data})


@app.route("/api/ticket_report_data")
def api_ticket_report_data():
    """Feeds the client-side PDF/PPTX report (pdfmake/pptxgenjs, see
    dashboard.html) for a client's Warranty or Maintenance tickets --
    generation happens entirely in the browser, this just hands back the
    JSON, scoped the same way tab_2/tab_7 themselves are."""
    ensure_schema()
    client = request.args.get("client") or ""
    category = request.args.get("category") or ""
    if not client or category not in ("Warranty", "Maintenance"):
        return jsonify({"success": False, "error": "client and category (Warranty or Maintenance) are required"}), 400
    try:
        data = build_ticket_report_data(client, category)
    except Exception as e:
        log(f"Ticket report data error: {e}", "ERROR")
        return jsonify({"success": False, "error": str(e)}), 500
    if data is None:
        return jsonify({"success": False, "error": f"No {category} data found for {client}"}), 404
    return jsonify({"success": True, "data": data})


def require_admin():
    return session.get("role") == "admin"


def require_cron_or_admin():
    """Admin session (manual "Sync now" button) OR Vercel Cron's own auth.

    Vercel Cron calls the endpoint with no session cookie, but -- as long
    as the CRON_SECRET env var is set on the project -- automatically adds
    `Authorization: Bearer <CRON_SECRET>` to the request, so that's what
    authenticates the scheduled path instead.
    """
    if require_admin():
        return True
    secret = os.environ.get("CRON_SECRET")
    if not secret:
        return False
    return request.headers.get("Authorization") == f"Bearer {secret}"


@app.route("/api/login", methods=["POST"])
def api_login():
    ensure_schema()
    data = request.get_json(silent=True) or request.form
    username = (data.get("username") or "").strip()
    password = data.get("password") or ""

    try:
        valid = db.verify_admin_credentials(username, password, conn=request_conn())
    except Exception as e:
        log(f"DB error checking admin credentials: {e}", "ERROR")
        return jsonify({"success": False, "error": "Login is temporarily unavailable"}), 503

    if valid:
        session["role"] = "admin"
        session.permanent = True
        return jsonify({"success": True, "role": "admin"})

    return jsonify({"success": False, "error": "Incorrect username or password"}), 401


@app.route("/api/login-viewer", methods=["POST"])
def api_login_viewer():
    session["role"] = "viewer"
    session.permanent = True
    return jsonify({"success": True, "role": "viewer"})


@app.route("/api/logout", methods=["POST"])
def api_logout():
    session.clear()
    return jsonify({"success": True})


@app.route("/api/upload", methods=["POST"])
def api_upload():
    if not require_admin():
        return jsonify({"success": False, "error": "Admin login required"}), 403
    ensure_schema()

    files = request.files.getlist("file")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"success": False, "error": "No file selected"}), 400

    form_client = request.form.get("client", "").strip()

    summary = {"files": [], "tickets_inserted": 0, "tickets_updated": 0,
               "projects_inserted": 0, "projects_updated": 0,
               "clients_inserted": 0, "clients_updated": 0,
               "milestones_inserted": 0, "milestones_updated": 0,
               "task_details_inserted": 0, "task_details_updated": 0,
               "errors": [], "rows_dropped": 0, "unmapped_columns": [], "notes": []}
    seen_unmapped = set()

    def note_diagnostics(diag):
        summary["rows_dropped"] += diag["rows_dropped"]
        for col in diag["unmapped_columns"]:
            if col not in seen_unmapped:
                seen_unmapped.add(col)
                summary["unmapped_columns"].append(col)

    def upsert_tickets_safely(parsed, label):
        """Commit each sheet's tickets as its own mini-transaction.

        The whole upload shares one connection (request_conn()) for
        speed, but that means an uncaught DB error -- e.g. a bad row
        that slipped past validation -- leaves the connection's
        transaction poisoned: every later statement on it fails too
        until rolled back, and a rollback with no earlier commit would
        also undo every *other* sheet already upserted in this same
        request. Committing per sheet on success and rolling back only
        that sheet on failure keeps one bad sheet from taking every
        other sheet in the file down with it.
        """
        conn = request_conn()
        try:
            ins, upd = db.upsert_tickets(parsed, conn=conn)
            conn.commit()
            summary["tickets_inserted"] += ins
            summary["tickets_updated"] += upd
            return True
        except Exception as e:
            conn.rollback()
            log(f"Upload error on {label}: {e}", "ERROR")
            summary["errors"].append(f"{label}: {str(e)[:300]}")
            return False

    def upsert_projects_safely(parsed_p, label):
        conn = request_conn()
        try:
            ins_p, upd_p = db.upsert_projects(parsed_p, conn=conn)
            # A brand-new task row (e.g. this module gained tasks since the
            # last upload) lands with no sort_order of its own, which would
            # otherwise put it at the very end of the whole table instead
            # of next to the rest of its module -- see
            # renumber_projects_sort_order()'s docstring for why that
            # splits a module into two separate-looking groups on the
            # Project Details page.
            db.renumber_projects_sort_order(conn=conn)
            conn.commit()
            summary["projects_inserted"] += ins_p
            summary["projects_updated"] += upd_p
        except Exception as e:
            conn.rollback()
            log(f"Upload error on {label}: {e}", "ERROR")
            summary["errors"].append(f"{label}: {str(e)[:300]}")

    def upsert_clients_safely(parsed_c, label):
        conn = request_conn()
        try:
            ins_c, upd_c = db.upsert_clients(parsed_c, conn=conn)
            # The source Client sheet always says "Development" for rows it
            # knows about, so this upsert can drag an already-transferred
            # row back out of Warranty. Push it straight back.
            reasserted = db.reassert_active_transfers(conn=conn)
            conn.commit()
            summary["clients_inserted"] += ins_c
            summary["clients_updated"] += upd_c
            if reasserted:
                summary["notes"].append(
                    f"{label}: {reasserted} row(s) kept in Warranty from a previous automatic transfer"
                )
        except Exception as e:
            conn.rollback()
            log(f"Upload error on {label}: {e}", "ERROR")
            summary["errors"].append(f"{label}: {str(e)[:300]}")

    def upsert_milestones_safely(parsed_m, label):
        """Same per-sheet commit/rollback isolation as the ticket helper
        above -- one bad milestone row fails only its own sheet."""
        conn = request_conn()
        try:
            ins_m, upd_m = db.upsert_project_milestones(parsed_m, conn=conn)
            conn.commit()
            summary["milestones_inserted"] += ins_m
            summary["milestones_updated"] += upd_m
        except Exception as e:
            conn.rollback()
            log(f"Upload error on {label}: {e}", "ERROR")
            summary["errors"].append(f"{label}: {str(e)[:300]}")

    def upsert_task_details_safely(parsed_t, label):
        """Same per-sheet commit/rollback isolation as the other helpers
        above -- one bad task-detail row fails only its own sheet."""
        conn = request_conn()
        try:
            ins_t, upd_t = db.upsert_project_task_details(parsed_t, conn=conn)
            conn.commit()
            summary["task_details_inserted"] += ins_t
            summary["task_details_updated"] += upd_t
        except Exception as e:
            conn.rollback()
            log(f"Upload error on {label}: {e}", "ERROR")
            summary["errors"].append(f"{label}: {str(e)[:300]}")

    for f in files:
        fname = f.filename
        ext = os.path.splitext(fname)[1].lower()
        raw = f.read()
        # Reset per file: otherwise a workbook without a Client sheet would
        # reuse the previous file's client mapping for its TASK DETAIL sheet.
        parsed_c = None
        try:
            if ext == ".csv":
                df = pd.read_csv(io.BytesIO(raw))
                client = form_client or (df["Client"].iloc[0] if "Client" in df.columns and len(df) else os.path.splitext(fname)[0])
                parsed, diag = parse_ticket_sheet(df, client=client, source_file=fname)
                note_diagnostics(diag)
                if upsert_tickets_safely(parsed, fname):
                    summary["files"].append({"name": fname, "rows_found": len(parsed)})

            elif ext in (".xlsx", ".xls"):
                buf = io.BytesIO(raw)
                sheets = detect_ticket_sheets(buf)
                rows_found = 0
                for sheet_name, header_row in sheets.items():
                    buf.seek(0)
                    df = pd.read_excel(buf, sheet_name=sheet_name, header=header_row, engine="openpyxl")
                    parsed, diag = parse_ticket_sheet(df, client=sheet_name, source_file=fname)
                    note_diagnostics(diag)
                    if parsed.empty:
                        continue
                    if upsert_tickets_safely(parsed, f"{fname} / {sheet_name}"):
                        rows_found += len(parsed)

                if not sheets:
                    summary["errors"].append(
                        f"{fname}: no sheet had a recognizable 'Ticket No' column (checked header rows 1, 0 and 2)"
                    )

                buf.seek(0)
                xl = pd.ExcelFile(buf, engine="openpyxl")
                if "Client Project" in xl.sheet_names:
                    buf.seek(0)
                    pdf = pd.read_excel(buf, sheet_name="Client Project", header=0, engine="openpyxl")
                    parsed_p, diag_p = parse_project_sheet(pdf, source_file=fname)
                    note_diagnostics({"rows_dropped": 0, "unmapped_columns": diag_p["unmapped_columns"]})
                    if not parsed_p.empty:
                        upsert_projects_safely(parsed_p, f"{fname} / Client Project")

                if "Client" in xl.sheet_names:
                    buf.seek(0)
                    cdf = pd.read_excel(buf, sheet_name="Client", header=0, engine="openpyxl")
                    parsed_c, diag_c = parse_client_sheet(cdf, source_file=fname)
                    note_diagnostics({"rows_dropped": diag_c["rows_dropped"], "unmapped_columns": diag_c["unmapped_columns"]})
                    if not parsed_c.empty:
                        upsert_clients_safely(parsed_c, f"{fname} / Client")

                # Matched case-insensitively so a re-titled sheet
                # ("Project Milestone", trailing space) still lands -- the
                # same way the other two sheets are looked up by name, but
                # without being locked to the source workbook's shouting.
                milestone_sheet = next(
                    (s for s in xl.sheet_names if str(s).strip().upper() == "PROJECT MILESTONE"),
                    None,
                )
                if milestone_sheet:
                    buf.seek(0)
                    mdf = pd.read_excel(buf, sheet_name=milestone_sheet, header=0, engine="openpyxl")
                    parsed_m, diag_m = parse_milestone_sheet(mdf, source_file=fname)
                    note_diagnostics(diag_m)
                    if not parsed_m.empty:
                        upsert_milestones_safely(parsed_m, f"{fname} / {milestone_sheet}")

                # "<project> TASK DETAIL": the sheet has no Client/Projek Name
                # column, so both are derived -- projek_name by stripping the
                # suffix, client by looking the project up in the Client sheet
                # of this same file (the clients table as fallback). Matched
                # case-insensitively so a re-titled sheet still lands, the same
                # way PROJECT MILESTONE is looked up above.
                task_sheet = next(
                    (s for s in xl.sheet_names if str(s).strip().upper().endswith(" TASK DETAIL")),
                    None,
                )
                if task_sheet:
                    projek_name = re.sub(
                        r"\s+TASK DETAIL\s*$", "", str(task_sheet).strip(), flags=re.IGNORECASE
                    ).strip()
                    client = lookup_client_for_project(projek_name, parsed_c)
                    if not client:
                        summary["notes"].append(
                            f'{fname} / {task_sheet}: no client found for project "{projek_name}" -- '
                            "add the project to the Client sheet and re-upload"
                        )
                    else:
                        buf.seek(0)
                        tdf = pd.read_excel(buf, sheet_name=task_sheet, header=0, engine="openpyxl")
                        parsed_t, diag_t = parse_task_detail_sheet(tdf, fname, client, projek_name)
                        note_diagnostics(diag_t)
                        if not parsed_t.empty:
                            upsert_task_details_safely(parsed_t, f"{fname} / {task_sheet}")

                summary["files"].append({"name": fname, "rows_found": rows_found})

            else:
                summary["errors"].append(f"{fname}: unsupported file type (use .csv or .xlsx)")

        except Exception as e:
            log(f"Upload error on {fname}: {e}", "ERROR")
            summary["errors"].append(f"{fname}: {str(e)[:300]}")
            try:
                request_conn().rollback()
            except Exception:
                pass

    summary["success"] = len(summary["errors"]) == 0
    return jsonify(summary)


@app.route("/api/restart", methods=["POST"])
def api_restart():
    if not require_admin():
        return jsonify({"success": False, "error": "Admin login required"}), 403
    ensure_schema()
    try:
        db.reset_all(conn=request_conn())
        return jsonify({"success": True})
    except Exception as e:
        log(f"Restart error: {e}", "ERROR")
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/sync_mysupport", methods=["GET", "POST"])
def api_sync_mysupport():
    """Pull tickets/projects/clients from the live mysupport MySQL DB and
    merge them into Postgres. Only overwrites the columns mysupport has
    data for (see mysupport_sync.*_SYNC_COLUMNS) -- Priority, SLA fields,
    Progress %, planned/actual dates, Assigned To, etc. are left exactly
    as they are, since mysupport has no equivalent for those.

    Triggered either by the "Sync from mysupport" button (admin session)
    or by the Vercel Cron schedule (see vercel.json), which is why this
    accepts GET too and checks require_cron_or_admin() instead of just
    require_admin().
    """
    if not require_cron_or_admin():
        return jsonify({"success": False, "error": "Admin login or cron secret required"}), 403
    ensure_schema()

    summary = {"tickets_inserted": 0, "tickets_updated": 0,
               "projects_inserted": 0, "projects_updated": 0,
               "clients_inserted": 0, "clients_updated": 0, "errors": []}

    try:
        mysupport_conn = mysupport_sync.get_mysupport_conn()
    except Exception as e:
        log(f"mysupport sync: connection failed: {e}", "ERROR")
        return jsonify({"success": False, "error": f"Could not connect to mysupport: {e}"}), 502

    try:
        try:
            tickets_df = mysupport_sync.fetch_mysupport_tickets_df(conn=mysupport_conn)
            ins, upd = db.upsert_tickets(
                tickets_df, conn=request_conn(), sync_columns=mysupport_sync.TICKET_SYNC_COLUMNS,
            )
            request_conn().commit()
            summary["tickets_inserted"] += ins
            summary["tickets_updated"] += upd
        except Exception as e:
            request_conn().rollback()
            log(f"mysupport sync: tickets failed: {e}", "ERROR")
            summary["errors"].append(f"tickets: {str(e)[:300]}")

        # NOTE: syncing into Postgres `projects` is intentionally disabled --
        # see the comment on mysupport_sync.fetch_mysupport_projects_df().
        # Postgres `projects` holds one row per *module/task line* (title,
        # description, plan/target/actual dates), which mysupport's
        # `projects` table (one row per project, no dates/description) does
        # not map onto without picking a source for those rows (tasks?
        # progress?) that hasn't been decided yet.

        # NOTE: syncing into Postgres `clients` is also intentionally
        # disabled -- see the comment on mysupport_sync.fetch_mysupport_clients_df().
        # It's a small, manually-curated set of Development/Warranty/
        # Maintenance engagement rows per client, not a raw project catalog;
        # this used to insert one row per mysupport project (15+ per client)
        # and every one of them displayed that client's *entire* ticket
        # total on the Home page, making totals look wildly inflated.
    finally:
        mysupport_conn.close()

    summary["success"] = len(summary["errors"]) == 0
    return jsonify(summary)


@app.route("/api/status")
def api_status():
    ensure_schema()
    try:
        return jsonify({"success": True, **db.get_counts(conn=request_conn())})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/save", methods=["POST"])
def api_save():
    if not require_admin():
        return jsonify({"success": False, "error": "Admin login required"}), 403
    data = request.get_json()
    row_idx = data.get("row_idx")
    column = data.get("column")
    value = data.get("value")
    sheet = data.get("sheet")

    if sheet == "Client":
        db_column = CLIENT_DB_COL_BY_DISPLAY.get(column)
        if not db_column:
            return {"success": False, "error": f"Column not editable: {column}"}
        try:
            db.update_client_field(int(row_idx), db_column, value, conn=request_conn())
            return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    if sheet == "Milestone":
        db_column = MILESTONE_DB_COL_BY_DISPLAY.get(column)
        if not db_column:
            return {"success": False, "error": f"Column not editable: {column}"}
        try:
            db.update_project_milestone_field(int(row_idx), db_column, value, conn=request_conn())
            return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    if sheet == "Client Project":
        db_column = PROJECT_DB_COL_BY_DISPLAY.get(column)
        if not db_column:
            return {"success": False, "error": f"Column not editable: {column}"}
        try:
            db.update_project_field(int(row_idx), db_column, value, conn=request_conn())
            return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    if sheet == "Task Detail":
        # Must come before the ticket fall-through below: an unmatched sheet
        # would otherwise be treated as a ticket table and silently miss.
        db_column = TASK_DETAIL_DB_COL_BY_DISPLAY.get(column)
        if not db_column:
            return {"success": False, "error": f"Column not editable: {column}"}
        try:
            db.update_project_task_detail_field(int(row_idx), db_column, value, conn=request_conn())
            return {"success": True}
        except Exception as e:
            # Roll back here rather than leaving it to teardown: a rejected
            # UPDATE (e.g. a percentage out of range) aborts the transaction,
            # and the next statement on this shared connection would fail.
            request_conn().rollback()
            return {"success": False, "error": str(e)}

    db_column = TICKET_DB_COL_BY_DISPLAY.get(column)
    if not db_column:
        return {"success": False, "error": f"Column not editable: {column}"}

    try:
        db.update_ticket_field(int(row_idx), db_column, value, conn=request_conn())
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def _task_detail_scope():
    """(clients, projek_name) the Task Detail tab is filtered to -- taken
    from the query string both Excel routes are called with, so an export
    never contains another project's rows."""
    filters = parse_filters(request.args)
    return filters["clients"] or None, filters.get("projek_name")


def _task_detail_export_values(row, headers):
    """One dataframe row -> the cell values the source sheet stores.

    Percentage is written as an 0-1 fraction with a 0% number format, which
    is how Excel's own percentage format represents 50% -- matching the
    source means a re-upload of the file round-trips through
    data_utils._scale_percentage back to 50 instead of 0.5.
    """
    values = []
    for c in headers:
        v = row.get(c) if hasattr(row, "get") else row[c]
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            values.append(None)
        elif c == "No":
            values.append(int(v))
        elif c == "Percentage":
            values.append(float(v) / 100.0)
        else:
            values.append(v)
    return values


def _style_task_detail_header(ws, headers):
    for i, header in enumerate(headers, start=1):
        ws.cell(row=1, column=i, value=header).font = Font(bold=True)


def _fill_task_detail_header(ws, headers):
    """Keep the sheet's own header text on row 1, filling only blanks.

    Used by Save to Excel: the source header is authoritative (it may label
    things differently than the app does), so a write-back must not rename
    columns out from under a human's existing layout.
    """
    for i, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=i)
        if cell.value in (None, ""):
            cell.value = header
        cell.font = Font(bold=True)


def _style_task_detail_row(ws, row_number, headers, date_format="DD/MM/YYYY"):
    """Number formats for the row just written: a bare 0.5 should read as
    50%, and a date cell should read the way the source/target formats it.

    `date_format` differs per caller on purpose: the download gets the
    unambiguous format the tab shows, while Save to Excel keeps the source
    sheet's own `mm-dd-yy` so a write-back doesn't reformat someone's file.
    """
    for i, header in enumerate(headers, start=1):
        cell = ws.cell(row=row_number, column=i)
        if header == "Percentage" and isinstance(cell.value, (int, float)):
            cell.number_format = "0.00%"
        elif header == "Target Date" and cell.value is not None:
            cell.number_format = date_format
        elif header == "Definisi":
            cell.alignment = Alignment(wrap_text=True, vertical="top")


@app.route("/api/task_detail/export")
def api_task_detail_export():
    """Download the scoped task-detail rows as a fresh .xlsx.

    Unlike save_to_excel this never touches a file on disk, so it works
    anywhere (including Vercel) and can't collide with a workbook someone
    has open in Excel.
    """
    ensure_schema()
    clients, projek_name = _task_detail_scope()
    headers = db.TASK_DETAIL_COLUMNS
    try:
        df = db.fetch_task_detail_for_export(
            clients=clients, projek_name=projek_name, conn=request_conn()
        )
    except Exception as e:
        log(f"Task detail export error: {e}", "ERROR")
        return jsonify({"success": False, "error": str(e)}), 500
    if df.empty:
        return jsonify({"success": False, "error": "No task detail rows to export"}), 404

    wb = Workbook()
    ws = wb.active
    # Excel caps sheet names at 31 chars and bans :\\/?*[]. Keep the source
    # sheet's "TASK DETAIL" suffix so a re-upload of the download matches.
    ws.title = re.sub(r"[:\\/?*\[\]]", " ", f"{projek_name or 'Task Detail'} TASK DETAIL")[:31]
    _style_task_detail_header(ws, headers)
    for _, row in df.iterrows():
        ws.append(_task_detail_export_values(row, headers))
        _style_task_detail_row(ws, ws.max_row, headers)
    for i, header in enumerate(headers, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = max(12, len(header) + 4)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    label = (projek_name or "task-detail").strip() or "task-detail"
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{label} TASK DETAIL.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/task_detail/save_to_excel", methods=["POST"])
def api_task_detail_save_to_excel():
    """Write the current task-detail rows back into their source sheet.

    The database is the source of truth after an edit, so this rewrites the
    sheet's data region wholesale rather than patching cells: module groups
    are rebuilt from the rows (with the blank separator line the source has
    between modules), column A is re-merged across each multi-row group, and
    everything above row 1 -- header, column widths, the other nine sheets --
    is left alone.
    """
    if not require_admin():
        return jsonify({"success": False, "error": "Admin login required"}), 403
    ensure_schema()

    clients, projek_name = _task_detail_scope()
    headers = db.TASK_DETAIL_COLUMNS
    try:
        df = db.fetch_task_detail_for_export(
            clients=clients, projek_name=projek_name, conn=request_conn()
        )
    except Exception as e:
        log(f"Task detail load error: {e}", "ERROR")
        return jsonify({"success": False, "error": str(e)}), 500
    if df.empty:
        return jsonify({"success": False, "error": "No task detail rows to write back"}), 404

    source_file = db.primary_task_detail_source_file(
        clients=clients, projek_name=projek_name, conn=request_conn()
    )
    path = resolve_source_workbook(source_file)
    if not path:
        return jsonify({
            "success": False,
            "error": f'Cannot find the source workbook "{source_file or "unknown"}" on this server',
        }), 404

    try:
        wb = load_workbook(path)
    except Exception as e:
        log(f"Task detail workbook open error: {e}", "ERROR")
        return jsonify({"success": False, "error": f"Cannot open the workbook: {e}"}), 400

    sheet_title = next(
        (s for s in wb.sheetnames if str(s).strip().upper().endswith(" TASK DETAIL")),
        None,
    )
    if not sheet_title:
        return jsonify({
            "success": False,
            "error": f'No sheet ending in "TASK DETAIL" inside {os.path.basename(path)}',
        }), 400
    ws = wb[sheet_title]

    # Rebuild column A's merges: unmerge first (delete_rows does not move
    # merged ranges), drop the old data region, then re-merge each module's
    # block once its rows are back on the sheet.
    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    # Leave row 1 alone. The source calls this column "% Progress" while the
    # app calls it "Percentage"; rewriting the header would mutate a file the
    # user only asked us to put data back into (and the parser accepts both).
    _fill_task_detail_header(ws, headers)

    groups = []
    for _, row in df.iterrows():
        key = (row.get("No"), row.get("Modul"))
        if not groups or groups[-1]["key"] != key:
            groups.append({"key": key, "rows": []})
        groups[-1]["rows"].append(row)

    row_number = 2
    for position, group in enumerate(groups):
        start = row_number
        for row in group["rows"]:
            for column, value in enumerate(_task_detail_export_values(row, headers), start=1):
                ws.cell(row=row_number, column=column, value=value)
            _style_task_detail_row(ws, row_number, headers, date_format="mm-dd-yy")
            row_number += 1
        if row_number - 1 > start:
            ws.merge_cells(start_row=start, start_column=1, end_row=row_number - 1, end_column=1)
        if position < len(groups) - 1:
            row_number += 1  # blank line between modules, as the sheet has

    try:
        wb.save(path)
    except PermissionError:
        return jsonify({
            "success": False,
            "error": "The workbook is open in Excel -- close it and try again",
        }), 400
    except Exception as e:
        log(f"Task detail workbook save error: {e}", "ERROR")
        return jsonify({"success": False, "error": f"Cannot write the workbook: {e}"}), 400

    log(f"Task detail written back to {os.path.basename(path)} / {sheet_title} ({len(df)} rows)")
    return jsonify({
        "success": True,
        "rows": len(df),
        "file": os.path.basename(path),
        "sheet": sheet_title,
    })


@app.route("/api/add_row", methods=["POST"])
def api_add_row():
    if not require_admin():
        return jsonify({"success": False, "error": "Admin login required"}), 403
    data = request.get_json()
    table = data.get("table")
    values = data.get("values") or {}

    col_by_display, insert_fn = {
        "clients": (CLIENT_DB_COL_BY_DISPLAY, db.insert_client_row),
        "projects": (PROJECT_DB_COL_BY_DISPLAY, db.insert_project_row),
        "tickets": (TICKET_DB_COL_BY_DISPLAY, db.insert_ticket_row),
        "milestones": (MILESTONE_DB_COL_BY_DISPLAY, db.insert_project_milestone_row),
    }.get(table, (None, None))
    if not insert_fn:
        return jsonify({"success": False, "error": f"Unknown table: {table}"}), 400

    db_values = {}
    for display, val in values.items():
        db_col = col_by_display.get(display)
        if not db_col:
            return jsonify({"success": False, "error": f"Column not editable: {display}"}), 400
        db_values[db_col] = val

    try:
        new_id = insert_fn(db_values, conn=request_conn())
        return jsonify({"success": True, "row_idx": new_id})
    except ValueError as e:
        return jsonify({"success": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400


@app.route("/api/reorder_rows", methods=["POST"])
def api_reorder_rows():
    if not require_admin():
        return jsonify({"success": False, "error": "Admin login required"}), 403
    data = request.get_json()
    table = data.get("table")
    ids = data.get("ids") or []

    if table != "projects":
        return jsonify({"success": False, "error": f"Reordering not supported for: {table}"}), 400

    try:
        db.reorder_project_rows(ids, conn=request_conn())
        return jsonify({"success": True})
    except ValueError as e:
        return jsonify({"success": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400


@app.route("/api/delete_row", methods=["POST"])
def api_delete_row():
    if not require_admin():
        return jsonify({"success": False, "error": "Admin login required"}), 403
    data = request.get_json()
    table = data.get("table")
    row_idx = data.get("row_idx")

    delete_fn = {
        "projects": db.delete_project_row,
        "tickets": db.delete_ticket_row,
        "clients": db.delete_client_row,
        "milestones": db.delete_project_milestone_row,
    }.get(table)
    if not delete_fn:
        return jsonify({"success": False, "error": f"Unknown table: {table}"}), 400

    try:
        deleted = delete_fn(int(row_idx), conn=request_conn())
        return jsonify({"success": True, "deleted": deleted})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400


@app.route("/api/revert_transfer", methods=["POST"])
def api_revert_transfer():
    """Undo an automatic Development -> Warranty transfer from the Home panel."""
    if not require_admin():
        return jsonify({"success": False, "error": "Admin login required"}), 403
    data = request.get_json() or {}
    history_id = data.get("history_id")
    try:
        result = db.revert_transfer(int(history_id), conn=request_conn())
    except (TypeError, ValueError):
        return jsonify({"success": False, "error": "Missing or invalid history_id"}), 400
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400
    if not result["success"]:
        return jsonify(result), 404
    log(f"Reverted transfer #{history_id} (row restored: {result['row_restored']})")
    return jsonify(result)


@app.route("/api/dismiss_transfer", methods=["POST"])
def api_dismiss_transfer():
    """Proceed on a Home panel transfer: keep it, stop listing it."""
    if not require_admin():
        return jsonify({"success": False, "error": "Admin login required"}), 403
    data = request.get_json() or {}
    history_id = data.get("history_id")
    try:
        result = db.dismiss_transfer(int(history_id), conn=request_conn())
    except (TypeError, ValueError):
        return jsonify({"success": False, "error": "Missing or invalid history_id"}), 400
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400
    if not result["success"]:
        return jsonify(result), 404
    log(f"Dismissed transfer #{history_id}")
    return jsonify(result)


if __name__ == "__main__":
    app.run(debug=True, port=8501)
