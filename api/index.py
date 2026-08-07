import io
import os
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
from flask import Flask, render_template, request, jsonify, send_from_directory, g, session

load_dotenv()
load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".env.local"), override=True)

import db
from data_utils import (
    COLORS, PRIORITY_COLORS, AGEING_COLORS,
    parse_ticket_sheet, parse_project_sheet, parse_client_sheet, detect_ticket_sheets,
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

PROJECT_ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")


@app.route("/sw.svg")
def brand_watermark():
    return send_from_directory(PROJECT_ROOT, "sw.svg", mimetype="image/svg+xml")


@app.route("/manifest.json")
def pwa_manifest():
    return send_from_directory(PROJECT_ROOT, "manifest.json", mimetype="application/manifest+json")


@app.route("/sw.js")
def pwa_service_worker():
    # Served from the root path (not /static/sw.js) so its default scope
    # covers the whole origin instead of just /static/.
    return send_from_directory(PROJECT_ROOT, "sw.js", mimetype="application/javascript")


@app.route("/static/<path:filename>")
def static_files(filename):
    return send_from_directory(os.path.join(PROJECT_ROOT, "static"), filename)


def log(msg, level="INFO"):
    print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {level} {msg}", flush=True)


log("=" * 50)
log("Dashboard starting (Flask + Neon Postgres)")
log(f"Python: {sys.version}")

pio.templates.default = "plotly_white"

TICKET_DB_COL_BY_DISPLAY = {display: col for display, col in db.TICKET_DB_COLUMNS}
CLIENT_DB_COL_BY_DISPLAY = {display: col for display, col in db.CLIENT_DB_COLUMNS}

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
    return filters


def load_data(filters=None):
    ensure_schema()
    df = db.fetch_tickets_df(filters, conn=request_conn())
    return df, []


def load_project_data():
    ensure_schema()
    return db.fetch_projects_df(conn=request_conn())


def load_client_data():
    ensure_schema()
    return db.fetch_clients_df(conn=request_conn())


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
        charts["task_type_bar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Project" in warranty_df.columns:
        pc = warranty_df["Project"].value_counts().reset_index()
        pc.columns = ["Project", "Count"]
        fig = px.bar(pc, x="Project", y="Count", title="Warranty Tickets by Project",
                      color="Project", text="Count")
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), showlegend=False, xaxis_tickangle=-45)
        charts["project_bar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    display_cols = ["Ticket No", "Task Type", "Project", "Company", "Ticket Title", "Priority", "Ticket Status", "Ticket Created Date", "Days"]
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

    overall_by_group = df.groupby(["Client", "Title"]).apply(lambda g: group_overall(g.index), include_groups=False)
    df["Overall Progress Task (%)"] = df.set_index(["Client", "Title"]).index.map(overall_by_group).to_numpy()
    return df


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

    if "Client" in df.columns:
        valid_clients = df.dropna(subset=["Client"])
        if valid_clients.empty:
            return charts
        cc = valid_clients["Client"].value_counts().reset_index()
        cc.columns = ["Client", "Count"]
        fig = px.bar(cc, x="Client", y="Count", title="Projects by Client",
                      color="Client", text="Count")
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), showlegend=False)
        charts["client_bar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Status Progress" in df.columns:
        valid_status = df.dropna(subset=["Status Progress"])
        if valid_status.empty:
            return charts
        sc = valid_status["Status Progress"].value_counts().reset_index()
        sc.columns = ["Status", "Count"]
        fig = px.pie(sc, names="Status", values="Count", title="Project Status Progress",
                      hole=0.3)
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
        charts["status_pie"] = fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})

    if "Start date" in df.columns and "Due date" in df.columns and "Title" in df.columns:
        valid = df.dropna(subset=["Start date", "Due date", "Title"]).copy()
        valid = valid[valid["Title"].astype(str).str.strip() != ""]
        if "Description" in valid.columns:
            valid["Task Label"] = valid["Description"].astype(str)
            valid["Task Label"] = valid["Task Label"].str.replace(r"^\d+\.\s*", "", regex=True)
            valid["Task Label"] = valid["Task Label"].str.split("\n").str[0].str.strip()
        else:
            valid["Task Label"] = valid["Title"].astype(str)
        if not valid.empty:
            timeline_charts_html = ""
            if "Client" in valid.columns:
                for client in sorted(valid["Client"].dropna().unique()):
                    cdf = valid[valid["Client"] == client]
                    if cdf.empty:
                        continue
                    fig = px.timeline(
                        cdf, x_start="Start date", x_end="Due date",
                        y="Task Label", color="Client",
                        title=f"{client} - PROJECT DEVELOPMENT TIMELINE",
                        color_discrete_sequence=px.colors.qualitative.Plotly,
                    )
                    fig.update_yaxes(autorange="reversed", title=None)
                    fig.update_xaxes(title="Tarikh")
                    fig.update_layout(
                        template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                        font=dict(color="#374151"), showlegend=False,
                        height=max(200, 30*len(cdf)),
                    )
                    timeline_charts_html += f'<div class="client-section"><h4>{client}</h4>{fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})}</div>'
            charts["timeline_chart"] = timeline_charts_html

    display_cols_p = ["Client", "Title", "Description", "Category", "Progress", "Priority", "Start date", "Due date", "Target Date", "Duration", "Assigned to", "Status Progress", "Percentage", "Overall Progress Task (%)"]
    avail_p = [c for c in display_cols_p if c in df.columns]
    meta_p = [c for c in ["_row_idx", "_source_file"] if c in df.columns]
    detail = df[avail_p + meta_p].copy()
    for c in ["Start date", "Due date", "Target Date"]:
        if c in detail.columns:
            detail[c] = detail[c].dt.strftime("%d/%m/%Y") if not detail[c].isna().all() else detail[c]
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


def build_overall_client_charts(df):
    charts = {}
    if df.empty:
        return charts

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

    display_cols = ["Client", "Projek ID", "Projek Name", "Projek Status", "Start Date", "End Date"]
    avail = [c for c in display_cols if c in df.columns]
    meta_cols = [c for c in ["_row_idx", "Source File"] if c in df.columns]
    detail = df[avail + meta_cols].copy()
    for c in ["Start Date", "End Date"]:
        if c in detail.columns and not detail[c].isna().all():
            detail[c] = detail[c].dt.strftime("%d/%m/%Y")
    detail = detail.fillna("")
    charts["detail_data"] = detail.to_dict("records")

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
                    "rows": sdf.to_dict("records"),
                }
        for status in sorted(set(statuses) - set(status_order), key=lambda s: str(s).lower()):
            sdf = detail[detail["Projek Status"] == status]
            if sdf.empty:
                continue
            charts["status_sections"][status] = {
                "count": int(len(sdf)),
                "rows": sdf.to_dict("records"),
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


@app.route("/")
def index():
    user_role = session.get("role")
    filters = parse_filters(request.args)
    try:
        df, load_errors = load_data(filters)
    except Exception as e:
        log(f"DB error loading tickets: {e}", "ERROR")
        df, load_errors = pd.DataFrame(), [str(e)]

    # Warranty tickets are stored under a shared "Client Warranty" sentinel
    # in the tickets table (not the real client name) -- the real client is
    # only recorded in the Company column. So filtering tickets by the real
    # client name (as every other category does) always returns zero rows
    # for Warranty. Detect that case here and fetch a Warranty-appropriate
    # dataframe instead, scoped by Company rather than Client.
    client_category = filters["task_types"][0] if len(filters["task_types"]) == 1 else None
    warranty_client = filters["clients"][0] if (len(filters["clients"]) == 1 and client_category == "Warranty") else None

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
    filtered_has_data = not df.empty

    data_info = {
        "total_raw": counts.get("tickets", len(df)),
        "total_filtered": len(df),
        "load_errors": load_errors,
        "columns": list(df.columns) if not df.empty else [],
        "counts": counts,
    }

    try:
        project_df = load_project_data()
    except Exception as e:
        log(f"DB error loading projects: {e}", "ERROR")
        project_df = pd.DataFrame()
    # The projects table is independent of the tickets table (no shared
    # filter query), so a client selected via the sidebar/Overall Client
    # table has to be applied here explicitly to scope the Project tab to
    # that client.
    if filters["clients"] and "Client" in project_df.columns:
        project_df = project_df[project_df["Client"].isin(filters["clients"])]
    project_df = recompute_overall_progress(project_df)
    has_project = not project_df.empty

    try:
        client_df = load_client_data()
    except Exception as e:
        log(f"DB error loading clients: {e}", "ERROR")
        client_df = pd.DataFrame()
    has_client = not client_df.empty

    if filtered_has_data:
        overview_charts = build_charts(df)
        priority_charts = build_priority_charts(df)
        ageing_charts = build_ageing_charts(df)
        comparison_charts = build_client_comparison_charts(df)
        timeline_charts = build_timeline_charts(df)
        sla_charts = build_sla_charts(df)
    else:
        overview_charts = priority_charts = {}
        comparison_charts = timeline_charts = sla_charts = {}
        # Matches the shape build_ageing_charts() always returns (even for
        # a genuinely empty df) -- the template unconditionally iterates
        # ageing_charts.ageing_clients.items(), so a bare {} here would
        # raise UndefinedError instead of just rendering zero rows.
        ageing_charts = {"age_order": ["1-30 Days", "31-60 Days", "> 60 Days"], "ageing_clients": {}}

    # Warranty is scoped by Company, not the tickets-table Client filter
    # (see the warranty_client comment above), so it's built independently
    # of filtered_has_data just like the Project tab.
    if warranty_client:
        try:
            warranty_df, _ = load_data({"clients": ["Client Warranty"], "priorities": [], "statuses": [], "task_types": [], "search": None})
        except Exception as e:
            log(f"DB error loading warranty tickets: {e}", "ERROR")
            warranty_df = pd.DataFrame()
        if "Company" in warranty_df.columns:
            warranty_df = warranty_df[warranty_df["Company"] == warranty_client]
        warranty_charts = build_warranty_charts(warranty_df)
    elif filtered_has_data:
        warranty_charts = build_warranty_charts(df)
    else:
        warranty_charts = {}

    # The Project tab is driven by project_df, not the tickets df, so it
    # must not be blanked out just because the current ticket filter (e.g.
    # a client viewed under their Development category, which has no
    # matching Task Type on tickets) happens to match zero tickets.
    project_charts = build_project_charts(project_df) if has_project else {}

    overall_client_charts = build_overall_client_charts(client_df) if has_client else {}

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

    if "Ticket Created Date" in detail_df.columns:
        detail_df["Ticket Created Date"] = detail_df["Ticket Created Date"].dt.strftime("%d/%m/%Y")
    if "Ticket Completed Date" in detail_df.columns:
        detail_df["Ticket Completed Date"] = detail_df["Ticket Completed Date"].dt.strftime("%d/%m/%Y")
    if "Ticket Closed Date" in detail_df.columns:
        detail_df["Ticket Closed Date"] = detail_df["Ticket Closed Date"].dt.strftime("%d/%m/%Y")

    detail_data = detail_df.to_dict("records") if filtered_has_data and not detail_df.empty else []
    detail_by_client = {}
    for row in detail_data:
        client = row.get("Client", "Unknown")
        detail_by_client.setdefault(client, []).append(row)

    ageing_list_data = {}
    if filtered_has_data and "Ageing" in df.columns and df["Ageing"].notna().sum() > 0:
        age_order = ["1-30 Days", "31-60 Days", "> 60 Days"]
        ageing_cols = [c for c in ["Client", "Ticket No", "Ticket Title", "Ticket Status", "Priority", "Ticket Created Date", "Days", "_row_idx", "Source File"] if c in df.columns]
        ageing_df = df.dropna(subset=["Ageing"]).copy()
        if "Ticket Created Date" in ageing_df.columns:
            ageing_df["Ticket Created Date"] = ageing_df["Ticket Created Date"].dt.strftime("%d/%m/%Y")
        ageing_list_data = {
            "total": int(ageing_df["Ageing"].notna().sum()),
            "buckets": {},
        }
        for bucket in age_order:
            bucket_df = ageing_df[ageing_df["Ageing"] == bucket]
            if bucket_df.empty:
                continue
            clients_in_bucket = {}
            for client in sorted(bucket_df["Client"].unique()):
                client_df = bucket_df[bucket_df["Client"] == client]
                clients_in_bucket[client] = {
                    "count": len(client_df),
                    "rows": client_df[ageing_cols].to_dict("records"),
                }
            ageing_list_data["buckets"][bucket] = clients_in_bucket

    return render_template(
        "dashboard.html",
        user_role=user_role,
        has_data=has_data,
        filtered_has_data=filtered_has_data,
        has_project=has_project,
        has_client=has_client,
        data_info=data_info,
        filter_options=filter_options,
        overview_charts=overview_charts,
        priority_charts=priority_charts,
        ageing_charts=ageing_charts,
        comparison_charts=comparison_charts,
        timeline_charts=timeline_charts,
        sla_charts=sla_charts,
        warranty_charts=warranty_charts,
        project_charts=project_charts,
        overall_client_charts=overall_client_charts,
        project_data=project_df.to_dict("records") if has_project else [],
        detail_by_client=detail_by_client,
        ageing_list_data=ageing_list_data,
        now=datetime.now().strftime("%d-%m-%Y %H:%M"),
    )


def require_admin():
    return session.get("role") == "admin"


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
               "clients_inserted": 0, "clients_updated": 0, "errors": [],
               "rows_dropped": 0, "unmapped_columns": []}
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
            conn.commit()
            summary["clients_inserted"] += ins_c
            summary["clients_updated"] += upd_c
        except Exception as e:
            conn.rollback()
            log(f"Upload error on {label}: {e}", "ERROR")
            summary["errors"].append(f"{label}: {str(e)[:300]}")

    for f in files:
        fname = f.filename
        ext = os.path.splitext(fname)[1].lower()
        raw = f.read()
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

    db_column = TICKET_DB_COL_BY_DISPLAY.get(column)
    if not db_column:
        return {"success": False, "error": f"Column not editable: {column}"}

    try:
        db.update_ticket_field(int(row_idx), db_column, value, conn=request_conn())
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


if __name__ == "__main__":
    app.run(debug=True, port=8501)
