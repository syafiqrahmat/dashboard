"""Read-only layer for the mysupport MySQL support/ticketing system.

mysupport is the source of truth for ticket/project identity and status;
this module only ever SELECTs from it. Every DataFrame returned here uses
the same *display-name* columns as db.TICKET_DB_COLUMNS / PROJECT_DB_COLUMNS
/ CLIENT_DB_COLUMNS, but only the subset mysupport actually has data for --
callers pass a matching `sync_columns` list to db.upsert_*() so fields that
only exist on the Postgres side (Priority, SLA_*, Progress, planned/actual
dates, Assigned To, ...) are never touched by a sync.
"""
import os

import pandas as pd
import pymysql

TICKET_STATUS_BY_CODE = {
    1: "Pending", 2: "Inprogress", 3: "Completed",
    4: "Closed", 5: "KIV", 6: "Deleted",
}
TICKET_CATEGORY_BY_CODE = {1: "Issue", 2: "Change Request"}
PROJECT_STATUS_BY_CODE = {1: "Active", 0: "Inactive"}


def get_mysupport_conn():
    host = os.environ.get("MYSUPPORT_DB_HOST")
    if not host:
        raise RuntimeError("MYSUPPORT_DB_HOST environment variable is not set")
    return pymysql.connect(
        host=host,
        port=int(os.environ.get("MYSUPPORT_DB_PORT", "3306")),
        database=os.environ.get("MYSUPPORT_DB_NAME", "mysupport"),
        user=os.environ.get("MYSUPPORT_DB_USER"),
        password=os.environ.get("MYSUPPORT_DB_PASSWORD"),
        connect_timeout=10,
        charset="utf8mb4",
    )


# Columns each fetch_* function actually populates -- passed straight
# through to db.upsert_*(sync_columns=...) by the /api/sync_mysupport route
# so nothing else on the Postgres row gets overwritten.
TICKET_SYNC_COLUMNS = [
    "client", "company", "project", "ticket_title", "ticket_detail",
    "ticket_category", "ticket_created_date", "ticket_completed_date",
    "ticket_closed_date", "ticket_status", "source_file",
    # "task_type" deliberately excluded -- fetch_mysupport_tickets_df()
    # still sets a "Maintenance" default in the DataFrame, but leaving it
    # out of the *update* set here means that default only ever applies
    # to a ticket's initial INSERT, never overwrites a value someone has
    # since hand-tagged on an existing row.
]
PROJECT_SYNC_COLUMNS = ["client", "title", "projek_name", "status_progress", "source_file"]
CLIENT_SYNC_COLUMNS = ["client", "projek_id", "projek_name", "projek_status", "source_file"]


def fetch_mysupport_tickets_df(conn=None):
    sql = """
        SELECT
            c.code AS `Client`,
            c.code AS `Company`,
            p.name AS `Project`,
            t.ticket_no AS `Ticket No`,
            t.title AS `Ticket Title`,
            t.detail AS `Ticket Detail`,
            t.category AS `_category_code`,
            t.created_at AS `Ticket Created Date`,
            t.completed_date AS `Ticket Completed Date`,
            t.closed_date AS `Ticket Closed Date`,
            t.overall_status AS `_status_code`
        FROM tickets t
        JOIN companies c ON t.company_id = c.id
        JOIN projects p ON t.project_id = p.id
    """
    owned = conn is None
    conn = conn or get_mysupport_conn()
    try:
        df = pd.read_sql(sql, conn)
    finally:
        if owned:
            conn.close()

    if df.empty:
        return df

    df["Ticket Category"] = df["_category_code"].map(TICKET_CATEGORY_BY_CODE)
    df["Ticket Status"] = df["_status_code"].map(TICKET_STATUS_BY_CODE)
    df["Source File"] = "mysupport-sync"
    # mysupport IS the maintenance-phase support system, so every ticket
    # coming through it defaults to Task Type = Maintenance. Only applied
    # to brand-new tickets, though -- "task_type" is deliberately left out
    # of TICKET_SYNC_COLUMNS below, so a ticket someone has already
    # hand-tagged (e.g. as "Daily") never gets silently overwritten back
    # to this default on a later sync.
    df["Task Type"] = "Maintenance"
    return df.drop(columns=["_category_code", "_status_code"])


def fetch_mysupport_projects_df(conn=None):
    """NOT wired into /api/sync_mysupport yet -- see that route for why.

    Postgres `projects` turns out to be a *module/task-line* table (one row
    per Implementation Timeline item, with its own title/description/plan
    dates), not a one-row-per-project table. mysupport's `projects` table
    is one row per project (name, status only), so upserting it here would
    just create 135 spurious top-level rows alongside the real per-module
    rows already in Postgres, rather than usefully merging into them. Left
    in place (correctly built, verified against the live DB) for when the
    real source of per-module rows -- likely mysupport's `tasks`/`progress`
    tables -- is decided.
    """
    sql = """
        SELECT
            c.code AS `Client`,
            p.name AS `Title`,
            p.name AS `Projek Name`,
            p.status AS `_status_code`
        FROM projects p
        JOIN companies c ON p.company_id = c.id
    """
    owned = conn is None
    conn = conn or get_mysupport_conn()
    try:
        df = pd.read_sql(sql, conn)
    finally:
        if owned:
            conn.close()

    if df.empty:
        return df

    df["Status Progress"] = df["_status_code"].map(PROJECT_STATUS_BY_CODE)
    df["Source File"] = "mysupport-sync"
    # Matches the dedup_seq every plain (non-duplicate) upload row gets --
    # mysupport has one row per project, so there's nothing to dedupe here,
    # but the column is NOT NULL and part of the ON CONFLICT target, so it
    # must be present and match what an existing row would already have.
    df["Dedup Seq"] = 0
    return df.drop(columns=["_status_code"])


def fetch_mysupport_clients_df(conn=None):
    """NOT wired into /api/sync_mysupport -- see that route for why.

    Postgres `clients` is a small, manually-curated set of Development/
    Warranty/Maintenance engagement rows per client (e.g. FRIM has one:
    "FRIMSAGA001" / "SAGA FRIM" / Maintenance), not a raw project catalog.
    mysupport's `projects` table is one row per project (FRIM alone has
    15), so upserting it here inserted 15+ duplicate rows per client, each
    displaying that client's *entire* ticket total on the Home page --
    which is what made totals look wildly inflated there. Left in place
    (correctly built, verified against the live DB) in case a *subset* or
    *aggregated* form of this data turns out to be useful later.
    """
    sql = """
        SELECT
            c.code AS `Client`,
            p.id AS `Projek ID`,
            p.name AS `Projek Name`,
            p.status AS `_status_code`
        FROM projects p
        JOIN companies c ON p.company_id = c.id
    """
    owned = conn is None
    conn = conn or get_mysupport_conn()
    try:
        df = pd.read_sql(sql, conn)
    finally:
        if owned:
            conn.close()

    if df.empty:
        return df

    df["Projek Status"] = df["_status_code"].map(PROJECT_STATUS_BY_CODE)
    df["Source File"] = "mysupport-sync"
    df["Projek ID"] = df["Projek ID"].astype(str)
    return df.drop(columns=["_status_code"])
