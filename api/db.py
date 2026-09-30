"""Neon/Postgres data layer.

All ticket & project data lives in Postgres now instead of bundled Excel
files (Vercel's filesystem is read-only/ephemeral anyway, so that never
would have worked in production). Uploads are *merged* in: existing rows
are matched by a natural key and updated in place, new rows are inserted,
nothing is ever silently overwritten by an older file.
"""
import os
import re
import warnings
from contextlib import contextmanager

import pandas as pd
import psycopg2
import psycopg2.extras
from werkzeug.security import check_password_hash, generate_password_hash

warnings.filterwarnings("ignore", message="pandas only supports SQLAlchemy")

# Only used to seed the admin_users table the very first time it's empty --
# after that, the password lives solely in the database (as a hash) and
# this constant is never read again. Change the password afterwards via
# set_admin_password(), not by editing this.
DEFAULT_ADMIN_USERNAME = "admin"
DEFAULT_ADMIN_PASSWORD = "AdminSW123"

TICKET_DB_COLUMNS = [
    ("Client", "client"),
    ("Ticket No", "ticket_no"),
    ("Task Type", "task_type"),
    ("Project", "project"),
    ("Company", "company"),
    ("Ticket Title", "ticket_title"),
    ("Ticket Detail", "ticket_detail"),
    ("Ticket Category", "ticket_category"),
    ("Priority", "priority"),
    ("Ticket Created Date", "ticket_created_date"),
    ("Ticket Completed Date", "ticket_completed_date"),
    ("Ticket Closed Date", "ticket_closed_date"),
    ("Ticket Status", "ticket_status"),
    ("SLA Dateline", "sla_dateline"),
    ("SLA Late", "sla_late"),
    ("Days", "days"),
    ("Ageing", "ageing"),
    ("Days to Close", "days_to_close"),
    ("SLA Breach", "sla_breach"),
    ("Source File", "source_file"),
]

PROJECT_DB_COLUMNS = [
    ("Client", "client"),
    ("Title", "title"),
    ("Projek Name", "projek_name"),
    ("Description", "description"),
    ("Category", "category"),
    ("Progress", "progress"),
    ("Priority", "priority"),
    ("Plan Start Date", "plan_start_date"),
    ("Plan End Date", "plan_end_date"),
    ("Target Start Date", "target_start_date"),
    ("Target End Date", "target_end_date"),
    ("Actual Start Date", "actual_start_date"),
    ("Actual End Date", "actual_end_date"),
    ("Duration", "duration"),
    ("Assigned to", "assigned_to"),
    ("Status Progress", "status_progress"),
    ("Percentage", "percentage"),
    ("Overall Progress Task (%)", "overall_progress_task"),
    ("Source File", "source_file"),
    ("Dedup Seq", "dedup_seq"),
]

CLIENT_DB_COLUMNS = [
    ("Client", "client"),
    ("Projek ID", "projek_id"),
    ("Projek Name", "projek_name"),
    ("Projek Status", "projek_status"),
    ("Progress Status", "progress_status"),
    ("Start Date", "start_date"),
    ("End Date", "end_date"),
    ("Technology", "technology"),
    ("Source File", "source_file"),
]

# Display label -> column for the projectmilestone table. The labels are the
# field names the table was specified with (Projectname, Startdate, ...);
# the columns themselves stay snake_case like every other table here.
# "Dedup Seq" is plumbing for the re-upload key (see idx_projectmilestone_
# dedup_key), not a user-facing field -- the Milestones tab filters it out
# of its column list, exactly like the Project Details tab does.
PROJECT_MILESTONE_DB_COLUMNS = [
    ("Client", "client"),
    ("Projectname", "project_name"),
    ("Taskname", "task_name"),
    ("Duration", "duration"),
    ("Startdate", "start_date"),
    ("Enddate", "end_date"),
    ("Progress", "progress"),
    ("Dedup Seq", "dedup_seq"),
]

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS tickets (
    id SERIAL PRIMARY KEY,
    client TEXT NOT NULL,
    ticket_no TEXT NOT NULL,
    task_type TEXT,
    project TEXT,
    company TEXT,
    ticket_title TEXT,
    ticket_detail TEXT,
    ticket_category TEXT,
    priority TEXT,
    ticket_created_date DATE,
    ticket_completed_date DATE,
    ticket_closed_date DATE,
    ticket_status TEXT,
    sla_dateline DATE,
    sla_late TEXT,
    days NUMERIC,
    ageing TEXT,
    days_to_close NUMERIC,
    sla_breach BOOLEAN DEFAULT FALSE,
    source_file TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE (client, ticket_no)
);

CREATE TABLE IF NOT EXISTS projects (
    id SERIAL PRIMARY KEY,
    client TEXT,
    title TEXT,
    projek_name TEXT,
    description TEXT,
    category TEXT,
    progress TEXT,
    priority TEXT,
    start_date DATE,
    due_date DATE,
    target_date DATE,
    duration TEXT,
    assigned_to TEXT,
    status_progress TEXT,
    percentage NUMERIC,
    overall_progress_task NUMERIC,
    target_start_date DATE,
    actual_start_date DATE,
    target_end_date DATE,
    actual_end_date DATE,
    source_file TEXT,
    dedup_seq INTEGER NOT NULL DEFAULT 0,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Columns added after the table already existed in production.
ALTER TABLE projects ADD COLUMN IF NOT EXISTS description TEXT;
ALTER TABLE projects ADD COLUMN IF NOT EXISTS duration TEXT;
ALTER TABLE projects ADD COLUMN IF NOT EXISTS dedup_seq INTEGER NOT NULL DEFAULT 0;
ALTER TABLE projects ADD COLUMN IF NOT EXISTS projek_name TEXT;
ALTER TABLE projects ADD COLUMN IF NOT EXISTS target_start_date DATE;
ALTER TABLE projects ADD COLUMN IF NOT EXISTS actual_start_date DATE;
ALTER TABLE projects ADD COLUMN IF NOT EXISTS target_end_date DATE;
ALTER TABLE projects ADD COLUMN IF NOT EXISTS actual_end_date DATE;

-- Split into three date "types" (Plan/Target/Actual); the old bare
-- start_date/due_date/target_date columns are retired from the app (see
-- PROJECT_DB_COLUMNS) but left in place rather than dropped, and their
-- data is copied forward once here. start_date/due_date always meant
-- "originally scheduled", so they become Plan Start/End; the old single
-- target_date becomes Target End Date (it was a one-sided deadline, not
-- a range). Only fills rows that haven't already been migrated/edited.
ALTER TABLE projects ADD COLUMN IF NOT EXISTS plan_start_date DATE;
ALTER TABLE projects ADD COLUMN IF NOT EXISTS plan_end_date DATE;
UPDATE projects SET plan_start_date = start_date WHERE plan_start_date IS NULL AND start_date IS NOT NULL;
UPDATE projects SET plan_end_date = due_date WHERE plan_end_date IS NULL AND due_date IS NOT NULL;
UPDATE projects SET target_end_date = target_date WHERE target_end_date IS NULL AND target_date IS NOT NULL;

-- Lets the Project Details table's row/module order be dragged around by
-- hand instead of being stuck at insertion (id) order. Backfilled from id
-- the first time this column exists so existing tables keep their current
-- order until someone actually reorders something; every row inserted
-- after that always gets an explicit value (see insert_project_row), so
-- this UPDATE only ever touches genuinely new/legacy NULLs, not rows
-- someone has already reordered.
ALTER TABLE projects ADD COLUMN IF NOT EXISTS sort_order BIGINT;
UPDATE projects SET sort_order = id WHERE sort_order IS NULL;
CREATE INDEX IF NOT EXISTS idx_projects_sort_order ON projects(sort_order);

CREATE TABLE IF NOT EXISTS clients (
    id SERIAL PRIMARY KEY,
    client TEXT,
    projek_id TEXT,
    projek_name TEXT,
    projek_status TEXT,
    start_date DATE,
    end_date DATE,
    technology TEXT,
    source_file TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE (client, projek_id)
);

-- Column added after the table already existed in production.
ALTER TABLE clients ADD COLUMN IF NOT EXISTS technology TEXT;

-- Hand-entered on the Home page's Development table (it's a status people
-- type, not something derived from the source Excel's Client sheet).
ALTER TABLE clients ADD COLUMN IF NOT EXISTS progress_status TEXT;

-- The source "Client Project" sheet has many rows with a blank title
-- and/or start/due date (sub-item description lines, section
-- separators). A plain UNIQUE constraint can't dedupe those on
-- re-upload: SQL NULL is never equal to NULL, even inside a composite
-- UNIQUE constraint, so a row with *any* NULL in the key columns is
-- exempt from the uniqueness check entirely and just inserts again
-- every time, no matter how many extra columns (dedup_seq included)
-- are added to a plain constraint. Wrapping the nullable columns in
-- COALESCE turns each NULL into a real, comparable sentinel value, so
-- this unique INDEX (not a table CONSTRAINT -- Postgres only allows
-- expressions in an index) is what actually makes ON CONFLICT match.
ALTER TABLE projects DROP CONSTRAINT IF EXISTS projects_client_title_start_date_key;
ALTER TABLE projects DROP CONSTRAINT IF EXISTS projects_client_title_start_date_due_date_description_key;
ALTER TABLE projects DROP CONSTRAINT IF EXISTS projects_client_title_start_date_due_date_descrip_dedup_key;
DROP INDEX IF EXISTS idx_projects_dedup_key;
CREATE UNIQUE INDEX idx_projects_dedup_key ON projects (
    COALESCE(client, ''),
    COALESCE(title, ''),
    COALESCE(plan_start_date, DATE '0001-01-01'),
    COALESCE(plan_end_date, DATE '0001-01-01'),
    COALESCE(description, ''),
    dedup_seq
);

-- (client, ticket_no) and (client, title, start_date, ...) already have a
-- backing index from the UNIQUE constraints above, which also serves
-- plain "WHERE client = ..." lookups since client is the leading column.
-- These cover the other columns the dashboard filters/sorts by, so those
-- queries hit an index instead of a sequential scan.
CREATE INDEX IF NOT EXISTS idx_tickets_status ON tickets(ticket_status);
CREATE INDEX IF NOT EXISTS idx_tickets_priority ON tickets(priority);
CREATE INDEX IF NOT EXISTS idx_tickets_task_type ON tickets(task_type);
CREATE INDEX IF NOT EXISTS idx_tickets_created_date ON tickets(ticket_created_date);

CREATE TABLE IF NOT EXISTS admin_users (
    id SERIAL PRIMARY KEY,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Audit trail for the automatic Development -> Warranty transfer (a
-- Development row whose Progress Status is "Completed" and whose End Date
-- has passed). Stores enough of the old state to put the row back on
-- revert, plus text snapshots so the log stays readable even after the
-- client row itself is deleted. Deliberately no FK to clients: reset_all()
-- TRUNCATEs clients, and truncating a table that a foreign key points at
-- requires CASCADE (or truncating both together), which would silently
-- start wiping this table too.
CREATE TABLE IF NOT EXISTS transfer_history (
    id SERIAL PRIMARY KEY,
    client_row_id INTEGER NOT NULL,
    client TEXT,
    projek_id TEXT,
    projek_name TEXT,
    from_status TEXT NOT NULL,
    to_status TEXT NOT NULL,
    prev_progress_status TEXT,
    transferred_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    reverted_at TIMESTAMPTZ,
    -- Set by the Proceed button: the transfer stands, the admin just
    -- doesn't want the row listed in the panel any more. Kept separate
    -- from reverted_at because the two mean opposite things -- a dismissed
    -- entry must keep re-asserting itself after an Excel upload (see
    -- reassert_active_transfers), a reverted one must not.
    dismissed_at TIMESTAMPTZ
);

-- Column added after the table already existed (same pattern as the other
-- post-hoc ALTERs above -- init_schema runs SCHEMA_SQL on every startup).
ALTER TABLE transfer_history ADD COLUMN IF NOT EXISTS dismissed_at TIMESTAMPTZ;

-- One *active* (not yet reverted) transfer per client row. This is what
-- makes the auto-transfer idempotent: a re-upload can push a row back to
-- Development, and the next Home load moves it again without writing a
-- second history entry (the INSERT ... ON CONFLICT DO NOTHING below).
-- Dismissed entries still count as active for the same reason: the move
-- happened and must not be logged twice -- it's only hidden from the panel.
CREATE UNIQUE INDEX IF NOT EXISTS idx_transfer_history_active
    ON transfer_history(client_row_id) WHERE reverted_at IS NULL;

-- Milestones/tasks per project (Client, Projectname, Taskname, Duration,
-- Startdate, Enddate, Progress). Column names follow the snake_case style
-- of the other tables -- the exact field names above live in
-- PROJECT_MILESTONE_DB_COLUMNS as the display labels. Duration stays TEXT
-- and Progress stays TEXT for the same reason projects.duration /
-- projects.progress do: the values arrive as free text ("5 days", "70%",
-- "Completed") rather than as a typed number. No FK to clients -- the other
-- tables deliberately don't reference each other either, so a row can be
-- uploaded/edited without constraint ordering to worry about.
CREATE TABLE IF NOT EXISTS projectmilestone (
    id SERIAL PRIMARY KEY,
    client TEXT,
    project_name TEXT,
    task_name TEXT,
    duration TEXT,
    start_date DATE,
    end_date DATE,
    progress TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_projectmilestone_client ON projectmilestone(client);

-- Same NULL-safe dedup pattern as idx_projects_dedup_key below: a plain
-- UNIQUE constraint can never match ON CONFLICT while any key column is
-- NULL (SQL NULL != NULL), so COALESCE turns every key part into a real
-- comparable value. Deliberately keyed on (client, project, task) and NOT
-- on the dates: milestone dates shift constantly as a rollout plan slips,
-- and a date in the key would turn every shifted re-upload into a second
-- copy of the task instead of an update. dedup_seq disambiguates two
-- otherwise-identical rows (the same task name listed twice in one
-- project), counting occurrences the same way projects does.
ALTER TABLE projectmilestone ADD COLUMN IF NOT EXISTS dedup_seq INTEGER NOT NULL DEFAULT 0;
DROP INDEX IF EXISTS idx_projectmilestone_dedup_key;
CREATE UNIQUE INDEX idx_projectmilestone_dedup_key ON projectmilestone (
    COALESCE(client, ''),
    COALESCE(project_name, ''),
    COALESCE(task_name, ''),
    dedup_seq
);
"""


def get_conn():
    url = os.environ.get("DATABASE_URL")
    if not url:
        raise RuntimeError("DATABASE_URL environment variable is not set")
    return psycopg2.connect(url, connect_timeout=10)


@contextmanager
def db_connection(conn=None):
    """A connection that actually gets closed -- unless the caller is
    already managing one, in which case we just reuse it.

    `with psycopg2_connection:` only wraps commit/rollback, it never
    closes the socket -- leaving that to garbage collection let each page
    load quietly leak several connections against Neon's connection cap,
    which is what made requests hang once it was exhausted rather than
    fail fast. Passing `conn` through (see request_connection() in
    index.py) also means a single page load, which needs several of the
    functions below, opens ONE connection instead of one per call --
    each fresh connect to Neon costs real round-trip time, so that's
    most of what makes a page load fast or slow.
    """
    if conn is not None:
        yield conn
        return
    owned = get_conn()
    try:
        yield owned
        owned.commit()
    except Exception:
        owned.rollback()
        raise
    finally:
        owned.close()


def init_schema(conn=None):
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute(SCHEMA_SQL)
            # Seed the one default admin account the first time this table
            # is created (ON CONFLICT DO NOTHING makes this a no-op on
            # every later startup, so an admin who changes the password
            # afterward never gets silently reset back to the default).
            cur.execute(
                "INSERT INTO admin_users (username, password_hash) VALUES (%s, %s) "
                "ON CONFLICT (username) DO NOTHING",
                (DEFAULT_ADMIN_USERNAME, generate_password_hash(DEFAULT_ADMIN_PASSWORD)),
            )


def verify_admin_credentials(username, password, conn=None):
    """Check a username/password against the admin_users table.

    Case-sensitive on username (matches how it was stored), and safe
    against timing-based username enumeration since check_password_hash
    is always run -- against a dummy hash if the username doesn't exist --
    rather than short-circuiting as soon as the lookup misses.
    """
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute("SELECT password_hash FROM admin_users WHERE username = %s", (username,))
            row = cur.fetchone()

    dummy_hash = generate_password_hash("not-a-real-password")
    stored_hash = row[0] if row else dummy_hash
    ok = check_password_hash(stored_hash, password)
    return ok and row is not None


def set_admin_password(username, new_password, conn=None):
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute(
                "UPDATE admin_users SET password_hash = %s, updated_at = now() WHERE username = %s",
                (generate_password_hash(new_password), username),
            )
            updated = cur.rowcount > 0
        c.commit()
    return updated


def _records_for_insert(df, columns):
    """DataFrame -> list of tuples in `columns` order, NaN/NaT -> None."""
    for display_col, _ in columns:
        if display_col not in df.columns:
            df[display_col] = None
    df = df[[c for c, _ in columns]].copy()
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.date
    df = df.astype(object).where(pd.notnull(df), None)
    return list(df.itertuples(index=False, name=None))


def upsert_tickets(df, conn=None, sync_columns=None):
    """Insert new tickets / update existing ones (matched by client + ticket no).

    `sync_columns`, when given, restricts which non-key columns get
    overwritten on an existing row (used by the mysupport sync, which only
    has data for a subset of fields -- Priority/SLA/Ageing/etc. are left
    exactly as they are instead of being blanked out). Defaults to every
    mapped column, i.e. the original full-overwrite behavior used by the
    manual Excel/CSV upload.

    Returns (inserted_count, updated_count).
    """
    if df.empty:
        return 0, 0

    records = _records_for_insert(df, TICKET_DB_COLUMNS)
    db_cols = [c for _, c in TICKET_DB_COLUMNS]
    update_cols = [c for c in db_cols if c not in ("client", "ticket_no")]
    if sync_columns is not None:
        update_cols = [c for c in update_cols if c in sync_columns]
    set_clause = ", ".join(f"{c} = EXCLUDED.{c}" for c in update_cols)

    sql = f"""
        INSERT INTO tickets ({', '.join(db_cols)})
        VALUES %s
        ON CONFLICT (client, ticket_no) DO UPDATE SET
            {set_clause},
            updated_at = now()
        RETURNING (xmax = 0) AS inserted
    """

    with db_connection(conn) as c:
        with c.cursor() as cur:
            results = psycopg2.extras.execute_values(cur, sql, records, page_size=500, fetch=True)

    inserted = sum(1 for r in results if r[0])
    updated = len(results) - inserted
    return inserted, updated


def upsert_projects(df, conn=None, sync_columns=None):
    """See upsert_tickets() for what `sync_columns` does."""
    if df.empty:
        return 0, 0

    records = _records_for_insert(df, PROJECT_DB_COLUMNS)
    db_cols = [c for _, c in PROJECT_DB_COLUMNS]
    key_cols = ("client", "title", "plan_start_date", "plan_end_date", "description", "dedup_seq")
    update_cols = [c for c in db_cols if c not in key_cols]
    if sync_columns is not None:
        update_cols = [c for c in update_cols if c in sync_columns]
    set_clause = ", ".join(f"{c} = EXCLUDED.{c}" for c in update_cols)

    # Must match idx_projects_dedup_key's expressions exactly for
    # Postgres to recognize it as the ON CONFLICT target.
    sql = f"""
        INSERT INTO projects ({', '.join(db_cols)})
        VALUES %s
        ON CONFLICT (
            COALESCE(client, ''),
            COALESCE(title, ''),
            COALESCE(plan_start_date, DATE '0001-01-01'),
            COALESCE(plan_end_date, DATE '0001-01-01'),
            COALESCE(description, ''),
            dedup_seq
        ) DO UPDATE SET
            {set_clause},
            updated_at = now()
        RETURNING (xmax = 0) AS inserted
    """

    with db_connection(conn) as c:
        with c.cursor() as cur:
            results = psycopg2.extras.execute_values(cur, sql, records, page_size=500, fetch=True)

    inserted = sum(1 for r in results if r[0])
    updated = len(results) - inserted
    return inserted, updated


def renumber_projects_sort_order(conn=None):
    """Re-group any project rows that share a (Client, Title) module but
    have drifted apart in sort_order back into one contiguous block.

    upsert_projects() intentionally leaves a brand-new row's sort_order
    unset -- SCHEMA_SQL's `UPDATE projects SET sort_order = id WHERE
    sort_order IS NULL` (which reruns on every app startup, see
    init_schema()) then "heals" it to that row's own id, which is always
    higher than everything already there. A module whose sheet gained
    extra tasks in a later upload therefore has its new tasks land at the
    very end of the whole table instead of next to the rest of that
    module -- and the Project Details page's Overall Progress Task (%)
    merged-cell rendering (see build_project_charts() in index.py) only
    merges *contiguous* same-(Client, Title) rows, so the module then
    displays as two separate groups.

    Also re-sorts *within* each module by the leading number in
    Description (e.g. "9. Integrasi..." before "10. Doc UAT..." before
    "12. FAT..."), instead of upload/arrival order -- a later upload's
    rows land after the module's existing ones (see above), which
    otherwise leaves e.g. "12. FAT" sitting before "13. Training" and
    "14. Go Live" but ahead of "10."/"11." simply because 10-14 came from
    an earlier upload than 12 did. Plain arrival order is kept as the
    fallback for a description with no leading number, so nothing
    disappears or gets pushed somewhere arbitrary.

    Fixes this generally, independent of insert/update history: read the
    table in its current (possibly split/misordered) order, stable-sort
    every row to (a) the position of its (Client, Title) group's *first*
    appearance, then (b) its own leading Description number if it has
    one, then renumber sequentially. A module that's already contiguous
    and numerically ordered is left exactly where it was. Call this once
    after any project upsert so neither problem can recur.
    """
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute("SELECT id, client, title, description FROM projects ORDER BY sort_order NULLS LAST, id")
            rows = cur.fetchall()
            if not rows:
                return

            group_first_pos = {}
            for pos, (row_id, client, title, description) in enumerate(rows):
                key = (client, title)
                if key not in group_first_pos:
                    group_first_pos[key] = pos

            def desc_number(description):
                m = re.match(r"\s*(\d+)\s*\.", description or "")
                return int(m.group(1)) if m else None

            indexed = list(enumerate(rows))
            indexed.sort(key=lambda item: (
                group_first_pos[(item[1][1], item[1][2])],
                (0, desc_number(item[1][3])) if desc_number(item[1][3]) is not None else (1, item[0]),
            ))

            updates = [(new_order + 1, row_id) for new_order, (_, (row_id, _, _, _)) in enumerate(indexed)]
            # Not touching updated_at here -- this is purely a display-order
            # repair, not a change to the row's actual data, and bumping it
            # for every project on every upload would falsely make
            # everything look freshly edited.
            psycopg2.extras.execute_values(
                cur,
                "UPDATE projects AS p SET sort_order = v.new_order "
                "FROM (VALUES %s) AS v(new_order, id) WHERE p.id = v.id",
                updates, page_size=500,
            )


def upsert_clients(df, conn=None, sync_columns=None):
    """Insert new client rows / update existing ones (matched by client + projek id).

    See upsert_tickets() for what `sync_columns` does.

    Returns (inserted_count, updated_count).
    """
    if df.empty:
        return 0, 0

    # Captured before _records_for_insert() backfills missing columns with
    # None: Progress Status is typed into the dashboard and isn't part of
    # the source Excel's Client sheet, so an upload that doesn't carry the
    # column must leave existing values alone instead of blanking them.
    has_progress_status = "Progress Status" in df.columns

    records = _records_for_insert(df, CLIENT_DB_COLUMNS)
    db_cols = [c for _, c in CLIENT_DB_COLUMNS]
    update_cols = [c for c in db_cols if c not in ("client", "projek_id")]
    if not has_progress_status:
        update_cols = [c for c in update_cols if c != "progress_status"]
    if sync_columns is not None:
        update_cols = [c for c in update_cols if c in sync_columns]
    set_clause = ", ".join(f"{c} = EXCLUDED.{c}" for c in update_cols)

    sql = f"""
        INSERT INTO clients ({', '.join(db_cols)})
        VALUES %s
        ON CONFLICT (client, projek_id) DO UPDATE SET
            {set_clause},
            updated_at = now()
        RETURNING (xmax = 0) AS inserted
    """

    with db_connection(conn) as c:
        with c.cursor() as cur:
            results = psycopg2.extras.execute_values(cur, sql, records, page_size=500, fetch=True)

    inserted = sum(1 for r in results if r[0])
    updated = len(results) - inserted
    return inserted, updated


TICKET_SEARCH_COLUMNS = ["ticket_detail", "ticket_title", "ticket_no", "ticket_category", "company", "project"]


def _build_ticket_where(filters):
    """Turn the parsed filter dict into a parameterized WHERE clause.

    Every branch here maps onto one of the indexes created in SCHEMA_SQL
    (ticket_status, priority, task_type, ticket_created_date) so filtered
    fetches hit an index scan instead of pulling the whole table into
    Python and filtering with pandas.
    """
    clauses = []
    params = []

    if filters.get("clients"):
        clauses.append("client = ANY(%s)")
        params.append(filters["clients"])
    if filters.get("priorities"):
        clauses.append("priority = ANY(%s)")
        params.append(filters["priorities"])
    if filters.get("statuses"):
        clauses.append("ticket_status = ANY(%s)")
        params.append(filters["statuses"])
    if filters.get("task_types"):
        clauses.append("task_type = ANY(%s)")
        params.append(filters["task_types"])
    if filters.get("date_start"):
        clauses.append("ticket_created_date >= %s")
        params.append(filters["date_start"])
    if filters.get("date_end"):
        clauses.append("ticket_created_date <= %s")
        params.append(filters["date_end"])
    if filters.get("search"):
        term = f"%{filters['search']}%"
        clauses.append("(" + " OR ".join(f"{c} ILIKE %s" for c in TICKET_SEARCH_COLUMNS) + ")")
        params.extend([term] * len(TICKET_SEARCH_COLUMNS))

    where_sql = (" WHERE " + " AND ".join(clauses)) if clauses else ""
    return where_sql, params


def fetch_tickets_df(filters=None, conn=None):
    db_cols = [c for _, c in TICKET_DB_COLUMNS]
    display_cols = [c for c, _ in TICKET_DB_COLUMNS]
    where_sql, params = _build_ticket_where(filters or {})
    sql = f"SELECT id, {', '.join(db_cols)} FROM tickets{where_sql} ORDER BY id"

    with db_connection(conn) as c:
        df = pd.read_sql_query(sql, c, params=params or None)

    if df.empty:
        return pd.DataFrame(columns=["_row_idx"] + display_cols)

    df = df.rename(columns=dict(zip(db_cols, display_cols)))
    df = df.rename(columns={"id": "_row_idx"})

    for col in ["Ticket Created Date", "Ticket Completed Date", "Ticket Closed Date", "SLA Dateline"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col])

    return df


def get_filter_metadata(conn=None):
    """Distinct filter-dropdown values and the ticket date range.

    Queried unfiltered (independent of whatever the user currently has
    selected) so dropdowns always show every possible option. Each SELECT
    DISTINCT ... ORDER BY hits the matching index, so it's an index scan
    rather than a sequential one even as the table grows.
    """
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute("SELECT DISTINCT client FROM tickets ORDER BY client")
            clients = [r[0] for r in cur.fetchall()]
            cur.execute("SELECT DISTINCT priority FROM tickets WHERE priority IS NOT NULL ORDER BY priority")
            priorities = [r[0] for r in cur.fetchall()]
            cur.execute("SELECT DISTINCT ticket_status FROM tickets WHERE ticket_status IS NOT NULL ORDER BY ticket_status")
            statuses = [r[0] for r in cur.fetchall()]
            cur.execute("SELECT DISTINCT task_type FROM tickets WHERE task_type IS NOT NULL ORDER BY task_type")
            task_types = [r[0] for r in cur.fetchall()]
            cur.execute("SELECT min(ticket_created_date), max(ticket_created_date) FROM tickets")
            min_date, max_date = cur.fetchone()

    return {
        "clients": clients,
        "priorities": priorities,
        "statuses": statuses,
        "task_types": task_types,
        "min_date": min_date.strftime("%Y-%m-%d") if min_date else None,
        "max_date": max_date.strftime("%Y-%m-%d") if max_date else None,
    }


def fetch_projects_df(conn=None):
    db_cols = [c for _, c in PROJECT_DB_COLUMNS]
    display_cols = [c for c, _ in PROJECT_DB_COLUMNS]
    sql = f"SELECT id, {', '.join(db_cols)} FROM projects ORDER BY sort_order NULLS LAST, id"

    with db_connection(conn) as c:
        df = pd.read_sql_query(sql, c)

    if df.empty:
        return pd.DataFrame(columns=["_row_idx"] + display_cols)

    df = df.rename(columns=dict(zip(db_cols, display_cols)))
    df = df.rename(columns={"id": "_row_idx", "Source File": "_source_file"})

    for col in ["Plan Start Date", "Plan End Date", "Target Start Date", "Target End Date", "Actual Start Date", "Actual End Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col])

    return df


def fetch_clients_df(conn=None):
    db_cols = [c for _, c in CLIENT_DB_COLUMNS]
    display_cols = [c for c, _ in CLIENT_DB_COLUMNS]
    sql = f"SELECT id, {', '.join(db_cols)} FROM clients ORDER BY client, projek_id"

    with db_connection(conn) as c:
        df = pd.read_sql_query(sql, c)

    if df.empty:
        return pd.DataFrame(columns=["_row_idx"] + display_cols)

    df = df.rename(columns=dict(zip(db_cols, display_cols)))
    df = df.rename(columns={"id": "_row_idx"})

    for col in ["Start Date", "End Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col])

    return df


def get_counts(conn=None):
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute("SELECT count(*) FROM tickets")
            tickets = cur.fetchone()[0]
            cur.execute("SELECT count(*) FROM projects")
            projects = cur.fetchone()[0]
            cur.execute("SELECT count(*) FROM clients")
            clients = cur.fetchone()[0]
            cur.execute("SELECT max(updated_at) FROM tickets")
            last_ticket_update = cur.fetchone()[0]
    return {
        "tickets": tickets,
        "projects": projects,
        "clients": clients,
        "last_updated": last_ticket_update.strftime("%d/%m/%Y %H:%M") if last_ticket_update else None,
    }


def reset_all(conn=None):
    """Wipe all ticket, project & client data. Used by the 'Restart' button."""
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute("TRUNCATE TABLE tickets RESTART IDENTITY")
            cur.execute("TRUNCATE TABLE projects RESTART IDENTITY")
            cur.execute("TRUNCATE TABLE clients RESTART IDENTITY")
            cur.execute("TRUNCATE TABLE transfer_history RESTART IDENTITY")
            cur.execute("TRUNCATE TABLE projectmilestone RESTART IDENTITY")


def update_ticket_field(row_id, db_column, value, conn=None):
    valid_cols = {c for _, c in TICKET_DB_COLUMNS}
    if db_column not in valid_cols:
        raise ValueError(f"Unknown column: {db_column}")
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute(
                f"UPDATE tickets SET {db_column} = %s, updated_at = now() WHERE id = %s",
                (value, row_id),
            )


def update_client_field(row_id, db_column, value, conn=None):
    valid_cols = {c for _, c in CLIENT_DB_COLUMNS}
    if db_column not in valid_cols:
        raise ValueError(f"Unknown column: {db_column}")
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute(
                f"UPDATE clients SET {db_column} = %s, updated_at = now() WHERE id = %s",
                (value, row_id),
            )


def update_project_field(row_id, db_column, value, conn=None):
    valid_cols = {c for _, c in PROJECT_DB_COLUMNS}
    if db_column not in valid_cols:
        raise ValueError(f"Unknown column: {db_column}")
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute(
                f"UPDATE projects SET {db_column} = %s, updated_at = now() WHERE id = %s",
                (value, row_id),
            )


# A Development row qualifies for the automatic transfer to Warranty when
# its hand-typed Progress Status says the work is finished AND its End Date
# has already passed. The date comparison mirrors applyEndDateBadges()'s
# "Expired" rule on the Home page (diffDays < 0): today itself still counts
# as running, only strictly-past end dates count as ended. Progress Status
# is compared case/whitespace-insensitively because it's free text typed by
# hand, not a constrained column.
TRANSFER_ELIGIBLE_PREDICATE = (
    "projek_status = 'Development' "
    "AND lower(btrim(progress_status)) = 'completed' "
    "AND end_date IS NOT NULL AND end_date < CURRENT_DATE"
)


def auto_transfer_development_rows(conn=None):
    """Move every eligible Development row to Warranty and log the change.

    Eligibility: Progress Status = "Completed" and End Date already passed.
    The move clears Progress Status (it's a Development-only field) but keeps
    the old value in transfer_history so revert can restore it.

    Three steps in one transaction: lock the eligible rows, insert history,
    then update. FOR UPDATE serialises two Home pages loading at once, and
    the partial unique index idx_transfer_history_active makes the history
    INSERT a no-op when a row already has an active (un-reverted) entry --
    which is what happens after an Excel upload drags the row back to
    Development while the cleared Progress Status would otherwise never
    match again.

    Returns the list of rows moved (possibly empty).
    """
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute(
                f"""SELECT id, client, projek_id, projek_name, progress_status
                    FROM clients
                    WHERE {TRANSFER_ELIGIBLE_PREDICATE}
                    ORDER BY id
                    FOR UPDATE"""
            )
            eligible = cur.fetchall()
            if not eligible:
                return []

            moved = []
            for row_id, client, projek_id, projek_name, prev_progress in eligible:
                cur.execute(
                    """INSERT INTO transfer_history
                           (client_row_id, client, projek_id, projek_name,
                            from_status, to_status, prev_progress_status)
                       VALUES (%s, %s, %s, %s, %s, %s, %s)
                       ON CONFLICT (client_row_id) WHERE reverted_at IS NULL
                       DO NOTHING""",
                    (row_id, client, projek_id, projek_name,
                     "Development", "Warranty", prev_progress),
                )
                cur.execute(
                    """UPDATE clients
                       SET projek_status = 'Warranty',
                           progress_status = NULL,
                           updated_at = now()
                       WHERE id = %s""",
                    (row_id,),
                )
                moved.append({
                    "client_row_id": row_id,
                    "client": client,
                    "projek_id": projek_id,
                    "projek_name": projek_name,
                    "prev_progress_status": prev_progress,
                })
            return moved


def fetch_transfer_history(conn=None, active_only=True):
    """Transfer log rows for the Home page panel, newest first.

    active_only keeps just the transfers that are still worth showing --
    neither reverted (undone) nor dismissed via Proceed (accepted and
    cleared from the list). Timestamps are formatted for display since this
    feeds the template directly.
    """
    sql = (
        "SELECT id, client_row_id, client, projek_id, projek_name, "
        "from_status, to_status, prev_progress_status, transferred_at, reverted_at "
        "FROM transfer_history"
    )
    if active_only:
        sql += " WHERE reverted_at IS NULL AND dismissed_at IS NULL"
    sql += " ORDER BY transferred_at DESC, id DESC"

    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute(sql)
            cols = [d[0] for d in cur.description]
            rows = [dict(zip(cols, r)) for r in cur.fetchall()]

    for r in rows:
        for key in ("transferred_at", "reverted_at"):
            val = r[key]
            # Postgres hands back TIMESTAMPTZ in the DB's own zone (UTC on
            # Neon), which would print 8 hours behind wall-clock time here --
            # shift to the server's local zone before formatting.
            if val is not None and val.tzinfo is not None:
                val = val.astimezone()
            r[key] = val.strftime("%d/%m/%Y %H:%M") if val else None
    return rows


def revert_transfer(history_id, conn=None):
    """Undo one transfer: put the row back to its previous phase/status.

    The client row is only touched if it still sits in the status the
    transfer moved it to -- if someone manually edited it to Maintenance
    (or deleted it) since, reverting just closes the history entry instead
    of clobbering that newer change.

    Returns {"success": bool, "row_restored": bool, "error": str|None}.
    """
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute(
                """SELECT client_row_id, from_status, to_status,
                          prev_progress_status, reverted_at
                   FROM transfer_history
                   WHERE id = %s
                   FOR UPDATE""",
                (history_id,),
            )
            h = cur.fetchone()
            if not h:
                return {"success": False, "row_restored": False, "error": "Transfer not found"}
            row_id, from_status, to_status, prev_progress, reverted_at = h
            if reverted_at is not None:
                return {"success": True, "row_restored": False, "error": None}

            cur.execute("SELECT projek_status FROM clients WHERE id = %s", (row_id,))
            current = cur.fetchone()
            restored = False
            if current and current[0] == to_status:
                cur.execute(
                    """UPDATE clients
                       SET projek_status = %s,
                           progress_status = %s,
                           updated_at = now()
                       WHERE id = %s""",
                    (from_status, prev_progress, row_id),
                )
                restored = True

            cur.execute(
                "UPDATE transfer_history SET reverted_at = now() WHERE id = %s",
                (history_id,),
            )
            return {"success": True, "row_restored": restored, "error": None}


def dismiss_transfer(history_id, conn=None):
    """Accept a transfer and drop it from the Home panel (the Proceed button).

    Only the history entry is closed (dismissed_at) -- the client row stays
    in Warranty. Deliberately NOT the same as reverted_at: a dismissed entry
    still keeps its row out of Development on the next Excel upload (see
    reassert_active_transfers) and still blocks a duplicate history entry
    via idx_transfer_history_active; all it changes is that the panel stops
    listing it.

    Returns {"success": bool, "already_dismissed": bool, "error": str|None}.
    """
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute(
                "SELECT dismissed_at FROM transfer_history WHERE id = %s FOR UPDATE",
                (history_id,),
            )
            row = cur.fetchone()
            if not row:
                return {"success": False, "already_dismissed": False, "error": "Transfer not found"}
            already = row[0] is not None
            if not already:
                cur.execute(
                    "UPDATE transfer_history SET dismissed_at = now() WHERE id = %s",
                    (history_id,),
                )
            return {"success": True, "already_dismissed": already, "error": None}


def reassert_active_transfers(conn=None):
    """Re-apply still-active transfers after an Excel upload.

    The source workbook's Client sheet always says "Development" and has no
    Progress Status column, so an upload both moves the row back to
    Development and leaves Progress Status untouched (upsert_clients
    excludes it). Without this, a row already transferred to Warranty would
    land back in Development and never match the eligibility rule again
    (its Progress Status was cleared on transfer). Restores to_status for
    every row whose active history entry says it was moved there.

    Returns the number of rows pushed back.
    """
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute(
                """UPDATE clients c
                   SET projek_status = h.to_status,
                       updated_at = now()
                   FROM transfer_history h
                   WHERE h.client_row_id = c.id
                     AND h.reverted_at IS NULL
                     AND c.projek_status = h.from_status"""
            )
            return cur.rowcount


def _clean_insert_values(db_values, valid_cols):
    """Keep only known columns, and drop blank/None ones so an untouched
    field gets its SQL default/NULL instead of an empty string that would
    raise a cast error on a DATE/NUMERIC column."""
    return {
        col: val for col, val in db_values.items()
        if col in valid_cols and val is not None and str(val).strip() != ""
    }


def insert_ticket_row(db_values, conn=None):
    values = _clean_insert_values(db_values, {c for _, c in TICKET_DB_COLUMNS})
    if not values.get("client"):
        raise ValueError("Client is required")
    if not values.get("ticket_no"):
        raise ValueError("Ticket No is required")
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute(
                "SELECT 1 FROM tickets WHERE client = %s AND ticket_no = %s",
                (values["client"], values["ticket_no"]),
            )
            if cur.fetchone():
                raise ValueError(f"Ticket No \"{values['ticket_no']}\" already exists for {values['client']}")
            cols = list(values.keys())
            col_list = ", ".join(cols)
            placeholders = ", ".join(["%s"] * len(cols))
            cur.execute(
                f"INSERT INTO tickets ({col_list}) VALUES ({placeholders}) RETURNING id",
                [values[c] for c in cols],
            )
            return cur.fetchone()[0]


def insert_project_row(db_values, conn=None):
    values = _clean_insert_values(db_values, {c for _, c in PROJECT_DB_COLUMNS} - {"dedup_seq"})
    if not values.get("client"):
        raise ValueError("Client is required")
    if not values.get("title"):
        raise ValueError("Title is required")
    with db_connection(conn) as c:
        with c.cursor() as cur:
            # Mirrors idx_projects_dedup_key: dedup_seq is "how many rows
            # already share this key," so a lone hand-typed row lands on
            # the next free slot instead of colliding with an existing one.
            cur.execute(
                """SELECT COUNT(*) FROM projects
                   WHERE COALESCE(client,'') = COALESCE(%s,'')
                     AND COALESCE(title,'') = COALESCE(%s,'')
                     AND COALESCE(plan_start_date, DATE '0001-01-01') = COALESCE(%s::date, DATE '0001-01-01')
                     AND COALESCE(plan_end_date, DATE '0001-01-01') = COALESCE(%s::date, DATE '0001-01-01')
                     AND COALESCE(description,'') = COALESCE(%s,'')""",
                (
                    values.get("client"), values.get("title"),
                    values.get("plan_start_date"), values.get("plan_end_date"),
                    values.get("description"),
                ),
            )
            dedup_seq = cur.fetchone()[0]
            cols = list(values.keys())
            # sort_order comes from a subquery, not a bound parameter, so a
            # freshly added row always lands at the end of the current
            # display order rather than defaulting to NULL (which would
            # sort first under NULLS LAST... no -- NULLS LAST already
            # keeps it last, but giving it a real value here means a
            # *later* reorder can freely move it without a NULL ever
            # comparing oddly against real sort_order values).
            col_list = ", ".join(cols + ["dedup_seq", "sort_order"])
            placeholders = ", ".join(["%s"] * (len(cols) + 1))
            cur.execute(
                f"INSERT INTO projects ({col_list}) VALUES "
                f"({placeholders}, (SELECT COALESCE(MAX(sort_order), 0) + 1 FROM projects)) RETURNING id",
                [values[c] for c in cols] + [dedup_seq],
            )
            return cur.fetchone()[0]


def reorder_project_rows(ids, conn=None):
    """Reassign sort_order for exactly this set of project row ids, in the
    order given, reusing the same set of sort_order values those rows
    already occupy (just permuted). That confines the change to swapping
    these rows/blocks among themselves -- every other row's sort_order,
    and therefore its position relative to rows outside this set, is left
    completely untouched.
    """
    ids = [int(i) for i in ids]
    if not ids:
        return
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute("SELECT id, sort_order FROM projects WHERE id = ANY(%s)", (ids,))
            rows = dict(cur.fetchall())
            missing = [i for i in ids if i not in rows]
            if missing:
                raise ValueError(f"Unknown project row id(s): {missing}")
            slots = sorted(v if v is not None else k for k, v in rows.items())
            for row_id, slot in zip(ids, slots):
                cur.execute(
                    "UPDATE projects SET sort_order = %s, updated_at = now() WHERE id = %s",
                    (slot, row_id),
                )


def delete_project_row(row_id, conn=None):
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute("DELETE FROM projects WHERE id = %s", (row_id,))
            return cur.rowcount > 0


def delete_ticket_row(row_id, conn=None):
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute("DELETE FROM tickets WHERE id = %s", (row_id,))
            return cur.rowcount > 0


def delete_client_row(row_id, conn=None):
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute("DELETE FROM clients WHERE id = %s", (row_id,))
            return cur.rowcount > 0


def insert_client_row(db_values, conn=None):
    values = _clean_insert_values(db_values, {c for _, c in CLIENT_DB_COLUMNS})
    if not values.get("client"):
        raise ValueError("Client is required")
    with db_connection(conn) as c:
        with c.cursor() as cur:
            if values.get("projek_id"):
                cur.execute(
                    "SELECT 1 FROM clients WHERE client = %s AND projek_id = %s",
                    (values["client"], values["projek_id"]),
                )
                if cur.fetchone():
                    raise ValueError(f"Projek ID \"{values['projek_id']}\" already exists for {values['client']}")
            cols = list(values.keys())
            col_list = ", ".join(cols)
            placeholders = ", ".join(["%s"] * len(cols))
            cur.execute(
                f"INSERT INTO clients ({col_list}) VALUES ({placeholders}) RETURNING id",
                [values[c] for c in cols],
            )
            return cur.fetchone()[0]


def upsert_project_milestones(df, conn=None):
    """Insert new milestone rows / update existing ones (matched by client +
    project + task, with dedup_seq counting same-key repeats) so re-uploading
    the PROJECT MILESTONE sheet refreshes dates/progress in place instead of
    appending a second copy of every task.

    Returns (inserted_count, updated_count).
    """
    if df.empty:
        return 0, 0

    if "Dedup Seq" not in df.columns:
        df = df.copy()
        df["Dedup Seq"] = 0

    records = _records_for_insert(df, PROJECT_MILESTONE_DB_COLUMNS)
    db_cols = [c for _, c in PROJECT_MILESTONE_DB_COLUMNS]
    key_cols = ("client", "project_name", "task_name", "dedup_seq")
    update_cols = [c for c in db_cols if c not in key_cols]
    set_clause = ", ".join(f"{c} = EXCLUDED.{c}" for c in update_cols)

    # Must match idx_projectmilestone_dedup_key's expressions exactly for
    # Postgres to recognize it as the ON CONFLICT target.
    sql = f"""
        INSERT INTO projectmilestone ({', '.join(db_cols)})
        VALUES %s
        ON CONFLICT (
            COALESCE(client, ''),
            COALESCE(project_name, ''),
            COALESCE(task_name, ''),
            dedup_seq
        ) DO UPDATE SET
            {set_clause},
            updated_at = now()
        RETURNING (xmax = 0) AS inserted
    """

    with db_connection(conn) as c:
        with c.cursor() as cur:
            results = psycopg2.extras.execute_values(cur, sql, records, page_size=500, fetch=True)

    inserted = sum(1 for r in results if r[0])
    updated = len(results) - inserted
    return inserted, updated


def fetch_project_milestone_df(conn=None):
    """Milestones as a dataframe shaped exactly like fetch_clients_df():
    an _row_idx (the table's id) plus the display column names, with the
    two date columns parsed to datetime so callers can format them."""
    db_cols = [c for _, c in PROJECT_MILESTONE_DB_COLUMNS]
    display_cols = [c for c, _ in PROJECT_MILESTONE_DB_COLUMNS]
    # start_date (not task_name) is the ordering key: the sheet lists a
    # project's tasks in rollout order (Kick-off -> ... -> GoLive), and
    # alphabetical-by-task would put "Development" before "Kick-off".
    sql = (
        f"SELECT id, {', '.join(db_cols)} FROM projectmilestone "
        "ORDER BY client, project_name, start_date NULLS LAST, id"
    )

    with db_connection(conn) as c:
        df = pd.read_sql_query(sql, c)

    if df.empty:
        return pd.DataFrame(columns=["_row_idx"] + display_cols)

    df = df.rename(columns=dict(zip(db_cols, display_cols)))
    df = df.rename(columns={"id": "_row_idx"})

    for col in ["Startdate", "Enddate"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col])

    return df


def insert_project_milestone_row(db_values, conn=None):
    values = _clean_insert_values(
        db_values, {c for _, c in PROJECT_MILESTONE_DB_COLUMNS} - {"dedup_seq"}
    )
    if not values.get("client"):
        raise ValueError("Client is required")
    if not values.get("task_name"):
        raise ValueError("Taskname is required")
    with db_connection(conn) as c:
        with c.cursor() as cur:
            # Mirrors idx_projectmilestone_dedup_key: dedup_seq is "how many
            # rows already share this key," so a lone hand-typed row lands on
            # the next free slot instead of violating the unique index.
            cur.execute(
                """SELECT COUNT(*) FROM projectmilestone
                   WHERE COALESCE(client,'') = COALESCE(%s,'')
                     AND COALESCE(project_name,'') = COALESCE(%s,'')
                     AND COALESCE(task_name,'') = COALESCE(%s,'')""",
                (values.get("client"), values.get("project_name"), values.get("task_name")),
            )
            dedup_seq = cur.fetchone()[0]
            cols = list(values.keys())
            col_list = ", ".join(cols + ["dedup_seq"])
            placeholders = ", ".join(["%s"] * (len(cols) + 1))
            cur.execute(
                f"INSERT INTO projectmilestone ({col_list}) VALUES ({placeholders}) RETURNING id",
                [values[c] for c in cols] + [dedup_seq],
            )
            return cur.fetchone()[0]


def update_project_milestone_field(row_id, db_column, value, conn=None):
    valid_cols = {c for _, c in PROJECT_MILESTONE_DB_COLUMNS}
    if db_column not in valid_cols:
        raise ValueError(f"Unknown column: {db_column}")
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute(
                f"UPDATE projectmilestone SET {db_column} = %s, updated_at = now() WHERE id = %s",
                (value, row_id),
            )


def delete_project_milestone_row(row_id, conn=None):
    with db_connection(conn) as c:
        with c.cursor() as cur:
            cur.execute("DELETE FROM projectmilestone WHERE id = %s", (row_id,))
            return cur.rowcount > 0
