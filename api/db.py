"""Neon/Postgres data layer.

All ticket & project data lives in Postgres now instead of bundled Excel
files (Vercel's filesystem is read-only/ephemeral anyway, so that never
would have worked in production). Uploads are *merged* in: existing rows
are matched by a natural key and updated in place, new rows are inserted,
nothing is ever silently overwritten by an older file.
"""
import os
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
    ("Description", "description"),
    ("Category", "category"),
    ("Progress", "progress"),
    ("Priority", "priority"),
    ("Start date", "start_date"),
    ("Due date", "due_date"),
    ("Target Date", "target_date"),
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
    ("Start Date", "start_date"),
    ("End Date", "end_date"),
    ("Source File", "source_file"),
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
    source_file TEXT,
    dedup_seq INTEGER NOT NULL DEFAULT 0,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Columns added after the table already existed in production.
ALTER TABLE projects ADD COLUMN IF NOT EXISTS description TEXT;
ALTER TABLE projects ADD COLUMN IF NOT EXISTS duration TEXT;
ALTER TABLE projects ADD COLUMN IF NOT EXISTS dedup_seq INTEGER NOT NULL DEFAULT 0;

CREATE TABLE IF NOT EXISTS clients (
    id SERIAL PRIMARY KEY,
    client TEXT,
    projek_id TEXT,
    projek_name TEXT,
    projek_status TEXT,
    start_date DATE,
    end_date DATE,
    source_file TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE (client, projek_id)
);

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
    COALESCE(start_date, DATE '0001-01-01'),
    COALESCE(due_date, DATE '0001-01-01'),
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


def upsert_tickets(df, conn=None):
    """Insert new tickets / update existing ones (matched by client + ticket no).

    Returns (inserted_count, updated_count).
    """
    if df.empty:
        return 0, 0

    records = _records_for_insert(df, TICKET_DB_COLUMNS)
    db_cols = [c for _, c in TICKET_DB_COLUMNS]
    update_cols = [c for c in db_cols if c not in ("client", "ticket_no")]
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


def upsert_projects(df, conn=None):
    if df.empty:
        return 0, 0

    records = _records_for_insert(df, PROJECT_DB_COLUMNS)
    db_cols = [c for _, c in PROJECT_DB_COLUMNS]
    key_cols = ("client", "title", "start_date", "due_date", "description", "dedup_seq")
    update_cols = [c for c in db_cols if c not in key_cols]
    set_clause = ", ".join(f"{c} = EXCLUDED.{c}" for c in update_cols)

    # Must match idx_projects_dedup_key's expressions exactly for
    # Postgres to recognize it as the ON CONFLICT target.
    sql = f"""
        INSERT INTO projects ({', '.join(db_cols)})
        VALUES %s
        ON CONFLICT (
            COALESCE(client, ''),
            COALESCE(title, ''),
            COALESCE(start_date, DATE '0001-01-01'),
            COALESCE(due_date, DATE '0001-01-01'),
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


def upsert_clients(df, conn=None):
    """Insert new client rows / update existing ones (matched by client + projek id).

    Returns (inserted_count, updated_count).
    """
    if df.empty:
        return 0, 0

    records = _records_for_insert(df, CLIENT_DB_COLUMNS)
    db_cols = [c for _, c in CLIENT_DB_COLUMNS]
    update_cols = [c for c in db_cols if c not in ("client", "projek_id")]
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
    sql = f"SELECT id, {', '.join(db_cols)} FROM projects ORDER BY id"

    with db_connection(conn) as c:
        df = pd.read_sql_query(sql, c)

    if df.empty:
        return pd.DataFrame(columns=["_row_idx"] + display_cols)

    df = df.rename(columns=dict(zip(db_cols, display_cols)))
    df = df.rename(columns={"id": "_row_idx", "Source File": "_source_file"})

    for col in ["Start date", "Due date", "Target Date"]:
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
                     AND COALESCE(start_date, DATE '0001-01-01') = COALESCE(%s::date, DATE '0001-01-01')
                     AND COALESCE(due_date, DATE '0001-01-01') = COALESCE(%s::date, DATE '0001-01-01')
                     AND COALESCE(description,'') = COALESCE(%s,'')""",
                (
                    values.get("client"), values.get("title"),
                    values.get("start_date"), values.get("due_date"),
                    values.get("description"),
                ),
            )
            dedup_seq = cur.fetchone()[0]
            cols = list(values.keys())
            col_list = ", ".join(cols + ["dedup_seq"])
            placeholders = ", ".join(["%s"] * (len(cols) + 1))
            cur.execute(
                f"INSERT INTO projects ({col_list}) VALUES ({placeholders}) RETURNING id",
                [values[c] for c in cols] + [dedup_seq],
            )
            return cur.fetchone()[0]


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
