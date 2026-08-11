"""Shared parsing/normalization helpers for ticket & project spreadsheets.

Used by both the Flask app (api/index.py) and the upload endpoint / migration
script, so a CSV upload and an Excel sheet go through the exact same
standardization before hitting the database.
"""
import re
import warnings

import pandas as pd

warnings.filterwarnings("ignore", message="Downcasting object dtype arrays")

HEADER_ROW = 1

COLORS = {
    "Completed": "#34d399",
    "Closed": "#60a5fa",
    "Pending": "#fbbf24",
    "InProgress": "#f87171",
    "Inprogress": "#f87171",
    "Open": "#a78bfa",
    "Cancelled": "#94a3b8",
    "On Hold": "#22d3ee",
}
PRIORITY_COLORS = {"High": "#f87171", "Medium": "#60a5fa", "Low": "#34d399"}
AGEING_COLORS = {"1-30 Days": "#34d399", "31-60 Days": "#fbbf24", "> 60 Days": "#f87171"}

COLUMN_MAPPING = {
    "ticket no": "Ticket No",
    "ticket number": "Ticket No",
    "ticket_no": "Ticket No",
    "ticket id": "Ticket No",
    "task type": "Task Type",
    "task_type": "Task Type",
    "project": "Project",
    "projek name": "Projek Name",
    "projek_name": "Projek Name",
    "company": "Company",
    "ticket title": "Ticket Title",
    "ticket_title": "Ticket Title",
    "ticket detail": "Ticket Detail",
    "ticket_detail": "Ticket Detail",
    "ticket desc": "Ticket Detail",
    "detail": "Ticket Detail",
    "description": "Ticket Detail",
    "ticket category": "Ticket Category",
    "ticket_category": "Ticket Category",
    "category": "Ticket Category",
    "priority": "Priority",
    "ticket created date": "Ticket Created Date",
    "ticket_created_date": "Ticket Created Date",
    "created date": "Ticket Created Date",
    "created": "Ticket Created Date",
    "date created": "Ticket Created Date",
    "ticket completed date": "Ticket Completed Date",
    "ticket_completed_date": "Ticket Completed Date",
    "completed date": "Ticket Completed Date",
    "completed": "Ticket Completed Date",
    "ticket closed date": "Ticket Closed Date",
    "ticket_closed_date": "Ticket Closed Date",
    "closed date": "Ticket Closed Date",
    "closed": "Ticket Closed Date",
    "ticket status": "Ticket Status",
    "ticket_status": "Ticket Status",
    "status": "Ticket Status",
    "sla dateline": "SLA Dateline",
    "sla_deadline": "SLA Dateline",
    "sla": "SLA Dateline",
    "sla late": "SLA Late",
    "sla_breach": "SLA Late",
    "days": "Days",
    "ageing": "Ageing",
    "aging": "Ageing",
    "client": "Client",
}

# Ticket dataframe columns, in the order they're stored in the DB.
TICKET_COLUMNS = [
    "Client", "Ticket No", "Task Type", "Project", "Company", "Ticket Title",
    "Ticket Detail", "Ticket Category", "Priority", "Ticket Created Date",
    "Ticket Completed Date", "Ticket Closed Date", "Ticket Status",
    "SLA Dateline", "SLA Late", "Days", "Ageing", "Days to Close",
    "SLA Breach", "Source File",
]

PROJECT_COLUMNS = [
    "Client", "Title", "Projek Name", "Description", "Category", "Progress", "Priority",
    "Start date", "Due date", "Target Date", "Duration", "Assigned to",
    "Status Progress", "Percentage", "Overall Progress Task (%)", "Source File",
    "Dedup Seq",
]

CLIENT_COLUMNS = [
    "Client", "Projek ID", "Projek Name", "Projek Status",
    "Start Date", "End Date", "Source File",
]


def standardize_columns(df):
    col_map = {}
    for col in df.columns:
        cl = str(col).lower().strip()
        if cl in COLUMN_MAPPING:
            col_map[col] = COLUMN_MAPPING[cl]
    df = df.rename(columns=col_map)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df = df.loc[:, ~df.columns.duplicated()]
    return df


def convert_dtypes(df):
    for date_col in ["Ticket Created Date", "Ticket Completed Date", "Ticket Closed Date", "SLA Dateline"]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)

    if "Ticket Status" in df.columns:
        ts = df["Ticket Status"]
        if isinstance(ts, pd.DataFrame):
            ts = ts.iloc[:, 0]
        df["Ticket Status"] = ts.astype(str).str.strip()
        df["Ticket Status"] = df["Ticket Status"].replace({
            "InProgress": "In Progress",
            "Inprogress": "In Progress",
            "nan": None,
        })

    if "Priority" in df.columns:
        pr = df["Priority"]
        if isinstance(pr, pd.DataFrame):
            pr = pr.iloc[:, 0]
        df["Priority"] = pr.astype(str).str.strip().str.title()
        df["Priority"] = df["Priority"].replace({"Nan": None, "None": None})

    if "Ticket Created Date" in df.columns and "Ticket Closed Date" in df.columns:
        mask = df["Ticket Created Date"].notna() & df["Ticket Closed Date"].notna()
        df.loc[mask, "Days to Close"] = (df.loc[mask, "Ticket Closed Date"] - df.loc[mask, "Ticket Created Date"]).dt.days
    elif "Ticket Created Date" in df.columns and "Ticket Completed Date" in df.columns:
        mask = df["Ticket Created Date"].notna() & df["Ticket Completed Date"].notna()
        df.loc[mask, "Days to Close"] = (df.loc[mask, "Ticket Completed Date"] - df.loc[mask, "Ticket Created Date"]).dt.days

    if "Days" in df.columns:
        df["Days"] = pd.to_numeric(df["Days"], errors="coerce")

    if "Ageing" in df.columns:
        df["Ageing"] = df["Ageing"].astype(str).str.strip().replace({
            "1-30 days": "1-30 Days",
            "1-30 Days": "1-30 Days",
            "30-60 Days": "31-60 Days",
            "30-60 days": "31-60 Days",
            "31-60 Days": "31-60 Days",
            "31-60 days": "31-60 Days",
            ">60 Days": "> 60 Days",
            ">60 days": "> 60 Days",
            "> 60 Days": "> 60 Days",
            "> 60 days": "> 60 Days",
            "nan": None,
        })
    if "Ageing" not in df.columns:
        if "Days" in df.columns:
            def classify(d):
                if pd.isna(d):
                    return None
                if d <= 30:
                    return "1-30 Days"
                elif d <= 60:
                    return "31-60 Days"
                else:
                    return "> 60 Days"
            df["Ageing"] = df["Days"].apply(classify)
        elif "Ticket Created Date" in df.columns:
            days_open = (pd.Timestamp.now() - df["Ticket Created Date"]).dt.days
            def classify2(d):
                if pd.isna(d):
                    return None
                if d <= 30:
                    return "1-30 Days"
                elif d <= 60:
                    return "31-60 Days"
                else:
                    return "> 60 Days"
            df["Ageing"] = days_open.apply(classify2)

    if "SLA Late" in df.columns:
        sla = df["SLA Late"]
        if isinstance(sla, pd.DataFrame):
            sla = sla.iloc[:, 0]
        df["SLA Late"] = sla.astype(str).str.strip()
        # SLA Late holds a numeric days-remaining value; negative means
        # the ticket blew past its deadline.
        sla_num = pd.to_numeric(df["SLA Late"].replace({"nan": None}), errors="coerce")
        df["SLA Breach"] = sla_num < 0
    else:
        df["SLA Breach"] = False

    return df


def parse_ticket_sheet(df, client, source_file):
    """Standardize a raw ticket sheet/CSV into the canonical ticket dataframe.

    Returns (parsed_df, info) where info reports what got left behind, so
    callers can surface it instead of rows/columns silently vanishing:
      - rows_dropped: rows with no Ticket No at all (blank separator/
        subtotal rows that survive the initial dropna(how="all") because
        some other cell in the row is filled in).
      - unmapped_columns: source columns that didn't match anything in
        COLUMN_MAPPING and aren't part of the fixed ticket schema, so
        their data isn't stored (e.g. a "Remarks" or "Assigned To" column
        the spreadsheet has that this dashboard has no field for).
    """
    df = df.dropna(how="all")
    df = standardize_columns(df)
    if df.empty:
        return df, {"rows_dropped": 0, "unmapped_columns": []}

    # "Projek Name" is the same concept as this sheet's own "Project"
    # column under a different header some source CSVs use -- fold it in
    # here rather than in the shared COLUMN_MAPPING, since the Client
    # sheet's own genuinely-distinct "Projek Name" field (a different
    # concept there) would otherwise get renamed away to "Project" too
    # and go missing on Client-sheet uploads.
    if "Projek Name" in df.columns:
        if "Project" in df.columns:
            df["Project"] = df["Project"].fillna(df["Projek Name"])
        else:
            df = df.rename(columns={"Projek Name": "Project"})
        df = df.drop(columns=["Projek Name"], errors="ignore")

    if "Client" not in df.columns or df["Client"].isna().all():
        df["Client"] = client
    else:
        df["Client"] = df["Client"].fillna(client)
    df["Source File"] = source_file

    df = convert_dtypes(df)

    unmapped_columns = [c for c in df.columns if c not in TICKET_COLUMNS]

    if "Ticket No" in df.columns:
        tn = df["Ticket No"]
        if isinstance(tn, pd.DataFrame):
            tn = tn.iloc[:, 0]
    else:
        tn = pd.Series([None] * len(df), index=df.index)

    # tn.notna() catches every pandas null sentinel (None, NaN, NaT,
    # pandas.NA) regardless of the column's inferred dtype; the string
    # checks additionally catch a cell that's literally the text "nan"/
    # "none". Relying on the string form alone previously let a
    # not-actually-empty-looking NA slip through on at least one real
    # upload.
    tn_str = tn.astype(str).str.strip()
    valid = tn.notna() & tn_str.ne("") & tn_str.str.lower().ne("nan") & tn_str.str.lower().ne("none")
    rows_dropped = int((~valid).sum())
    df = df[valid]

    for col in TICKET_COLUMNS:
        if col not in df.columns:
            df[col] = None

    parsed = df[TICKET_COLUMNS]

    # Belt-and-suspenders: tickets.ticket_no is NOT NULL in Postgres, and
    # a single row violating that fails the *entire* batch insert -- not
    # just that row -- silently taking every other valid row in the file
    # down with it. Guarantee nothing null reaches the database no
    # matter what slipped past the check above.
    still_null = parsed["Ticket No"].isna()
    if still_null.any():
        rows_dropped += int(still_null.sum())
        parsed = parsed[~still_null]

    return parsed, {"rows_dropped": rows_dropped, "unmapped_columns": unmapped_columns}


def _scale_percentage(series):
    """Excel's native percentage format stores 50% as the float 0.5, not
    50 -- there's no "%" character to strip, so pd.to_numeric alone leaves
    it as a 0-1 fraction. Scale fractional values up to the 0-100 range
    the UI expects; leave anything already >1 (e.g. entered as "80")
    alone.
    """
    numeric = pd.to_numeric(series.astype(str).str.replace("%", ""), errors="coerce")
    return numeric.apply(lambda v: v * 100 if pd.notna(v) and v <= 1 else v)


def parse_project_sheet(df, source_file):
    """Standardize a raw 'Client Project' sheet.

    Returns (parsed_df, info); info["unmapped_columns"] lists source
    columns with no matching field (dropped silently before), so a
    column like a spreadsheet's own "Notes" shows up instead of just
    disappearing.
    """
    df = df.dropna(how="all").copy()
    # parse_project_sheet doesn't go through standardize_columns (unlike
    # ticket/client sheets) since its column set is fixed/positional, but
    # "Projek Name" is a late addition some source sheets may not spell
    # exactly right -- normalize any casing/underscore variant.
    rename_map = {c: "Projek Name" for c in df.columns if str(c).strip().lower() in ("projek name", "projek_name")}
    if rename_map:
        df = df.rename(columns=rename_map)
    if "Client" in df.columns:
        df["Client"] = df["Client"].ffill()
    df["Source File"] = source_file

    # A project's Title, and every other block-level field (Category,
    # Priority, dates, Assigned to, Status Progress, the overall
    # percentages...) only appear on the row where that task starts.
    # The sub-item/checklist rows underneath it (e.g. a "Pre UAT"
    # breakdown split across several rows, one bullet per row) leave
    # every column but Description blank -- same as a merged cell would.
    # Two problems came from not carrying those down: (1) two
    # *different* projects that happen to share identical boilerplate
    # checklist text ("Sign-off - 18/06-21/06") both collapsed to
    # Title=NULL and became indistinguishable from each other, and
    # (2) those rows showed as a wall of empty cells in the UI even
    # though they clearly belong to a specific task with a real
    # category/priority/date range/owner.
    #
    # Capture which rows were originally sub-items *before* filling
    # anything in -- a row with no Title and no Description is a pure
    # spacer (leftover row-height padding, not real data) and needs
    # dropping, but that distinction disappears once Title is filled.
    was_subitem = df["Title"].isna() if "Title" in df.columns else pd.Series(False, index=df.index)

    if "Title" in df.columns and "Client" in df.columns:
        df["Title"] = df.groupby("Client")["Title"].ffill()

    is_pure_spacer = was_subitem & (df["Description"].isna() if "Description" in df.columns else True)
    df = df[~is_pure_spacer]

    # Now that Title reflects the real parent task, fill the rest of
    # that task's block-level columns down onto its sub-item rows too
    # -- scoped to (Client, Title) so it can never bleed from one
    # project into the next. ffill only ever fills a cell that's
    # already blank, so a row with its own genuine value (e.g. every
    # numbered task already has its own date range) is left untouched.
    block_cols = [c for c in [
        "Projek Name", "Category", "Progress", "Priority", "Start date", "Due date", "Tempoh",
        "Target Date", "Assigned to", "Status Progress", "Percentage",
        "Overall Progress Task (%)",
    ] if c in df.columns]
    if block_cols and "Title" in df.columns and "Client" in df.columns:
        df[block_cols] = df.groupby(["Client", "Title"])[block_cols].transform(lambda s: s.ffill())

    if "Tempoh" in df.columns:
        df["Duration"] = df["Tempoh"].astype(str)
        df.loc[df["Tempoh"].isna(), "Duration"] = None

    for c in ["Start date", "Due date", "Target Date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)

    # SQL's NULL is never equal to NULL, even inside a UNIQUE constraint --
    # so rows with a blank title/date (common for sub-item description
    # lines, e.g. a checklist under a numbered task) never match an
    # existing row on re-upload no matter how many content columns are
    # in the key, and just keep multiplying. A per-sheet occurrence
    # counter (never NULL) fixes that: as long as the sheet's row order
    # is unchanged between uploads, the same row gets the same sequence
    # number and updates in place instead of inserting a duplicate.
    key_basis = df[["Title", "Start date", "Due date", "Description"]] if "Description" in df.columns else df[["Title", "Start date", "Due date"]]
    df["Dedup Seq"] = key_basis.astype(str).groupby(list(key_basis.columns)).cumcount()

    if "Percentage" in df.columns:
        df["Percentage"] = _scale_percentage(df["Percentage"])
    if "Overall Progress Task (%)" in df.columns:
        df["Overall Progress Task (%)"] = _scale_percentage(df["Overall Progress Task (%)"])

    unmapped_columns = [
        c for c in df.columns
        if c not in PROJECT_COLUMNS and c != "Tempoh" and not str(c).startswith("Unnamed")
    ]

    for col in PROJECT_COLUMNS:
        if col not in df.columns:
            df[col] = None

    return df[PROJECT_COLUMNS], {"unmapped_columns": unmapped_columns}


def parse_client_sheet(df, source_file):
    """Standardize the workbook's 'Client' sheet into the canonical client dataframe.

    The sheet lists every project under its owning client (one row per
    Projek ID) with its status and, where filled, contract start/end
    dates. Returns (parsed_df, info) using the same diagnostics shape as
    the ticket/project parsers so uploads report what got left behind.
    """
    df = df.dropna(how="all").copy()
    df = standardize_columns(df)
    if df.empty:
        return df, {"rows_dropped": 0, "unmapped_columns": []}

    df["Source File"] = source_file

    if "Client" in df.columns:
        df["Client"] = df["Client"].ffill()

    for c in ["Start Date", "End Date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)

    unmapped_columns = [c for c in df.columns if c not in CLIENT_COLUMNS]

    rows_dropped = 0
    if "Projek ID" in df.columns:
        pid = df["Projek ID"]
        if isinstance(pid, pd.DataFrame):
            pid = pid.iloc[:, 0]
        pid_str = pid.astype(str).str.strip()
        valid = pid.notna() & pid_str.ne("") & pid_str.str.lower().ne("nan")
        rows_dropped = int((~valid).sum())
        df = df[valid]
    else:
        df["Projek ID"] = None

    for col in CLIENT_COLUMNS:
        if col not in df.columns:
            df[col] = None

    return df[CLIENT_COLUMNS], {"rows_dropped": rows_dropped, "unmapped_columns": unmapped_columns}


def detect_ticket_sheets(filepath_or_buffer):
    """Return {sheet_name: header_row} for sheets that look like ticket sheets.

    Tries HEADER_ROW first (row 1, matching the original bundled workbook's
    layout) then falls back to rows 0 and 2, since a sheet exported from a
    different tool can put the real header one row up or down -- previously
    a sheet like that wasn't recognized as a ticket sheet at all and its
    entire client's worth of tickets went missing from the upload with no
    indication why.
    """
    xl = pd.ExcelFile(filepath_or_buffer, engine="openpyxl")
    ticket_sheets = {}
    for name in xl.sheet_names:
        for header_row in (HEADER_ROW, 0, 2):
            try:
                df_head = pd.read_excel(filepath_or_buffer, sheet_name=name, header=header_row, engine="openpyxl", nrows=3)
            except Exception:
                continue
            # pandas dedupes repeated header names as "Ticket No.1",
            # "Ticket No.2", ... -- strip that suffix before checking so a
            # printed/aggregate report with the same block of columns
            # repeated side by side several times is actually recognized
            # as repeated, not read as a single "Ticket No" column.
            base_names = [re.sub(r"\.\d+$", "", c) for c in (str(c).lower().strip() for c in df_head.columns)]
            ticket_no_matches = sum(1 for c in base_names if COLUMN_MAPPING.get(c) == "Ticket No")
            # More than one match means a repeated-block layout, which
            # isn't a one-row-per-ticket sheet and would wrongly become a
            # "client" named after the sheet -- skip it.
            if ticket_no_matches == 1:
                ticket_sheets[name] = header_row
                break
    return ticket_sheets
