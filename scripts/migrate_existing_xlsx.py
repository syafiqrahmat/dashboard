"""One-time seed: load the bundled Excel workbook into Neon Postgres.

Run once locally (with DATABASE_URL pointed at your Neon database) to
carry over whatever data already exists in api/*.xlsx. After this, the
app reads and writes the database, and only touches the workbook again
through the Task Detail tab's explicit "Save to Excel" button.

Usage:
    pip install -r requirements.txt
    python scripts/migrate_existing_xlsx.py                # every api/*.xlsx
    python scripts/migrate_existing_xlsx.py path/to/file.xlsx
"""
import glob
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "api"))

from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), "..", ".env.local"))

import pandas as pd

import db
from data_utils import (
    detect_ticket_sheets, parse_ticket_sheet, parse_project_sheet,
    parse_client_sheet, parse_milestone_sheet, parse_task_detail_sheet,
)


def find_default_workbooks():
    """Every workbook under api/, so a bare `python scripts/migrate_existing_xlsx.py`
    seeds all of them -- the ticket data and the Task Detail sheet live in
    different files, and seeding only one would leave the other empty."""
    api_dir = os.path.join(os.path.dirname(__file__), "..", "api")
    matches = glob.glob(os.path.join(api_dir, "*.xlsx"))
    matches = [m for m in matches if not os.path.basename(m).startswith("~")]
    return sorted(matches)


def lookup_client_for_project(projek_name, parsed_client_df=None):
    """Client that owns a project, for the '<project> TASK DETAIL' sheet.
    Prefers the Client sheet just parsed from this same workbook (it is the
    newer mapping) and falls back to the clients table."""
    if parsed_client_df is not None and not parsed_client_df.empty and "Projek Name" in parsed_client_df.columns:
        names = parsed_client_df["Projek Name"].astype(str).str.strip()
        hit = parsed_client_df.loc[names == projek_name, "Client"]
        if len(hit):
            value = str(hit.iloc[0]).strip()
            if value and value.lower() not in ("nan", "none"):
                return value
    return db.find_client_for_project(projek_name)


def seed_workbook(filepath):
    fname = os.path.basename(filepath)
    sheets = detect_ticket_sheets(filepath)
    print(f"  Ticket sheets found: {sheets}")

    total_ins = total_upd = 0
    all_unmapped = set()
    for sheet_name, header_row in sheets.items():
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=header_row, engine="openpyxl")
        parsed, diag = parse_ticket_sheet(df, client=sheet_name, source_file=fname)
        all_unmapped.update(diag["unmapped_columns"])
        if parsed.empty:
            print(f"  {sheet_name}: 0 rows, skipping")
            continue
        ins, upd = db.upsert_tickets(parsed)
        total_ins += ins
        total_upd += upd
        dropped_note = f", {diag['rows_dropped']} row(s) skipped (no Ticket No)" if diag["rows_dropped"] else ""
        print(f"  {sheet_name}: {len(parsed)} rows -> {ins} inserted, {upd} updated{dropped_note}")

    if all_unmapped:
        print(f"\nColumns present in the workbook but not stored (no field for them): {sorted(all_unmapped)}")

    xl = pd.ExcelFile(filepath, engine="openpyxl")
    if "Client Project" in xl.sheet_names:
        pdf = pd.read_excel(filepath, sheet_name="Client Project", header=0, engine="openpyxl")
        parsed_p, diag_p = parse_project_sheet(pdf, source_file=fname)
        if diag_p["unmapped_columns"]:
            print(f"  Client Project: columns not stored: {diag_p['unmapped_columns']}")
        if not parsed_p.empty:
            ins_p, upd_p = db.upsert_projects(parsed_p)
            print(f"  Client Project: {len(parsed_p)} rows -> {ins_p} inserted, {upd_p} updated")

    parsed_c = None
    if "Client" in xl.sheet_names:
        cdf = pd.read_excel(filepath, sheet_name="Client", header=0, engine="openpyxl")
        parsed_c, diag_c = parse_client_sheet(cdf, source_file=fname)
        if diag_c["unmapped_columns"]:
            print(f"  Client: columns not stored: {diag_c['unmapped_columns']}")
        dropped_note = f", {diag_c['rows_dropped']} row(s) skipped (no Projek ID)" if diag_c["rows_dropped"] else ""
        if not parsed_c.empty:
            ins_c, upd_c = db.upsert_clients(parsed_c)
            print(f"  Client: {len(parsed_c)} rows -> {ins_c} inserted, {upd_c} updated{dropped_note}")

    milestone_sheet = next(
        (s for s in xl.sheet_names if str(s).strip().upper() == "PROJECT MILESTONE"),
        None,
    )
    if milestone_sheet:
        mdf = pd.read_excel(filepath, sheet_name=milestone_sheet, header=0, engine="openpyxl")
        parsed_m, diag_m = parse_milestone_sheet(mdf, source_file=fname)
        if diag_m["unmapped_columns"]:
            print(f"  {milestone_sheet}: columns not stored: {diag_m['unmapped_columns']}")
        if not parsed_m.empty:
            ins_m, upd_m = db.upsert_project_milestones(parsed_m)
            dropped_note = f", {diag_m['rows_dropped']} row(s) skipped (no Task Name)" if diag_m["rows_dropped"] else ""
            print(f"  {milestone_sheet}: {len(parsed_m)} rows -> {ins_m} inserted, {upd_m} updated{dropped_note}")

    # "<project> TASK DETAIL" -- no Client/Projek Name column on the sheet,
    # so both are derived from the sheet name + the Client mapping above.
    task_sheet = next(
        (s for s in xl.sheet_names if str(s).strip().upper().endswith(" TASK DETAIL")),
        None,
    )
    if task_sheet:
        projek_name = re.sub(r"\s+TASK DETAIL\s*$", "", str(task_sheet).strip(), flags=re.IGNORECASE).strip()
        client = lookup_client_for_project(projek_name, parsed_c)
        if not client:
            print(f"  {task_sheet}: no client found for project {projek_name!r} -- skipped")
        else:
            tdf = pd.read_excel(filepath, sheet_name=task_sheet, header=0, engine="openpyxl")
            parsed_t, diag_t = parse_task_detail_sheet(tdf, fname, client, projek_name)
            if diag_t["unmapped_columns"]:
                print(f"  {task_sheet}: columns not stored: {diag_t['unmapped_columns']}")
            if not parsed_t.empty:
                ins_t, upd_t = db.upsert_project_task_details(parsed_t)
                dropped_note = f", {diag_t['rows_dropped']} row(s) skipped (no Proses)" if diag_t["rows_dropped"] else ""
                print(f"  {task_sheet}: {len(parsed_t)} rows -> {ins_t} inserted, {upd_t} updated{dropped_note}")

    print(f"\nDone. Tickets: {total_ins} inserted, {total_upd} updated.")


def main():
    paths = sys.argv[1:] if len(sys.argv) > 1 else find_default_workbooks()
    paths = [p for p in paths if os.path.exists(p)]
    if not paths:
        print("No workbook found. Pass a path: python scripts/migrate_existing_xlsx.py <file.xlsx>")
        sys.exit(1)

    print("Connecting to database and ensuring schema exists...")
    db.init_schema()

    for filepath in paths:
        print(f"\n=== Reading: {filepath} ===")
        seed_workbook(filepath)

    counts = db.get_counts()
    print(f"\nDatabase now has {counts['tickets']} tickets, {counts['projects']} project rows "
          f"and {counts['clients']} client rows.")


if __name__ == "__main__":
    main()
