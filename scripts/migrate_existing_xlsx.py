"""One-time seed: load the bundled Excel workbook into Neon Postgres.

Run once locally (with DATABASE_URL pointed at your Neon database) to
carry over whatever data already exists in api/*.xlsx. After this, the
app reads/writes the database exclusively -- the workbook is no longer
touched at runtime.

Usage:
    pip install -r requirements.txt
    python scripts/migrate_existing_xlsx.py [path/to/file.xlsx]
"""
import glob
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "api"))

from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), "..", ".env.local"))

import pandas as pd

import db
from data_utils import (
    detect_ticket_sheets, parse_ticket_sheet, parse_project_sheet,
    parse_client_sheet, parse_milestone_sheet,
)


def find_default_workbook():
    api_dir = os.path.join(os.path.dirname(__file__), "..", "api")
    matches = glob.glob(os.path.join(api_dir, "*.xlsx"))
    matches = [m for m in matches if not os.path.basename(m).startswith("~")]
    return matches[0] if matches else None


def main():
    filepath = sys.argv[1] if len(sys.argv) > 1 else find_default_workbook()
    if not filepath or not os.path.exists(filepath):
        print("No workbook found. Pass a path: python scripts/migrate_existing_xlsx.py <file.xlsx>")
        sys.exit(1)

    fname = os.path.basename(filepath)
    print(f"Connecting to database and ensuring schema exists...")
    db.init_schema()

    print(f"Reading: {filepath}")
    sheets = detect_ticket_sheets(filepath)
    print(f"Ticket sheets found: {sheets}")

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

    print(f"\nDone. Tickets: {total_ins} inserted, {total_upd} updated.")
    counts = db.get_counts()
    print(f"Database now has {counts['tickets']} tickets, {counts['projects']} project rows and {counts['clients']} client rows.")


if __name__ == "__main__":
    main()
