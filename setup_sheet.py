"""
setup_sheet.py
──────────────────────────────────────────────────────────────────────────────
One-time setup script. Run this ONCE from your local machine to:
  1. Create a new Google Spreadsheet (or open an existing one by ID)
  2. Create all 5 tabs with correct headers and sample/seed data
  3. Apply light formatting (bold headers, frozen first row, column widths)
  4. Print the Sheet ID to put in your .env / st.secrets

Usage:
    python setup_sheet.py                        # creates a new sheet
    python setup_sheet.py --sheet-id <ID>        # adds tabs to existing sheet

Requirements:
    pip install gspread google-auth pandas openpyxl

Auth:
    Set GOOGLE_CREDENTIALS_PATH in .env (defaults to credentials.json).
    The service account must have Editor access to create/write sheets.
    Share the resulting sheet with the service account email as Editor too.
"""

import argparse
import os
import time

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1
from gspread_formatting import (
    CellFormat, TextFormat, Color,
    format_cell_range, set_frozen,
    set_column_width,
)

_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEET_TITLE = "DEL × TransGrid SPV — Project Dashboard"


# ── colour palette ────────────────────────────────────────────────────────────
NAVY  = Color(0.051, 0.106, 0.165)   # #0D1B2A
WHITE = Color(1, 1, 1)
GOLD  = Color(0.784, 0.588, 0.243)   # #C8963E
LGREY = Color(0.941, 0.945, 0.976)   # #F0F1F9


# ── seed data ─────────────────────────────────────────────────────────────────

CONFIG_ROWS = [
    ["Key",                  "Value"],
    ["Project Name",         "DEL × TransGrid SPV"],
    ["Client",               "DEL / TransGrid"],
    ["VAP Lead",             ""],
    ["Start Date",           "27/04/2026"],
    ["End Date",             "30/09/2026"],
    ["Project Status",       "In Progress"],
    ["Notes",                ""],
]

WORKPLAN_HEADERS = [
    "Task ID", "Workstream", "Phase", "Main Task",
    "Description of Activity", "Responsible Owner",
    "Key Deliverable", "Start Date", "End Date",
    "Duration (Wks)", "Dependencies", "Critical Path", "Status", "Notes",
]

MEETINGS_HEADERS = [
    "Title", "Date", "Time", "Type", "Location/Link", "Notes",
]

MEETINGS_SAMPLE = [
    ["JDT Working Meeting",       "16/06/2026", "10:00 AM", "In-Person",  "",  ""],
    ["Quarterly Check-In Review", "03/07/2026", "2:00 PM",  "Virtual",    "",  ""],
    ["Steering Committee Update", "07/07/2026", "9:00 AM",  "In-Person",  "",  ""],
    ["Mid-Runway Gate Review",    "24/06/2026", "11:00 AM", "Virtual",    "",  ""],
]

DELIVERABLES_HEADERS = [
    "ID", "Deliverable", "Workstream", "Owner",
    "Due Date", "Status", "Notes",
]

DELIVERABLES_SAMPLE = [
    ["D-01", "JDT Charter",               "Project Governance & Coordination", "JDT Leads",        "28/04/2026", "Completed",   ""],
    ["D-02", "Baseline Workplan",          "Project Governance & Coordination", "Fortune Beredam",  "05/05/2026", "In Progress", ""],
    ["D-03", "Feeder Simulation Report",   "Technical Design & Engineering",    "Engr Tope Opelusi","05/05/2026", "Completed",   ""],
    ["D-04", "Transformer Assessment",     "Technical Design & Engineering",    "Engr Ajibade",     "05/05/2026", "In Progress", ""],
    ["D-05", "Site Assessment Report",     "Site Assessment & Preparation",     "",                 "30/06/2026", "Not Started", ""],
    ["D-06", "Financial Model v1",         "Financing & Lender Relations",      "",                 "31/07/2026", "Not Started", ""],
    ["D-07", "Regulatory Approval Filing", "Regulatory Approvals & Permits",    "",                 "15/08/2026", "Not Started", ""],
]

INSTRUCTIONS_ROWS = [
    ["DEL × TransGrid SPV — Project Dashboard: Instructions for the Project Manager"],
    [""],
    ["IMPORTANT: Do not rename any of the sheet tabs or change column header names."],
    ["The dashboard reads data by exact column name — any change will break the connection."],
    [""],
    ["── CONFIG sheet ─────────────────────────────────────────────────────────────"],
    ["Key",           "What it does"],
    ["Project Name",  "Displayed in the dashboard header"],
    ["Client",        "Displayed in the footer"],
    ["VAP Lead",      "Internal reference — not currently displayed"],
    ["Start Date",    "Used to calculate project day number and timeline bar (DD/MM/YYYY)"],
    ["End Date",      "Used for the timeline bar (DD/MM/YYYY)"],
    ["Project Status","Free text — not currently displayed"],
    [""],
    ["── WORKPLAN sheet ───────────────────────────────────────────────────────────"],
    ["Column",            "Notes"],
    ["Task ID",           "Unique ID e.g. PG-01. Required — rows without it are ignored."],
    ["Workstream",        "The workstream group this task belongs to."],
    ["Phase",             "Phase 1 / Phase 2 etc."],
    ["Main Task",         "Short task name shown in the dashboard table."],
    ["Description",       "Longer description — not currently shown in the dashboard."],
    ["Responsible Owner", "Person or team accountable for this task."],
    ["Key Deliverable",   "Output of this task — not currently shown."],
    ["Start Date",        "DD/MM/YYYY or MM/DD/YYYY — both are parsed correctly."],
    ["End Date",          "DD/MM/YYYY or MM/DD/YYYY."],
    ["Duration (Wks)",    "Free text — not used by the dashboard."],
    ["Dependencies",      "Free text — not used by the dashboard."],
    ["Critical Path",     "Write 'GATE' here to flag a milestone task with a star badge."],
    ["Status",            "VALID VALUES: Completed | In Progress | Overdue | Not Started"],
    ["",                  "Leave blank to auto-derive from today vs Start/End Date."],
    ["Notes",             "Internal notes — not displayed."],
    [""],
    ["── MEETINGS sheet ───────────────────────────────────────────────────────────"],
    ["Column",       "Notes"],
    ["Title",        "Meeting name shown on the dashboard."],
    ["Date",         "e.g. 16/06/2026 or 'Monday, 16 June 2026' — both work."],
    ["Time",         "e.g. 10:00 AM"],
    ["Type",         "In-Person / Virtual / Hybrid"],
    ["Location/Link","Venue or video call URL — not currently shown in dashboard."],
    ["Notes",        "Internal notes."],
    [""],
    ["── DELIVERABLES sheet ───────────────────────────────────────────────────────"],
    ["Column",       "Notes"],
    ["ID",           "Unique ID e.g. D-01."],
    ["Deliverable",  "Name of the formal deliverable."],
    ["Workstream",   "Which workstream this belongs to."],
    ["Owner",        "Person responsible."],
    ["Due Date",     "DD/MM/YYYY"],
    ["Status",       "VALID VALUES: Completed | In Progress | Overdue | Not Started"],
    ["Notes",        "Internal notes."],
]


# ── formatting helpers ────────────────────────────────────────────────────────

def _fmt_header(ws, n_cols: int):
    """Bold white text on navy background for row 1."""
    hdr_range = f"A1:{rowcol_to_a1(1, n_cols)}"
    fmt = CellFormat(
        backgroundColor=NAVY,
        textFormat=TextFormat(bold=True, foregroundColor=WHITE, fontSize=10),
    )
    format_cell_range(ws, hdr_range, fmt)
    set_frozen(ws, rows=1)


def _write_tab(ss, title: str, rows: list[list], col_widths: dict | None = None):
    """
    Create or overwrite a worksheet tab with `rows` data.
    col_widths: {1: 200, 2: 120, ...}  (1-indexed column → pixel width)
    """
    try:
        ws = ss.worksheet(title)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=title, rows=max(len(rows) + 20, 100), cols=20)

    ws.update(rows, value_input_option="USER_ENTERED")

    if len(rows) > 0:
        _fmt_header(ws, len(rows[0]))

    if col_widths:
        for col_idx, width in col_widths.items():
            try:
                set_column_width(ws, rowcol_to_a1(1, col_idx)[:-1], width)
            except Exception:
                pass  # non-critical

    time.sleep(1)  # avoid quota limits
    return ws


def _seed_workplan(ss, csv_path: str = "data/workplan.csv"):
    """
    Seed the Workplan tab from the existing CSV if available,
    otherwise write headers + one example row only.
    """
    header_row = WORKPLAN_HEADERS

    if os.path.exists(csv_path):
        try:
            raw = pd.read_csv(csv_path, skiprows=2, header=1)
            raw = raw.loc[:, ~raw.columns.str.startswith("Unnamed")]
            raw.columns = raw.columns.str.strip()
            raw = raw.dropna(subset=["Task ID"])
            raw = raw[raw["Task ID"].astype(str).str.strip() != ""]

            # Normalise workstream names
            if "Workstream" in raw.columns:
                raw["Workstream"] = raw["Workstream"].astype(str).str.replace(
                    r"^\d+\.\s*", "", regex=True
                )

            # Add empty Status and Notes columns if missing
            if "Status" not in raw.columns:
                raw["Status"] = ""
            if "Notes" not in raw.columns:
                raw["Notes"] = ""

            # Select and reorder to match WORKPLAN_HEADERS
            existing_cols = [c for c in WORKPLAN_HEADERS if c in raw.columns]
            df_out = raw[existing_cols].copy()
            for col in WORKPLAN_HEADERS:
                if col not in df_out.columns:
                    df_out[col] = ""
            df_out = df_out[WORKPLAN_HEADERS]
            df_out = df_out.fillna("")

            rows = [WORKPLAN_HEADERS] + df_out.values.tolist()
            print(f"  Seeding Workplan with {len(df_out)} tasks from CSV…")
        except Exception as e:
            print(f"  Warning: could not parse CSV ({e}), using example row only.")
            rows = [header_row, ["PG-01", "Project Governance & Coordination",
                                 "Phase 1", "JDT Kick-off", "", "JDT Leads",
                                 "Minutes; signed JDT Charter", "21/04/2026",
                                 "28/04/2026", "2 wks", "", "Critical", "", ""]]
    else:
        rows = [header_row, ["PG-01", "Project Governance & Coordination",
                             "Phase 1", "JDT Kick-off", "", "JDT Leads",
                             "Minutes; signed JDT Charter", "21/04/2026",
                             "28/04/2026", "2 wks", "", "Critical", "", ""]]

    _write_tab(ss, "Workplan", rows, col_widths={
        1: 80,   # Task ID
        2: 220,  # Workstream
        3: 80,   # Phase
        4: 240,  # Main Task
        5: 300,  # Description
        6: 160,  # Responsible Owner
        7: 200,  # Key Deliverable
        8: 110,  # Start Date
        9: 110,  # End Date
        10: 90,  # Duration
        11: 120, # Dependencies
        12: 110, # Critical Path
        13: 120, # Status
        14: 200, # Notes
    })


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Populate a Google Sheet with all 5 dashboard tabs.",
        epilog="""
BEFORE RUNNING THIS SCRIPT:
  1. Go to https://sheets.google.com and create a blank spreadsheet manually.
     Name it: DEL x TransGrid SPV — Project Dashboard
  2. Copy the Sheet ID from the URL:
        https://docs.google.com/spreadsheets/d/  THIS_PART  /edit
  3. Share the sheet with your service account email (Editor access):
        You can find the service account email in credentials.json → client_email
  4. Run this script with the --sheet-id argument:
        python setup_sheet.py --sheet-id 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms
        """,
    )
    parser.add_argument(
        "--sheet-id",
        required=True,
        help="ID of the Google Sheet you created manually (from the URL).",
    )
    args = parser.parse_args()

    creds_path = os.getenv("GOOGLE_CREDENTIALS_PATH", "credentials.json")
    if not os.path.exists(creds_path):
        raise FileNotFoundError(
            f"credentials.json not found at '{creds_path}'.\n"
            "Download it from Google Cloud Console → IAM → Service Accounts → Keys.\n"
            f"Then set GOOGLE_CREDENTIALS_PATH in your .env file."
        )

    print("Authenticating…")
    creds = Credentials.from_service_account_file(creds_path, scopes=_SCOPES)
    gc = gspread.authorize(creds)

    print(f"Opening sheet: {args.sheet_id}")
    try:
        ss = gc.open_by_key(args.sheet_id)
    except gspread.exceptions.APIError as e:
        if "403" in str(e):
            print("\n❌  Permission denied (403).")
            print(f"   Make sure you shared the sheet with the service account as Editor:")
            print(f"   → {creds.service_account_email}")
        else:
            print(f"\n❌  Could not open sheet: {e}")
        raise SystemExit(1)

    print(f"\nSpreadsheet: {ss.title}")
    print(f"URL:         {ss.url}\n")

    # ── write each tab ──
    print("Writing Config…")
    _write_tab(ss, "Config", CONFIG_ROWS, col_widths={1: 180, 2: 300})

    print("Writing Workplan…")
    _seed_workplan(ss)

    print("Writing Meetings…")
    _write_tab(ss, "Meetings",
               [MEETINGS_HEADERS] + MEETINGS_SAMPLE,
               col_widths={1: 240, 2: 120, 3: 90, 4: 100, 5: 220, 6: 200})

    print("Writing Deliverables…")
    _write_tab(ss, "Deliverables",
               [DELIVERABLES_HEADERS] + DELIVERABLES_SAMPLE,
               col_widths={1: 70, 2: 260, 3: 230, 4: 160, 5: 100, 6: 120, 7: 200})

    print("Writing Instructions…")
    _write_tab(ss, "Instructions", INSTRUCTIONS_ROWS, col_widths={1: 260, 2: 460})

    # Remove the default blank Sheet1 if it still exists
    try:
        ss.del_worksheet(ss.worksheet("Sheet1"))
    except Exception:
        pass

    print("\n✅  All done!")
    print(f"\nAdd this to your .env file:")
    print(f"    GOOGLE_SHEET_ID={ss.id}")
    print(f"\nService account email (keep this — you need it to share the sheet):")
    print(f"    {creds.service_account_email}")


if __name__ == "__main__":
    main()