"""
data_loader.py
Receives raw DataFrames from sheets_client and applies:
  - date parsing
  - status resolution (Sheets Status column → auto-derive fallback)
  - derived fields (is_gate, resolved_status)
  - aggregation helpers (KPIs, workstream progress, upcoming items)

No file I/O here — all data comes from sheets_client.load_all_sheets().
The overrides.json mechanism is retired; the PM now edits Status directly
in the Workplan sheet.
"""

from datetime import date, datetime

import pandas as pd

# ── valid status values ───────────────────────────────────────────────────────
VALID_STATUSES = {"Completed", "In Progress", "Overdue", "Not Started"}


# ── date helpers ──────────────────────────────────────────────────────────────

def _parse_date(val) -> date | None:
    if pd.isna(val) or str(val).strip() == "":
        return None
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%b-%Y", "%d %b %Y"):
        try:
            return datetime.strptime(str(val).strip(), fmt).date()
        except ValueError:
            continue
    return None


# ── status logic ──────────────────────────────────────────────────────────────

def _auto_status(row, today: date) -> str:
    """Derive status from dates when no explicit Status value is set."""
    end   = row.get("_end_date")
    start = row.get("_start_date")

    if end is None:
        return "Not Started"
    if today > end:
        return "Overdue"
    if start and today >= start:
        return "In Progress"
    return "Not Started"


def _resolve_status(row, today: date) -> str:
    """
    Priority:
      1. Explicit 'Status' column in sheet (if valid value)
      2. Auto-derived from dates
    """
    raw = str(row.get("Status", "")).strip()
    if raw in VALID_STATUSES:
        return raw
    return _auto_status(row, today)


# ── workplan processor ────────────────────────────────────────────────────────

def process_workplan(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()

    # Strip whitespace from all string columns
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()

    # Drop rows with no Task ID
    df = df[df.get("Task ID", pd.Series(dtype=str)).replace("", pd.NA).notna()]
    df = df[df["Task ID"] != "nan"]

    # Normalise workstream name (strip leading "1. " numbering if present)
    if "Workstream" in df.columns:
        df["Workstream"] = df["Workstream"].str.replace(r"^\d+\.\s*", "", regex=True)

    # Parse dates into hidden columns
    df["_start_date"] = df["Start Date"].apply(_parse_date) if "Start Date" in df.columns else None
    df["_end_date"]   = df["End Date"].apply(_parse_date)   if "End Date"   in df.columns else None

    today = date.today()
    df["resolved_status"] = df.apply(lambda r: _resolve_status(r, today), axis=1)

    # Gate flag
    df["is_gate"] = df.get("Critical Path", pd.Series(dtype=str)).str.contains(
        "GATE", case=False, na=False
    )

    return df


# ── aggregations ──────────────────────────────────────────────────────────────

def kpi_counts(df: pd.DataFrame) -> dict:
    total       = len(df)
    completed   = int((df["resolved_status"] == "Completed").sum())
    overdue     = int((df["resolved_status"] == "Overdue").sum())
    in_progress = int((df["resolved_status"] == "In Progress").sum())
    outstanding = total - completed
    return {
        "total":       total,
        "completed":   completed,
        "outstanding": outstanding,
        "overdue":     overdue,
        "in_progress": in_progress,
    }


def workstream_progress(df: pd.DataFrame) -> pd.DataFrame:
    summary = (
        df.groupby("Workstream")
        .apply(lambda g: pd.Series({
            "total":       len(g),
            "completed":   int((g["resolved_status"] == "Completed").sum()),
            "in_progress": int((g["resolved_status"] == "In Progress").sum()),
            "overdue":     int((g["resolved_status"] == "Overdue").sum()),
        }), include_groups=False)
        .reset_index()
    )
    summary["pct_complete"] = (
        summary["completed"] / summary["total"] * 100
    ).round(1).fillna(0)
    return summary


def upcoming_deliverables(df: pd.DataFrame, n: int = 6) -> pd.DataFrame:
    today = date.today()
    future = df[
        df["_end_date"].notna() &
        (df["_end_date"] >= today) &
        (df["resolved_status"] != "Completed")
    ].copy()
    future = future.sort_values("_end_date").head(n)
    cols = ["Task ID", "Main Task", "Workstream", "_end_date",
            "Responsible Owner", "resolved_status"]
    return future[[c for c in cols if c in future.columns]]


# ── meetings processor ────────────────────────────────────────────────────────

def process_meetings(df_raw: pd.DataFrame) -> list[dict]:
    """
    Returns a list of meeting dicts ready for rendering.
    Expected columns: Title | Date | Time | Type | Location/Link | Notes
    """
    df = df_raw.copy()
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()

    # Drop blank rows
    df = df[df.get("Title", pd.Series(dtype=str)).replace("", pd.NA).notna()]
    df = df[df["Title"] != "nan"]

    meetings = []
    for _, row in df.iterrows():
        meetings.append({
            "title":    row.get("Title", ""),
            "date":     row.get("Date", ""),
            "time":     row.get("Time", ""),
            "type":     row.get("Type", ""),
            "location": row.get("Location/Link", ""),
            "notes":    row.get("Notes", ""),
        })
    return meetings


# ── deliverables processor ────────────────────────────────────────────────────

def process_deliverables(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Formal deliverables sheet — separate from workplan tasks.
    Expected columns: ID | Deliverable | Workstream | Owner | Due Date |
                      Status | Notes
    """
    df = df_raw.copy()
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()

    df = df[df.get("Deliverable", pd.Series(dtype=str)).replace("", pd.NA).notna()]
    df = df[df["Deliverable"] != "nan"]

    if "Due Date" in df.columns:
        df["_due_date"] = df["Due Date"].apply(_parse_date)
    else:
        df["_due_date"] = None

    today = date.today()

    def _del_status(row):
        raw = str(row.get("Status", "")).strip()
        if raw in VALID_STATUSES:
            return raw
        due = row.get("_due_date")
        if due and today > due:
            return "Overdue"
        return "Not Started"

    df["resolved_status"] = df.apply(_del_status, axis=1)
    return df
