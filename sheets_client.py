"""
sheets_client.py
Thin wrapper around gspread that reads all 5 sheets into DataFrames.

Auth:  service account via credentials.json  (or st.secrets on Cloud)
Env:   GOOGLE_CREDENTIALS_PATH=credentials.json   (default)
       GOOGLE_SHEET_ID=<your-sheet-id>             (required)
"""

import json
import os

# Load .env file for local development (no-op on Streamlit Cloud)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # python-dotenv not installed — relying on env vars being set another way

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

# ── scopes ────────────────────────────────────────────────────────────────────
_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ── sheet tab names (must match exactly) ──────────────────────────────────────
TAB_CONFIG        = "Config"
TAB_WORKPLAN      = "Workplan"
TAB_MEETINGS      = "Meetings"
TAB_DELIVERABLES  = "Deliverables"
TAB_INSTRUCTIONS  = "Instructions"


# ── auth ──────────────────────────────────────────────────────────────────────

def _get_credentials() -> Credentials:
    """
    Try st.secrets first (Streamlit Community Cloud), then fall back to
    local credentials.json path from env / default.
    """
    # 1. Streamlit secrets (Cloud deployment)
    try:
        info = dict(st.secrets["gcp_service_account"])
        return Credentials.from_service_account_info(info, scopes=_SCOPES)
    except (KeyError, FileNotFoundError):
        pass

    # 2. Local credentials.json
    creds_path = os.getenv("GOOGLE_CREDENTIALS_PATH", "credentials.json")
    if not os.path.exists(creds_path):
        raise FileNotFoundError(
            f"credentials.json not found at '{creds_path}'.\n"
            "Set GOOGLE_CREDENTIALS_PATH in your .env or place credentials.json "
            "in the project root."
        )
    return Credentials.from_service_account_file(creds_path, scopes=_SCOPES)


def _get_sheet_id() -> str:
    try:
        return st.secrets["GOOGLE_SHEET_ID"]
    except (KeyError, FileNotFoundError):
        pass
    sid = os.getenv("GOOGLE_SHEET_ID", "")
    if not sid:
        raise ValueError(
            "GOOGLE_SHEET_ID is not set.\n"
            "Add it to your .env file or st.secrets."
        )
    return sid


@st.cache_resource
def _get_client() -> gspread.Client:
    return gspread.authorize(_get_credentials())


def _open_sheet() -> gspread.Spreadsheet:
    return _get_client().open_by_key(_get_sheet_id())


# ── readers ───────────────────────────────────────────────────────────────────

def _ws_to_df(spreadsheet: gspread.Spreadsheet, tab: str) -> pd.DataFrame:
    """Read a worksheet tab into a DataFrame using the first row as header."""
    ws = spreadsheet.worksheet(tab)
    records = ws.get_all_records(expected_headers=[], numericise_ignore=["all"])
    return pd.DataFrame(records)


def read_config(spreadsheet: gspread.Spreadsheet) -> dict:
    """
    Config sheet has two columns: Key | Value
    Returns a plain dict.
    """
    df = _ws_to_df(spreadsheet, TAB_CONFIG)
    df.columns = [c.strip() for c in df.columns]
    df = df[df["Key"].astype(str).str.strip() != ""]
    return dict(zip(df["Key"].str.strip(), df["Value"].astype(str).str.strip()))


def read_workplan(spreadsheet: gspread.Spreadsheet) -> pd.DataFrame:
    return _ws_to_df(spreadsheet, TAB_WORKPLAN)


def read_meetings(spreadsheet: gspread.Spreadsheet) -> pd.DataFrame:
    return _ws_to_df(spreadsheet, TAB_MEETINGS)


def read_deliverables(spreadsheet: gspread.Spreadsheet) -> pd.DataFrame:
    return _ws_to_df(spreadsheet, TAB_DELIVERABLES)


# ── main cached loader ────────────────────────────────────────────────────────

@st.cache_data(show_spinner="Loading data from Google Sheets…")
def load_all_sheets() -> dict:
    """
    Returns a dict with keys: config, workplan, meetings, deliverables.
    Call st.cache_data.clear() to force a refresh.
    """
    ss = _open_sheet()
    return {
        "config":       read_config(ss),
        "workplan":     read_workplan(ss),
        "meetings":     read_meetings(ss),
        "deliverables": read_deliverables(ss),
    }