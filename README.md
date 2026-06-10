# DEL × TransGrid SPV — Monitoring Dashboard

A Streamlit monitoring dashboard for the DEL × TransGrid SPV project,
built for Vista Advisory Partners.

## Project structure

```
del_transgrid_dashboard/
├── app.py               # Main Streamlit app
├── data_loader.py       # CSV parsing, status logic, override management
├── requirements.txt     # Python dependencies
├── overrides.json       # Auto-created on first in-app status override
├── data/
│   └── workplan.csv     # Source workplan (update this to refresh data)
└── .streamlit/
    └── config.toml      # Theme & server config
```

## Status logic

Status is resolved in this priority order:

1. **In-app manual override** — saved to `overrides.json`, wins always
2. **CSV `Status` column** — if the column exists and the cell is non-empty
3. **Auto-derived from dates:**
   - `Completed` → `Completed` flag in CSV is `yes/true/1/✓/done`
   - Past end date & not completed → `Overdue`
   - Between start and end date → `In Progress`
   - Before start date → `Not Started`

## Adding a Status column to the CSV

Add a column named `Status` to `workplan.csv`. Valid values:
`Completed`, `In Progress`, `Overdue`, `Not Started`

Leave the cell blank to fall back to auto-derive.

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploying to Streamlit Community Cloud

1. Push this folder to a **public GitHub repo** (or private with Streamlit Cloud access)
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Point to your repo, branch `main`, main file `app.py`
4. Click **Deploy** — done

> **Note on overrides.json:** Streamlit Community Cloud has an ephemeral
> filesystem — manual overrides saved via the sidebar will reset on each
> redeploy. For persistence across deploys, consider adding a `Status`
> column directly to the CSV and committing it, or wiring up
> `st.secrets` + a lightweight DB (e.g. Supabase free tier).

## Updating the workplan

Replace `data/workplan.csv` with the latest export from your workplan
Excel file. The app re-parses on each page load (cached for 60 seconds).
