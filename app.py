"""
app.py  —  DEL × TransGrid SPV Monitoring Dashboard
Data source: Google Sheets via sheets_client.load_all_sheets()
"""

import base64
import math
import os
from datetime import date, datetime

import pandas as pd
import streamlit as st

from sheets_client import load_all_sheets
from data_loader import (
    process_workplan,
    process_meetings,
    process_deliverables,
    kpi_counts,
    workstream_progress,
    upcoming_deliverables,
)

# ── date helpers (cross-platform, no %-d) ─────────────────────────────────────
def fmt_date(d) -> str:
    return f"{d.day} {d.strftime('%B %Y')}"

def fmt_short(d) -> str:
    return f"{d.day} {d.strftime('%b')}, {d.strftime('%Y')}"

def fmt_weekday_short(d) -> str:
    return f"{d.strftime('%a')}, {d.day} {d.strftime('%b %Y')}"

def _parse_config_date(val: str) -> date | None:
    for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(val.strip(), fmt).date()
        except (ValueError, AttributeError):
            continue
    return None

# ── logo helper ───────────────────────────────────────────────────────────────
def _logo_img_tag(height: int = 36) -> str:
    for fname in ("logo.png", "logo.svg", "logo.jpg", "logo.jpeg"):
        fpath = os.path.join(os.path.dirname(__file__), fname)
        if os.path.exists(fpath):
            ext = fname.rsplit(".", 1)[-1].lower()
            mime = "image/svg+xml" if ext == "svg" else f"image/{ext}"
            with open(fpath, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            return f'<img src="data:{mime};base64,{b64}" style="height:{height}px;width:auto;display:block" alt="Logo">'
    return '<div class="navbar-logo">VAP</div>'

# ── page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DEL × TransGrid | Monitoring Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700&display=swap');

html, body, [class*="css"], .stApp { font-family: 'Inter', sans-serif; }
.stApp { background: #F4F6F9; }
.block-container { padding: 3rem 2rem 2rem 2rem !important; max-width: 100% !important; }

/* hide only Streamlit branding */
#MainMenu { visibility: hidden; }
footer    { visibility: hidden; }
[data-testid="stToolbar"]    { visibility: hidden; }
[data-testid="stDecoration"] { visibility: hidden; }
.stDeployButton { display: none !important; }

.meta-strip {
    background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 10px;
    padding: 14px 24px; display: flex; align-items: center; gap: 40px; margin-bottom: 20px;
}
.meta-field label { font-size: 0.68rem; text-transform: uppercase; letter-spacing: 0.08em; color: #94A3B8; display: block; margin-bottom: 2px; }
.meta-field span  { font-size: 0.88rem; font-weight: 600; color: #0D1B2A; }

/* compact top strip */
.topstrip {
    background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 10px;
    padding: 10px 20px; display: flex; align-items: center;
    gap: 0; margin-bottom: 20px;
}
.topstrip-brand { display: flex; align-items: center; gap: 10px; flex: 0 0 auto; padding-right: 24px; border-right: 1px solid #E2E8F0; }
.topstrip-brand-title { font-family: 'DM Sans',sans-serif; font-size: 0.95rem; font-weight: 600; color: #0D1B2A; white-space: nowrap; }
.topstrip-fields { display: flex; align-items: center; gap: 0; flex: 1; }
.topstrip-field { padding: 0 20px; border-right: 1px solid #E2E8F0; }
.topstrip-field:last-child { border-right: none; }
.topstrip-field-label { font-size: 0.65rem; text-transform: uppercase; letter-spacing: 0.08em; color: #94A3B8; margin-bottom: 2px; }
.topstrip-field-value { font-size: 0.82rem; font-weight: 600; color: #0D1B2A; white-space: nowrap; }
.topstrip-date { margin-left: auto; font-size: 0.75rem; color: #64748B; white-space: nowrap; padding-left: 20px; }

.kpi-row { display: flex; gap: 16px; margin-bottom: 20px; }
.kpi-card { flex: 1; background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 10px; padding: 18px 20px; position: relative; overflow: hidden; }
.kpi-card::after { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; }
.kpi-card.c-gold::after  { background: #C8963E; }
.kpi-card.c-green::after { background: #22C55E; }
.kpi-card.c-blue::after  { background: #3B82F6; }
.kpi-card.c-red::after   { background: #EF4444; }
.kpi-card-label { font-size: 0.7rem; text-transform: uppercase; letter-spacing: 0.08em; color: #94A3B8; font-weight: 600; margin-bottom: 6px; }
.kpi-card-value { font-family: 'DM Sans', sans-serif; font-size: 2rem; font-weight: 700; color: #0D1B2A; line-height: 1; }
.kpi-card-sub   { font-size: 0.71rem; color: #94A3B8; margin-top: 5px; }

.card { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 10px; padding: 22px 24px; margin-bottom: 20px; height: 100%; box-sizing: border-box; }
.card-title { font-family: 'DM Sans', sans-serif; font-size: 0.9rem; font-weight: 600; color: #0D1B2A; margin-bottom: 18px; }

.donut-wrap { display: flex; align-items: center; gap: 28px; }
.legend { display: flex; flex-direction: column; gap: 10px; }
.legend-row { display: flex; align-items: center; gap: 9px; font-size: 0.82rem; color: #374151; }
.legend-dot { width: 9px; height: 9px; border-radius: 50%; flex-shrink: 0; }
.legend-pct { font-weight: 700; margin-left: 4px; }

.ws-bar-wrap { margin-bottom: 14px; }
.ws-bar-label { display: flex; justify-content: space-between; font-size: 0.8rem; color: #374151; margin-bottom: 5px; font-weight: 500; }
.ws-bar-pct { font-weight: 700; color: #1A3A5C; }
.ws-bg  { background: #EEF2FF; border-radius: 4px; height: 7px; }
.ws-fill { height: 7px; border-radius: 4px; background: linear-gradient(90deg, #1E3A5C 0%, #2D6DA4 100%); }

.wp-wrap { overflow-x: auto; margin-top: 4px; }
table.wp { width: 100%; border-collapse: collapse; font-size: 0.8rem; }
table.wp thead tr { background: #0D1B2A; position: sticky; top: 0; z-index: 2; }
table.wp thead th { padding: 10px 14px; font-weight: 500; text-align: left; color: #FFFFFF; font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.06em; white-space: nowrap; }
table.wp tbody tr:nth-child(even) { background: #F8FAFC; }
table.wp tbody tr:hover { background: #EFF6FF; }
table.wp td { padding: 9px 14px; color: #374151; border-bottom: 1px solid #F1F5F9; vertical-align: middle; }
tr.ws-header td { background: #EEF2FF !important; color: #1E3A5C; font-weight: 700; font-size: 0.78rem; padding: 7px 14px; border-bottom: 1px solid #C7D7EE; text-transform: uppercase; letter-spacing: 0.04em; }

.badge { display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: 0.7rem; font-weight: 700; letter-spacing: 0.02em; white-space: nowrap; }
.b-completed  { background: #DCFCE7; color: #166534; }
.b-overdue    { background: #FEE2E2; color: #991B1B; }
.b-inprogress { background: #FEF9C3; color: #854D0E; }
.b-notstarted { background: #F1F5F9; color: #475569; }

.gate  { background: #EDE9FE; color: #5B21B6; border-radius: 4px; padding: 1px 6px; font-size: 0.65rem; font-weight: 700; margin-left: 5px; vertical-align: middle; }
.dchip { background: #F1F5F9; color: #475569; border-radius: 4px; padding: 2px 7px; font-size: 0.74rem; white-space: nowrap; }
.tick  { color: #22C55E; font-size: 1rem; }
.cross { color: #EF4444; font-size: 0.85rem; }

.up-item { display: flex; align-items: flex-start; gap: 12px; padding: 12px 0; border-bottom: 1px solid #F1F5F9; }
.up-item:last-child { border-bottom: none; }
.up-icon { width: 34px; height: 34px; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 0.95rem; flex-shrink: 0; }
.up-icon.mtg  { background: #EFF6FF; }
.up-icon.dlv  { background: #FFF7ED; }
.up-title { font-weight: 600; color: #0D1B2A; font-size: 0.86rem; margin-bottom: 2px; }
.up-meta  { font-size: 0.74rem; color: #94A3B8; }
.up-type  { font-size: 0.7rem; color: #6B7280; background: #F1F5F9; border-radius: 4px; padding: 1px 6px; margin-left: 6px; }

.wp-filter-bar { display: flex; align-items: center; gap: 14px; background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 16px; margin-bottom: 12px; }
.wp-filter-label { font-size: 0.78rem; font-weight: 600; color: #374151; white-space: nowrap; }
.wp-filter-count { font-size: 0.78rem; color: #6B7280; }
.wp-legend { display: flex; gap: 20px; margin-top: 14px; font-size: 0.76rem; color: #64748B; }
.wp-legend span { display: flex; align-items: center; gap: 5px; }

.del-table { width:100%; border-collapse:collapse; font-size:0.8rem; }
.del-table thead tr { background:#0D1B2A; }
.del-table thead th { padding:9px 12px; color:#FFFFFF; font-size:0.7rem; text-transform:uppercase; letter-spacing:0.06em; font-weight:500; text-align:left; }
.del-table tbody tr:nth-child(even) { background:#F8FAFC; }
.del-table td { padding:8px 12px; color:#374151; border-bottom:1px solid #F1F5F9; font-size:0.79rem; }
</style>
""", unsafe_allow_html=True)

# ── load data ─────────────────────────────────────────────────────────────────
try:
    raw = load_all_sheets()
    load_error = None
except Exception as e:
    raw = None
    load_error = str(e)

if load_error:
    st.error(f"⚠️ Could not connect to Google Sheets: {load_error}")
    st.info("Make sure `GOOGLE_SHEET_ID` and `GOOGLE_CREDENTIALS_PATH` are set correctly.")
    st.stop()

# Process each sheet
cfg   = raw["config"]
df    = process_workplan(raw["workplan"])
meets = process_meetings(raw["meetings"])
df_dl = process_deliverables(raw["deliverables"])

# Config values
project_name  = cfg.get("Project Name",  "DEL × TransGrid SPV")
client_name   = cfg.get("Client",        "DEL / TransGrid")
project_start = _parse_config_date(cfg.get("Start Date", "")) or date(2026, 4, 27)
project_end   = _parse_config_date(cfg.get("End Date",   "")) or date(2026, 9, 30)
today         = date.today()
day_num       = max((today - project_start).days + 1, 1)
total_days    = (project_end - project_start).days
progress_pct  = min(day_num / total_days * 100, 100)

all_ws = sorted(df["Workstream"].unique().tolist())

# ── filtered df ───────────────────────────────────────────────────────────────
df_f = df.copy()
kpi  = kpi_counts(df_f)

# ── Row 1: logo | title | date pill | refresh button ─────────────────────────
logo_tag = _logo_img_tag(height=30)
row1_left, row1_right = st.columns([4, 1])

with row1_left:
    st.markdown(f"""
<div style="background:#FFFFFF;border:1px solid #E2E8F0;border-radius:10px;
            padding:10px 20px;display:flex;align-items:center;gap:12px">
    {logo_tag}
    <span style="font-family:'DM Sans',sans-serif;font-size:1.05rem;
                 font-weight:700;color:#0D1B2A">Monitoring Dashboard</span>
    <span style="margin-left:auto;background:#F4F6F9;border:1px solid #E2E8F0;
                 border-radius:20px;padding:4px 14px;font-size:0.76rem;
                 color:#4A5568;white-space:nowrap">📅 {fmt_date(today)}</span>
</div>
""", unsafe_allow_html=True)

with row1_right:
    if st.button("🔄 Refresh data", use_container_width=True):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.rerun()
    st.caption(f"Loaded: {today.strftime('%d %b %Y')}")

# ── Row 2: project meta fields ────────────────────────────────────────────────
def _field(label, value, border=True):
    sep = "border-right:1px solid #E2E8F0;" if border else ""
    return f"""
<div style="padding:0 20px;{sep}">
    <div style="font-size:0.62rem;text-transform:uppercase;letter-spacing:0.08em;
                color:#94A3B8;margin-bottom:3px">{label}</div>
    <div style="font-size:0.82rem;font-weight:600;color:#0D1B2A;
                white-space:nowrap">{value}</div>
</div>"""

timeline_bar = f"""
<div style="padding:0 20px">
    <div style="font-size:0.62rem;text-transform:uppercase;letter-spacing:0.08em;
                color:#94A3B8;margin-bottom:6px">Timeline</div>
    <div style="width:140px;background:#EEF2FF;border-radius:4px;height:6px">
        <div style="background:#C8963E;width:{progress_pct:.1f}%;
                    height:6px;border-radius:4px"></div>
    </div>
    <div style="font-size:0.65rem;color:#94A3B8;margin-top:3px">
        {progress_pct:.0f}% of project elapsed</div>
</div>"""

st.markdown(f"""
<div style="background:#FFFFFF;border:1px solid #E2E8F0;border-radius:10px;
            padding:10px 0;display:flex;align-items:center;margin-bottom:20px;
            margin-top:8px">
    {_field("Project",     project_name)}
    {_field("Client",      client_name)}
    {_field("Start Date",  fmt_date(project_start))}
    {_field("End Date",    fmt_date(project_end))}
    {_field("Project Day", f"Day {day_num} of {total_days}")}
    {timeline_bar}
</div>
""", unsafe_allow_html=True)

# ── KPI row ───────────────────────────────────────────────────────────────────
total_pct = f"{kpi['completed']/kpi['total']*100:.0f}% done" if kpi['total'] else "—"
st.markdown(f"""
<div class="kpi-row">
    <div class="kpi-card c-gold">
        <div class="kpi-card-label">Total Actions</div>
        <div class="kpi-card-value">{kpi['total']}</div>
        <div class="kpi-card-sub">across {len(all_ws)} workstreams</div>
    </div>
    <div class="kpi-card c-green">
        <div class="kpi-card-label">Total Completed Actions</div>
        <div class="kpi-card-value">{kpi['completed']}</div>
        <div class="kpi-card-sub">{total_pct}</div>
    </div>
    <div class="kpi-card c-blue">
        <div class="kpi-card-label">Total Outstanding</div>
        <div class="kpi-card-value">{kpi['outstanding']}</div>
        <div class="kpi-card-sub">{kpi['in_progress']} in progress</div>
    </div>
    <div class="kpi-card c-red">
        <div class="kpi-card-label">Total Delayed</div>
        <div class="kpi-card-value">{kpi['overdue']}</div>
        <div class="kpi-card-sub">past end date</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── milestone donut + workstream progress ────────────────────────────────────
col_l, col_r = st.columns(2, gap="medium")

with col_l:
    pct_done = kpi['completed']   / kpi['total'] * 100 if kpi['total'] else 0
    pct_prog = kpi['in_progress'] / kpi['total'] * 100 if kpi['total'] else 0
    pct_out  = (kpi['outstanding'] - kpi['in_progress']) / kpi['total'] * 100 if kpi['total'] else 0

    def arc(cx, cy, r, s, e):
        sa = s / 100 * 2 * math.pi - math.pi / 2
        ea = e / 100 * 2 * math.pi - math.pi / 2
        x1, y1 = cx + r * math.cos(sa), cy + r * math.sin(sa)
        x2, y2 = cx + r * math.cos(ea), cy + r * math.sin(ea)
        lg = 1 if (e - s) > 50 else 0
        return f"M{x1:.2f},{y1:.2f} A{r},{r} 0 {lg},1 {x2:.2f},{y2:.2f}"

    CX, CY, R = 85, 85, 60
    arcs_html, cur = "", 0
    for pct, col in [(pct_done, "#22C55E"), (pct_prog, "#F59E0B"), (pct_out, "#3B82F6")]:
        if pct > 0.5:
            arcs_html += f'<path d="{arc(CX,CY,R,cur,cur+pct)}" stroke="{col}" stroke-width="14" fill="none" stroke-linecap="butt"/>'
            cur += pct

    donut = f"""<svg width="170" height="170" viewBox="0 0 170 170">
  <circle cx="{CX}" cy="{CY}" r="{R}" stroke="#E2E8F0" stroke-width="14" fill="none"/>
  {arcs_html}
  <text x="{CX}" y="{CY-6}" text-anchor="middle" font-family="DM Sans,sans-serif" font-size="24" font-weight="700" fill="#0D1B2A">{day_num}</text>
  <text x="{CX}" y="{CY+13}" text-anchor="middle" font-family="Inter,sans-serif" font-size="10" fill="#94A3B8">days</text>
</svg>"""

    st.markdown(f"""
<div class="card">
  <div class="card-title">Milestone Status</div>
  <div class="donut-wrap">
    {donut}
    <div>
      <div style="font-size:0.78rem;font-weight:600;color:#0D1B2A;margin-bottom:12px">Day {day_num}</div>
      <div class="legend">
        <div class="legend-row"><div class="legend-dot" style="background:#22C55E"></div>Completed<span class="legend-pct">{pct_done:.0f}%</span></div>
        <div class="legend-row"><div class="legend-dot" style="background:#F59E0B"></div>In Progress<span class="legend-pct">{pct_prog:.0f}%</span></div>
        <div class="legend-row"><div class="legend-dot" style="background:#3B82F6"></div>Outstanding<span class="legend-pct">{pct_out:.0f}%</span></div>
      </div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

with col_r:
    ws_df = workstream_progress(df_f).sort_values("pct_complete", ascending=False)
    bars = ""
    for _, row in ws_df.iterrows():
        p = row["pct_complete"]
        bars += f"""<div class="ws-bar-wrap">
  <div class="ws-bar-label"><span>{row['Workstream']}</span><span class="ws-bar-pct">{p:.0f}%</span></div>
  <div class="ws-bg"><div class="ws-fill" style="width:{max(p,1)}%"></div></div>
</div>"""

    st.markdown(f"""
<div class="card" style="max-height:300px;overflow-y:auto">
  <div class="card-title">Workstream Progress</div>
  {bars}
</div>
""", unsafe_allow_html=True)

# ── workplan table ────────────────────────────────────────────────────────────
def badge(s):
    m = {"Completed":"b-completed","Overdue":"b-overdue","In Progress":"b-inprogress","Not Started":"b-notstarted"}
    cls = m.get(s, "b-notstarted")
    return f'<span class="badge {cls}">{s}</span>'

# ── workplan filters: workstream checkboxes + status dropdown side by side ────

# Initialise individual checkbox keys (source of truth — not a derived list)
for ws in all_ws:
    if f"wp_cb_{ws}" not in st.session_state:
        st.session_state[f"wp_cb_{ws}"] = True

# Handle All / Clear flags set by buttons on the PREVIOUS run
if st.session_state.pop("_do_select_all", False):
    for ws in all_ws:
        st.session_state[f"wp_cb_{ws}"] = True
if st.session_state.pop("_do_clear_all", False):
    for ws in all_ws:
        st.session_state[f"wp_cb_{ws}"] = False

# Derive current selection from the (now up-to-date) checkbox keys
new_sel = [ws for ws in all_ws if st.session_state.get(f"wp_cb_{ws}", True)]
n_sel = len(new_sel)
ws_count_label = (
    f"All {len(all_ws)} selected"
    if n_sel == len(all_ws)
    else ("None selected" if n_sel == 0 else f"{n_sel} of {len(all_ws)} selected")
)

filter_col, status_col = st.columns([3, 1], gap="medium")

with filter_col:
    st.markdown(f"""
    <div class="wp-filter-bar">
        <span class="wp-filter-label">Workstreams:</span>
        <span class="wp-filter-count">{ws_count_label}</span>
    </div>
    """, unsafe_allow_html=True)

    cb_cols = st.columns(3)
    for i, ws in enumerate(all_ws):
        cb_cols[i % 3].checkbox(ws, key=f"wp_cb_{ws}")

    sa_col, sc_col, _ = st.columns([1, 1, 4])
    if sa_col.button("✅ All", key="wp_sel_all"):
        st.session_state["_do_select_all"] = True
        st.rerun()
    if sc_col.button("☐ Clear", key="wp_clr_all"):
        st.session_state["_do_clear_all"] = True
        st.rerun()

with status_col:
    st.markdown("""
    <div style="font-size:0.78rem;font-weight:600;color:#374151;margin-bottom:6px;margin-top:2px">
        Filter by Status
    </div>
    """, unsafe_allow_html=True)
    all_statuses = ["Completed", "In Progress", "Overdue", "Not Started"]
    status_filter = st.multiselect(
        "status_filter",
        options=all_statuses,
        default=all_statuses,
        label_visibility="collapsed",
        key="wp_status_filter",
    )

# Apply both filters
df_wp = df_f[
    df_f["Workstream"].isin(new_sel) &
    df_f["resolved_status"].isin(status_filter if status_filter else all_statuses)
] if new_sel else pd.DataFrame(columns=df_f.columns)

rows_html = ""
for ws_name, grp in df_wp.groupby("Workstream", sort=False):
    rows_html += f'<tr class="ws-header"><td colspan="5">{ws_name}</td></tr>'
    for _, r in grp.iterrows():
        end_s     = fmt_short(r["_end_date"]) if pd.notna(r.get("_end_date")) and r.get("_end_date") else "—"
        gate      = '<span class="gate">★ GATE</span>' if r.get("is_gate") else ""
        comp_s    = r["resolved_status"]
        comp_icon = '<span class="tick">✓</span>' if comp_s == "Completed" else '<span class="cross">✗</span>'
        rows_html += f"""<tr>
  <td>{r.get('Main Task','')}{gate}</td>
  <td style="text-align:center">{comp_icon}</td>
  <td><span class="dchip">{end_s}</span></td>
  <td>{str(r.get('Responsible Owner','—'))}</td>
  <td>{badge(comp_s)}</td>
</tr>"""

if not rows_html:
    rows_html = '<tr><td colspan="5" style="text-align:center;color:#94A3B8;padding:20px">No tasks match the selected filters.</td></tr>'

st.markdown(f"""
<div class="card" style="padding-bottom:16px">
  <div class="card-title">Workplan</div>
  <div class="wp-wrap" style="max-height:420px;overflow-y:auto">
    <table class="wp">
      <thead style="position:sticky;top:0;z-index:2"><tr>
        <th style="width:34%">Task</th>
        <th style="width:8%;text-align:center">Done</th>
        <th style="width:15%">Due Date</th>
        <th style="width:20%">Responsible Owner</th>
        <th style="width:13%">Status</th>
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>
  </div>
  <div class="wp-legend">
    <span><span class="tick">✓</span> Completed</span>
    <span><span class="cross">✗</span> Not yet complete</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ── upcoming: meetings (left) + top 10 urgent deliverables (right) ────────────
up_l, up_r = st.columns(2, gap="medium")

with up_l:
    mtg_html = ""
    for m in meets:
        type_chip = f'<span class="up-type">{m["type"]}</span>' if m.get("type") else ""
        mtg_html += f"""<div class="up-item">
  <div class="up-icon mtg">📅</div>
  <div>
    <div class="up-title">{m['title']}{type_chip}</div>
    <div class="up-meta">{m['date']} &nbsp;·&nbsp; {m['time']}</div>
  </div>
</div>"""
    if not mtg_html:
        mtg_html = "<div style='color:#94A3B8;font-size:0.83rem'>No upcoming meetings.</div>"
    st.markdown(f'<div class="card"><div class="card-title">Upcoming — Meetings</div>{mtg_html}</div>',
                unsafe_allow_html=True)

with up_r:
    # Top 10 most urgent deliverables: sort by days until due (ascending),
    # overdue tasks first (negative days), exclude completed
    today_d = date.today()

    def _urgency_days(row):
        d = row.get("_due_date")
        if d and pd.notna(d):
            return (d - today_d).days
        return 9999  # no date → push to end

    if not df_dl.empty:
        urgent = df_dl[df_dl["resolved_status"] != "Completed"].copy()
        urgent["_days_left"] = urgent.apply(_urgency_days, axis=1)
        urgent = urgent.sort_values("_days_left").head(10)
    else:
        urgent = pd.DataFrame()

    del_rows = ""
    for _, r in urgent.iterrows():
        due_d = r.get("_due_date")
        days_left = r.get("_days_left", 9999)
        due_str = fmt_short(due_d) if due_d and pd.notna(due_d) else "—"

        if days_left < 0:
            urgency_chip = f'<span style="font-size:0.68rem;background:#FEE2E2;color:#991B1B;border-radius:4px;padding:1px 6px;white-space:nowrap">{abs(int(days_left))}d overdue</span>'
        elif days_left == 0:
            urgency_chip = '<span style="font-size:0.68rem;background:#FEF9C3;color:#854D0E;border-radius:4px;padding:1px 6px">Due today</span>'
        elif days_left <= 7:
            urgency_chip = f'<span style="font-size:0.68rem;background:#FEF9C3;color:#854D0E;border-radius:4px;padding:1px 6px">{int(days_left)}d left</span>'
        else:
            urgency_chip = f'<span style="font-size:0.68rem;background:#F1F5F9;color:#475569;border-radius:4px;padding:1px 6px">{int(days_left)}d left</span>'

        del_rows += f"""<tr>
  <td>{r.get('Deliverable','')}</td>
  <td>{r.get('Owner','—')}</td>
  <td><span class="dchip">{due_str}</span>&nbsp;{urgency_chip}</td>
  <td>{badge(r['resolved_status'])}</td>
</tr>"""

    if not del_rows:
        del_rows = '<tr><td colspan="4" style="text-align:center;color:#94A3B8;padding:16px">No deliverables found.</td></tr>'

    st.markdown(f"""
<div class="card" style="padding-bottom:16px">
  <div class="card-title">Upcoming — Top 10 Urgent Deliverables</div>
  <div style="max-height:340px;overflow-y:auto">
    <table class="del-table">
      <thead style="position:sticky;top:0;z-index:2"><tr>
        <th style="width:32%">Deliverable</th>
        <th style="width:20%">Owner</th>
        <th style="width:28%">Due Date</th>
        <th style="width:16%">Status</th>
      </tr></thead>
      <tbody>{del_rows}</tbody>
    </table>
  </div>
</div>
""", unsafe_allow_html=True)

# ── footer ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="text-align:center;padding:16px 0 8px;font-size:0.7rem;color:#CBD5E1">
    Vista Advisory Partners &nbsp;·&nbsp; {project_name} &nbsp;·&nbsp; {client_name} &nbsp;·&nbsp; {fmt_short(today)}
</div>
""", unsafe_allow_html=True)