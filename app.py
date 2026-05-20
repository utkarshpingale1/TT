"""
GHR & TT Workforce Dashboard
------------------------------
Usage:
    python dashboard.py

Requirements:
    pip install pandas openpyxl

Place dashboard.py one level above the 'data' folder:
    project/
    ├── dashboard.py
    └── data/
        ├── GHR.xlsx
        └── TT.xlsx

Then run: python dashboard.py
A browser tab will open automatically.
"""
from flask import Flask
import json
import os
import webbrowser
import http.server
import threading
import pandas as pd
import numpy as np
from pathlib import Path

# ─────────────────────────────────────────────
# 1. DATA LOADING & PROCESSING
# ─────────────────────────────────────────────

BASE = Path(__file__).parent / "data"

def load_data():
    ghr = pd.read_excel(BASE / "GHR.xlsx")

    tt_raw = pd.read_excel(BASE / "TT.xlsx", sheet_name="Data", header=None)
    # find header row that contains 'EmpCd'
    header_row = next(i for i, row in tt_raw.iterrows() if "EmpCd" in str(list(row)))
    tt_data = pd.read_excel(BASE / "TT.xlsx", sheet_name="Data", header=header_row)
    tt_data.columns = [str(c).strip() for c in tt_data.columns]
    # drop duplicate 'Note' columns
    tt_data = tt_data.loc[:, ~tt_data.columns.duplicated()]
    tt_data["Hours"] = pd.to_numeric(tt_data["Hours"], errors="coerce")
    tt_data["Val"]   = pd.to_numeric(tt_data["Val"],   errors="coerce")
    tt_data["Date"]  = pd.to_datetime(tt_data["Date"], errors="coerce")
    tt_data["Dept1_clean"] = tt_data["Dept1"].astype(str).str.strip()

    # ── GHR stats ──
    active = ghr[ghr["Employee Status.1"] == "Active"].copy()
    active["EmpCd"] = active["Emp ID"].astype(str).str.zfill(4)
    active["Dept_clean"] = active["Department"].astype(str).str.strip()

    ghr_stats = {
        "total":      int(len(ghr)),
        "active":     int(len(active)),
        "left":       int(len(ghr[ghr["Employee Status.1"] == "Left"])),
        "confirmed":  int(len(active[active["Employee Status"] == "Confirmed"])),
        "consultant": int(len(active[active["Employee Status"] == "Consultant"])),
        "intern":     int(len(active[active["Employee Status"] == "Intern"])),
        "male":       int(len(active[active["Gender"] == "Male"])),
        "female":     int(len(active[active["Gender"] == "Female"])),
        "dept_breakdown": [
            {"dept": k, "count": int(v)}
            for k, v in active.groupby("Dept_clean").size()
                              .sort_values(ascending=False).head(12).items()
        ],
    }

    # ── TT stats ──
    tt_emps = set(tt_data["EmpCd"].dropna().unique())
    emp_hours = tt_data.groupby("EmpCd")["Hours"].sum()

    weekly = (
        tt_data.assign(Week=tt_data["Date"].dt.to_period("W"))
               .groupby("Week")["Hours"].sum()
    )
    weekly_list = [
        {"week": str(p.start_time.strftime("%d %b")), "hours": round(float(h), 1)}
        for p, h in weekly.items()
    ]

    dept_hours = (
        tt_data.groupby("Dept1_clean")["Hours"].sum()
               .sort_values(ascending=False).head(10)
    )

    top_proj = (
        tt_data.groupby("CoNo")["Hours"].sum()
               .sort_values(ascending=False).head(10)
    )

    # ── Employee → Project breakdown ──
    emp_proj_grp = (
        tt_data.groupby(["EmpCd", "EmpName", "Dept1_clean", "CoNo", "ProjRef"])["Hours"]
               .sum().reset_index()
    )
    emp_proj_grp["Hours"] = emp_proj_grp["Hours"].round(1)

    emp_total_df = (
        tt_data.groupby(["EmpCd", "EmpName", "Dept1_clean"])["Hours"]
               .sum().reset_index()
    )
    emp_total_df.columns = ["EmpCd", "Name", "Dept", "TotalHours"]
    emp_total_df["TotalHours"] = emp_total_df["TotalHours"].round(1)
    emp_total_df = emp_total_df[emp_total_df["Name"].str.contains(" ", na=False)]
    emp_total_df = emp_total_df[~emp_total_df["Name"].str.match(r"^[A-Z0-9\-]+$")]

    proj_by_emp = {}
    for _, r in emp_proj_grp.iterrows():
        proj_by_emp.setdefault(r["EmpCd"], []).append({
            "proj":  str(r["CoNo"]).strip(),
            "ref":   str(r["ProjRef"]).strip()[:70] if pd.notna(r["ProjRef"]) else str(r["CoNo"]).strip(),
            "hours": float(r["Hours"]),
        })
    for k in proj_by_emp:
        proj_by_emp[k].sort(key=lambda x: x["hours"], reverse=True)

    emp_proj_list = [
        {
            "EmpCd":        row["EmpCd"],
            "Name":         row["Name"].strip(),
            "Dept":         row["Dept"],
            "TotalHours":   float(row["TotalHours"]),
            "ProjectCount": len(proj_by_emp.get(row["EmpCd"], [])),
            "projects":     proj_by_emp.get(row["EmpCd"], []),
        }
        for _, row in emp_total_df.sort_values("TotalHours", ascending=False).iterrows()
    ]

    # All unique departments for filter chips
    all_depts = sorted(list(set(str(d) for d in tt_data["Dept1_clean"].dropna().unique())))

    tt_stats = {
        "unique_emps":     int(len(tt_emps)),
        "total_hours":     round(float(tt_data["Hours"].sum()), 1),
        "total_val":       round(float(tt_data["Val"].sum()), 0),
        "unique_projects": int(tt_data["CoNo"].nunique()),
        "date_from":       "26 Mar 2026",
        "date_to":         "18 May 2026",
        "avg_hrs_per_emp": round(float(emp_hours.mean()), 1),
        "weekly_trend": weekly_list,
        "dept_hours": [
            {"dept": k, "hours": round(float(v), 1)}
            for k, v in dept_hours.items()
        ],
        "top_projects": [
            {"proj": str(k), "hours": round(float(v), 1)}
            for k, v in top_proj.items()
        ],
        "emp_projects": emp_proj_list,
        "all_depts":    all_depts,
    }

    # ── Compare stats ──
    active["in_TT"] = active["EmpCd"].isin(tt_emps)
    non_consultant  = active[active["Employee Status"] != "Consultant"]
    defaulters_df   = non_consultant[~non_consultant["in_TT"]].copy()
    fillers_df      = active[active["in_TT"]].copy()

    emp_hours_map = emp_hours.reset_index()
    emp_hours_map.columns = ["EmpCd", "TT_Hours"]
    fillers_df = fillers_df.merge(emp_hours_map, on="EmpCd", how="left")
    fillers_df["TT_Hours"] = fillers_df["TT_Hours"].fillna(0).round(1)

    def to_records(df, cols, rename):
        return df[cols].rename(columns=rename).fillna("—").to_dict("records")

    fillers_list = to_records(
        fillers_df,
        ["EmpCd", "Name", "Dept_clean", "Employee Status", "Manager Name", "TT_Hours"],
        {"Dept_clean": "Department", "Employee Status": "Status",
         "Manager Name": "Manager", "TT_Hours": "Hours"},
    )
    defaulters_list = to_records(
        defaulters_df,
        ["EmpCd", "Name", "Dept_clean", "Employee Status", "Manager Name"],
        {"Dept_clean": "Department", "Employee Status": "Status",
         "Manager Name": "Manager"},
    )

    dept_default = (
        defaulters_df.groupby("Dept_clean").size()
                     .sort_values(ascending=False).head(10)
    )

    compare_stats = {
        "active_total":       int(len(active)),
        "active_consultants": int(len(active[active["Employee Status"] == "Consultant"])),
        "filling_tt":         int(len(fillers_df)),
        "defaulters":         int(len(defaulters_df)),
        "defaulters_by_dept": [
            {"dept": k, "count": int(v)} for k, v in dept_default.items()
        ],
        "fillers":            fillers_list,
        "defaulters_list":    defaulters_list,
    }

    return {"ghr": ghr_stats, "tt": tt_stats, "compare": compare_stats}


# ─────────────────────────────────────────────
# 2. HTML GENERATION
# ─────────────────────────────────────────────

def build_html(data: dict) -> str:
    payload = json.dumps(data, default=str)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GHR & TT Workforce Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',system-ui,sans-serif;background:#f4f6f9;color:#1a1a2e;min-height:100vh}}
.topbar{{background:#fff;border-bottom:1px solid #e2e8f0;padding:14px 28px;display:flex;align-items:center;justify-content:space-between}}
.topbar h1{{font-size:18px;font-weight:600;color:#1a1a2e}}
.topbar .sub{{font-size:12px;color:#64748b;margin-top:2px}}
.badge-period{{background:#EFF6FF;color:#1d4ed8;font-size:11px;padding:4px 12px;border-radius:20px;font-weight:500}}
.tabs{{display:flex;gap:0;background:#fff;padding:0 28px;border-bottom:1px solid #e2e8f0}}
.tab{{padding:13px 22px;font-size:13px;font-weight:500;cursor:pointer;border:none;background:none;color:#64748b;border-bottom:3px solid transparent;margin-bottom:-1px;transition:all .15s;display:flex;align-items:center;gap:6px}}
.tab.active{{color:#1d4ed8;border-bottom-color:#1d4ed8}}
.tab:hover:not(.active){{color:#1a1a2e}}
.tab-badge{{background:#f1f5f9;border-radius:10px;font-size:11px;padding:1px 8px;color:#64748b}}
.tab-badge.red{{background:#fef2f2;color:#991b1b}}
.tab-badge.blue{{background:#eff6ff;color:#1d4ed8}}
.tab-badge.green{{background:#f0fdf4;color:#166534}}
.panel{{display:none;padding:24px 28px}}
.panel.active{{display:block}}
.metric-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:12px;margin-bottom:24px}}
.metric{{background:#fff;border-radius:10px;padding:16px;border:1px solid #e2e8f0}}
.metric-label{{font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px}}
.metric-value{{font-size:26px;font-weight:600;color:#1a1a2e}}
.metric-sub{{font-size:11px;color:#94a3b8;margin-top:3px}}
.metric.blue .metric-value{{color:#1d4ed8}}
.metric.green .metric-value{{color:#166534}}
.metric.red .metric-value{{color:#991b1b}}
.metric.amber .metric-value{{color:#92400e}}
.chart-row{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}}
.chart-box{{background:#fff;border-radius:10px;padding:16px 20px;border:1px solid #e2e8f0}}
.chart-title{{font-size:12px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.05em;margin-bottom:14px}}
.prog-row{{display:flex;align-items:center;gap:10px;margin-bottom:9px}}
.prog-label{{font-size:12px;color:#334155;width:150px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.prog-track{{flex:1;height:8px;background:#f1f5f9;border-radius:4px;overflow:hidden}}
.prog-fill{{height:100%;border-radius:4px;background:#3b82f6;transition:width .4s ease}}
.prog-val{{font-size:12px;color:#64748b;min-width:48px;text-align:right}}
.legend-row{{display:flex;gap:14px;flex-wrap:wrap;margin-top:10px}}
.legend-item{{display:flex;align-items:center;gap:5px;font-size:12px;color:#64748b}}
.legend-dot{{width:10px;height:10px;border-radius:2px;display:inline-block}}

/* ── TT Filter Bar ── */
.tt-filter-bar{{
  background:#fff;border:1px solid #e2e8f0;border-radius:12px;
  padding:14px 18px;margin-bottom:16px;display:flex;align-items:flex-start;
  gap:14px;flex-wrap:wrap;box-shadow:0 1px 4px rgba(0,0,0,.04);
}}
.fi-group{{display:flex;align-items:center;gap:8px;min-width:220px;flex:1}}
.fi-icon{{
  width:36px;height:36px;border-radius:9px;background:#f1f5f9;
  display:flex;align-items:center;justify-content:center;font-size:17px;flex-shrink:0;
}}
.fi-input{{
  flex:1;border:1.5px solid #e2e8f0;border-radius:8px;padding:8px 12px;
  font-size:13px;background:#fafbfc;color:#1a1a2e;outline:none;
  transition:border .15s,box-shadow .15s;
}}
.fi-input:focus{{border-color:#3b82f6;background:#fff;box-shadow:0 0 0 3px rgba(59,130,246,.1)}}
.fi-divider{{width:1px;align-self:stretch;background:#e2e8f0;flex-shrink:0;margin:2px 0}}
.fi-dept-section{{display:flex;flex-direction:column;gap:6px;flex:2;min-width:240px}}
.fi-dept-label{{font-size:11px;font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:.05em}}
.dept-chips-wrap{{display:flex;flex-wrap:wrap;gap:5px}}
.dept-chip{{
  display:inline-flex;align-items:center;gap:3px;padding:4px 11px;border-radius:20px;
  font-size:11px;font-weight:500;cursor:pointer;border:1.5px solid #e2e8f0;
  background:#fff;color:#475569;transition:all .15s;user-select:none;white-space:nowrap;
}}
.dept-chip:hover{{border-color:#3b82f6;color:#1d4ed8;background:#eff6ff}}
.dept-chip.on{{background:#1d4ed8;color:#fff;border-color:#1d4ed8;padding-right:8px}}
.chip-x{{font-size:14px;line-height:1;margin-left:1px;opacity:.75}}
.fi-actions{{display:flex;flex-direction:column;align-items:flex-end;gap:8px;flex-shrink:0}}
.fi-result-badge{{
  background:#eff6ff;color:#1d4ed8;font-size:11px;font-weight:600;
  padding:3px 11px;border-radius:20px;white-space:nowrap;
}}
.fi-clear{{
  padding:6px 14px;border-radius:8px;font-size:12px;font-weight:500;
  border:1.5px solid #e2e8f0;background:#fff;color:#64748b;
  cursor:pointer;transition:all .15s;white-space:nowrap;
}}
.fi-clear:hover{{background:#fef2f2;border-color:#fca5a5;color:#991b1b}}

.search-bar{{width:100%;padding:9px 14px;border:1px solid #e2e8f0;border-radius:8px;font-size:13px;background:#fff;color:#1a1a2e;outline:none;margin-bottom:10px}}
.search-bar:focus{{border-color:#3b82f6;box-shadow:0 0 0 3px rgba(59,130,246,.1)}}
.filter-row{{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px}}
.flt-btn{{padding:5px 14px;border:1px solid #e2e8f0;border-radius:20px;font-size:12px;cursor:pointer;background:#fff;color:#64748b;transition:all .15s;font-weight:500}}
.flt-btn.on{{background:#1d4ed8;color:#fff;border-color:#1d4ed8}}
.view-toggle{{display:flex;gap:0;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;margin-bottom:12px;width:fit-content}}
.vt-btn{{padding:7px 18px;font-size:13px;font-weight:500;cursor:pointer;border:none;background:#fff;color:#64748b;transition:all .15s}}
.vt-btn.on{{background:#1d4ed8;color:#fff}}
.tbl-wrap{{max-height:400px;overflow-y:auto;border:1px solid #e2e8f0;border-radius:10px}}
table{{width:100%;border-collapse:collapse;font-size:12px;table-layout:fixed}}
th{{text-align:left;padding:9px 12px;font-weight:600;color:#64748b;background:#f8fafc;border-bottom:1px solid #e2e8f0;font-size:11px;text-transform:uppercase;letter-spacing:.04em;position:sticky;top:0;z-index:2}}
td{{padding:8px 12px;border-bottom:1px solid #f1f5f9;color:#334155;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
tr:last-child td{{border-bottom:none}}
tr:hover td{{background:#f8fafc}}
.pill{{display:inline-block;padding:2px 9px;border-radius:12px;font-size:11px;font-weight:500}}
.pill-confirmed{{background:#eff6ff;color:#1d4ed8}}
.pill-intern{{background:#fefce8;color:#854d0e}}
.pill-default{{background:#fef2f2;color:#991b1b}}
.pill-ok{{background:#f0fdf4;color:#166534}}
.pg-info{{font-size:12px;color:#94a3b8;text-align:right;margin-bottom:6px}}
.pg-btns{{display:flex;justify-content:center;gap:6px;margin-top:10px}}
.pg-btn{{padding:5px 12px;border:1px solid #e2e8f0;border-radius:6px;font-size:12px;cursor:pointer;background:#fff;color:#334155}}
.pg-btn.on{{background:#1d4ed8;color:#fff;border-color:#1d4ed8}}
.pg-btn:hover:not(.on){{background:#f1f5f9}}
@media(max-width:680px){{.chart-row{{grid-template-columns:1fr}}.metric-grid{{grid-template-columns:repeat(2,1fr)}}}}
</style>
</head>
<body>

<div class="topbar">
  <div>
    <h1>&#128198; Workforce Dashboard</h1>
    <div class="sub">Monarch Surveyors &amp; Engineering Consultants Ltd.</div>
  </div>
  <span class="badge-period">TT period: 26 Mar – 18 May 2026</span>
</div>

<div class="tabs">
  <button class="tab active" onclick="showTab('ghr',this)">&#128100; GHR <span class="tab-badge blue" id="t-ghr">—</span></button>
  <button class="tab" onclick="showTab('tt',this)">&#128336; TT <span class="tab-badge green" id="t-tt">—</span></button>
  <button class="tab" onclick="showTab('compare',this)">&#9878; Compare <span class="tab-badge red" id="t-cmp">—</span></button>
</div>

<!-- ═══════════════ GHR PANEL ═══════════════ -->
<div id="panel-ghr" class="panel active">
  <div class="metric-grid" id="ghr-metrics"></div>
  <div class="chart-row">
    <div class="chart-box">
      <div class="chart-title">Active employees by department</div>
      <div id="ghr-dept-bars"></div>
    </div>
    <div class="chart-box">
      <div class="chart-title">Employment status — all records</div>
      <div style="position:relative;height:210px">
        <canvas id="ghrTypeChart"></canvas>
      </div>
      <div class="legend-row" id="ghr-type-legend"></div>
    </div>
  </div>
  <div class="chart-box" style="background:#fff;border-radius:10px;padding:16px 20px;border:1px solid #e2e8f0">
    <div class="chart-title">Gender split (active employees)</div>
    <div id="ghr-gender-bars"></div>
  </div>
</div>

<!-- ═══════════════ TT PANEL ═══════════════ -->
<div id="panel-tt" class="panel">
  <div class="metric-grid" id="tt-metrics"></div>
  <div class="chart-row">
    <div class="chart-box">
      <div class="chart-title">Weekly hours logged</div>
      <div style="position:relative;height:220px">
        <canvas id="weeklyChart"></canvas>
      </div>
    </div>
    <div class="chart-box">
      <div class="chart-title">Hours by department</div>
      <div style="position:relative;height:220px">
        <canvas id="deptHoursChart"></canvas>
      </div>
    </div>
  </div>
  <div class="chart-box" style="background:#fff;border-radius:10px;padding:16px 20px;border:1px solid #e2e8f0;margin-bottom:16px">
    <div class="chart-title">Top 10 projects by hours</div>
    <div id="proj-bars"></div>
  </div>

  <!-- Employee Project Breakdown -->
  <div style="background:#fff;border-radius:10px;padding:16px 20px;border:1px solid #e2e8f0">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px">
      <div class="chart-title" style="margin-bottom:0">Employee &#8594; project breakdown</div>
    </div>

    <!-- ── FILTER BAR ── -->
    <div class="tt-filter-bar">
      <!-- Search group -->
      <div class="fi-group">
        <div class="fi-icon">🔍</div>
        <input class="fi-input" type="text" id="empProjSearch"
          placeholder="Name, emp ID, or project…"
          oninput="EP_PAGE=1;renderEmpProj()">
      </div>

      <div class="fi-divider"></div>

      <!-- Dept chips -->
      <div class="fi-dept-section">
        <div class="fi-dept-label">🏢 Department</div>
        <div class="dept-chips-wrap" id="deptChips"></div>
      </div>

      <div class="fi-divider"></div>

      <!-- Actions -->
      <div class="fi-actions">
        <span class="fi-result-badge" id="epResultBadge">— emps</span>
        <button class="fi-clear" onclick="clearTTFilters()">✕ Clear filters</button>
      </div>
    </div>

    <div class="pg-info" id="epPgInfo"></div>
    <div class="tbl-wrap">
      <table style="width:100%;border-collapse:collapse;font-size:12px;table-layout:fixed">
        <thead>
          <tr>
            <th style="width:70px">Emp ID</th>
            <th style="width:180px">Name</th>
            <th style="width:120px">Department</th>
            <th style="width:75px;text-align:right">Total hrs</th>
            <th style="width:60px;text-align:right">Projs</th>
            <th>Project details &nbsp;&#x25BC; sorted by hours</th>
          </tr>
        </thead>
        <tbody id="empProjBody"></tbody>
      </table>
    </div>
    <div class="pg-btns" id="epPgBtns"></div>
  </div>
</div>

<!-- ═══════════════ COMPARE PANEL ═══════════════ -->
<div id="panel-compare" class="panel">
  <div class="metric-grid" id="cmp-metrics"></div>
  <div class="chart-row" style="margin-bottom:16px">
    <div class="chart-box">
      <div class="chart-title">TT compliance (non-consultants)</div>
      <div style="position:relative;height:190px">
        <canvas id="complianceChart"></canvas>
      </div>
      <div class="legend-row" id="cmp-legend"></div>
    </div>
    <div class="chart-box">
      <div class="chart-title">Defaulters by department</div>
      <div id="default-dept-bars"></div>
    </div>
  </div>

  <div class="view-toggle">
    <button class="vt-btn on" id="vt-def" onclick="switchView('defaulters')">&#128308; Defaulters</button>
    <button class="vt-btn" id="vt-fil" onclick="switchView('fillers')">&#128994; Filling TT</button>
  </div>

  <input class="search-bar" type="text" id="cmpSearch" placeholder="Search by name, department, manager, or emp ID…" oninput="renderTable()">

  <div class="filter-row" id="statusFilters">
    <button class="flt-btn on" onclick="setStatus('all',this)">All</button>
    <button class="flt-btn" onclick="setStatus('Confirmed',this)">Confirmed</button>
    <button class="flt-btn" onclick="setStatus('Intern',this)">Intern</button>
  </div>

  <div class="pg-info" id="pgInfo"></div>
  <div class="tbl-wrap">
    <table><thead id="tblHead"></thead><tbody id="tblBody"></tbody></table>
  </div>
  <div class="pg-btns" id="pgBtns"></div>
</div>

<script>
const DATA = {payload};

// ─── Tab switching ───
function showTab(name, el) {{
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('panel-' + name).classList.add('active');
  el.classList.add('active');
}}

// ─── Helpers ───
function fmt(n) {{ return Number(n).toLocaleString('en-IN'); }}
function pct(a,b) {{ return b ? Math.round(a/b*100) + '%' : '—'; }}

function progBars(containerId, items, keyLabel, keyVal, color) {{
  const el = document.getElementById(containerId);
  const max = Math.max(...items.map(i => i[keyVal]));
  el.innerHTML = items.map(i => `
    <div class="prog-row">
      <div class="prog-label" title="${{i[keyLabel]}}">${{i[keyLabel]}}</div>
      <div class="prog-track"><div class="prog-fill" style="width:${{Math.round(i[keyVal]/max*100)}}%;background:${{color}}"></div></div>
      <div class="prog-val">${{fmt(i[keyVal])}}</div>
    </div>`).join('');
}}

function metricCard(label, value, sub, cls='') {{
  return `<div class="metric ${{cls}}">
    <div class="metric-label">${{label}}</div>
    <div class="metric-value">${{value}}</div>
    <div class="metric-sub">${{sub}}</div>
  </div>`;
}}

// ─── GHR panel ───
function buildGHR() {{
  const g = DATA.ghr;
  document.getElementById('t-ghr').textContent = fmt(g.active) + ' active';
  document.getElementById('ghr-metrics').innerHTML = [
    metricCard('Total in GHR',    fmt(g.total),     'All-time records',  'blue'),
    metricCard('Active',          fmt(g.active),    'Currently employed','green'),
    metricCard('Left / Separated',fmt(g.left),      'Inactive',          'red'),
    metricCard('Confirmed',       fmt(g.confirmed), 'Active confirmed',  ''),
    metricCard('Consultants',     fmt(g.consultant),'Active consultants','amber'),
    metricCard('Interns',         fmt(g.intern),    'Active interns',    ''),
    metricCard('Male',            fmt(g.male),      pct(g.male,g.active) + ' of active',''),
    metricCard('Female',          fmt(g.female),    pct(g.female,g.active) + ' of active',''),
  ].join('');

  progBars('ghr-dept-bars', g.dept_breakdown, 'dept', 'count', '#3b82f6');

  // ── Donut: Confirmed + Consultant + Intern + Left/Separated ──
  new Chart(document.getElementById('ghrTypeChart'), {{
    type: 'doughnut',
    data: {{
      labels: ['Confirmed', 'Consultant', 'Intern', 'Left / Sep.'],
      datasets: [{{
        data: [g.confirmed, g.consultant, g.intern, g.left],
        backgroundColor: ['#3b82f6', '#f59e0b', '#22c55e', '#ef4444'],
        borderWidth: 2,
        borderColor: '#fff',
        hoverOffset: 6,
      }}]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{
          callbacks: {{
            label: ctx => ' ' + ctx.label + ': ' + fmt(ctx.raw) +
                          ' (' + Math.round(ctx.raw / g.total * 100) + '%)'
          }}
        }}
      }}
    }}
  }});

  document.getElementById('ghr-type-legend').innerHTML = [
    ['Confirmed',   g.confirmed,  '#3b82f6'],
    ['Consultant',  g.consultant, '#f59e0b'],
    ['Intern',      g.intern,     '#22c55e'],
    ['Left / Sep.', g.left,       '#ef4444'],
  ].map(([l, n, c]) =>
    `<span class="legend-item">
       <span class="legend-dot" style="background:${{c}}"></span>
       ${{l}} — <strong>${{fmt(n)}}</strong>
       <span style="color:#94a3b8">(${{Math.round(n / g.total * 100)}}%)</span>
     </span>`
  ).join('');

  document.getElementById('ghr-gender-bars').innerHTML = [
    ['Male',   g.male,   g.active, '#3b82f6'],
    ['Female', g.female, g.active, '#ec4899'],
  ].map(([l,n,tot,c]) => `
    <div class="prog-row">
      <div class="prog-label">${{l}}</div>
      <div class="prog-track"><div class="prog-fill" style="width:${{Math.round(n/tot*100)}}%;background:${{c}}"></div></div>
      <div class="prog-val">${{fmt(n)}} — ${{pct(n,tot)}}</div>
    </div>`).join('');
}}

// ─── TT panel ───
function buildTT() {{
  const t = DATA.tt;
  document.getElementById('t-tt').textContent = fmt(t.unique_emps) + ' emps';
  document.getElementById('tt-metrics').innerHTML = [
    metricCard('Employees in TT',   fmt(t.unique_emps),    'Logged hours',      'blue'),
    metricCard('Total hours',       fmt(t.total_hours),    t.date_from + ' – ' + t.date_to, 'green'),
    metricCard('Total value (₹)',   '₹' + (t.total_val/1e7).toFixed(2) + 'Cr', 'Billed value',''),
    metricCard('Active projects',   fmt(t.unique_projects),'Unique CO nos',     ''),
    metricCard('Avg hrs / emp',     t.avg_hrs_per_emp,     'Over the period',   ''),
  ].join('');

  new Chart(document.getElementById('weeklyChart'), {{
    type: 'line',
    data: {{
      labels: t.weekly_trend.map(w => w.week),
      datasets: [{{ label:'Hours', data: t.weekly_trend.map(w => w.hours),
        borderColor:'#3b82f6', backgroundColor:'rgba(59,130,246,.08)',
        tension:.35, fill:true, pointRadius:4, pointBackgroundColor:'#3b82f6' }}]
    }},
    options: {{ responsive:true, maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{ y:{{ ticks:{{ callback: v => fmt(v) }} }} }} }}
  }});

  const dh = t.dept_hours;
  new Chart(document.getElementById('deptHoursChart'), {{
    type: 'bar',
    data: {{
      labels: dh.map(d => d.dept),
      datasets: [{{ label:'Hours', data: dh.map(d => d.hours),
        backgroundColor:'#6366f1', borderRadius:4 }}]
    }},
    options: {{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{ x:{{ ticks:{{ callback: v => fmt(v) }} }}, y:{{ ticks:{{ font:{{size:11}} }} }} }}
    }}
  }});

  progBars('proj-bars', t.top_projects, 'proj', 'hours', '#6366f1');

  // ── Build dept chips ──
  const chipsWrap = document.getElementById('deptChips');
  (t.all_depts || []).forEach(d => {{
    const chip = document.createElement('span');
    chip.className = 'dept-chip';
    chip.dataset.dept = d;
    chip.innerHTML = d + '<span class="chip-x"> ×</span>';
    chip.title = d;
    chip.onclick = () => {{
      chip.classList.toggle('on');
      EP_PAGE = 1;
      renderEmpProj();
    }};
    chipsWrap.appendChild(chip);
  }});

  renderEmpProj();
}}

function clearTTFilters() {{
  document.getElementById('empProjSearch').value = '';
  document.querySelectorAll('#deptChips .dept-chip').forEach(c => c.classList.remove('on'));
  EP_PAGE = 1;
  renderEmpProj();
}}

// ─── Employee-Project table ───
let EP_PAGE = 1;
const EP_PER = 25;

function renderEmpProj() {{
  const q        = (document.getElementById('empProjSearch').value || '').toLowerCase();
  const selDepts = [...document.querySelectorAll('#deptChips .dept-chip.on')].map(c => c.dataset.dept);
  const src      = DATA.tt.emp_projects;

  const filtered = src.filter(e => {{
    const mDept = selDepts.length === 0 || selDepts.includes(e.Dept);
    const mQ    = !q || e.Name.toLowerCase().includes(q) ||
                  e.EmpCd.toLowerCase().includes(q) ||
                  e.projects.some(p => p.proj.toLowerCase().includes(q) || p.ref.toLowerCase().includes(q));
    return mDept && mQ;
  }});

  const total  = filtered.length;
  const pages  = Math.max(1, Math.ceil(total / EP_PER));
  EP_PAGE      = Math.min(EP_PAGE, pages);
  const slice  = filtered.slice((EP_PAGE-1)*EP_PER, EP_PAGE*EP_PER);

  // badges
  document.getElementById('epResultBadge').textContent =
    fmt(total) + ' emp' + (total !== 1 ? 's' : '');
  document.getElementById('epPgInfo').textContent =
    `Showing ${{Math.min((EP_PAGE-1)*EP_PER+1,total)}}–${{Math.min(EP_PAGE*EP_PER,total)}} of ${{total}} employees`;

  document.getElementById('empProjBody').innerHTML = slice.map(e => {{
    const projHtml = e.projects.map(p => {{
      const label = p.ref && p.ref !== p.proj ? p.ref : p.proj;
      return `<span style="display:inline-flex;align-items:center;gap:4px;background:#f1f5f9;border-radius:5px;padding:2px 7px;margin:2px 3px 2px 0;font-size:11px;white-space:nowrap">
        <span style="font-weight:600;color:#1d4ed8">${{p.proj}}</span>
        <span style="color:#475569;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{label}}">${{label.length>35?label.slice(0,35)+'…':label}}</span>
        <span style="background:#1d4ed8;color:#fff;border-radius:3px;padding:1px 5px;font-weight:600">${{fmt(p.hours)}}h</span>
      </span>`;
    }}).join('');

    return `<tr>
      <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;color:#334155;font-size:12px;vertical-align:top">${{e.EmpCd}}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;color:#1a1a2e;font-weight:500;font-size:12px;vertical-align:top;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:180px" title="${{e.Name}}">${{e.Name}}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;color:#475569;font-size:12px;vertical-align:top;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${{e.Dept}}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;text-align:right;font-weight:600;color:#166534;font-size:13px;vertical-align:top">${{fmt(e.TotalHours)}}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;text-align:right;color:#64748b;font-size:12px;vertical-align:top">${{e.ProjectCount}}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;vertical-align:top;line-height:1.6">${{projHtml}}</td>
    </tr>`;
  }}).join('');

  // pagination
  let pgHtml = '';
  if (pages > 1) {{
    pgHtml += `<button class="pg-btn" onclick="epPage(${{EP_PAGE-1}})" ${{EP_PAGE===1?'disabled':''}}>&#8249;</button>`;
    for (let p = Math.max(1,EP_PAGE-2); p <= Math.min(pages,EP_PAGE+2); p++) {{
      pgHtml += `<button class="pg-btn ${{p===EP_PAGE?'on':''}}" onclick="epPage(${{p}})">${{p}}</button>`;
    }}
    pgHtml += `<button class="pg-btn" onclick="epPage(${{EP_PAGE+1}})" ${{EP_PAGE===pages?'disabled':''}}>&#8250;</button>`;
  }}
  document.getElementById('epPgBtns').innerHTML = pgHtml;
}}

function epPage(p) {{
  EP_PAGE = p;
  renderEmpProj();
}}

// ─── Compare panel ───
let CMP_VIEW   = 'defaulters';
let CMP_STATUS = 'all';
let CMP_PAGE   = 1;
const PER_PAGE = 20;

function buildCompare() {{
  const c = DATA.compare;
  const nonCons = c.active_total - c.active_consultants;
  document.getElementById('t-cmp').textContent = c.defaulters + ' defaulters';
  document.getElementById('cmp-metrics').innerHTML = [
    metricCard('Active (GHR)',        fmt(c.active_total),        'From GHR',               'blue'),
    metricCard('Consultants (exempt)',fmt(c.active_consultants),  'Excluded from check',    ''),
    metricCard('Filling TT',          fmt(c.filling_tt),          'Logged at least 1 hr',   'green'),
    metricCard('Defaulters',          fmt(c.defaulters),          'Not in TT (non-consul.)', 'red'),
  ].join('');

  new Chart(document.getElementById('complianceChart'), {{
    type: 'doughnut',
    data: {{
      labels: ['Filling TT','Defaulters'],
      datasets: [{{ data:[c.filling_tt, c.defaulters],
        backgroundColor:['#22c55e','#ef4444'], borderWidth:2 }}]
    }},
    options:{{ responsive:true, maintainAspectRatio:false,
      plugins:{{ legend:{{display:false}},
        tooltip:{{ callbacks:{{ label: ctx => ctx.label + ': ' + fmt(ctx.raw) }} }} }} }}
  }});
  document.getElementById('cmp-legend').innerHTML = [
    ['Filling TT', c.filling_tt, '#22c55e'],
    ['Defaulters', c.defaulters, '#ef4444']
  ].map(([l,n,c]) => `<span class="legend-item"><span class="legend-dot" style="background:${{c}}"></span>${{l}} — ${{fmt(n)}} (${{pct(n, nonCons)}})</span>`).join('');

  progBars('default-dept-bars', c.defaulters_by_dept, 'dept', 'count', '#ef4444');
  renderTable();
}}

function switchView(v) {{
  CMP_VIEW = v; CMP_PAGE = 1;
  document.getElementById('vt-def').classList.toggle('on', v==='defaulters');
  document.getElementById('vt-fil').classList.toggle('on', v==='fillers');
  setStatusSilent('all');
  renderTable();
}}

function setStatusSilent(s) {{
  CMP_STATUS = s;
  document.querySelectorAll('#statusFilters .flt-btn').forEach(b => {{
    b.classList.toggle('on', b.textContent.trim() === (s==='all'?'All':s));
  }});
}}

function setStatus(s, el) {{
  CMP_STATUS = s; CMP_PAGE = 1;
  document.querySelectorAll('#statusFilters .flt-btn').forEach(b => b.classList.remove('on'));
  el.classList.add('on');
  renderTable();
}}

function renderTable() {{
  const q   = document.getElementById('cmpSearch').value.toLowerCase();
  const src = CMP_VIEW === 'defaulters' ? DATA.compare.defaulters_list : DATA.compare.fillers;

  let rows = src.filter(r => {{
    const matchQ = !q || [r.Name,r.Department,r.Manager,r.EmpCd].some(v => String(v).toLowerCase().includes(q));
    const matchS = CMP_STATUS === 'all' || r.Status === CMP_STATUS;
    return matchQ && matchS;
  }});

  const total = rows.length;
  const pages = Math.max(1, Math.ceil(total / PER_PAGE));
  CMP_PAGE    = Math.min(CMP_PAGE, pages);
  const slice = rows.slice((CMP_PAGE-1)*PER_PAGE, CMP_PAGE*PER_PAGE);
  const isDefault = CMP_VIEW === 'defaulters';

  document.getElementById('pgInfo').textContent =
    `Showing ${{Math.min((CMP_PAGE-1)*PER_PAGE+1, total)}}–${{Math.min(CMP_PAGE*PER_PAGE, total)}} of ${{total}} ${{isDefault ? 'defaulters' : 'employees filing TT'}}`;

  document.getElementById('tblHead').innerHTML = `<tr>
    <th style="width:70px">Emp ID</th>
    <th style="width:200px">Name</th>
    <th style="width:140px">Department</th>
    <th style="width:80px">Status</th>
    <th>Manager</th>
    ${{isDefault ? '' : '<th style="width:80px;text-align:right">Hours</th>'}}
  </tr>`;

  const pillClass = r => r.Status === 'Intern' ? 'pill-intern' : 'pill-confirmed';
  document.getElementById('tblBody').innerHTML = slice.map(r => `<tr>
    <td>${{r.EmpCd}}</td>
    <td title="${{r.Name}}">${{r.Name}}</td>
    <td title="${{r.Department}}">${{r.Department}}</td>
    <td><span class="pill ${{pillClass(r)}}">${{r.Status}}</span></td>
    <td title="${{r.Manager}}">${{r.Manager}}</td>
    ${{isDefault ? '' : `<td style="text-align:right;font-weight:500">${{fmt(r.Hours)}}</td>`}}
  </tr>`).join('');

  let pgHtml = '';
  if (pages > 1) {{
    pgHtml += `<button class="pg-btn" onclick="goPage(${{CMP_PAGE-1}})" ${{CMP_PAGE===1?'disabled':''}}>&#8249;</button>`;
    for (let p = Math.max(1,CMP_PAGE-2); p <= Math.min(pages,CMP_PAGE+2); p++) {{
      pgHtml += `<button class="pg-btn ${{p===CMP_PAGE?'on':''}}" onclick="goPage(${{p}})">${{p}}</button>`;
    }}
    pgHtml += `<button class="pg-btn" onclick="goPage(${{CMP_PAGE+1}})" ${{CMP_PAGE===pages?'disabled':''}}>&#8250;</button>`;
  }}
  document.getElementById('pgBtns').innerHTML = pgHtml;
}}

function goPage(p) {{
  const src   = CMP_VIEW === 'defaulters' ? DATA.compare.defaulters_list : DATA.compare.fillers;
  const pages = Math.ceil(src.length / PER_PAGE);
  if (p < 1 || p > pages) return;
  CMP_PAGE = p;
  renderTable();
}}

// ─── Init ───
buildGHR();
buildTT();
buildCompare();
</script>
</body>
</html>"""



app = Flask(__name__)

@app.route("/")
def dashboard():
    data = load_data()
    return build_html(data)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)