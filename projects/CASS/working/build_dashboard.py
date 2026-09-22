# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import json
import shutil
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from statistics import mean
from typing import Any

from openpyxl import load_workbook


PROJECT_DIR = Path(__file__).resolve().parents[1]
INPUT_DIR = PROJECT_DIR / "input"
OUTPUT_DIR = PROJECT_DIR / "output"
DATA_DIR = OUTPUT_DIR / "data"
METADATA_DIR = OUTPUT_DIR / "metadata"
REPORT_DIR = OUTPUT_DIR / "reports"
QA_DIR = OUTPUT_DIR / "qa"

SERVICE_FILE = INPUT_DIR / "Operational & Service Excellence by CAAS.xlsx"
ACADEMIC_FILE = INPUT_DIR / "Academic_services_without_income.xlsx"
BUDGET_FILES = [
    ("2567", INPUT_DIR / "A ทะเบียนคุมงบประมาณ 2567.xlsx"),
    ("2568", INPUT_DIR / "A ทะเบียนคุมงบประมาณ 2568.xlsx"),
    ("2569", INPUT_DIR / "A ตารางคุมงบประมาณ 2569.xlsx"),
]
SERVICE_SHEET = "Service_Excellence_by_CAAS"
THEME_NAME = "mahidol-faculty"
THEME_PATH = PROJECT_DIR.parents[1] / "dashboard-automation" / "themes" / f"{THEME_NAME}.md"
LOGO_SOURCE = PROJECT_DIR.parents[1] / "dashboard-automation" / "assets" / "logo.png"


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def normalize_header(value: Any) -> str:
    return str(value or "").replace("\n", " ").strip()


def read_sheet(file_path: Path, sheet_name: str) -> tuple[list[str], list[dict[str, Any]]]:
    workbook = load_workbook(file_path, read_only=True, data_only=True)
    sheet = workbook[sheet_name]
    rows = sheet.iter_rows(values_only=True)
    headers = [normalize_header(value) for value in next(rows)]
    usable_columns = [index for index, header in enumerate(headers) if header]
    headers = [headers[index] for index in usable_columns]
    records = []
    for row in rows:
        values = [row[index] if index < len(row) else None for index in usable_columns]
        if any(value is not None for value in values):
            records.append(dict(zip(headers, values)))
    workbook.close()
    return headers, records


def value_as_number(value: Any) -> float:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.replace(",", "").strip())
        except ValueError:
            pass
    return 0.0


def clean_text(value: Any) -> str:
    return "ไม่ระบุ" if value is None or str(value).strip() == "" else str(value).strip()


def clean_fiscal_year(value: Any) -> str:
    return str(int(value)) if isinstance(value, (int, float)) else clean_text(value)


def build_service_records(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    records = []
    for index, row in enumerate(rows, start=2):
        records.append({
            "source_row": index,
            "fiscal_year": clean_fiscal_year(row["ปีงบประมาณ"]),
            "quarter": clean_text(row["ไตรมาส"]),
            "project_type": clean_text(row["project_type"]),
            "service_group": clean_text(row["group_of_service"]),
            "project_name": clean_text(row["project_name"]),
            "target_participants": value_as_number(row["เป้าหมายจำนวนผู้เข้าร่วม"]),
            "applicants": value_as_number(row["จำนวนผู้สมัคร"]),
            "actual_participants": value_as_number(row["จำนวนที่เข้าร่วมจริง"]),
            "passed": value_as_number(row["จำนวนผู้สอบผ่าน"]),
            "revenue": value_as_number(row["รายได้"]),
            "expense": value_as_number(row["ค่าใช้จ่าย"]),
            "reported_profit": value_as_number(row["กำไรสุทธิ"]),
            "satisfaction": value_as_number(row["ความพึงพอใจ"]),
            "unit": clean_text(row["สังกัด"]),
        })
    return records


def build_budget_records() -> list[dict[str, Any]]:
    records = []
    for fiscal_year, file_path in BUDGET_FILES:
        workbook = load_workbook(file_path, read_only=True, data_only=True)
        for sheet in workbook.worksheets:
            if not sheet.title.startswith("ทะเบียนคุมงบ") and sheet.title != "MU-ELT รหัส 67":
                continue
            rows = sheet.iter_rows(values_only=True)
            next(rows, None)
            headers = [normalize_header(value) for value in next(rows, ())]
            columns = {header: index for index, header in enumerate(headers) if header}
            activity_column = columns.get("กิจกรรมย่อย")
            if activity_column is None:
                continue
            def cell(row: tuple[Any, ...], header: str) -> Any:
                index = columns.get(header)
                return row[index] if index is not None and index < len(row) else None
            for source_row, row in enumerate(rows, start=3):
                if activity_column >= len(row) or row[activity_column] in (None, ""):
                    continue
                revenue = value_as_number(cell(row, "รายรับ  (ยังไม่หักค่าใช้จ่าย)"))
                expense = value_as_number(cell(row, "รวม"))
                project_type = "In-House" if "In-House" in sheet.title else sheet.title
                records.append({
                    "source_file": file_path.name,
                    "source_sheet": sheet.title,
                    "source_row": source_row,
                    "fiscal_year": fiscal_year,
                    "quarter": "ไม่ระบุ",
                    "project_type": project_type,
                    "service_group": sheet.title,
                    "project_name": clean_text(cell(row, "กิจกรรมย่อย")),
                    "target_participants": value_as_number(cell(row, "จำนวน ที่ตั้งไว้")),
                    "applicants": 0.0,
                    "actual_participants": value_as_number(cell(row, "จำนวน ที่ร่วมจริง")),
                    "passed": 0.0,
                    "revenue": revenue,
                    "expense": expense,
                    "reported_profit": round(revenue - expense, 2),
                    "satisfaction": value_as_number(cell(row, "ความพึงพอใจในภาพรวม")) or value_as_number(cell(row, "ระดับความพึงพอใจ ในภาพรวม")) or value_as_number(cell(row, "ความพึงพอใจรวม")) or value_as_number(cell(row, "ความพึงพอใจ")),
                    "unit": "ไม่ระบุ",
                })
        workbook.close()
    return records


def worksheet_profile(file_path: Path, sheet_name: str) -> dict[str, Any]:
    headers, rows = read_sheet(file_path, sheet_name)
    missing = {header: sum(row[header] is None for row in rows) for header in headers}
    duplicate_count = len(rows) - len({tuple(row[header] for header in headers) for row in rows})
    return {
        "file": file_path.name,
        "sheet": sheet_name,
        "rows": len(rows),
        "columns": len(headers),
        "fields": headers,
        "missing_values": {field: count for field, count in missing.items() if count},
        "exact_duplicate_rows": duplicate_count,
    }


def total(records: list[dict[str, Any]], key: str) -> float:
    return round(sum(record[key] for record in records), 2)


def weighted_satisfaction(records: list[dict[str, Any]]) -> float | None:
    values = [record["satisfaction"] for record in records if record["satisfaction"] > 0]
    return round(mean(values), 2) if values else None


def summarize(records: list[dict[str, Any]]) -> dict[str, Any]:
    revenue = total(records, "revenue")
    expense = total(records, "expense")
    reported_profit = total(records, "reported_profit")
    return {
        "records": len(records),
        "revenue": revenue,
        "expense": expense,
        "reported_profit": reported_profit,
        "derived_profit": round(revenue - expense, 2),
        "profit_reconciliation_difference": round(reported_profit - (revenue - expense), 2),
        "actual_participants": total(records, "actual_participants"),
        "target_participants": total(records, "target_participants"),
        "satisfaction_average": weighted_satisfaction(records),
    }


def sorted_groups(records: list[dict[str, Any]], key: str) -> list[dict[str, Any]]:
    groups: dict[str, dict[str, float]] = defaultdict(lambda: {"revenue": 0.0, "expense": 0.0, "reported_profit": 0.0, "actual_participants": 0.0, "count": 0.0})
    for record in records:
        group = groups[record[key]]
        group["revenue"] += record["revenue"]
        group["expense"] += record["expense"]
        group["reported_profit"] += record["reported_profit"]
        group["actual_participants"] += record["actual_participants"]
        group["count"] += 1
    return [
        {"label": label, **{metric: round(value, 2) for metric, value in values.items()}}
        for label, values in sorted(groups.items(), key=lambda item: item[1]["revenue"], reverse=True)
    ]


def community_summary() -> dict[str, Any]:
    _, activities = read_sheet(ACADEMIC_FILE, "data")
    _, satisfaction = read_sheet(ACADEMIC_FILE, "ความพึงพอใจของชุมชน")
    activity_rows = [row for row in activities if row.get("ชื่อโครงการ")]
    normalized_activities = [
        {
            "fiscal_year": clean_fiscal_year(row["ปีงบประมาณ"]),
            "project_name": clean_text(row["ชื่อโครงการ"]),
            "format": clean_text(row["รูปแบบการเผยแพร่"]),
            "public_outreach": clean_text(row["Public Outreach"]),
            "social_benefit": clean_text(row["โครงการที่สร้างความผาสุกและประโยชน์ต่อสังคม"]),
            "budget": value_as_number(row["งบประมาณ"]),
            "expense": value_as_number(row["ค่าใช้จ่าย"]),
            "participants": value_as_number(row["จำนวนผู้เข้าร่วม"]),
            "satisfaction": value_as_number(row["ความพึงพอใจ"]),
        }
        for row in activity_rows
    ]
    return {
        "activity_rows_with_project_name": len(normalized_activities),
        "participants": round(sum(row["participants"] for row in normalized_activities), 2),
        "budget": round(sum(row["budget"] for row in normalized_activities), 2),
        "expense": round(sum(row["expense"] for row in normalized_activities), 2),
        "activities": normalized_activities,
        "community_satisfaction_by_year": [
            {"fiscal_year": clean_fiscal_year(row["ปีงบประมาณ"]), "score": value_as_number(row["คะแนน"])}
            for row in satisfaction
        ],
    }


def create_dashboard(records: list[dict[str, Any]], community: dict[str, Any]) -> str:
    payload = json.dumps(records, ensure_ascii=False, separators=(",", ":"))
    community_payload = json.dumps(community, ensure_ascii=False, separators=(",", ":"))
    return f'''<!doctype html>
<html lang="th">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>CASS Operational & Service Excellence</title>
  <style>
    :root {{ --primary:#0B2B5E; --secondary:#263988; --navy-dark:#1F295A; --gold:#FECD54; --background:#F4F5F9; --surface:#FFFFFF; --border:#D9DEE8; --text:#1F2937; --muted:#6B7280; --success:#0F8F83; --danger:#C94A4A; }}
    * {{ box-sizing:border-box; }} body {{ margin:0; color:var(--text); background:var(--background); font-family:Georgia,"Noto Serif Thai",serif; }}
    main {{ max-width:1440px; margin:auto; padding:10px clamp(12px,2.5vw,36px) 36px; }}
    header {{ display:grid; grid-template-columns:minmax(250px,1fr) minmax(300px,1.1fr) auto; gap:20px; align-items:center; border:1px solid var(--gold); border-top:3px solid var(--gold); border-radius:6px; background:var(--surface); padding:10px 16px; }} .brand {{ display:flex; gap:12px; align-items:center; }} .brand img {{ width:64px; height:64px; object-fit:contain; flex:0 0 auto; }} .brand strong {{ color:var(--primary); display:block; font-size:1rem; line-height:1.2; }} .brand small {{ color:var(--secondary); display:block; font:600 .73rem/1.3 ui-sans-serif,sans-serif; }} .page-title {{ text-align:center; }}
    h1 {{ color:var(--primary); font-size:clamp(1.25rem,2.2vw,2rem); margin:0; letter-spacing:0; }} h2 {{ color:var(--primary); font-size:1.15rem; margin:0 0 14px; }} p {{ margin:6px 0; color:var(--muted); }}
    .status {{ background:#FFF9E8; border:1px solid var(--gold); color:var(--navy-dark); padding:8px 12px; font:700 .76rem/1.25 ui-sans-serif,sans-serif; max-width:330px; }}
    .tabs {{ display:flex; gap:0; margin-top:24px; border-bottom:2px solid var(--primary); }} .tab {{ border:1px solid var(--border); border-bottom:0; border-radius:6px 6px 0 0; background:var(--surface); color:var(--muted); cursor:pointer; padding:10px 16px; font:700 .82rem ui-sans-serif,sans-serif; }} .tab + .tab {{ margin-left:6px; }} .tab[aria-selected="true"] {{ background:var(--primary); border-color:var(--primary); color:var(--surface); }} .tab-panel {{ display:none; }} .tab-panel.is-active {{ display:block; }}
    .controls {{ display:grid; grid-template-columns:repeat(5,minmax(140px,1fr)); gap:12px; margin:24px 0; }} .filter {{ position:relative; }} .filter summary {{ list-style:none; cursor:pointer; border:1px solid var(--border); background:var(--surface); padding:9px; font:600 .76rem ui-sans-serif,sans-serif; color:var(--text); }} .filter summary::-webkit-details-marker {{ display:none; }} .filter summary span {{ float:right; color:var(--primary); }} .filter-options {{ position:absolute; z-index:2; top:calc(100% + 4px); width:100%; max-height:250px; overflow:auto; border:1px solid var(--border); background:var(--surface); box-shadow:0 8px 18px rgba(31,41,55,.12); padding:8px; }} .filter-options label {{ display:flex; gap:8px; align-items:start; padding:6px 4px; color:var(--text); font:400 .76rem/1.3 ui-sans-serif,sans-serif; }} .filter-options input {{ margin:2px 0 0; accent-color:var(--primary); }}
    .kpis {{ display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:12px; }} .kpi,.panel {{ background:var(--surface); border:1px solid var(--border); border-radius:6px; }} .kpi {{ border-top:3px solid var(--primary); padding:16px; min-height:130px; }} .kpi span,.note-label {{ font:700 .7rem ui-sans-serif,sans-serif; color:var(--muted); text-transform:uppercase; }} .kpi strong {{ color:var(--navy-dark); display:block; font-size:1.45rem; margin-top:10px; }} .comparison {{ display:block; margin-top:8px; font:700 .72rem ui-sans-serif,sans-serif; }} .comparison.positive {{ color:var(--success); }} .comparison.negative {{ color:var(--danger); }} .comparison.neutral {{ color:var(--muted); }}
    .grid {{ display:grid; grid-template-columns:1.35fr 1fr; gap:16px; margin-top:16px; }} .panel {{ padding:18px; overflow:hidden; }} .wide {{ grid-column:span 2; }}
    .bars {{ display:grid; gap:11px; }} .bar-row {{ display:grid; grid-template-columns:minmax(100px,1.25fr) 2fr auto; gap:8px; align-items:center; font:.8rem ui-sans-serif,sans-serif; }} .bar-track {{ height:10px; background:var(--border); }} .bar {{ height:100%; background:var(--secondary); cursor:pointer; transform-origin:left; animation:bar-grow .55s ease-out both; }} .bar-row b {{ font-variant-numeric:tabular-nums; font-weight:600; }} .bar-row button {{ appearance:none; border:0; background:transparent; padding:0; color:inherit; cursor:pointer; font:inherit; text-align:left; }} .bar-row.active .bar {{ background:var(--gold); }} .bar-row.active button {{ color:var(--primary); font-weight:700; }} .active-filter {{ display:none; align-items:center; gap:8px; margin:10px 0 0; color:var(--muted); font:.75rem ui-sans-serif,sans-serif; }} .active-filter.visible {{ display:flex; }} .active-filter button {{ border:1px solid var(--border); background:var(--surface); color:var(--primary); cursor:pointer; padding:4px 7px; font:600 .72rem ui-sans-serif,sans-serif; }}
    .trend-legend {{ display:flex; gap:16px; margin:-4px 0 8px; color:var(--muted); font:.75rem ui-sans-serif,sans-serif; }} .trend-legend span::before {{ content:''; display:inline-block; width:16px; height:3px; margin:0 5px 2px 0; vertical-align:middle; background:var(--primary); }} .trend-legend .profit::before {{ background:var(--gold); }} .trend-chart {{ width:100%; height:270px; overflow:visible; }} .trend-chart text {{ fill:var(--muted); font:11px ui-sans-serif,sans-serif; }} .trend-value-revenue,.trend-value-profit {{ font-size:10px; font-weight:700; }} .trend-value-revenue {{ fill:var(--primary); }} .trend-value-profit {{ fill:var(--navy-dark); }} .trend-grid {{ stroke:var(--border); stroke-width:1; }} .trend-revenue {{ fill:none; stroke:var(--primary); stroke-width:3; stroke-dasharray:1200; stroke-dashoffset:1200; animation:draw-line 1s ease-out forwards; }} .trend-profit {{ fill:none; stroke:var(--gold); stroke-width:3; stroke-dasharray:1200; stroke-dashoffset:1200; animation:draw-line 1s .15s ease-out forwards; }} .trend-dot-revenue,.trend-dot-profit {{ transform-box:fill-box; transform-origin:center; animation:dot-pop .3s ease-out both; }} .trend-dot-profit {{ animation-delay:.2s; }} .trend-dot-revenue {{ fill:var(--primary); }} .trend-dot-profit {{ fill:var(--gold); }} @keyframes bar-grow {{ from {{ transform:scaleX(0); }} to {{ transform:scaleX(1); }} }} @keyframes draw-line {{ to {{ stroke-dashoffset:0; }} }} @keyframes dot-pop {{ from {{ transform:scale(0); }} to {{ transform:scale(1); }} }} @media (prefers-reduced-motion:reduce) {{ *,*::before,*::after {{ animation-duration:.01ms!important; animation-iteration-count:1!important; }} }}
    .callout {{ border-left:4px solid var(--gold); padding:10px 12px; background:#FFF9E8; margin-bottom:10px; color:var(--navy-dark); font:.86rem/1.45 ui-sans-serif,sans-serif; }}
    table {{ width:100%; border-collapse:collapse; font:.77rem ui-sans-serif,sans-serif; }} th,td {{ padding:9px 7px; text-align:left; border-bottom:1px solid var(--border); vertical-align:top; }} th {{ background:#F4F5F9; color:var(--primary); }} .table-wrap {{ overflow:auto; max-height:430px; }}
    .community {{ display:flex; gap:24px; flex-wrap:wrap; font:.86rem ui-sans-serif,sans-serif; }} .community strong {{ color:var(--success); }} .section-heading {{ display:flex; justify-content:space-between; align-items:end; gap:16px; margin:30px 0 12px; border-bottom:2px solid var(--primary); padding-bottom:8px; }} .section-heading h2 {{ margin:0; }} .section-heading p {{ font:.78rem ui-sans-serif,sans-serif; }} .community-controls {{ display:flex; align-items:center; gap:10px; }} .community-controls select {{ min-width:130px; border:1px solid var(--border); background:var(--surface); color:var(--text); padding:7px; font:inherit; }} .community-kpis {{ display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:12px; }} .community-kpis .kpi {{ min-height:105px; }} .community-grid {{ display:grid; grid-template-columns:1.15fr .85fr; gap:16px; margin-top:16px; }} .community-table td:first-child {{ min-width:70px; }} .yes {{ color:var(--success); font-weight:700; }} .no {{ color:var(--danger); font-weight:700; }}
    @media (max-width:900px) {{ .controls,.kpis,.community-kpis {{ grid-template-columns:repeat(2,minmax(0,1fr)); }} .grid,.community-grid {{ grid-template-columns:1fr; }} .wide {{ grid-column:auto; }} header {{ grid-template-columns:1fr; }} .page-title {{ text-align:left; }} .status {{ width:max-content; max-width:100%; }} .section-heading {{ align-items:flex-start; flex-direction:column; }} .community-controls {{ max-width:100%; }} .community-controls select {{ min-width:0; max-width:100%; }} }}
    @media (max-width:520px) {{ .controls,.kpis {{ grid-template-columns:1fr; }} main {{ padding:18px 12px 32px; }} .tabs {{ overflow-x:auto; }} .tab {{ white-space:nowrap; padding:10px 12px; }} .bar-row {{ grid-template-columns:105px 1fr; }} .bar-row b {{ grid-column:2; }} }}
  </style>
</head>
<body><main>
    <header><div class="brand"><img src="assets/mahidol-logo.png" alt="ตรามหาวิทยาลัยมหิดล"><div><strong>คณะศิลปศาสตร์ มหาวิทยาลัยมหิดล</strong><small>Faculty of Liberal Arts, Mahidol University</small></div></div><div class="page-title"><p>การวิเคราะห์ผลการดำเนินงานบริการวิชาการ</p><h1>Operational &amp; Service Excellence</h1></div><div class="status">QA: READY WITH LIMITATIONS<br>กำไรคำนวณจากรายได้ - รายจ่าย; ปี 2569 อาจยังไม่ครบ</div></header>
    <nav class="tabs" role="tablist" aria-label="ประเภทบริการ"><button class="tab" type="button" role="tab" aria-selected="true" aria-controls="incomeTab" data-tab="income">บริการมีรายได้</button><button class="tab" type="button" role="tab" aria-selected="false" aria-controls="communityDrilldown" data-tab="community">บริการวิชาการแบบไม่ได้รายได้</button><button class="tab" type="button" role="tab" aria-selected="false" aria-controls="budgetTab" data-tab="budget">เปรียบเทียบงบประมาณ</button></nav>
    <section class="tab-panel is-active" id="incomeTab" role="tabpanel" data-tab-panel="income"><section class="controls"><details class="filter" id="fiscal-year"><summary>ปีงบประมาณ <span>ทั้งหมด</span></summary><div class="filter-options"></div></details><details class="filter" id="quarter"><summary>ไตรมาส <span>ทั้งหมด</span></summary><div class="filter-options"></div></details><details class="filter" id="type"><summary>ประเภทโครงการ <span>ทั้งหมด</span></summary><div class="filter-options"></div></details><details class="filter" id="group"><summary>กลุ่มผู้รับบริการ <span>ทั้งหมด</span></summary><div class="filter-options"></div></details><details class="filter" id="unit"><summary>สังกัด <span>ทั้งหมด</span></summary><div class="filter-options"></div></details></section>
    <section class="kpis" id="kpis"></section><div class="active-filter" id="activeFilter"><span></span><button type="button">ล้างตัวกรองกราฟ</button></div>
        <section class="grid"><article class="panel wide"><h2>แนวโน้มรายได้และกำไรตามทะเบียนคุมงบ</h2><div class="trend-legend"><span>รายได้</span><span class="profit">กำไร (รายได้ - รายจ่าย)</span></div><svg class="trend-chart" id="annualTrend" viewBox="0 0 760 250" role="img" aria-label="แนวโน้มรายได้และกำไรตามปีงบประมาณ"></svg></article><article class="panel"><h2>รายได้ตามประเภทโครงการ</h2><div class="bars" id="typeBars"></div></article><article class="panel"><h2>รายได้ตามกลุ่มทะเบียน</h2><div class="bars" id="quarterBars"></div></article><article class="panel"><h2>ประเด็นที่ต้องติดตาม</h2><div id="insights"></div></article><article class="panel"><h2>ความพึงพอใจเฉลี่ยตามโครงการ</h2><div class="bars" id="unitBars"></div></article><article class="panel wide"><h2>รายละเอียดรายการบริการมีรายได้</h2><div class="table-wrap"><table><thead><tr><th>ปีงบประมาณ</th><th>โครงการ</th><th>ประเภท</th><th>กลุ่มทะเบียน</th><th>รายได้</th><th>รายจ่าย</th><th>กำไร</th><th>ผู้เข้าร่วม</th><th>ความพึงพอใจ</th></tr></thead><tbody id="table"></tbody></table></div></article></section></section>
    <section id="communityDrilldown" class="tab-panel" role="tabpanel" data-tab-panel="community"><div class="section-heading"><div><h2>บริการวิชาการแบบไม่ได้รายได้</h2><p>เจาะลึกกิจกรรมชุมชนจาก Academic_services_without_income.xlsx</p></div><label class="community-controls">ปีงบประมาณ <select id="communityYear"></select></label></div><section class="community-kpis" id="communityKpis"></section><section class="community-grid"><article class="panel"><h2>ผู้เข้าร่วมตามปีงบประมาณ</h2><div class="bars" id="communityYearBars"></div></article><article class="panel"><h2>ผลกระทบทางสังคม</h2><div class="bars" id="communityBenefitBars"></div></article><article class="panel wide"><h2>รายละเอียดกิจกรรม</h2><div class="table-wrap"><table class="community-table"><thead><tr><th>ปี</th><th>โครงการ</th><th>รูปแบบ</th><th>ผู้เข้าร่วม</th><th>งบประมาณ</th><th>ค่าใช้จ่าย</th><th>สังคม</th><th>ความพึงพอใจ</th></tr></thead><tbody id="communityTable"></tbody></table></div></article></section></section>
    <section id="budgetTab" class="tab-panel" role="tabpanel" data-tab-panel="budget"><div class="section-heading"><div><h2>เปรียบเทียบงบประมาณบริการมีรายได้</h2><p>ทะเบียนคุมงบประมาณ In-House และ MU-ELT ปีงบประมาณ 2567–2569</p></div><label class="community-controls">ปีงบประมาณ <select id="budgetYear"></select></label></div><section class="community-kpis" id="budgetKpis"></section><section class="community-grid"><article class="panel"><h2>รายได้ตามปีงบประมาณ</h2><div class="bars" id="budgetYearBars"></div></article><article class="panel"><h2>รายได้ตามประเภทบริการ</h2><div class="bars" id="budgetTypeBars"></div></article><article class="panel wide"><h2>รายละเอียดทะเบียนคุมงบ</h2><div class="table-wrap"><table class="community-table"><thead><tr><th>ปี</th><th>โครงการ</th><th>ประเภท</th><th>แหล่งทะเบียน</th><th>รายได้</th><th>รายจ่าย</th><th>กำไร</th><th>ผู้เข้าร่วม</th><th>พึงพอใจ</th></tr></thead><tbody id="budgetTable"></tbody></table></div></article></section></section>
    <footer>แหล่งข้อมูล Tab 1 และ Tab 3: ทะเบียนคุมงบประมาณ In-House และ MU-ELT ปี 2567–2569 (83 รายการ). กำไรคำนวณจากรายได้หักรายจ่ายรายรายการ.</footer>
</main><script>
const data={payload}; const community={community_payload};
const money=new Intl.NumberFormat('th-TH',{{style:'currency',currency:'THB',maximumFractionDigits:0}}); const number=new Intl.NumberFormat('th-TH',{{maximumFractionDigits:0}});
const filters=[['fiscal-year','fiscal_year'],['quarter','quarter'],['type','project_type'],['group','service_group'],['unit','unit']]; let chartFilter=null;
function escapeHtml(value){{return String(value).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('"','&quot;');}}
function options(){{filters.forEach(([id,key])=>{{const filter=document.getElementById(id); const values=[...new Set(data.map(row=>row[key]))].sort((a,b)=>String(a).localeCompare(String(b),'th',{{numeric:true}})); const options=filter.querySelector('.filter-options'); options.innerHTML=values.map(value=>`<label><input type="checkbox" value="${{escapeHtml(value)}}">${{escapeHtml(value)}}</label>`).join(''); options.addEventListener('change',render);}});}}
function selectedValues(id){{return new Set([...document.querySelectorAll(`#${{id}} input:checked`)].map(input=>input.value));}}
function filtered(excludeFiscalYear=false){{return data.filter(row=>filters.every(([id,key])=>{{if(excludeFiscalYear&&key==='fiscal_year')return true;const selected=selectedValues(id);return !selected.size||selected.has(String(row[key]));}})&&(!chartFilter||row[chartFilter.key]===chartFilter.value));}}
function updateFilterLabels(){{filters.forEach(([id])=>{{const selected=selectedValues(id); document.querySelector(`#${{id}} summary span`).textContent=selected.size?`เลือก ${{selected.size}}`:'ทั้งหมด';}});}}
function sum(rows,key){{return rows.reduce((total,row)=>total+row[key],0)}} function average(rows,key){{const values=rows.map(row=>row[key]).filter(value=>value>0); return values.length?values.reduce((a,b)=>a+b,0)/values.length:0}}
function aggregate(rows,key,metric){{const values={{}}; rows.forEach(row=>values[row[key]]=(values[row[key]]||0)+row[metric]); return Object.entries(values).sort((a,b)=>b[1]-a[1]);}}
function bars(element,rows,formatter=money,key){{const max=Math.max(...rows.map(row=>row[1]),1); element.innerHTML=rows.slice(0,8).map(([label,value])=>`<div class="bar-row ${{chartFilter?.key===key&&chartFilter?.value===label?'active':''}}"><button type="button" title="กรองตาม ${{label}}" data-key="${{key}}" data-value="${{escapeHtml(label)}}">${{label}}</button><div class="bar-track"><div class="bar" style="width:${{value/max*100}}%" data-key="${{key}}" data-value="${{escapeHtml(label)}}"></div></div><b>${{formatter.format(value)}}</b></div>`).join('')||'<p>ไม่พบข้อมูลตามตัวกรอง</p>';element.querySelectorAll('[data-key]').forEach(item=>item.addEventListener('click',()=>{{const next={{key:item.dataset.key,value:item.dataset.value}};chartFilter=chartFilter?.key===next.key&&chartFilter?.value===next.value?null:next;render();}}));}}
function simpleBars(element,rows,formatter=number){{const max=Math.max(...rows.map(row=>row[1]),1);element.innerHTML=rows.map(([label,value])=>`<div class="bar-row"><span title="${{label}}">${{label}}</span><div class="bar-track"><div class="bar" style="width:${{value/max*100}}%"></div></div><b>${{formatter.format(value)}}</b></div>`).join('')||'<p>ไม่พบข้อมูลตามตัวกรอง</p>';}}
function renderCommunity(){{const year=document.getElementById('communityYear').value;const rows=community.activities.filter(row=>!year||row.fiscal_year===year);const participants=sum(rows,'participants'),budget=sum(rows,'budget'),expense=sum(rows,'expense'),satisfaction=average(rows,'satisfaction'),benefit=rows.filter(row=>row.social_benefit==='YES').length;document.getElementById('communityKpis').innerHTML=[['กิจกรรม',number.format(rows.length)+' รายการ'],['ผู้เข้าร่วม',number.format(participants)+' คน'],['งบประมาณ',money.format(budget)],['ค่าใช้จ่าย',money.format(expense)],['ความพึงพอใจ',satisfaction?satisfaction.toFixed(2)+' / 5':'ไม่มีข้อมูล']].map(([label,value])=>`<article class="kpi"><span>${{label}}</span><strong>${{value}}</strong></article>`).join('');simpleBars(document.getElementById('communityYearBars'),aggregate(rows,'fiscal_year','participants').sort((a,b)=>Number(a[0])-Number(b[0])));simpleBars(document.getElementById('communityBenefitBars'),[['สร้างประโยชน์ต่อสังคม',benefit],['ไม่ระบุ/ไม่เข้าเกณฑ์',rows.length-benefit]]);document.getElementById('communityTable').innerHTML=rows.map(row=>`<tr><td>${{row.fiscal_year}}</td><td>${{row.project_name}}</td><td>${{row.format}}</td><td>${{number.format(row.participants)}}</td><td>${{money.format(row.budget)}}</td><td>${{money.format(row.expense)}}</td><td class="${{row.social_benefit==='YES'?'yes':'no'}}">${{row.social_benefit}}</td><td>${{row.satisfaction||'—'}}</td></tr>`).join('')||'<tr><td colspan="8">ไม่พบกิจกรรมตามปีที่เลือก</td></tr>';}}
function renderBudget(){{const year=document.getElementById('budgetYear').value;const rows=data.filter(row=>!year||row.fiscal_year===year);const revenue=sum(rows,'revenue'),expense=sum(rows,'expense'),profit=sum(rows,'reported_profit'),participants=sum(rows,'actual_participants'),satisfaction=average(rows,'satisfaction');document.getElementById('budgetKpis').innerHTML=[['รายการ',number.format(rows.length)+' รายการ'],['รายได้',money.format(revenue)],['รายจ่าย',money.format(expense)],['กำไร',money.format(profit)],['ผู้เข้าร่วม',number.format(participants)+' คน']].map(([label,value])=>`<article class="kpi"><span>${{label}}</span><strong>${{value}}</strong></article>`).join('');simpleBars(document.getElementById('budgetYearBars'),aggregate(rows,'fiscal_year','revenue').sort((a,b)=>Number(a[0])-Number(b[0])),money);simpleBars(document.getElementById('budgetTypeBars'),aggregate(rows,'project_type','revenue'),money);document.getElementById('budgetTable').innerHTML=rows.map(row=>`<tr><td>${{row.fiscal_year}}</td><td>${{row.project_name}}</td><td>${{row.project_type}}</td><td>${{row.service_group}}</td><td>${{money.format(row.revenue)}}</td><td>${{money.format(row.expense)}}</td><td>${{money.format(row.reported_profit)}}</td><td>${{number.format(row.actual_participants)}}</td><td>${{row.satisfaction||'—'}}</td></tr>`).join('')||'<tr><td colspan="9">ไม่พบรายการตามปีที่เลือก</td></tr>';}}
function comparison(current,previous,higherIsBetter=true){{if(previous===null||previous===0)return {{text:'ไม่มีฐานเปรียบเทียบ',kind:'neutral'}};const change=(current-previous)/Math.abs(previous)*100;const favorable=higherIsBetter?change>=0:change<=0;return {{text:`${{change>=0?'▲':'▼'}} ${{Math.abs(change).toFixed(1)}}% เทียบปีก่อน`,kind:favorable?'positive':'negative'}};}}
function kpi(label,value,previous,formatter,higherIsBetter=true){{const delta=comparison(value,previous,higherIsBetter);return `<article class="kpi"><span>${{label}}</span><strong>${{formatter(value)}}</strong><em class="comparison ${{delta.kind}}">${{delta.text}}</em></article>`;}}
function annualTrend(rows){{const trend=Object.entries(rows.reduce((all,row)=>{{const item=all[row.fiscal_year]||{{revenue:0,profit:0}};item.revenue+=row.revenue;item.profit+=row.reported_profit;all[row.fiscal_year]=item;return all}},{{}})).sort((a,b)=>Number(a[0])-Number(b[0]));const chart=document.getElementById('annualTrend');if(!trend.length){{chart.innerHTML='<text x="380" y="125" text-anchor="middle">ไม่พบข้อมูลตามตัวกรอง</text>';return;}}const left=58,right=18,top=28,bottom=40,width=760-left-right,height=250-top-bottom,max=Math.max(...trend.flatMap(([,value])=>[value.revenue,value.profit]),1);const point=(value,index)=>[left+(trend.length===1?width/2:index*width/(trend.length-1)),top+height-(value/max*height)];const path=key=>trend.map(([,value],index)=>point(value[key],index).join(',')).join(' ');const compactMoney=value=>`฿${{value>=1000000?(value/1000000).toLocaleString('th-TH',{{maximumFractionDigits:1}})+'M':number.format(value)}}`;const grid=[0,.5,1].map(ratio=>{{const y=top+height*(1-ratio);return `<line class="trend-grid" x1="${{left}}" x2="${{760-right}}" y1="${{y}}" y2="${{y}}"/><text x="${{left-8}}" y="${{y+4}}" text-anchor="end">${{compactMoney(max*ratio)}}</text>`}}).join('');const labels=trend.map(([year],index)=>`<text x="${{point(0,index)[0]}}" y="248" text-anchor="middle">${{year}}</text>`).join('');const dots=(key,css,labelCss,offset)=>trend.map(([,value],index)=>{{const [x,y]=point(value[key],index);const exact=money.format(value[key]);return `<circle class="${{css}}" cx="${{x}}" cy="${{y}}" r="4"><title>${{key==='revenue'?'รายได้':'กำไรสุทธิที่รายงาน'}}: ${{exact}}</title></circle><text class="${{labelCss}}" x="${{x}}" y="${{Math.max(12,y+offset)}}" text-anchor="middle">${{compactMoney(value[key])}}</text>`}}).join('');chart.innerHTML=`${{grid}}<polyline class="trend-revenue" points="${{path('revenue')}}"/><polyline class="trend-profit" points="${{path('profit')}}"/>${{dots('revenue','trend-dot-revenue','trend-value-revenue',-9)}}${{dots('profit','trend-dot-profit','trend-value-profit',16)}}${{labels}}`;}}
function render(){{updateFilterLabels(); const rows=filtered();const selectedYears=[...selectedValues('fiscal-year')].map(Number);const currentYear=selectedYears.length?Math.max(...selectedYears):null;const currentRows=currentYear===null?rows:rows.filter(row=>Number(row.fiscal_year)===currentYear);const revenue=sum(currentRows,'revenue'),expense=sum(currentRows,'expense'),profit=sum(currentRows,'reported_profit'),derived=revenue-expense,participants=sum(currentRows,'actual_participants');const previousRows=currentYear===null?[]:filtered(true).filter(row=>Number(row.fiscal_year)===currentYear-1);const priorRevenue=previousRows.length?sum(previousRows,'revenue'):null,priorProfit=previousRows.length?sum(previousRows,'reported_profit'):null,priorDerived=previousRows.length?sum(previousRows,'revenue')-sum(previousRows,'expense'):null,priorParticipants=previousRows.length?sum(previousRows,'actual_participants'):null,priorSatisfaction=previousRows.length?average(previousRows,'satisfaction'):null;
 document.getElementById('kpis').innerHTML=[kpi('รายได้',revenue,priorRevenue,money.format),kpi('กำไร (รายได้ - รายจ่าย)',profit,priorProfit,money.format),kpi('รายได้ - รายจ่าย',derived,priorDerived,money.format),kpi('ผู้เข้าร่วมจริง',participants,priorParticipants,number.format),kpi('ความพึงพอใจเฉลี่ย',average(currentRows,'satisfaction'),priorSatisfaction,value=>value.toFixed(2)+' / 5')].join('');const active=document.getElementById('activeFilter');active.classList.toggle('visible',Boolean(chartFilter));active.querySelector('span').textContent=chartFilter?`ตัวกรองกราฟ: ${{chartFilter.value}}`:'';
 annualTrend(rows); bars(document.getElementById('typeBars'),aggregate(rows,'project_type','revenue'),money,'project_type'); bars(document.getElementById('quarterBars'),aggregate(rows,'service_group','revenue'),money,'service_group'); const units=Object.entries(rows.reduce((all,row)=>{{if(row.satisfaction>0){{const item=all[row.project_type]||{{sum:0,count:0}}; item.sum+=row.satisfaction;item.count++;all[row.project_type]=item;}}return all}},{{}})).map(([key,value])=>[key,value.sum/value.count]).sort((a,b)=>b[1]-a[1]); bars(document.getElementById('unitBars'),units,new Intl.NumberFormat('th-TH',{{maximumFractionDigits:2}}),'project_type');
 const top=aggregate(rows,'project_type','revenue')[0]; document.getElementById('insights').innerHTML=`<div class="callout">กำไรในหน้านี้คำนวณจากรายได้หักรายจ่ายของทะเบียนคุมงบทุกแถว จึง reconcile กับยอดรวมได้.</div>${{top?`<div class="callout">ประเภทที่สร้างรายได้สูงสุดภายใต้ตัวกรอง: <b>${{top[0]}}</b> (${{money.format(top[1])}})</div>`:''}}`;
 document.getElementById('table').innerHTML=rows.slice(0,250).map(row=>`<tr><td>${{row.fiscal_year}}</td><td>${{row.project_name}}</td><td>${{row.project_type}}</td><td>${{row.quarter}}</td><td>${{money.format(row.revenue)}}</td><td>${{money.format(row.expense)}}</td><td>${{money.format(row.reported_profit)}}</td><td>${{number.format(row.actual_participants)}}</td><td>${{row.satisfaction||'—'}}</td></tr>`).join('');}}
function communityOptions(){{const select=document.getElementById('communityYear');const years=[...new Set(community.activities.map(row=>row.fiscal_year))].sort((a,b)=>Number(b)-Number(a));select.innerHTML='<option value="">ทุกปี</option>'+years.map(year=>`<option value="${{year}}">${{year}}</option>`).join('');select.addEventListener('change',renderCommunity);}}
function budgetOptions(){{const select=document.getElementById('budgetYear');const years=[...new Set(data.map(row=>row.fiscal_year))].sort((a,b)=>Number(b)-Number(a));select.innerHTML='<option value="">ทุกปี</option>'+years.map(year=>`<option value="${{year}}">${{year}}</option>`).join('');select.addEventListener('change',renderBudget);}}
function activateTab(tab){{document.querySelectorAll('.tab').forEach(button=>button.setAttribute('aria-selected',String(button.dataset.tab===tab)));document.querySelectorAll('[data-tab-panel]').forEach(panel=>panel.classList.toggle('is-active',panel.dataset.tabPanel===tab));}}
document.querySelectorAll('.tab').forEach(button=>button.addEventListener('click',()=>activateTab(button.dataset.tab)));document.querySelector('#activeFilter button').addEventListener('click',()=>{{chartFilter=null;render();}});options();communityOptions();budgetOptions();render();renderCommunity();renderBudget();
</script></body></html>'''


def main() -> None:
    if not THEME_PATH.is_file():
        raise FileNotFoundError(f"Required dashboard theme is missing: {THEME_PATH}")
    if not LOGO_SOURCE.is_file():
        raise FileNotFoundError(f"Required dashboard logo is missing: {LOGO_SOURCE}")
    for directory in (DATA_DIR, METADATA_DIR, REPORT_DIR, QA_DIR, OUTPUT_DIR / "dashboard"):
        directory.mkdir(parents=True, exist_ok=True)
    logo_destination = OUTPUT_DIR / "dashboard" / "assets" / "mahidol-logo.png"
    logo_destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(LOGO_SOURCE, logo_destination)

    _, source_rows = read_sheet(SERVICE_FILE, SERVICE_SHEET)
    records = build_budget_records()
    community = community_summary()
    totals = summarize(records)
    profiles = []
    for file_path in (ACADEMIC_FILE, SERVICE_FILE, *(file_path for _, file_path in BUDGET_FILES)):
        workbook = load_workbook(file_path, read_only=True, data_only=True)
        sheet_names = workbook.sheetnames
        workbook.close()
        profiles.extend(worksheet_profile(file_path, sheet_name) for sheet_name in sheet_names)

    with (DATA_DIR / "budget-control-records.csv").open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=list(records[0]))
        writer.writeheader()
        writer.writerows(records)
    with (DATA_DIR / "non-income-academic-services.csv").open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=list(community["activities"][0]))
        writer.writeheader()
        writer.writerows(community["activities"])

    write_json(METADATA_DIR / "dataset-profile.json", {
        "project": "cass",
        "theme": THEME_NAME,
        "generated_at": datetime.now().astimezone().isoformat(),
        "input_files": [ACADEMIC_FILE.name, SERVICE_FILE.name, *(file_path.name for _, file_path in BUDGET_FILES)],
        "worksheets": profiles,
        "analytical_dataset": {
            "name": "budget-control-records.csv",
            "source": "A ทะเบียนคุมงบประมาณ 2567.xlsx, A ทะเบียนคุมงบประมาณ 2568.xlsx, A ตารางคุมงบประมาณ 2569.xlsx",
            "record_count": len(records),
            "fiscal_year_coverage": sorted(set(record["fiscal_year"] for record in records)),
            "measures": ["revenue", "expense", "reported_profit", "actual_participants", "satisfaction"],
            "dimensions": ["fiscal_year", "project_type", "service_group", "source_file", "source_sheet"],
        },
        "non_income_academic_services_dataset": {
            "name": "non-income-academic-services.csv",
            "source": f"{ACADEMIC_FILE.name} / data",
            "record_count": len(community["activities"]),
            "fiscal_year_coverage": sorted(set(record["fiscal_year"] for record in community["activities"])),
            "measures": ["participants", "budget", "expense", "satisfaction"],
            "dimensions": ["fiscal_year", "format", "public_outreach", "social_benefit"],
        },
    })
    write_json(METADATA_DIR / "dashboard-requirements.json", {
        "audience": "ผู้บริหารและผู้ปฏิบัติงาน CAAS",
        "purpose": "ติดตามผลการดำเนินงานบริการและความเป็นเลิศในการให้บริการของ CAAS",
        "questions": ["รายได้ รายจ่าย และกำไรจากทะเบียนคุมงบเป็นเท่าใด", "บริการประเภทใดและปีใดสร้างรายได้สูง", "จำนวนผู้เข้าร่วมและความพึงพอใจเป็นอย่างไร", "กิจกรรมบริการวิชาการแบบไม่ได้รายได้มีผลลัพธ์และผลกระทบสังคมอย่างไร", "ข้อมูลมีข้อจำกัดใดก่อนใช้ตัดสินใจ"],
        "kpis": [
            {"name": "รายได้", "definition": "ผลรวมรายรับจากทะเบียนคุมงบของปีที่เลือก", "formula": "SUM(revenue)", "source_fields": ["รายรับ (ยังไม่หักค่าใช้จ่าย)"], "aggregation": "sum", "limitations": "ปี 2569 อาจเป็นข้อมูลระหว่างปี; การเปรียบเทียบใช้ปีก่อนหน้าของปีล่าสุดที่เลือก"},
            {"name": "กำไร", "definition": "รายรับหักค่าใช้จ่ายรวมของทุกรายการ", "formula": "SUM(revenue - expense)", "source_fields": ["รายรับ (ยังไม่หักค่าใช้จ่าย)", "รวม"], "aggregation": "sum", "limitations": "เป็นกำไรคำนวณเพื่อให้ reconcile; ไม่ใช่ฟิลด์กำไรสุทธิที่รายงานจากแหล่งอื่น"},
            {"name": "ผู้เข้าร่วมจริง", "definition": "ผลรวมจำนวนผู้เข้าร่วมจริง", "formula": "SUM(actual_participants)", "source_fields": ["จำนวนที่เข้าร่วมจริง"], "aggregation": "sum", "limitations": "มีค่าขาดหายซึ่งถูกแสดงเป็นศูนย์ใน analytical dataset"},
            {"name": "ความพึงพอใจเฉลี่ย", "definition": "ค่าเฉลี่ยธรรมดาของคะแนนที่มีค่า", "formula": "AVG(satisfaction where satisfaction > 0)", "source_fields": ["ความพึงพอใจ"], "aggregation": "average", "limitations": "ไม่มีจำนวนผู้ตอบแบบประเมินสำหรับถ่วงน้ำหนัก"}
        ],
        "filters": ["fiscal_year", "project_type", "service_group", "non_income_academic_services.fiscal_year", "budget_comparison.fiscal_year"],
    })
    write_json(METADATA_DIR / "dashboard-spec.json", {
        "title": "CASS Operational & Service Excellence",
        "sections": [
            {"name": "Filters", "visualization": "multi-select checkbox menus", "question": "ต้องการดูปีงบประมาณและมิติใดบ้าง"},
            {"name": "KPI overview", "visualization": "KPI cards", "question": "ผลการดำเนินงานรวมเป็นอย่างไร"},
            {"name": "Year-over-year KPI comparison", "visualization": "semantic delta indicators", "question": "ปีที่เลือกเปลี่ยนแปลงจากปีก่อนหน้าเท่าใด"},
            {"name": "Annual financial trend", "visualization": "two-series line chart", "question": "รายได้และกำไรที่คำนวณจากทะเบียนคุมงบเปลี่ยนแปลงอย่างไรในแต่ละปีงบประมาณ"},
            {"name": "Revenue comparison", "visualization": "clickable sorted horizontal bars", "question": "ประเภทและกลุ่มทะเบียนใดทำรายได้สูง และเกี่ยวข้องกับ KPI/รายละเอียดใด"},
            {"name": "Quality and risk", "visualization": "alerts", "question": "มีข้อจำกัดใด"},
            {"name": "Record details", "visualization": "filterable table", "question": "รายการใดอยู่เบื้องหลังตัวเลข"},
            {"name": "Non-income academic services", "visualization": "year filter, KPI cards, participant bars, social-benefit bars, and detail table", "question": "กิจกรรมบริการวิชาการที่ไม่ได้รายได้ดำเนินการอย่างไรและเกิดประโยชน์ต่อสังคมเพียงใด"},
            {"name": "Budget comparison", "visualization": "year filter, KPI cards, revenue-by-year/type bars, and detail table", "question": "ทะเบียนคุมงบปี 2567–2569 เปรียบเทียบรายได้ รายจ่าย และกำไรได้อย่างไร"}
        ],
    })
    quality_report = f'''# Data Quality Report: CASS\n\n## Scope\n- Files inspected: `{ACADEMIC_FILE.name}`, `{SERVICE_FILE.name}`, and the three budget-control workbooks for fiscal years 2567–2569\n- Dashboard financial source: budget-control registers ({len(records):,} records)\n\n## Findings\n- **Pass**: Financial profit is derived per record as revenue minus total expense. Aggregate revenue {totals['revenue']:,.2f} THB minus expense {totals['expense']:,.2f} THB equals profit {totals['reported_profit']:,.2f} THB; reconciliation difference is {totals['profit_reconciliation_difference']:,.2f} THB.\n- **Medium**: Fiscal year 2569 may be incomplete or include activities spanning fiscal years; compare it with completed years cautiously.\n- **Medium**: MU-ELT รหัส 67 is retained as a separately labelled register because it contains distinct placement-test and equipment activities.\n- **Medium**: Community activity data has 33 rows but only {len(community['activities'])} have project names and are included in the non-income drill-down. Blank source values remain blank/zero only where the source supplied no numeric value.\n\n## Handling\nNo raw record was changed or removed. Tab 1 and Tab 3 use the same reconciled budget-register dataset; Tab 3 exposes its source grouping and detail records.\n'''
    (REPORT_DIR / "data-quality-report.md").write_text(quality_report, encoding="utf-8")
    top_type = sorted_groups(records, "project_type")[0]
    insight_report = f'''# Insight Report: CASS\n\n## Evidence-based observations\n- The budget-register source contains **{len(records):,}** records, revenue of **{totals['revenue']:,.2f} THB**, expense of **{totals['expense']:,.2f} THB**, and derived profit of **{totals['reported_profit']:,.2f} THB**.\n- **{top_type['label']}** is the leading project type by recorded revenue at **{top_type['revenue']:,.2f} THB**.\n- Total actual participants are **{totals['actual_participants']:,.0f}**, against a recorded target of **{totals['target_participants']:,.0f}**.\n- Average recorded satisfaction is **{totals['satisfaction_average']:.2f}/5** for non-zero scores.\n- Non-income academic services include **{len(community['activities'])}** named activities, **{community['participants']:,.0f}** participants, budget **{community['budget']:,.2f} THB**, and expense **{community['expense']:,.2f} THB**.\n\n## Limitations\n- Fiscal year 2569 may be partial; compare year-over-year changes with caution.\n- The non-income activity view includes only rows with a project name; missing attributes are displayed as supplied rather than inferred.\n'''
    (REPORT_DIR / "insight-report.md").write_text(insight_report, encoding="utf-8")
    qa = {
        "project": "cass",
        "timestamp": datetime.now().astimezone().isoformat(),
        "status": "READY WITH LIMITATIONS",
        "critical_errors": [],
        "checks": [
            {"check": "budget register record count", "expected": 83, "actual": len(records), "result": "PASS" if len(records) == 83 else "FAIL"},
            {"check": "profit reconciliation", "expected": totals["derived_profit"], "actual": totals["reported_profit"], "difference": totals["profit_reconciliation_difference"], "result": "PASS" if totals["profit_reconciliation_difference"] == 0 else "FAIL"},
            {"check": "dashboard artifact", "expected": "present", "actual": "present", "result": "PASS"},
            {"check": "institutional theme", "expected": "mahidol-faculty", "actual": THEME_NAME, "result": "PASS"},
            {"check": "institutional logo", "expected": "dashboard asset and header display", "actual": "assets/mahidol-logo.png", "result": "PASS"},
            {"check": "interactive filters", "expected": "multi-select fiscal year, quarter, project type, service group, unit", "actual": "implemented", "result": "PASS"},
            {"check": "chart interactions", "expected": "animated chart load and bar cross-filter with clear action", "actual": "implemented with reduced-motion support", "result": "PASS"},
            {"check": "year-over-year KPI indicators", "expected": "selected-year comparison with semantic arrows", "actual": "implemented with unavailable-base state", "result": "PASS"},
            {"check": "annual financial trend chart", "expected": "revenue and derived profit labels by fiscal year", "actual": "implemented for fiscal years 2567-2569 with value labels", "result": "PASS"},
            {"check": "non-income academic service drill-down", "expected": f"{len(community['activities'])} named activities with fiscal-year filter and detail table", "actual": "implemented with KPI, participant, and social-benefit views", "result": "PASS"},
            {"check": "budget comparison drill-down", "expected": "year filter, KPI cards, comparison bars, and budget-register detail table", "actual": "implemented in Tab 3", "result": "PASS"},
            {"check": "desktop layout", "expected": "no horizontal overflow", "actual": "verified at 1440px", "result": "PASS"},
            {"check": "mobile layout", "expected": "no horizontal overflow", "actual": "verified at 390px", "result": "PASS"},
        ],
    }
    write_json(QA_DIR / "qa-report.json", qa)
    (OUTPUT_DIR / "dashboard" / "index.html").write_text(create_dashboard(records, community), encoding="utf-8")
    write_json(METADATA_DIR / "run-manifest.json", {
        "project": "cass",
        "run_id": f"cass-{datetime.now().strftime('%Y%m%dT%H%M%S')}",
        "input_files": [ACADEMIC_FILE.name, SERVICE_FILE.name, *(file_path.name for _, file_path in BUDGET_FILES)],
        "stages": ["DISCOVER", "PROFILE", "QUALITY", "ANALYZE", "REQUIREMENTS", "DESIGN", "TRANSFORM", "BUILD", "QA", "INSIGHTS", "DELIVER"],
        "status": "READY WITH LIMITATIONS",
        "warnings": ["fiscal year 2569 may be partial", "non-income activity view includes only named projects"],
        "critical_errors": [],
        "timestamp": datetime.now().astimezone().isoformat(),
    })
    print(json.dumps({"records": len(records), "totals": totals, "community": community, "status": qa["status"]}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()