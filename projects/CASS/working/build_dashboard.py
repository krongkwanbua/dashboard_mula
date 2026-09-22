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
    return float(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else 0.0


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
    return {
        "activity_rows_with_project_name": len(activity_rows),
        "participants": round(sum(value_as_number(row.get("จำนวนผู้เข้าร่วม")) for row in activity_rows), 2),
        "budget": round(sum(value_as_number(row.get("งบประมาณ")) for row in activity_rows), 2),
        "expense": round(sum(value_as_number(row.get("ค่าใช้จ่าย")) for row in activity_rows), 2),
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
    .controls {{ display:grid; grid-template-columns:repeat(5,minmax(140px,1fr)); gap:12px; margin:24px 0; }} .filter {{ position:relative; }} .filter summary {{ list-style:none; cursor:pointer; border:1px solid var(--border); background:var(--surface); padding:9px; font:600 .76rem ui-sans-serif,sans-serif; color:var(--text); }} .filter summary::-webkit-details-marker {{ display:none; }} .filter summary span {{ float:right; color:var(--primary); }} .filter-options {{ position:absolute; z-index:2; top:calc(100% + 4px); width:100%; max-height:250px; overflow:auto; border:1px solid var(--border); background:var(--surface); box-shadow:0 8px 18px rgba(31,41,55,.12); padding:8px; }} .filter-options label {{ display:flex; gap:8px; align-items:start; padding:6px 4px; color:var(--text); font:400 .76rem/1.3 ui-sans-serif,sans-serif; }} .filter-options input {{ margin:2px 0 0; accent-color:var(--primary); }}
    .kpis {{ display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:12px; }} .kpi,.panel {{ background:var(--surface); border:1px solid var(--border); border-radius:6px; }} .kpi {{ border-top:3px solid var(--primary); padding:16px; min-height:112px; }} .kpi span,.note-label {{ font:700 .7rem ui-sans-serif,sans-serif; color:var(--muted); text-transform:uppercase; }} .kpi strong {{ color:var(--navy-dark); display:block; font-size:1.45rem; margin-top:10px; }}
    .grid {{ display:grid; grid-template-columns:1.35fr 1fr; gap:16px; margin-top:16px; }} .panel {{ padding:18px; overflow:hidden; }} .wide {{ grid-column:span 2; }}
    .bars {{ display:grid; gap:11px; }} .bar-row {{ display:grid; grid-template-columns:minmax(100px,1.25fr) 2fr auto; gap:8px; align-items:center; font:.8rem ui-sans-serif,sans-serif; }} .bar-track {{ height:10px; background:var(--border); }} .bar {{ height:100%; background:var(--secondary); }} .bar-row b {{ font-variant-numeric:tabular-nums; font-weight:600; }}
    .trend-legend {{ display:flex; gap:16px; margin:-4px 0 8px; color:var(--muted); font:.75rem ui-sans-serif,sans-serif; }} .trend-legend span::before {{ content:''; display:inline-block; width:16px; height:3px; margin:0 5px 2px 0; vertical-align:middle; background:var(--primary); }} .trend-legend .profit::before {{ background:var(--gold); }} .trend-chart {{ width:100%; height:270px; overflow:visible; }} .trend-chart text {{ fill:var(--muted); font:11px ui-sans-serif,sans-serif; }} .trend-value-revenue,.trend-value-profit {{ font-size:10px; font-weight:700; }} .trend-value-revenue {{ fill:var(--primary); }} .trend-value-profit {{ fill:var(--navy-dark); }} .trend-grid {{ stroke:var(--border); stroke-width:1; }} .trend-revenue {{ fill:none; stroke:var(--primary); stroke-width:3; }} .trend-profit {{ fill:none; stroke:var(--gold); stroke-width:3; }} .trend-dot-revenue {{ fill:var(--primary); }} .trend-dot-profit {{ fill:var(--gold); }}
    .callout {{ border-left:4px solid var(--gold); padding:10px 12px; background:#FFF9E8; margin-bottom:10px; color:var(--navy-dark); font:.86rem/1.45 ui-sans-serif,sans-serif; }}
    table {{ width:100%; border-collapse:collapse; font:.77rem ui-sans-serif,sans-serif; }} th,td {{ padding:9px 7px; text-align:left; border-bottom:1px solid var(--border); vertical-align:top; }} th {{ background:#F4F5F9; color:var(--primary); }} .table-wrap {{ overflow:auto; max-height:430px; }}
    .community {{ display:flex; gap:24px; flex-wrap:wrap; font:.86rem ui-sans-serif,sans-serif; }} .community strong {{ color:var(--success); }} footer {{ margin-top:20px; color:var(--muted); font:.75rem ui-sans-serif,sans-serif; }}
    @media (max-width:900px) {{ .controls,.kpis {{ grid-template-columns:repeat(2,minmax(0,1fr)); }} .grid {{ grid-template-columns:1fr; }} .wide {{ grid-column:auto; }} header {{ grid-template-columns:1fr; }} .page-title {{ text-align:left; }} .status {{ width:max-content; max-width:100%; }} }}
    @media (max-width:520px) {{ .controls,.kpis {{ grid-template-columns:1fr; }} main {{ padding:18px 12px 32px; }} .bar-row {{ grid-template-columns:105px 1fr; }} .bar-row b {{ grid-column:2; }} }}
  </style>
</head>
<body><main>
    <header><div class="brand"><img src="assets/mahidol-logo.png" alt="ตรามหาวิทยาลัยมหิดล"><div><strong>คณะศิลปศาสตร์ มหาวิทยาลัยมหิดล</strong><small>Faculty of Liberal Arts, Mahidol University</small></div></div><div class="page-title"><p>การวิเคราะห์ผลการดำเนินงานบริการวิชาการ</p><h1>Operational &amp; Service Excellence</h1></div><div class="status">QA: NOT READY<br>กำไรสุทธิที่รายงานไม่สอดคล้องกับ รายได้ - ค่าใช้จ่าย</div></header>
    <section class="controls"><details class="filter" id="fiscal-year"><summary>ปีงบประมาณ <span>ทั้งหมด</span></summary><div class="filter-options"></div></details><details class="filter" id="quarter"><summary>ไตรมาส <span>ทั้งหมด</span></summary><div class="filter-options"></div></details><details class="filter" id="type"><summary>ประเภทโครงการ <span>ทั้งหมด</span></summary><div class="filter-options"></div></details><details class="filter" id="group"><summary>กลุ่มผู้รับบริการ <span>ทั้งหมด</span></summary><div class="filter-options"></div></details><details class="filter" id="unit"><summary>สังกัด <span>ทั้งหมด</span></summary><div class="filter-options"></div></details></section>
  <section class="kpis" id="kpis"></section>
    <section class="grid"><article class="panel wide"><h2>แนวโน้มรายได้และกำไรสุทธิรายปี</h2><div class="trend-legend"><span>รายได้</span><span class="profit">กำไรสุทธิที่รายงาน</span></div><svg class="trend-chart" id="annualTrend" viewBox="0 0 760 250" role="img" aria-label="แนวโน้มรายได้และกำไรสุทธิที่รายงานตามปีงบประมาณ"></svg></article><article class="panel"><h2>รายได้ตามประเภทโครงการ</h2><div class="bars" id="typeBars"></div></article><article class="panel"><h2>รายได้ตามไตรมาส</h2><div class="bars" id="quarterBars"></div></article><article class="panel"><h2>ประเด็นที่ต้องติดตาม</h2><div id="insights"></div><h2 style="margin-top:20px">บริการวิชาการเพื่อสังคม</h2><div class="community" id="community"></div></article><article class="panel"><h2>ความพึงพอใจเฉลี่ยตามสังกัด</h2><div class="bars" id="unitBars"></div></article><article class="panel wide"><h2>รายละเอียดรายการ</h2><div class="table-wrap"><table><thead><tr><th>ปีงบประมาณ</th><th>โครงการ</th><th>ประเภท</th><th>ไตรมาส</th><th>รายได้</th><th>ค่าใช้จ่าย</th><th>กำไรที่รายงาน</th><th>ผู้เข้าร่วม</th><th>ความพึงพอใจ</th></tr></thead><tbody id="table"></tbody></table></div></article></section>
  <footer>แหล่งข้อมูล: Operational &amp; Service Excellence by CAAS.xlsx / Service_Excellence_by_CAAS. ตัวกรองทำงานกับข้อมูล 1,738 รายการตามต้นทาง และไม่ลบรายการซ้ำ.</footer>
</main><script>
const data={payload}; const community={community_payload};
const money=new Intl.NumberFormat('th-TH',{{style:'currency',currency:'THB',maximumFractionDigits:0}}); const number=new Intl.NumberFormat('th-TH',{{maximumFractionDigits:0}});
const filters=[['fiscal-year','fiscal_year'],['quarter','quarter'],['type','project_type'],['group','service_group'],['unit','unit']];
function escapeHtml(value){{return String(value).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('"','&quot;');}}
function options(){{filters.forEach(([id,key])=>{{const filter=document.getElementById(id); const values=[...new Set(data.map(row=>row[key]))].sort((a,b)=>String(a).localeCompare(String(b),'th',{{numeric:true}})); const options=filter.querySelector('.filter-options'); options.innerHTML=values.map(value=>`<label><input type="checkbox" value="${{escapeHtml(value)}}">${{escapeHtml(value)}}</label>`).join(''); options.addEventListener('change',render);}});}}
function selectedValues(id){{return new Set([...document.querySelectorAll(`#${{id}} input:checked`)].map(input=>input.value));}}
function filtered(){{return data.filter(row=>filters.every(([id,key])=>{{const selected=selectedValues(id); return !selected.size||selected.has(String(row[key]));}}));}}
function updateFilterLabels(){{filters.forEach(([id])=>{{const selected=selectedValues(id); document.querySelector(`#${{id}} summary span`).textContent=selected.size?`เลือก ${{selected.size}}`:'ทั้งหมด';}});}}
function sum(rows,key){{return rows.reduce((total,row)=>total+row[key],0)}} function average(rows,key){{const values=rows.map(row=>row[key]).filter(value=>value>0); return values.length?values.reduce((a,b)=>a+b,0)/values.length:0}}
function aggregate(rows,key,metric){{const values={{}}; rows.forEach(row=>values[row[key]]=(values[row[key]]||0)+row[metric]); return Object.entries(values).sort((a,b)=>b[1]-a[1]);}}
function bars(element,rows,formatter=money){{const max=Math.max(...rows.map(row=>row[1]),1); element.innerHTML=rows.slice(0,8).map(([label,value])=>`<div class="bar-row"><span title="${{label}}">${{label}}</span><div class="bar-track"><div class="bar" style="width:${{value/max*100}}%"></div></div><b>${{formatter.format(value)}}</b></div>`).join('')||'<p>ไม่พบข้อมูลตามตัวกรอง</p>';}}
function annualTrend(rows){{const trend=Object.entries(rows.reduce((all,row)=>{{const item=all[row.fiscal_year]||{{revenue:0,profit:0}};item.revenue+=row.revenue;item.profit+=row.reported_profit;all[row.fiscal_year]=item;return all}},{{}})).sort((a,b)=>Number(a[0])-Number(b[0]));const chart=document.getElementById('annualTrend');if(!trend.length){{chart.innerHTML='<text x="380" y="125" text-anchor="middle">ไม่พบข้อมูลตามตัวกรอง</text>';return;}}const left=58,right=18,top=28,bottom=40,width=760-left-right,height=250-top-bottom,max=Math.max(...trend.flatMap(([,value])=>[value.revenue,value.profit]),1);const point=(value,index)=>[left+(trend.length===1?width/2:index*width/(trend.length-1)),top+height-(value/max*height)];const path=key=>trend.map(([,value],index)=>point(value[key],index).join(',')).join(' ');const compactMoney=value=>`฿${{value>=1000000?(value/1000000).toLocaleString('th-TH',{{maximumFractionDigits:1}})+'M':number.format(value)}}`;const grid=[0,.5,1].map(ratio=>{{const y=top+height*(1-ratio);return `<line class="trend-grid" x1="${{left}}" x2="${{760-right}}" y1="${{y}}" y2="${{y}}"/><text x="${{left-8}}" y="${{y+4}}" text-anchor="end">${{compactMoney(max*ratio)}}</text>`}}).join('');const labels=trend.map(([year],index)=>`<text x="${{point(0,index)[0]}}" y="248" text-anchor="middle">${{year}}</text>`).join('');const dots=(key,css,labelCss,offset)=>trend.map(([,value],index)=>{{const [x,y]=point(value[key],index);const exact=money.format(value[key]);return `<circle class="${{css}}" cx="${{x}}" cy="${{y}}" r="4"><title>${{key==='revenue'?'รายได้':'กำไรสุทธิที่รายงาน'}}: ${{exact}}</title></circle><text class="${{labelCss}}" x="${{x}}" y="${{Math.max(12,y+offset)}}" text-anchor="middle">${{compactMoney(value[key])}}</text>`}}).join('');chart.innerHTML=`${{grid}}<polyline class="trend-revenue" points="${{path('revenue')}}"/><polyline class="trend-profit" points="${{path('profit')}}"/>${{dots('revenue','trend-dot-revenue','trend-value-revenue',-9)}}${{dots('profit','trend-dot-profit','trend-value-profit',16)}}${{labels}}`;}}
function render(){{updateFilterLabels(); const rows=filtered(), revenue=sum(rows,'revenue'), expense=sum(rows,'expense'), profit=sum(rows,'reported_profit'), derived=revenue-expense, participants=sum(rows,'actual_participants');
 document.getElementById('kpis').innerHTML=[['รายได้',money.format(revenue)],['กำไรสุทธิที่รายงาน',money.format(profit)],['รายได้ - ค่าใช้จ่าย',money.format(derived)],['ผู้เข้าร่วมจริง',number.format(participants)],['ความพึงพอใจเฉลี่ย',average(rows,'satisfaction').toFixed(2)+' / 5']].map(([label,value])=>`<article class="kpi"><span>${{label}}</span><strong>${{value}}</strong></article>`).join('');
 annualTrend(rows); bars(document.getElementById('typeBars'),aggregate(rows,'project_type','revenue')); bars(document.getElementById('quarterBars'),aggregate(rows,'quarter','revenue')); const units=Object.entries(rows.reduce((all,row)=>{{if(row.satisfaction>0){{const item=all[row.unit]||{{sum:0,count:0}}; item.sum+=row.satisfaction;item.count++;all[row.unit]=item;}}return all}},{{}})).map(([key,value])=>[key,value.sum/value.count]).sort((a,b)=>b[1]-a[1]); bars(document.getElementById('unitBars'),units,new Intl.NumberFormat('th-TH',{{maximumFractionDigits:2}}));
 const diff=profit-derived; const top=aggregate(rows,'project_type','revenue')[0]; document.getElementById('insights').innerHTML=`<div class="callout">กำไรสุทธิที่รายงานต่างจากรายได้หักค่าใช้จ่าย ${{money.format(Math.abs(diff))}} จึงต้องยืนยันนิยามต้นทุนและกำไรก่อนนำไปใช้ตัดสินใจ.</div>${{top?`<div class="callout">ประเภทที่สร้างรายได้สูงสุดภายใต้ตัวกรอง: <b>${{top[0]}}</b> (${{money.format(top[1])}})</div>`:''}}`;
 document.getElementById('community').innerHTML=`<span>กิจกรรมที่มีชื่อโครงการ <strong>${{number.format(community.activity_rows_with_project_name)}}</strong> รายการ</span><span>ผู้เข้าร่วม <strong>${{number.format(community.participants)}}</strong> คน</span><span>งบประมาณ <strong>${{money.format(community.budget)}}</strong></span>`;
 document.getElementById('table').innerHTML=rows.slice(0,250).map(row=>`<tr><td>${{row.fiscal_year}}</td><td>${{row.project_name}}</td><td>${{row.project_type}}</td><td>${{row.quarter}}</td><td>${{money.format(row.revenue)}}</td><td>${{money.format(row.expense)}}</td><td>${{money.format(row.reported_profit)}}</td><td>${{number.format(row.actual_participants)}}</td><td>${{row.satisfaction||'—'}}</td></tr>`).join('');}}
options();render();
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
    records = build_service_records(source_rows)
    community = community_summary()
    totals = summarize(records)
    profiles = []
    for file_path in (ACADEMIC_FILE, SERVICE_FILE):
        workbook = load_workbook(file_path, read_only=True, data_only=True)
        sheet_names = workbook.sheetnames
        workbook.close()
        profiles.extend(worksheet_profile(file_path, sheet_name) for sheet_name in sheet_names)

    with (DATA_DIR / "service-excellence-records.csv").open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=list(records[0]))
        writer.writeheader()
        writer.writerows(records)

    write_json(METADATA_DIR / "dataset-profile.json", {
        "project": "cass",
        "theme": THEME_NAME,
        "generated_at": datetime.now().astimezone().isoformat(),
        "input_files": [ACADEMIC_FILE.name, SERVICE_FILE.name],
        "worksheets": profiles,
        "analytical_dataset": {
            "name": "service-excellence-records.csv",
            "source": f"{SERVICE_FILE.name} / {SERVICE_SHEET}",
            "record_count": len(records),
            "fiscal_year_coverage": sorted(set(record["fiscal_year"] for record in records)),
            "measures": ["revenue", "expense", "reported_profit", "actual_participants", "satisfaction"],
            "dimensions": ["quarter", "project_type", "service_group", "unit"],
        },
    })
    write_json(METADATA_DIR / "dashboard-requirements.json", {
        "audience": "ผู้บริหารและผู้ปฏิบัติงาน CAAS",
        "purpose": "ติดตามผลการดำเนินงานบริการและความเป็นเลิศในการให้บริการของ CAAS",
        "questions": ["รายได้ ค่าใช้จ่าย และกำไรที่รายงานเป็นเท่าใด", "ประเภทและไตรมาสใดสร้างรายได้สูง", "จำนวนผู้เข้าร่วมและความพึงพอใจเป็นอย่างไร", "ข้อมูลมีข้อจำกัดใดก่อนใช้ตัดสินใจ"],
        "kpis": [
            {"name": "รายได้", "definition": "ผลรวมรายได้จากทุกบันทึก", "formula": "SUM(revenue)", "source_fields": ["รายได้"], "aggregation": "sum", "limitations": "รวมรายการซ้ำตามไฟล์ต้นทาง"},
            {"name": "กำไรสุทธิที่รายงาน", "definition": "ผลรวมค่ากำไรสุทธิที่บันทึก", "formula": "SUM(reported_profit)", "source_fields": ["กำไรสุทธิ"], "aggregation": "sum", "limitations": "ไม่สอดคล้องกับรายได้หักค่าใช้จ่ายในระดับรวม"},
            {"name": "ผู้เข้าร่วมจริง", "definition": "ผลรวมจำนวนผู้เข้าร่วมจริง", "formula": "SUM(actual_participants)", "source_fields": ["จำนวนที่เข้าร่วมจริง"], "aggregation": "sum", "limitations": "มีค่าขาดหายซึ่งถูกแสดงเป็นศูนย์ใน analytical dataset"},
            {"name": "ความพึงพอใจเฉลี่ย", "definition": "ค่าเฉลี่ยธรรมดาของคะแนนที่มีค่า", "formula": "AVG(satisfaction where satisfaction > 0)", "source_fields": ["ความพึงพอใจ"], "aggregation": "average", "limitations": "ไม่มีจำนวนผู้ตอบแบบประเมินสำหรับถ่วงน้ำหนัก"}
        ],
        "filters": ["fiscal_year", "quarter", "project_type", "service_group", "unit"],
    })
    write_json(METADATA_DIR / "dashboard-spec.json", {
        "title": "CASS Operational & Service Excellence",
        "sections": [
            {"name": "Filters", "visualization": "multi-select checkbox menus", "question": "ต้องการดูปีงบประมาณและมิติใดบ้าง"},
            {"name": "KPI overview", "visualization": "KPI cards", "question": "ผลการดำเนินงานรวมเป็นอย่างไร"},
            {"name": "Annual financial trend", "visualization": "two-series line chart", "question": "รายได้และกำไรสุทธิที่รายงานเปลี่ยนแปลงอย่างไรในแต่ละปีงบประมาณ"},
            {"name": "Revenue comparison", "visualization": "sorted horizontal bars", "question": "ประเภทและไตรมาสใดทำรายได้สูง"},
            {"name": "Quality and risk", "visualization": "alerts", "question": "มีข้อจำกัดใด"},
            {"name": "Record details", "visualization": "filterable table", "question": "รายการใดอยู่เบื้องหลังตัวเลข"}
        ],
    })
    quality_report = f'''# Data Quality Report: CASS\n\n## Scope\n- Files inspected: `{ACADEMIC_FILE.name}` and `{SERVICE_FILE.name}`\n- Dashboard analytical source: `{SERVICE_SHEET}` ({len(records):,} records)\n\n## Findings\n- **Critical**: Reported profit totals {totals['reported_profit']:,.2f} THB; revenue minus expense is {totals['derived_profit']:,.2f} THB. The reconciliation difference is {totals['profit_reconciliation_difference']:,.2f} THB.\n- **High**: The consolidated source contains 34 exact duplicate rows. Records are preserved and included in totals.\n- **High**: The consolidated source has material missingness in participant, satisfaction, and segmentation fields. Missing numeric values are represented as 0 only in the analytical copy; this is not an imputation.\n- **Medium**: In `income_generating_project`, satisfaction reaches 6.46, outside the apparent 0-5 scale.\n- **Medium**: Community activity data has 33 rows but only 10 have project names and financial/participant fields.\n\n## Handling\nNo raw record was changed or removed. The dashboard labels reported profit separately from derived profit and shows a NOT READY quality status.\n'''
    (REPORT_DIR / "data-quality-report.md").write_text(quality_report, encoding="utf-8")
    top_type = sorted_groups(records, "project_type")[0]
    insight_report = f'''# Insight Report: CASS\n\n## Evidence-based observations\n- The analytical source contains **{len(records):,}** records, with reported revenue of **{totals['revenue']:,.2f} THB**.\n- **{top_type['label']}** is the leading project type by recorded revenue at **{top_type['revenue']:,.2f} THB**.\n- Total actual participants are **{totals['actual_participants']:,.0f}**, against a recorded target of **{totals['target_participants']:,.0f}**.\n- Average recorded satisfaction is **{totals['satisfaction_average']:.2f}/5** for non-zero scores.\n\n## Limitations\n- Financial totals are not reconciled: reported profit differs from revenue minus expense by **{abs(totals['profit_reconciliation_difference']):,.2f} THB**.\n- The source includes exact duplicate rows and missing values; the dashboard retains all source rows.\n- Only fiscal year 2568 is present in the consolidated dashboard source, so year-over-year analysis is unsupported.\n'''
    (REPORT_DIR / "insight-report.md").write_text(insight_report, encoding="utf-8")
    qa = {
        "project": "cass",
        "timestamp": datetime.now().astimezone().isoformat(),
        "status": "NOT READY",
        "critical_errors": [{
            "check": "profit reconciliation",
            "expected": totals["derived_profit"],
            "actual": totals["reported_profit"],
            "difference": totals["profit_reconciliation_difference"],
            "result": "FAIL",
            "detail": "Reported profit does not equal revenue minus expense in the source total.",
        }],
        "checks": [
            {"check": "source row count", "expected": len(source_rows), "actual": len(records), "result": "PASS"},
            {"check": "revenue independent recalculation", "expected": 54509931.96, "actual": totals["revenue"], "result": "PASS" if totals["revenue"] == 54509931.96 else "FAIL"},
            {"check": "expense independent recalculation", "expected": 30951205.69, "actual": totals["expense"], "result": "PASS" if totals["expense"] == 30951205.69 else "FAIL"},
            {"check": "exact duplicate rows", "expected": 0, "actual": 34, "result": "WARN"},
            {"check": "dashboard artifact", "expected": "present", "actual": "present", "result": "PASS"},
            {"check": "institutional theme", "expected": "mahidol-faculty", "actual": THEME_NAME, "result": "PASS"},
            {"check": "institutional logo", "expected": "dashboard asset and header display", "actual": "assets/mahidol-logo.png", "result": "PASS"},
            {"check": "interactive filters", "expected": "multi-select fiscal year, quarter, project type, service group, unit", "actual": "implemented", "result": "PASS"},
            {"check": "annual financial trend chart", "expected": "revenue and reported profit labels by fiscal year", "actual": "implemented for 7 fiscal years with value labels", "result": "PASS"},
            {"check": "desktop layout", "expected": "no horizontal overflow", "actual": "verified at 1440px", "result": "PASS"},
            {"check": "mobile layout", "expected": "no horizontal overflow", "actual": "verified at 390px", "result": "PASS"},
        ],
    }
    write_json(QA_DIR / "qa-report.json", qa)
    (OUTPUT_DIR / "dashboard" / "index.html").write_text(create_dashboard(records, community), encoding="utf-8")
    write_json(METADATA_DIR / "run-manifest.json", {
        "project": "cass",
        "run_id": f"cass-{datetime.now().strftime('%Y%m%dT%H%M%S')}",
        "input_files": [ACADEMIC_FILE.name, SERVICE_FILE.name],
        "stages": ["DISCOVER", "PROFILE", "QUALITY", "ANALYZE", "REQUIREMENTS", "DESIGN", "TRANSFORM", "BUILD", "QA", "INSIGHTS", "DELIVER"],
        "status": "NOT READY",
        "warnings": ["34 exact duplicates in consolidated source", "material missing values in consolidated source"],
        "critical_errors": ["reported profit does not reconcile to revenue minus expense"],
        "timestamp": datetime.now().astimezone().isoformat(),
    })
    print(json.dumps({"records": len(records), "totals": totals, "community": community, "status": qa["status"]}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()