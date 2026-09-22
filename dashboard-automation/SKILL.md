---
name: dashboard-automation
description: >
  Production-ready multi-project dashboard automation. Turns Excel/CSV files
  in a selected project's input folder into a validated dashboard. Use when
  the user asks to create, update, analyze, validate, or publish a dashboard.
---

# Dashboard Automation Skill

## Core goal

Turn project-specific Excel/CSV files into a usable dashboard while keeping
raw data unchanged and producing traceable analysis, requirements, QA, and
documentation.

## Project structure

Each dashboard is an independent project:

projects/<project-name>/
├── input/
├── output/
├── working/
├── config.json
└── dashboard-request.md   # optional

Never mix datasets between projects unless the user explicitly asks for it.

## Default workflow

DISCOVER
→ PROFILE
→ QUALITY
→ ANALYZE
→ REQUIREMENTS
→ DESIGN
→ TRANSFORM
→ BUILD
→ QA
→ INSIGHTS
→ DELIVER

## Rules

1. Never modify original files in `input/`.
2. Never fabricate data or KPI values.
3. Never silently delete records.
4. Every KPI needs a definition, formula, and source fields.
5. Every visualization must answer an analytical question.
6. Record important transformations.
7. Validate major KPI values independently.
8. If critical QA fails, status is NOT READY.
9. State assumptions and data limitations.
10. Do not publish sensitive raw data.

## Selecting a project

When the user says:
- "สร้าง dashboard ใหม่" → create a new project under `projects/`.
- "อัปเดต dashboard X" → use `projects/X/`.
- "วิเคราะห์ข้อมูล X" → use `projects/X/`.
- If a project name is ambiguous, inspect available projects and ask only if necessary.

Project names should be lowercase kebab-case, e.g.:
- sales
- student-analytics
- worktrack
- finance-2026

## Creating a new project

Create:

projects/<project-name>/
├── input/
├── output/
├── working/
└── config.json

Use this default config:

{
  "name": "<project-name>",
  "language": "th",
  "timezone": "Asia/Bangkok",
  "dashboard_type": "auto",
  "audience": "auto",
  "show_insights": true,
  "show_data_table": true,
  "responsive": true
}

Do not ask for information that can reasonably be inferred from the data.

## Optional dashboard-request.md

If present, treat explicit requirements as higher priority than AI inference.

Example:

# Dashboard Request

Audience: ผู้บริหาร
Purpose: ติดตามผลการดำเนินงาน
Important metrics:
- revenue
- orders
- profit

## Input discovery

Inspect all `.csv`, `.xlsx`, and `.xls` files in the selected project's `input/`.

Identify:
- row/column counts
- sheets
- data types
- date fields
- measures
- dimensions
- IDs
- relationships
- date coverage

Create:
`output/metadata/dataset-profile.json`

## Data quality

Check:
- missing values
- duplicates
- invalid dates
- inconsistent categories
- suspicious numeric values
- unexpected ranges
- coverage gaps

Create:
`output/reports/data-quality-report.md`

Do not silently repair data.

## Business analysis

Identify supported:
- KPIs
- trends
- rankings
- comparisons
- segments
- anomalies

For every KPI record:
- name
- definition
- formula
- source fields
- aggregation
- limitations

## Dashboard requirements

Determine:
- audience
- purpose
- primary questions
- KPIs
- filters
- charts
- alerts
- detail views

Create:
`output/metadata/dashboard-requirements.json`

## Dashboard design

Preferred visualization:
- KPI → KPI card
- trend → line
- category comparison → bar
- ranking → sorted horizontal bar
- exact records → table
- relationship → scatter
- 2D pattern → heatmap

Avoid decorative and misleading charts.

Create:
`output/metadata/dashboard-spec.json`

## Data transformation

Create analytical data under:
`output/data/`

Never overwrite raw input.

Document important transformations.

## Dashboard build

If an existing dashboard implementation exists in the project, update it rather
than replacing it unnecessarily.

If none exists, build a simple responsive local web dashboard that can be opened
from the generated project.

Recommended sections:
1. Header + filters
2. KPI cards
3. Main trend
4. Comparisons
5. Rankings
6. Insights/alerts
7. Detail table where useful

## QA

Validate:
- major KPI values
- row counts
- nulls
- duplicates
- date ranges
- categories
- charts
- labels
- layout
- responsiveness
- filters and sorting

Create:
`output/qa/qa-report.json`

Critical numerical mismatch => NOT READY.

## Insights

Create:
`output/reports/insight-report.md`

Every major insight must include evidence and limitations.

## Run manifest

Create:
`output/metadata/run-manifest.json`

Include:
- project
- run_id
- input files
- stages
- status
- warnings
- critical errors
- timestamp

## Completion response

After running, report:
- project
- files analyzed
- dashboard location
- QA status
- important warnings
- how to open the dashboard

Do not claim production-ready if critical QA failed.

## Commands understood

The agent should understand these natural-language commands:

### New
"สร้าง dashboard ใหม่จาก sales.xlsx"

### Run
"รัน dashboard ของ sales"

### Update
"อัปเดต dashboard sales จากไฟล์ใหม่"

### Analyze
"วิเคราะห์ข้อมูล student"

### QA
"ตรวจสอบ dashboard sales"

### Publish preparation
"เตรียม dashboard sales สำหรับ publish"

### List
"แสดง dashboard ทั้งหมด"

When possible, execute the requested command without asking redundant questions.
