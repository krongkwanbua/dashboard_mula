# CASS Dashboard Playbook

เอกสารนี้กำหนดวิธีอัปเดต CASS Operational & Service Excellence เพื่อให้การสร้าง dashboard รอบถัดไปใช้ความหมายข้อมูลและกติกาเดิมอย่างสม่ำเสมอ

## ขอบเขต Dashboard

- กลุ่มผู้ใช้: ผู้บริหารและผู้ปฏิบัติงาน CAAS
- ภาษา: ไทย
- Theme: Mahidol Faculty (`mahidol-faculty`)
- Output หลัก: `output/dashboard/index.html`
- ห้ามแก้ไฟล์ต้นฉบับใต้ `input/`

## แท็บและแหล่งข้อมูล

### Tab 1: บริการมีรายได้

ใช้ข้อมูลผสมตามปีงบประมาณ:

| ปีงบประมาณ | แหล่งข้อมูล | การใช้งาน |
|---|---|---|
| 2563-2566 | `Operational & Service Excellence by CAAS.xlsx` / `Service_Excellence_by_CAAS` | คงข้อมูลเดิม |
| 2567-2569 | ไฟล์ทะเบียนคุมงบประมาณ | ใช้แทนข้อมูลเดิมของปีนั้น |

ไฟล์ทะเบียนคุมงบ:

- `A ทะเบียนคุมงบประมาณ 2567.xlsx`
- `A ทะเบียนคุมงบประมาณ 2568.xlsx`
- `A ตารางคุมงบประมาณ 2569.xlsx`

กติกาการเงินของปี 2567-2569:

- รายได้ = `รายรับ (ยังไม่หักค่าใช้จ่าย)`
- รายจ่าย = `รวม`
- กำไร = `รายได้ - รายจ่าย` ต่อรายการ
- ต้องตรวจให้ยอดกำไรรวม reconcile กับ `รายได้รวม - รายจ่ายรวม` เป็นศูนย์ก่อนเผยแพร่

สำหรับปี 2563-2566 ให้คง `กำไรสุทธิ` จาก Operational source และต้องระบุใน QA/insight ว่าไม่ได้ถูกแทนด้วยทะเบียนคุมงบ

### Tab 2: บริการวิชาการแบบไม่ได้รายได้

แหล่งข้อมูล: `Academic_services_without_income.xlsx` / sheet `data`

- ใช้เฉพาะแถวที่มี `ชื่อโครงการ`
- ไม่อนุมานค่าที่ว่าง
- แสดง KPI จำนวนกิจกรรม, ผู้เข้าร่วม, งบประมาณ, ค่าใช้จ่าย และความพึงพอใจเมื่อมีข้อมูล
- ใช้ตัวกรองปีงบประมาณแยกจาก Tab 1

### Tab 3: เปรียบเทียบงบประมาณ

ใช้เฉพาะทะเบียนคุมงบปี 2567-2569 เพื่อดูรายละเอียดที่ reconcile แล้ว

- `ทะเบียนคุมงบ MU-ELT` และ `MU-ELT รหัส 67` เป็นประเภทเดียวกับ `การจัดสอบ MU-ELT`
- เก็บชื่อชีตต้นทางไว้เป็น `source_sheet` หรือแสดงเป็น `แหล่งทะเบียน` ในตาราง เพื่อการตรวจสอบย้อนกลับ
- `In-House` เป็นอีกประเภทบริการหนึ่ง
- ต้องไม่แสดง `ทะเบียนคุมงบ MU-ELT` หรือ `MU-ELT รหัส 67` เป็นประเภทโครงการแยกในกราฟ

## Analytical Outputs

ทุกครั้งที่สร้างใหม่ ให้ regenerate อย่างน้อย:

- `output/data/service-excellence-records.csv`: ชุดผสมสำหรับ Tab 1
- `output/data/non-income-academic-services.csv`: ชุด Tab 2
- `output/data/budget-control-records.csv`: ชุด Tab 3
- `output/metadata/dataset-profile.json`
- `output/metadata/dashboard-requirements.json`
- `output/metadata/dashboard-spec.json`
- `output/reports/data-quality-report.md`
- `output/reports/insight-report.md`
- `output/qa/qa-report.json`

## ขั้นตอนอัปเดต

1. วางไฟล์ใหม่ใน `projects/CASS/input/` และรักษาไฟล์ต้นฉบับไว้
2. อ่าน `dashboard-automation/SKILL.md` และเอกสารนี้ก่อนแก้ generator
3. Profile ทุก workbook/sheet: จำนวนแถว, หัวคอลัมน์, ปี, null, duplicate และประเภทโครงการ
4. ปรับ `working/build_dashboard.py` โดยรักษา data provenance (`source_file`, `source_sheet`, `source_row`)
5. Run generator:

```bash
python3 projects/CASS/working/build_dashboard.py
```

6. ตรวจอย่างน้อย:

```bash
python3 -m py_compile projects/CASS/working/build_dashboard.py
python3 -m json.tool projects/CASS/output/qa/qa-report.json
```

7. ตรวจ reconciliation ของปีที่ใช้ทะเบียนคุมงบ:

```text
SUM(กำไร) = SUM(รายได้) - SUM(รายจ่าย)
```

8. ตรวจ UI: ทั้งสามแท็บ, ตัวกรอง, cross-filter, mobile ไม่มี horizontal overflow, และค่า KPI ตรงกับ CSV
9. หากมี critical numerical mismatch ให้ตั้ง QA เป็น `NOT READY` และไม่ publish

## Prompt รอบถัดไป

```text
ใช้ dashboard-automation skill และอ่าน projects/CASS/dashboard-playbook.md ก่อนเริ่ม
อัปเดต dashboard CASS จากไฟล์ใหม่ใน projects/CASS/input
รักษาข้อมูล Operational เดิมของปีที่ไม่มีทะเบียนคุมงบ
ใช้ทะเบียนคุมงบแทนเฉพาะปีที่มีข้อมูลและทำ reconciliation รายได้-รายจ่าย-กำไร
normalize ทะเบียนคุมงบ MU-ELT และ MU-ELT รหัส 67 เป็นประเภท การจัดสอบ MU-ELT
regenerate output, ทำ QA ทั้งข้อมูลและ UI แล้วสรุปข้อจำกัด
```
