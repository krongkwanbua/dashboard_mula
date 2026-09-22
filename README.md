# Dashboard Factory

Multi-project AI dashboard automation for VS Code + GitHub Copilot Agent.

Workflow:
1. Create or select a project.
2. Put Excel/CSV files into `projects/<project>/input/`.
3. Put optional `config.json` and `dashboard-request.md` in the project folder.
4. Ask Copilot Agent to run the dashboard automation.
5. Final dashboard is written to `projects/<project>/output/dashboard/`.

Do not put sensitive raw data into a public repository.

สร้าง dashboard ใหม่จาก projects/sales/input/sales.xlsx
ใช้ dashboard-automation skill
วิเคราะห์ข้อมูล ออกแบบ dashboard สร้าง dashboard และทำ QA ให้ครบ

อัปเดต dashboard sales จากข้อมูลใหม่ใน projects/sales/input
วิเคราะห์ข้อมูลใหม่ สร้าง dashboard ใหม่ และตรวจสอบ KPI

คำสั่งที่ใช้บ่อย
ดู Dashboard ทั้งหมด
แสดง dashboard ทั้งหมดใน projects
สร้างใหม่
สร้าง dashboard ใหม่จากไฟล์นี้
วิเคราะห์
วิเคราะห์ข้อมูลของ dashboard sales
อัปเดต
อัปเดต dashboard sales จากข้อมูลใหม่
ตรวจสอบ
ตรวจสอบ dashboard sales และทำ QA
เตรียม Publish
เตรียม dashboard sales สำหรับ publish


อัปเดต dashboard CASS จากข้อมูลใหม่ใน projects/CASS/input
วิเคราะห์ข้อมูลใหม่ สร้าง dashboard ใหม่ และตรวจสอบ KPI

## อัปเดต CASS รอบถัดไป

1. วางไฟล์ Excel หรือ CSV ใหม่ใน `projects/CASS/input/` โดยไม่แก้ไฟล์เก่า
	หากต้องการให้ข้อมูลใหม่แทนชุดเดิม ให้ย้ายไฟล์เดิมออกจากโฟลเดอร์ก่อน
	สำหรับ Tab 1 และ Tab 3 ให้ใช้ไฟล์ทะเบียนคุมงบประมาณของแต่ละปี; ระบบคำนวณกำไรเป็น `รายรับ - ค่าใช้จ่ายรวม` จากแต่ละรายการ
2. สั่ง Copilot Agent ด้วยข้อความนี้:

```text
ใช้ dashboard-automation skill
อัปเดต dashboard CASS จากข้อมูลใหม่ใน projects/CASS/input
วิเคราะห์ข้อมูลใหม่ สร้าง dashboard ใหม่ และตรวจสอบ KPI ให้ครบ
```

3. ตรวจสถานะใน `projects/CASS/output/qa/qa-report.json`
	หากมี critical error จะเป็น `NOT READY` และไม่ควรใช้ตัวเลขตัดสินใจ
	จนกว่าจะตรวจสอบข้อมูลต้นทาง
4. เมื่อผลลัพธ์พร้อมเผยแพร่ ให้ push dashboard ที่สร้างใหม่:

```bash
git add projects/CASS
git commit -m "Update CASS dashboard data"
git push
```

GitHub Pages จะ deploy อัตโนมัติหลัง push ไฟล์ dashboard

มาตรฐานที่ระบบสร้างให้:
- Mahidol Faculty theme และตรามหาวิทยาลัยใน header
- animation ตอนเปิดหน้า โดยรองรับการลดการเคลื่อนไหวของระบบ
- คลิกแท่งกราฟเพื่อ cross-filter KPI, กราฟ และตาราง แล้วล้างได้
- เลือกปีงบประมาณเพื่อดู KPI เทียบปีก่อนด้วยลูกศรและเปอร์เซ็นต์
- แยก 3 แท็บ: บริการมีรายได้, บริการวิชาการแบบไม่ได้รายได้, และเปรียบเทียบงบประมาณ
