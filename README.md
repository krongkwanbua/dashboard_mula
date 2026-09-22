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
