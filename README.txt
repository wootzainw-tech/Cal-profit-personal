ExpenseTracker Pro - Complete Fixed Version

ไฟล์ภายใน
1) index.html
   - นำไปแทน index.html เดิมใน GitHub
   - รองรับดึงข้อมูลล่าสุดทุก 5 วินาที
   - ไม่ล้างฟอร์มขณะกำลังกรอก
   - เก็บรายการรอซิงก์ไว้จนกว่าจะยืนยันว่าบันทึกเข้า Google Sheets แล้ว
   - แดชบอร์ดใช้เดือนและปีปัจจุบันเป็นค่าเริ่มต้น

2) google_apps_script_code.js
   - นำไปแทน Code.gs ใน Google Apps Script
   - บันทึก/แก้ไขเฉพาะรายการตาม ID
   - ใช้ LockService ป้องกันหลายอุปกรณ์เขียนชนกัน
   - รองรับ Category ID สำหรับแดชบอร์ด

วิธีติดตั้ง
A. GitHub / Vercel
- อัปโหลด index.html แทนไฟล์เดิม
- Commit และรอ Vercel Deploy สำเร็จ

B. Google Apps Script
- เปิด Google Sheets > ส่วนขยาย > Apps Script
- ลบโค้ดเดิมใน Code.gs แล้ววางโค้ดจาก google_apps_script_code.js
- กดบันทึก
- เลือกฟังก์ชัน setup แล้วกด Run 1 ครั้ง
- ไปที่ Deploy > Manage deployments > Edit
- เลือก New version แล้ว Deploy
- Execute as: Me
- Who has access: Anyone

C. หลังติดตั้ง
- เปิดเว็บแล้วกด Ctrl + F5
- ตรวจว่า PC และมือถือใช้ Google Apps Script Web App URL เดียวกัน
- ทดสอบบันทึกรายการจากทั้งสองอุปกรณ์
