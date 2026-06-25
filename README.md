# RECAP Auto-Filler

เว็บแอปสำหรับเติมข้อมูล **DIVISION / DEPT / SUB-DEPT / Class / PLANOGRAM**
ในไฟล์ RECAP.xlsx โดยอัตโนมัติจากไฟล์ 100 ช่อง (.xlsb)

---

## วิธี Deploy บน Vercel (ครั้งเดียว)

```bash
# 1. Clone หรือ init repo ใหม่
git init
git add .
git commit -m "initial commit"

# 2. Push ขึ้น GitHub
git remote add origin https://github.com/YOUR_USER/recap-filler.git
git push -u origin main

# 3. ไปที่ vercel.com → Import Repository → เลือก repo นี้
# Vercel จะ auto-detect Next.js และ deploy ให้เลย
```

## รัน Local

```bash
npm install
npm run dev
# เปิด http://localhost:3000
```

## วิธีใช้งาน

1. **Step 1** — อัปโหลดไฟล์ RECAP.xlsx
2. **Step 2** — อัปโหลดไฟล์ 100 ช่อง (.xlsb) ได้หลายไฟล์พร้อมกัน
3. **Step 3** — อัปโหลด DATA_SPACEMAN.xlsx (อัปโหลดใหม่ทุกสัปดาห์)
4. กดปุ่ม **ประมวลผล** — ระบบจะค้นหาและเติมข้อมูลอัตโนมัติ
5. **ตรวจสอบ**ผลลัพธ์ แก้ไขรายการที่ต้องการ
6. **ดาวน์โหลด** RECAP_filled.xlsx

## หลักการทำงาน

```
Barcode (จาก RECAP) 
  → Base sheet (ใน xlsb) → Sub-Class Code (10 หลัก)
  → Sh_ProdStructure (ใน xlsb) → DIVISION / DEPT / SUB-DEPT / Class
  → DATA_SPACEMAN → PLANOGRAM
```

## สีบอกสถานะ

| สี | ความหมาย |
|----|----------|
| 🟢 ยืนยันแล้ว | พบข้อมูลครบทั้ง 5 คอลัมน์ |
| 🟡 อนุมาน | พบจาก xlsb แต่ PLANOGRAM ไม่พบใน DATA_SPACEMAN |
| 🔴 ไม่พบ | ไม่พบบาร์โค้ดในไฟล์ 100 ช่อง — ต้องกรอกเอง |
