# 📦 วิธีติดตั้งและใช้งาน BossPOS

---

## ✅ ข้อกำหนดเบื้องต้น

- บัญชี Google (Gmail)
- Google Sheets
- Google Apps Script
- เบราว์เซอร์ที่ทันสมัย (Chrome, Firefox, Safari)

---

## 🔧 ขั้นตอนติดตั้ง

### **ขั้นตอนที่ 1: สร้าง Google Sheet**

1. ไปที่ [Google Sheets](https://sheets.google.com)
2. สร้าง Sheet ใหม่ (New Spreadsheet)
3. ตั้งชื่อ: **"BossPOS"**

### **ขั้นตอนที่ 2: สร้าง Sheets ที่ต้องใช้**

คัดลอกรูปแบบนี้สำหรับแต่ละ Sheet:

#### **2.1 Sheet: Products** (สินค้า)
```
| barcode | name | price | cost | stock |
|---------|------|-------|------|-------|
| 001     | สินค้า 1 | 100 | 50 | 10 |
| 002     | สินค้า 2 | 200 | 100 | 5 |
```

#### **2.2 Sheet: Sales_Header** (ใบเสร็จ)
```
| billNo | date | total | type | memberId | ch_เงินสด | ch_เงินโอน | ch_บัตรสวัสดิการแห่งรัฐ |
|--------|------|-------|------|----------|-----------|-----------|----------------------|
```

#### **2.3 Sheet: Sales_Detail** (รายการสินค้า)
```
| billNo | barcode | name | price | cost | qty | total | profit |
|--------|---------|------|-------|------|-----|-------|--------|
```

#### **2.4 Sheet: Settings** (ตั้งค่า)
```
| A | B |
|---|---|
| billNo | 1 |
| shopName | ร้านของคุณ |
| shopAddress | ที่อยู่ร้าน |
| shopTel | เบอร์โทร |
| footer | ขอบคุณที่ใช้บริการ |
| logoUrl | (URL รูปโลโก้ หรือว่างไว้) |
```

#### **2.5 Sheet: Members** (สมาชิก)
```
| memberId | name | phone | points | joinDate | note | status |
|----------|------|-------|--------|----------|------|--------|
```

#### **2.6 Sheet: Shifts** (กะการทำงาน)
```
| shiftNo | openTime | closeTime | duration | startCash | totalSales | billCount | ... |
|---------|----------|-----------|----------|-----------|------------|-----------|-----|
```

---

### **ขั้นตอนที่ 3: เพิ่ม Google Apps Script**

1. ที่ Google Sheet ของคุณ → **Tools → Apps Script**
2. ลบโค้ดตัวอย่างทั้งหมด
3. Copy ไฟล์ `code.gs` จาก Repository นี้
4. Paste เข้ากไป

### **ขั้นตอนที่ 4: เพิ่ม HTML Files**

1. **ในหน้า Apps Script Editor:**
   - Click **"+" → HTML file**
   - สร้างไฟล์ชื่อ: `index`
   - Paste โค้ด `index.html`

2. ทำเหมือนกับไฟล์อื่นๆ:
   - `receipt`
   - `display`
   - `local_display`

### **ขั้นตอนที่ 5: Deploy Web App**

1. **ใน Apps Script:**
   - Click **"Deploy → New deployment"**
   - เลือก Type: **"Web app"**
   - Execute as: **คุณ**
   - Who has access: **Anyone**
   - Click **"Deploy"**

2. **คัดลอก URL ที่ได้มา**
   ```
   https://script.google.com/macros/d/{DEPLOYMENT_ID}/usercontent
   ```

---

## 🎯 วิธีใช้งาน

### **หน้าหลักขาย** (index.html)
1. เปิด URL ที่ deploy
2. สแกน Barcode หรือค้นหาสินค้า
3. เลือกวิธีการชำระเงิน
4. คลิก **"ชำระเงิน"**

### **ใบเสร็จ** (receipt.html)
1. เข้า URL: `?page=receipt`
2. ค้นหาบิลเลข
3. ดูและพิมพ์ใบเสร็จ

### **จอแสดงลูกค้า** (display.html)
- เข้า URL: `?page=display`
- แสดงข้อมูลการขายแบบ Real-time

### **ตั้งค่า**
- ไปที่ Apps Script → Properties ตั้งค่า payment channels
- ไปที่ Sheet Settings แก้ไขข้อมูลร้าน

---

## 🔑 Key Functions

### **getProducts()**
ดึงรายการสินค้าทั้งหมด

### **saveSale(data)**
บันทึกการขายใหม่
```javascript
saveSale({
  total: 500,
  type: 'Cash',
  items: [{barcode: '001', name: 'สินค้า', price: 500, qty: 1}],
  memberId: 'M001',
  channels: [{key: 'cash', amount: 500}]
})
```

### **getBill(billNo)**
ดึงข้อมูลใบเสร็จ

### **getDashboard()**
ดึงข้อมูล Dashboard (ยอดขาย, กำไร, สินค้าขายดี)

### **openShift() / closeShift()**
เปิด/ปิดกะการทำงาน

---

## 🛠️ การแก้ไขปัญหา

### **ปัญหา: ไม่ได้ data สินค้า**
- ✅ ตรวจสอบ Sheet "Products" มีข้อมูลหรือไม่
- ✅ ตรวจสอบชื่อคอลัมน์ barcode, name, price

### **ปัญหา: ไม่บันทึกการขาย**
- ✅ ตรวจสอบ Sheet "Sales_Header" และ "Sales_Detail" มีไหม
- ✅ ตรวจสอบ Settings sheet มีค่า billNo ไหม

### **ปัญหา: ใบเสร็จพิมพ์ไม่ได้**
- ✅ ตรวจสอบชื่อไฟล์ HTML (ต้องตรงกับเรียกจาก code.gs)

### **ปัญหา: payment channels ไม่แสดง**
- ✅ ตรวจสอบ PropertiesService มีการบันทึก 'PAYMENT_CHANNELS' ไหม
- ✅ ลอง resetPaymentChannels() ใน Apps Script

---

## 📱 URL Routes

```
/                     → หน้าหลักขาย (index.html)
?page=receipt         → ใบเสร็จ (receipt.html)
?page=display         → จอลูกค้า (display.html)
?page=local_display   → จอเครื่อง (local_display.html)
```

---

## 🔐 ความปลอดภัย

- ⚠️ ระบบนี้ใช้ Google Sheets ซึ่ง share ตาม URL
- ✅ ตั้งค่า Sheet ให้เป็น "Restricted" ถ้าต้องการ
- ✅ ใช้ Google Sign-in เพิ่มเติมได้ (ต้องแก้ code)

---

## 📊 Backup ข้อมูล

- ✅ Google Sheets ทำ backup อัตโนมัติ
- ✅ สามารถ Export ข้อมูลเป็น CSV ได้
- ✅ ใช้ Version history ของ Sheets

---

## 🎓 คำแนะนำเพิ่มเติม

- เพิ่ม User authentication ได้ (Google Sign-in)
- เชื่อมต่อกับระบบ Accounting อื่นๆ ได้
- เพิ่มระบบ Discount / Promo ได้
- เพิ่มระบบ Multi-store ได้

---

## 📞 ติดต่อและสนับสนุน

สำหรับปัญหาหรือคำแนะนำ:
- 📧 GitHub Issues: https://github.com/bosskungsub1-png/bosspos/issues
- 💬 Discussions

---

**เตรียมพร้อม!** 🚀 ระบบของคุณพร้อมสำหรับใช้งานแล้ว!
