# Taxwise — Tax Data Comparison

เปรียบเทียบข้อมูลจากไฟล์ Excel สองชุด จัดกลุ่มตามชื่อผู้มีหน้าที่หักและประเภทเงินได้ แสดงผลบนหน้าเว็บ และดาวน์โหลดรายงาน Excel

## การประมวลผลและข้อมูล

- ไฟล์ Excel ถูกอ่านในเบราว์เซอร์ของผู้ใช้ ไม่ได้อัปโหลดไปยังเซิร์ฟเวอร์หรือ Firebase
- ผลเปรียบเทียบอยู่ในหน่วยความจำของหน้าเว็บเท่านั้น และสร้างไฟล์ Excel ดาวน์โหลดในเบราว์เซอร์
- Cloud Firestore เก็บเฉพาะกฎปรับชื่อที่ใช้ร่วมกัน
- ผู้ใช้ทั่วไปเปรียบเทียบไฟล์ได้โดยไม่ต้องล็อกอิน ผู้ดูแลต้องล็อกอินเฉพาะเมื่อต้องเพิ่มหรือลบกฎชื่อ
- ถ้ายังไม่ตั้งค่า Firebase ระบบจะใช้ local storage สำหรับทดลอง กฎในโหมดนี้ไม่แชร์ข้ามเครื่อง

## รันบน Windows 11

ติดตั้ง Node.js 20+ จากนั้นดับเบิลคลิก `start.bat` ที่โฟลเดอร์หลัก ระบบจะติดตั้งแพ็กเกจที่จำเป็น เปิดเว็บ และเปิดที่อยู่ `http://localhost:5173` ให้อัตโนมัติ คงหน้าต่างเซิร์ฟเวอร์ไว้ระหว่างใช้งาน

## เผยแพร่ด้วย GitHub Pages และ Firebase

### ตั้งค่า Firebase

1. สร้าง Firebase project และเพิ่ม Web App ใน Firebase Console
2. เปิด **Authentication → Email/Password** และสร้างบัญชีผู้ดูแลหนึ่งบัญชี
3. สร้าง Cloud Firestore database
4. ในไฟล์ `firestore.rules` เปลี่ยน `REPLACE_WITH_ADMIN_EMAIL` เป็นอีเมลบัญชีผู้ดูแล แล้วนำกฎไป Publish ในหน้า Firestore Rules
5. ใน Firebase Authentication → Settings → Authorized domains เพิ่มโดเมน GitHub Pages เช่น `ชื่อผู้ใช้.github.io`

กฎ Firestore อนุญาตให้อ่านรายการชื่อได้โดยไม่ล็อกอิน เพื่อให้ผู้ใช้ทั่วไปเปรียบเทียบได้ และอนุญาตเพิ่ม/ลบกฎเฉพาะบัญชีผู้ดูแลที่กำหนด ดังนั้นรายการชื่อใน Firestore จะอ่านได้จากเว็บสาธารณะ ห้ามใส่ข้อมูลลับไว้ในกฎชื่อ

### ตั้งค่า GitHub

1. Push โปรเจกต์นี้ไปยัง GitHub repository
2. ที่ **Settings → Secrets and variables → Actions → Variables** เพิ่มตัวแปรต่อไปนี้จาก Firebase Web App config และผู้ดูแล:
   - `VITE_FIREBASE_API_KEY`
   - `VITE_FIREBASE_AUTH_DOMAIN`
   - `VITE_FIREBASE_PROJECT_ID`
   - `VITE_FIREBASE_APP_ID`
   - `VITE_FIREBASE_ADMIN_EMAIL`
3. ที่ **Settings → Pages** ตั้ง Source เป็น **GitHub Actions**
4. Push ไปยัง branch `main` หรือ `master` เพื่อเริ่ม workflow ใน `.github/workflows/pages.yml`

เมื่อ workflow สำเร็จ GitHub จะแสดง URL ของเว็บไซต์ในหน้า **Settings → Pages** โดยทั่วไปอยู่ในรูป `https://ชื่อผู้ใช้.github.io/ชื่อ-repository/`

### ทดลอง Firebase ในเครื่อง

คัดลอก `frontend/.env.example` เป็น `frontend/.env` แล้วใส่ค่า Firebase config และอีเมลผู้ดูแล จากนั้นรัน `start.bat` ค่าใน `.env` ใช้สำหรับ build หน้าเว็บเท่านั้น และไม่ควรใส่รหัสผ่านหรือ service account key ลงในไฟล์นี้

## รูปแบบไฟล์ Excel

ระบบอ่านชีตแรก โดยมีหัวตารางแถวแรก และรองรับ `.xlsx` ขนาดไม่เกิน 20 MB ต่อไฟล์

| คอลัมน์ | ตัวอย่างหัวตาราง |
|---|---|
| ชื่อผู้มีหน้าที่หัก | `ชื่อผู้มีหน้าที่หัก`, `ผู้มีหน้าที่หัก`, `ชื่อผู้จ่าย`, `payer` |
| ประเภทเงินได้ | `ประเภทเงินได้`, `ประเภทเงินได้พึงประเมิน`, `income type`, `type` |
| จำนวนเงิน | `จำนวนเงิน`, `ยอดเงิน`, `amount`, `income amount` |
| ภาษี (ไม่บังคับ) | `ภาษีหัก ณ ที่จ่าย`, `จำนวนภาษี`, `ภาษี`, `tax` |

หากไม่มีคอลัมน์ภาษี ระบบถือว่ายอดภาษีเป็นศูนย์ ผลต่างคำนวณเป็น **ชุด B − ชุด A**

## Build ตรวจสอบ

```bat
cd frontend
npm.cmd ci
npm.cmd run build
```

