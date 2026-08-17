# ============================================================
# Deploy Script - ระบบนิเทศภายในโรงเรียนนราศึกษาธิการ
# รันคำสั่งนี้ใน PowerShell เพื่อ Deploy Apps Script
# ============================================================

$CLASP = "C:\Users\NARAPEO\AppData\Roaming\npm\clasp.cmd"
$SCRIPT_DIR = "D:\nited\apps-script"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  ระบบนิเทศภายในโรงเรียนนราศึกษาธิการ" -ForegroundColor Cyan
Write-Host "  Deploy Apps Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Login
Write-Host "[1/4] กำลังเข้าสู่ระบบ Google..." -ForegroundColor Yellow
& $CLASP login 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "เกิดข้อผิดพลาดในการล็อกอิน" -ForegroundColor Red
    exit 1
}
Write-Host "ล็อกอินสำเร็จ!" -ForegroundColor Green
Write-Host ""

# Step 2: สร้าง .clasp.json
Write-Host "[2/4] กำลังตั้งค่า Project..." -ForegroundColor Yellow
Set-Location $SCRIPT_DIR

# Step 3: Push โค้ด
Write-Host "[3/4] กำลังอัพโหลดโค้ด..." -ForegroundColor Yellow
& $CLASP push --force 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "เกิดข้อผิดพลาดในการ Push โค้ด" -ForegroundColor Red
    exit 1
}
Write-Host "อัพโหลดโค้ดสำเร็จ!" -ForegroundColor Green
Write-Host ""

# Step 4: Deploy
Write-Host "[4/4] กำลัง Deploy เป็น Web App..." -ForegroundColor Yellow
& $CLASP deploy --description "Deployed via script" --force 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "เกิดข้อผิดพลาดในการ Deploy" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Deploy สำเร็จ!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# แสดง URL
$DEPLOY_URL = "https://script.google.com/macros/s/AKfycbz-NmIUQHTrpy5os8Sui90jVghZUOSkC0sgXehvy0h8n620U_gC6rkJXhbK6gR6EK0vkg/exec"
Write-Host "Web App URL: $DEPLOY_URL" -ForegroundColor Cyan
Write-Host ""
Write-Host "ขั้นตอนถัดไป:" -ForegroundColor Yellow
Write-Host "1. คัดลอก URL ด้านบนไปวางใน js/config.js" -ForegroundColor White
Write-Host "2. แก้ไข APPS_SCRIPT_URL ใน config.js" -ForegroundColor White
Write-Host "3. Push ขึ้น GitHub Pages" -ForegroundColor White
