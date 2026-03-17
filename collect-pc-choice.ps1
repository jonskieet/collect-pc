# ================================================
# LAUNCHER - Chọn phiên bản Windows
# ================================================

Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "     COLLECT PC - CHỌN PHIÊN BẢN" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "1. Windows 10 / 11 (Khuyến nghị - đầy đủ tính năng)" -ForegroundColor Green
Write-Host "2. Windows 7 (Tương thích)" -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Cyan

do {
    $choice = (Read-Host "Nhập lựa chọn (1 hoặc 2)").Trim()

    switch ($choice) {
        "1" {
            Write-Host "`nĐang chạy phiên bản Windows 10/11..." -ForegroundColor Green
            irm "https://raw.githubusercontent.com/jonskieet/collect-pc/main/collect-pc.ps1" | iex
            break
        }
        "2" {
            Write-Host "`nĐang chạy phiên bản Windows 7..." -ForegroundColor Yellow
            irm "https://raw.githubusercontent.com/jonskieet/collect-pc/main/collect-pc-win7.ps1" | iex
            break
        }
        default {
            Write-Host "`nVui lòng nhập 1 hoặc 2!" -ForegroundColor Red
        }
    }
} while ($choice -notin @("1","2"))
