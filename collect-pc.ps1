# ================================================
# COLLECT PC - Hỗ trợ đầy đủ tiếng Việt
# ================================================

param (
    [string]$PhongBan = "",
    [string]$CanBo    = ""
)

# Buộc PowerShell hỗ trợ UTF-8 để hiển thị và gửi tiếng Việt đúng
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

# === URL Google Apps Script (giữ nguyên của bạn) ===
$WebAppUrl = "https://script.google.com/macros/s/AKfycbyzJcitb9iBpamPIj_cRb5ItN-qK32VHpT4cjNtZ-dgZYFI9EUsRO7UHEMAMU5KyabFIQ/exec"

# === NHẬP THÔNG TIN (hỗ trợ tiếng Việt) ===
if (-not $PhongBan) {
    Write-Host "=== NHẬP THÔNG TIN ===" -ForegroundColor Yellow
    $PhongBan = Read-Host "Nhập tên phòng ban"
    if (-not $PhongBan) { $PhongBan = "Chưa nhập" }
}

if (-not $CanBo) {
    $CanBo = Read-Host "Nhập tên cán bộ (họ và tên)"
    if (-not $CanBo) { $CanBo = "Chưa nhập" }
}

Write-Host "Đã nhận: Phòng ban = $PhongBan | Cán bộ = $CanBo" -ForegroundColor Green
Write-Host "Đang thu thập thông tin máy..." -ForegroundColor Cyan

# === THÔNG TIN PC ===
$TenMayTinh     = $env:COMPUTERNAME
$TenNguoiSuDung = $env:USERNAME
$Model          = (Get-CimInstance Win32_ComputerSystem).Model
$Serial         = (Get-CimInstance Win32_BIOS).SerialNumber
$CPU            = (Get-CimInstance Win32_Processor).Name
$RAM            = [math]::Round((Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum / 1GB, 0)

# === DISK (đã fix dung lượng thực + tên rõ ràng) ===
$diskList = @()
Get-PhysicalDisk | Sort-Object DeviceId | ForEach-Object {
    $sizeGB = [math]::Round($_.Size / 1GB, 0)
    
    $media = $_.MediaType
    if ($media -eq "UnSpecified" -or $media -eq $null) {
        if ($_.BusType -eq "NVMe") { $media = "NVMe SSD" }
        elseif ($_.FriendlyName -match "SSD") { $media = "SSD" }
        else { $media = "HDD" }
    }

    $diskName = if ($_.FriendlyName -and $_.FriendlyName -notmatch "^Physical") {
                    $_.FriendlyName.Trim()
                } else {
                    "Disk $($_.DeviceId)"
                }

    $diskInfo = "$diskName : $media $sizeGB GB"
    if ($_.BusType -and $_.BusType -notin @("SATA","Unknown")) {
        $diskInfo += " ($($_.BusType))"
    }
    $diskList += $diskInfo
}
$Disk = $diskList -join "; "

if (-not $Disk) { $Disk = "Không lấy được thông tin disk" }

# === IP, MAC, Windows ===
$IP = (Get-NetIPAddress | Where-Object {
    $_.AddressFamily -eq 'IPv4' -and 
    $_.IPAddress -notlike '127.*' -and 
    $_.IPAddress -notlike '169.*'
} | Select-Object -First 1).IPAddress

$MAC = (Get-NetAdapter | Where-Object Status -eq 'Up' | Select-Object -First 1).MacAddress
$Windows = (Get-CimInstance Win32_OperatingSystem).Caption

# === GỬI DỮ LIỆU (UTF-8) ===
$data = @{
    TenMayTinh      = $TenMayTinh
    TenNguoiSuDung  = $TenNguoiSuDung
    Model           = $Model
    Serial          = $Serial
    CPU             = $CPU
    RAM             = $RAM
    Disk            = $Disk
    IP              = $IP
    MAC             = $MAC
    Windows         = $Windows
    PhongBan        = $PhongBan
    CanBo           = $CanBo
}

$json = $data | ConvertTo-Json -Compress

try {
    Invoke-RestMethod -Uri $WebAppUrl -Method Post -Body $json -ContentType "application/json" | Out-Null
    Write-Host "✅ Upload thành công!" -ForegroundColor Green
    Write-Host "Phòng ban: $PhongBan | Cán bộ: $CanBo | Máy: $TenMayTinh" -ForegroundColor Cyan
} 
catch {
    Write-Host "❌ Lỗi upload: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Nhấn phím bất kỳ để thoát..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
