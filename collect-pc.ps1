# ================================================
# COLLECT PC - Fix hoàn chỉnh tiếng Việt
# ================================================

param (
    [string]$PhongBan = "",
    [string]$CanBo    = ""
)

# Buộc UTF-8 toàn script
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

$WebAppUrl = "https://script.google.com/macros/s/AKfycbyzC2NKTE1rQa0tgBuXXoSUT_Q4OmqD14tUf9uN_VJj6BVG40L4chAegL-JtBWNswYjmw/exec"  

# Nhập thông tin
if (-not $PhongBan) {
    $PhongBan = Read-Host "Nhập tên phòng ban"
    if (-not $PhongBan) { $PhongBan = "Chưa nhập" }
}
if (-not $CanBo) {
    $CanBo = Read-Host "Nhập tên cán bộ"
    if (-not $CanBo) { $CanBo = "Chưa nhập" }
}

Write-Host "Đang thu thập thông tin PC..." -ForegroundColor Cyan

# Thông tin PC
$TenMayTinh     = $env:COMPUTERNAME
$TenNguoiSuDung = $env:USERNAME
$Model          = (Get-CimInstance Win32_ComputerSystem).Model
$Serial         = (Get-CimInstance Win32_BIOS).SerialNumber
$CPU            = (Get-CimInstance Win32_Processor).Name
$RAM            = [math]::Round((Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum / 1GB, 0)

# Disk (giữ nguyên phiên bản sạch)
$diskList = @()
Get-PhysicalDisk | Sort-Object DeviceId | ForEach-Object {
    $sizeGB = [math]::Round($_.Size / 1GB, 0)
    $media = if ($_.MediaType -eq "UnSpecified" -or $_.MediaType -eq $null) {
                 if ($_.BusType -eq "NVMe") { "NVMe SSD" } else { "SSD" }
             } else { $_.MediaType }
    $diskName = if ($_.FriendlyName -and $_.FriendlyName -notmatch "^Physical") { $_.FriendlyName.Trim() } else { "Disk $($_.DeviceId)" }
    $diskInfo = "$diskName : $media $sizeGB GB"
    if ($_.BusType -and $_.BusType -notin @("SATA","Unknown")) { $diskInfo += " ($($_.BusType))" }
    $diskList += $diskInfo
}
$Disk = $diskList -join "; " ; if (-not $Disk) { $Disk = "Không lấy được" }

# IP, MAC, Windows
$IP = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.*' } | Select-Object -First 1).IPAddress
$MAC = (Get-NetAdapter | Where-Object Status -eq 'Up' | Select-Object -First 1).MacAddress
$Windows = (Get-CimInstance Win32_OperatingSystem).Caption

# === PHẦN QUAN TRỌNG: Gửi JSON đúng UTF-8 ===
$data = @{
    PhongBan        = $PhongBan
    CanBo           = $CanBo
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
}

$json = $data | ConvertTo-Json -Compress -Depth 10

# Gửi với UTF-8 bytes (fix mojibake)
try {
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    Invoke-RestMethod -Uri $WebAppUrl -Method Post -Body $bodyBytes -ContentType "application/json; charset=utf-8" | Out-Null
    
    Write-Host "✅ Upload thành công!" -ForegroundColor Green
    Write-Host "Phòng ban: $PhongBan | Cán bộ: $CanBo" -ForegroundColor Cyan
} 
catch {
    Write-Host "❌ Lỗi: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Nhấn phím bất kỳ để thoát..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
