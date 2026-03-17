# ================================================
# COLLECT PC - PHIÊN BẢN WINDOWS 7
# ================================================

param (
    [string]$PhongBan = "",
    [string]$CanBo    = ""
)

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

$WebAppUrl = "https://script.google.com/macros/s/AKfycbzktruXAyFtyH0Jxyyh_2Fe6x2OOfsY9Yq_psM86gCwYiDS4eXSTvU34NuljbgYuI6Z5A/exec"  

# ==================== NHẬP TAY ====================
if (-not $PhongBan) { 
    $PhongBan = Read-Host "Nhập tên phòng ban" 
    if (-not $PhongBan) { $PhongBan = "Chưa nhập" } 
}
if (-not $CanBo) { 
    $CanBo = Read-Host "Nhập tên cán bộ" 
    if (-not $CanBo) { $CanBo = "Chưa nhập" } 
}

Write-Host "Đang thu thập thông tin PC (phiên bản Windows 7)..." -ForegroundColor Cyan

# ==================== THÔNG TIN CƠ BẢN ====================
$TenMayTinh     = $env:COMPUTERNAME
$TenNguoiSuDung = $env:USERNAME
$Model          = (Get-WmiObject Win32_ComputerSystem).Model
$Serial         = (Get-WmiObject Win32_BIOS).SerialNumber
$CPU            = (Get-WmiObject Win32_Processor).Name

# ==================== NGÀY SỬ DỤNG PC ====================
$NgaySuDung = "Không lấy được"
try {
    $installDate = (Get-WmiObject Win32_OperatingSystem).InstallDate
    $NgaySuDung = [Management.ManagementDateTimeConverter]::ToDateTime($installDate).ToString("dd/MM/yyyy")
}
catch { }

# ==================== RAM (GỘP 1 CỘT) ====================
$ramModules = Get-WmiObject Win32_PhysicalMemory

$ramList = @()
foreach ($ram in $ramModules) {
    $gb = [math]::Round($ram.Capacity / 1GB, 0)
    $speed = if ($ram.Speed) { $ram.Speed } else { "N/A" }
    $ramList += "$gb GB $($speed)MHz"
}

$RAM_ChiTiet = if ($ramList.Count -gt 0) { $ramList -join " + " } else { "Không lấy được" }
$RAM_SoThanh = $ramModules.Count
$RAM_TocDo   = if ($ramModules) { ($ramModules | ForEach-Object { $_.Speed }) -join " / " } else { "N/A" }
$RAM_GB      = [math]::Round(($ramModules | Measure-Object Capacity -Sum).Sum / 1GB, 0)

$Ram = "$RAM_GB GB - $RAM_SoThanh thanh - $RAM_TocDo MHz - $RAM_ChiTiet"

# ==================== DISK (Dùng WMI cũ cho Win7) ====================
$diskList = @()
$disks = Get-WmiObject Win32_DiskDrive | Where-Object { $_.MediaType -match "Fixed" -or $_.InterfaceType -match "IDE|SATA|SCSI" }

foreach ($disk in $disks) {
    $sizeGB = [math]::Round($disk.Size / 1GB, 0)
    $media = if ($disk.Model -match "SSD") { "SSD" } else { "HDD" }
    $diskName = if ($disk.Model) { $disk.Model.Trim() } else { "Disk $($disk.Index)" }
    $diskList += "$diskName : $media $sizeGB GB"
}
$Disk = if ($diskList.Count -gt 0) { $diskList -join "; " } else { "Không lấy được" }

# ==================== IP, MAC, Windows ====================
$IP = (Get-WmiObject Win32_NetworkAdapterConfiguration | 
       Where-Object { $_.IPEnabled -eq $true -and $_.IPAddress -ne $null } | 
       Select-Object -First 1).IPAddress

$MAC = (Get-WmiObject Win32_NetworkAdapter | 
        Where-Object { $_.PhysicalAdapter -eq $true -and $_.MACAddress -ne $null } | 
        Select-Object -First 1).MACAddress

$Windows = (Get-WmiObject Win32_OperatingSystem).Caption
$Mainboard = (Get-WmiObject Win32_BaseBoard).Manufacturer + " " + (Get-WmiObject Win32_BaseBoard).Product

# ==================== GỬI DỮ LIỆU ====================
$data = @{
    PhongBan      = $PhongBan
    CanBo         = $CanBo
    NgaySuDung    = $NgaySuDung
    TenMayTinh    = $TenMayTinh
    TenNguoiSuDung= $TenNguoiSuDung
    Model         = $Model
    Serial        = $Serial
    CPU           = $CPU
    Ram           = $Ram
    Mainboard     = $Mainboard
    Disk          = $Disk
    IP            = $IP
    MAC           = $MAC
    Windows       = $Windows
}

$json = $data | ConvertTo-Json -Compress -Depth 10
$bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($json)

try {
    Invoke-RestMethod -Uri $WebAppUrl -Method Post -Body $bodyBytes -ContentType "application/json; charset=utf-8" | Out-Null
    Write-Host "✅ Upload thành công! (Windows 7)" -ForegroundColor Green
    Write-Host "Ngày sử dụng: $NgaySuDung" -ForegroundColor Cyan
} catch {
    Write-Host "❌ Lỗi upload: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nNhấn phím bất kỳ để thoát..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
