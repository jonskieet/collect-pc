# ================================================
# COLLECT PC - Ngày sử dụng PC = Original Install Date
# ================================================

param (
    [string]$PhongBan = "",
    [string]$CanBo    = ""
)

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

# ===== PROGRESS FUNCTION =====
function Update-Progress {
    param(
        [string]$Status,
        [int]$Percent
    )
    Write-Progress -Activity "ACIC - Đang thu thập dữ liệu" -Status $Status -PercentComplete $Percent
}

$WebAppUrl = "https://script.google.com/macros/s/AKfycbzktruXAyFtyH0Jxyyh_2Fe6x2OOfsY9Yq_psM86gCwYiDS4eXSTvU34NuljbgYuI6Z5A/exec"  

# ==================== INPUT ====================
Update-Progress "Đang nhập thông tin người dùng..." 5

if (-not $PhongBan) { $PhongBan = Read-Host "Nhập tên phòng ban" ; if (-not $PhongBan) { $PhongBan = "Chưa nhập" } }
if (-not $CanBo)    { $CanBo    = Read-Host "Nhập tên cán bộ"    ; if (-not $CanBo)    { $CanBo    = "Chưa nhập" } }

# ==================== BASIC INFO ====================
Update-Progress "Đang lấy tên máy và user..." 10
$TenMayTinh     = $env:COMPUTERNAME
$TenNguoiSuDung = $env:USERNAME

Update-Progress "Đang lấy Model máy..." 15
$Model          = (Get-CimInstance Win32_ComputerSystem).Model

Update-Progress "Đang lấy Serial BIOS..." 20
$Serial         = (Get-CimInstance Win32_BIOS).SerialNumber

Update-Progress "Đang lấy thông tin CPU..." 25
$CPU            = (Get-CimInstance Win32_Processor).Name


# ==================== NGÀY SỬ DỤNG PC ====================
Update-Progress "Đang xác định ngày bắt đầu sử dụng PC..." 30

$NgaySuDung = "Không lấy được"

try {
    $NgaySuDung = (Get-Item "C:\Windows").CreationTime.ToString("dd/MM/yyyy")
}
catch { }

if ($NgaySuDung -eq "Không lấy được") {
    try {
        $sources = Get-ChildItem "HKLM:\SYSTEM\Setup\Source*" -ErrorAction SilentlyContinue |
                   ForEach-Object { Get-ItemProperty $_.PSPath }

        $allDates = $sources | Where-Object { $_.InstallDate } |
                    ForEach-Object { [DateTime]::FromFileTime($_.InstallDate) }

        if ($allDates) {
            $NgaySuDung = ($allDates | Sort-Object | Select-Object -First 1).ToString("dd/MM/yyyy")
        }
    }
    catch { }
}

# ==================== RAM ====================
Update-Progress "Đang phân tích RAM..." 40

$ramModules = Get-CimInstance Win32_PhysicalMemory
$ramList = @()

foreach ($ram in $ramModules) {
    $gb    = [math]::Round($ram.Capacity / 1GB, 0)
    $speed = if ($ram.ConfiguredClockSpeed) { $ram.ConfiguredClockSpeed } else { $ram.Speed }
    $ramList += "$gb GB $($speed)MHz"
}

$RAM_ChiTiet = $ramList -join " + "
$RAM_SoThanh = $ramModules.Count
$RAM_TocDo   = if ($ramModules) { ($ramModules | ForEach-Object { if($_.ConfiguredClockSpeed){$_.ConfiguredClockSpeed}else{$_.Speed} }) -join " / " } else { "N/A" }
$RAM_GB      = [math]::Round(($ramModules | Measure-Object Capacity -Sum).Sum / 1GB, 0)

$Ram = "$RAM_GB GB - $RAM_SoThanh thanh - $RAM_TocDo MHz - $RAM_ChiTiet"


# ==================== DISK ====================
Update-Progress "Đang quét ổ đĩa..." 55

$diskList = @()
Get-PhysicalDisk | Sort-Object DeviceId | ForEach-Object {

    $sizeGB = [math]::Round($_.Size / 1GB, 0)

    $media = if ($_.MediaType -eq "UnSpecified" -or $_.MediaType -eq $null) {
        if ($_.BusType -eq "NVMe") { "NVMe SSD" } else { "SSD" }
    } else { $_.MediaType }

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
if (-not $Disk) { $Disk = "Không lấy được" }


# ==================== NETWORK ====================
Update-Progress "Đang lấy IP và MAC..." 70

$IP = (Get-NetIPAddress | Where-Object {
    $_.AddressFamily -eq 'IPv4' -and
    $_.IPAddress -notlike '127.*' -and
    $_.IPAddress -notlike '169.*'
} | Select-Object -First 1).IPAddress

$MAC = (Get-NetAdapter | Where-Object Status -eq 'Up' | Select-Object -First 1).MacAddress


# ==================== OS + MAINBOARD ====================
Update-Progress "Đang lấy thông tin Windows..." 80
$Windows = (Get-CimInstance Win32_OperatingSystem).Caption

Update-Progress "Đang lấy thông tin Mainboard..." 85
$Mainboard = (Get-CimInstance Win32_BaseBoard).Manufacturer + " " + (Get-CimInstance Win32_BaseBoard).Product


# ==================== SEND DATA ====================
Update-Progress "Đang chuẩn bị dữ liệu..." 90

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

Update-Progress "Đang gửi dữ liệu lên server..." 95

try {
    Invoke-RestMethod -Uri $WebAppUrl -Method Post -Body $bodyBytes -ContentType "application/json; charset=utf-8" | Out-Null

    Write-Progress -Activity "COLLECT PC - Đang thu thập dữ liệu" -Completed

    Write-Host "✅ Upload thành công! Ngày sử dụng PC: $NgaySuDung" -ForegroundColor Green
}
catch {
    Write-Progress -Activity "COLLECT PC - Đang thu thập dữ liệu" -Completed
    Write-Host "❌ Lỗi: $($_.Exception.Message)" -ForegroundColor Red
}

$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
