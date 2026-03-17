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

$WebAppUrl = "https://script.google.com/macros/s/AKfycbzktruXAyFtyH0Jxyyh_2Fe6x2OOfsY9Yq_psM86gCwYiDS4eXSTvU34NuljbgYuI6Z5A/exec"  

# Nhập tay
if (-not $PhongBan) { $PhongBan = Read-Host "Nhập tên phòng ban" ; if (-not $PhongBan) { $PhongBan = "Chưa nhập" } }
if (-not $CanBo)    { $CanBo    = Read-Host "Nhập tên cán bộ"    ; if (-not $CanBo)    { $CanBo    = "Chưa nhập" } }

Write-Host "Đang thu thập thông tin PC..." -ForegroundColor Cyan

# Thông tin cơ bản
$TenMayTinh     = $env:COMPUTERNAME
$TenNguoiSuDung = $env:USERNAME
$Model          = (Get-CimInstance Win32_ComputerSystem).Model
$Serial         = (Get-CimInstance Win32_BIOS).SerialNumber
$CPU            = (Get-CimInstance Win32_Processor).Name

# ==================== NGÀY SỬ DỤNG PC ====================
$NgaySuDung = "Không lấy được"

# Cách 1 (Ưu tiên): Lấy thời gian tạo thư mục Windows
try {
    $NgaySuDung = (Get-Item "C:\Windows").CreationTime.ToString("dd/MM/yyyy")
}
catch { }

# Cách 2: Lấy từ Registry Source (nếu cách 1 lỗi)
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

# Cách 3: Registry InstallDate
if ($NgaySuDung -eq "Không lấy được") {
    try {
        $installTime = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name InstallDate).InstallDate
        $NgaySuDung = [DateTime]::FromFileTime($installTime).ToString("dd/MM/yyyy")
    }
    catch { }
}

# Cách 4: WMI
if ($NgaySuDung -eq "Không lấy được") {
    try {
        $NgaySuDung = ([WMI]'').ConvertToDateTime((Get-CimInstance Win32_OperatingSystem).InstallDate).ToString("dd/MM/yyyy")
    }
    catch { }
}

# ==================== RAM (gộp 1 cột) ====================
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

# ==================== Disk, IP, MAC... ====================
$diskList = @()
Get-PhysicalDisk | Sort-Object DeviceId | ForEach-Object {
    $sizeGB = [math]::Round($_.Size / 1GB, 0)
    $media = if ($_.MediaType -eq "UnSpecified" -or $_.MediaType -eq $null) { if ($_.BusType -eq "NVMe") { "NVMe SSD" } else { "SSD" } } else { $_.MediaType }
    $diskName = if ($_.FriendlyName -and $_.FriendlyName -notmatch "^Physical") { $_.FriendlyName.Trim() } else { "Disk $($_.DeviceId)" }
    $diskInfo = "$diskName : $media $sizeGB GB"
    if ($_.BusType -and $_.BusType -notin @("SATA","Unknown")) { $diskInfo += " ($($_.BusType))" }
    $diskList += $diskInfo
}
$Disk = $diskList -join "; " ; if (-not $Disk) { $Disk = "Không lấy được" }

$IP = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.*' } | Select-Object -First 1).IPAddress
$MAC = (Get-NetAdapter | Where-Object Status -eq 'Up' | Select-Object -First 1).MacAddress
$Windows = (Get-CimInstance Win32_OperatingSystem).Caption
$Mainboard = (Get-CimInstance Win32_BaseBoard).Manufacturer + " " + (Get-CimInstance Win32_BaseBoard).Product

# Gửi dữ liệu
$data = @{
    PhongBan      = $PhongBan
    CanBo         = $CanBo
    NgaySuDung    = $NgaySuDung          # <-- Cột này là ngày sử dụng PC
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
    Write-Host "✅ Upload thành công! Ngày sử dụng PC: $NgaySuDung" -ForegroundColor Green
} catch {
    Write-Host "❌ Lỗi: $($_.Exception.Message)" -ForegroundColor Red
}

$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
