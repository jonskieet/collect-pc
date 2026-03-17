# === THAY ĐỔI DÒNG NÀY BẰNG URL CỦA BẠN ===
$WebAppUrl = "https://script.google.com/macros/s/AKfycbyV3JChqfiYTfROjxi_QhpALuNSnlt4TfQDmq_7WWUykI8OZFEH9bSD_DsSrcjXd3QSoQ/exec"

# === NHẬP TAY 2 CỘT MỚI ===
Write-Host "=== NHẬP THÔNG TIN PHÒNG BAN VÀ CÁN BỘ ===" -ForegroundColor Yellow
$PhongBan = Read-Host "Tên phòng ban (ví dụ: Kinh doanh, IT, Kế toán)"
if (-not $PhongBan) { $PhongBan = "Chưa nhập" }

$CanBo = Read-Host "Tên cán bộ (họ tên đầy đủ)"
if (-not $CanBo) { $CanBo = "Chưa nhập" }

Write-Host "Bạn đã nhập: Phòng ban = $PhongBan | Cán bộ = $CanBo" -ForegroundColor Green
Write-Host "Bắt đầu thu thập thông tin PC..." -ForegroundColor Cyan

# === THÔNG TIN PC ===
$TenMayTinh     = $env:COMPUTERNAME
$TenNguoiSuDung = $env:USERNAME
$Model          = (Get-CimInstance Win32_ComputerSystem).Model
$Serial         = (Get-CimInstance Win32_BIOS).SerialNumber
$CPU            = (Get-CimInstance Win32_Processor).Name
$RAM            = [math]::Round((Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum / 1GB, 0)

# ==================== PHẦN DISK ĐÃ SỬA ====================
$diskList = @()

Get-PhysicalDisk | Sort-Object DeviceId | ForEach-Object {
    # Dung lượng thực tế (không khấu trừ) - làm tròn chuẩn
    $sizeGB = [math]::Round($_.Size / 1GB, 0)
    
    # Xác định loại ổ (ưu tiên MediaType, fallback BusType + Model)
    $media = $_.MediaType
    if ($media -eq "UnSpecified" -or $media -eq $null) {
        if ($_.BusType -eq "NVMe") { $media = "NVMe SSD" }
        elseif ($_.BusType -eq "SATA" -and $_.FriendlyName -match "SSD") { $media = "SSD" }
        elseif ($_.BusType -eq "USB") { $media = "USB" }
        else { $media = "HDD" }
    }

    # Sửa lỗi hiển thị "?" → dùng DeviceId hoặc FriendlyName
    $diskName = if ($_.FriendlyName -and $_.FriendlyName -notmatch "^PhysicalDisk") {
                    $_.FriendlyName
                } else {
                    "Disk $($_.DeviceId)"
                }

    $diskInfo = "$diskName : $media $sizeGB GB"

    # Thêm BusType nếu là NVMe hoặc USB để chi tiết hơn
    if ($_.BusType -and $_.BusType -notin @("SATA","Unknown")) {
        $diskInfo += " ($($_.BusType))"
    }

    $diskList += $diskInfo
}

$Disk = $diskList -join "; "

# Fallback nếu không lấy được PhysicalDisk
if (-not $Disk) {
    $total = 0
    Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object { $total += $_.Size }
    $Disk = [math]::Round($total / 1GB, 0) + " GB (tổng)"
}
# ========================================================

# IP và các trường còn lại
$IP = (Get-NetIPAddress | Where-Object {
    $_.AddressFamily -eq 'IPv4' -and $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.*'
} | Select-Object -First 1).IPAddress

$MAC = (Get-NetAdapter | Where-Object Status -eq 'Up' | Select-Object -First 1).MacAddress
$Windows = (Get-CimInstance Win32_OperatingSystem).Caption

# === GỬI DỮ LIỆU ===
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
    $response = Invoke-RestMethod -Uri $WebAppUrl -Method Post -Body $json -ContentType "application/json"
    Write-Host "✅ Upload thành công! Đã thêm dòng mới vào Google Sheets." -ForegroundColor Green
    Write-Host "Thông tin: $CanBo - $PhongBan - Máy: $TenMayTinh" -ForegroundColor Cyan
} catch {
    Write-Host "❌ Lỗi upload: $($_.Exception.Message)" -ForegroundColor Red
}
