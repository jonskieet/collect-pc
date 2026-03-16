# === THAY ĐỔI DÒNG NÀY BẰNG URL CỦA BẠN ===
$WebAppUrl = "https://script.google.com/macros/s/AKfycbx4dwSloGihLlYBExIPCk81z3R6MO1Yre1Xlp_uD7aYQaGfOBQ4iYPXOsTr0h____RRQg/exec"

$TenMayTinh     = $env:COMPUTERNAME
$TenNguoiSuDung = $env:USERNAME
$Model          = (Get-CimInstance Win32_ComputerSystem).Model
$Serial         = (Get-CimInstance Win32_BIOS).SerialNumber
$CPU            = (Get-CimInstance Win32_Processor).Name
$RAM            = [math]::Round((Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum / 1GB, 0)

# Disk = tổng dung lượng tất cả ổ cứng
$Disk = 0
Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object { $Disk += $_.Size }
$Disk = [math]::Round($Disk / 1GB, 0)

# IP (local IPv4 chính, không phải 169.x hay loopback)
$IP = (Get-NetIPAddress | Where-Object {
    $_.AddressFamily -eq 'IPv4' -and
    $_.IPAddress -notlike '127.*' -and
    $_.IPAddress -notlike '169.*'
} | Select-Object -First 1).IPAddress

$MAC = (Get-NetAdapter | Where-Object Status -eq 'Up' | Select-Object -First 1).MacAddress
$Windows = (Get-CimInstance Win32_OperatingSystem).Caption

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
}

$json = $data | ConvertTo-Json

try {
    $response = Invoke-RestMethod -Uri $WebAppUrl -Method Post -Body $json -ContentType "application/json"
    Write-Host "✅ Upload thành công! Dòng mới đã được thêm vào Google Sheets." -ForegroundColor Green
} catch {
    Write-Host "❌ Lỗi: $($_.Exception.Message)" -ForegroundColor Red
}