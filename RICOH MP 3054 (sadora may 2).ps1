# Kiểm tra xem script có đang chạy với quyền admin không
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if (-not $IsAdmin) {
    Write-Host "Script chưa chạy với quyền admin. Đang khởi động lại..."
    $scriptPath = $MyInvocation.MyCommand.Definition
    Start-Process powershell -ArgumentList "-File `"$scriptPath`"" -Verb RunAs
    exit
}

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$nameFolder = "MP 2554-3054-3554-4054-5054-6054 series"
# Thư mục driver gốc (nằm cùng chỗ với script)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$sourceDriverFolder = Join-Path $scriptDir $nameFolder

# Thư mục tạm
$tempDir = "$env:TEMP\PrinterDriver"
$driverFolder = Join-Path $tempDir $nameFolder
# Nếu chưa có thư mục driver trong temp thì copy từ thư mục chạy script
if (-not (Test-Path $driverFolder)) {
    if (-not (Test-Path $tempDir)) {
        New-Item -ItemType Directory -Path $tempDir | Out-Null
    }
    Write-Host "Đang copy driver từ thư mục $sourceDriverFolder sang thư mục tạm $driverFolder..."
    Copy-Item -Path $sourceDriverFolder -Destination $driverFolder -Recurse -Force
} else {
    Write-Host "Đã có thư mục driver trong temp, bỏ qua bước copy."
}

# Đường dẫn đến file INF
$driverPath = Join-Path $driverFolder "MP_2554_.INF"
Write-Host "Đường dẫn đến file INF: $driverPath"

# Thông tin máy in
$driverName = "RICOH MP 3054 PCL 6"
$printerName = "RICOH MP 3054 (sadora may 2)"
$portName = "IP_10.10.108.52"
$portAddress = "10.10.108.52"
$paperSize = "A4"

# Cài driver từ INF bằng pnputil
pnputil /add-driver $driverPath /install
Add-PrinterDriver -Name $driverName

# Tạo cổng nếu chưa có
if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
}

# Thêm máy in nếu chưa có
if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
    Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
} else {
    Write-Host "Máy in $printerName đã tồn tại, bỏ qua bước thêm."
}

# Cấu hình khổ giấy
Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize
Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
Pause
