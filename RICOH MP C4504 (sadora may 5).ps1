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

# Đường dẫn driver trên GitHub (raw link đến file ZIP)
$githubDriverUrl = "https://github.com/nvluan3/driver/raw/main/MP%20C4504ex%20-%20C6004ex%20series.zip"
$tempDir = "$env:TEMP\PrinterDriver"
$driverFolder = Join-Path $tempDir "MP C4504ex - C6004ex series"

# Nếu chưa có thư mục driver thì tải về và giải nén
if (-not (Test-Path $driverFolder)) {
    if (-not (Test-Path $tempDir)) { 
        New-Item -ItemType Directory -Path $tempDir | Out-Null 
    }
    $zipPath = Join-Path $tempDir "driver.zip"
    # Write-Host "Đang tải driver từ GitHub..."
    # Invoke-WebRequest -Uri $githubDriverUrl -OutFile $zipPath
    # Kiểm tra nếu file chưa tồn tại thì tải về, nếu đã tồn tại thì bỏ qua
    Write-Host "Đang tải driver từ GitHub bằng Start-BitsTransfer..."
    Start-BitsTransfer -Source $githubDriverUrl -Destination $zipPath
    Write-Host "Đang giải nén driver..."
    Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force
    if (Test-Path $zipPath) {
        Remove-Item $zipPath -Force
        Write-Host "Đã xóa file driver.zip sau khi giải nén."
    }
} else {
    Write-Host "Đã có thư mục driver, bỏ qua bước tải."
}

# Đường dẫn đến file INF
$driverPath1 = Join-Path $driverFolder "oemsetup.inf"
Write-Host "Đường dẫn đến file INF: $driverPath1"

# Thông tin máy in
$driverName1 = "RICOH MP C4504 PCL 6"
$printer1Name = "RICOH MP C4504 (sadora may 5)"
$port1Name = "IP_10.10.109.22"
$port1Address = "10.10.109.22"
$paperSize = "A4"

# Cài driver từ INF bằng pnputil
# Nếu driver đã tồn tại thì bỏ qua bước cài đặt
$existingDriver = Get-PrinterDriver -Name $driverName1 -ErrorAction SilentlyContinue
if ($existingDriver) {
    Write-Host "Driver $driverName1 đã tồn tại, bỏ qua bước cài đặt."
} else {
    Write-Host "Đang cài đặt driver $driverName1..."
    pnputil /add-driver $driverPath1 /install
    Add-PrinterDriver -Name $driverName1
}

# Tạo cổng nếu chưa có
if (-not (Get-PrinterPort -Name $port1Name -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name $port1Name -PrinterHostAddress $port1Address
}

# Thêm máy in nếu chưa có
if (-not (Get-Printer -Name $printer1Name -ErrorAction SilentlyContinue)) {
    Add-Printer -Name $printer1Name -PortName $port1Name -DriverName $driverName1
} else {
    Write-Host "Máy in $printer1Name đã tồn tại, bỏ qua bước thêm."
}

# Cấu hình khổ giấy
Set-PrintConfiguration -PrinterName $printer1Name -PaperSize $paperSize
Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
Pause
