[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 > $null
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$driverPath = Join-Path $scriptDir "MP C4504ex - C6004ex series\oemsetup.inf"
# $driverPath = Join-Path $scriptDir "MP C4504ex - C6004ex series\disk1\MPC4504e.inf"
$driverName = "RICOH MP C4504ex PCL 6"
$portName = "IP_10.18.10.10"
$portName2 = "IP_10.18.10.11"
$printerName = "RICOH MP C4504ex-01 (Sarimi)"
$printerName2 = "RICOH MP C4504ex-02 (Sarimi)"
$printerIP = "10.18.10.10"
$printerIP2 = "10.18.10.11"
$paperSize = "A4"
# Kiem tra xem driver da ton tai chua
$existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
if ($existingDriver) {
    Write-Host "Driver $driverName đã tồn tại, bỏ qua bước cài đặt."
} 
else {
    Write-Host "Đang cài đặt driver $driverName..."
    pnputil /add-driver $driverPath /install
    Add-PrinterDriver -Name $driverName
}
if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name $portName -PrinterHostAddress $printerIP
}
if (-not (Get-PrinterPort -Name $portName2 -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name $portName2 -PrinterHostAddress $printerIP2
}
if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
    Write-Host "Đang cài đặt máy in $printerName..."
    Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
}
if (-not (Get-Printer -Name $printerName2 -ErrorAction SilentlyContinue)) {
    Write-Host "Đang cài đặt máy in $printerName2..."
    Add-Printer -Name $printerName2 -PortName $portName2 -DriverName $driverName
}
Write-Host "Đang cấu hình máy in $printerName với khổ giấy $paperSize..."
Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize -Color $false
Write-Host "Đang cấu hình máy in $printerName2 với khổ giấy $paperSize..."
Set-PrintConfiguration -PrinterName $printerName2 -PaperSize $paperSize -Color $false
# Đặt máy in mặc định
Write-Host "Đang đặt $printerName làm máy in mặc định..."
$printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'"
Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
write-Host "Đã đặt $printerName làm máy in mặc định."
Write-Host "Đang mở thư mục Devices and Printers..."
Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
Pause