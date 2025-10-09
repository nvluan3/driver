# Đặt đường dẫn đến thư mục script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Đường dẫn driver chung
$driverPath = Join-Path $scriptDir "IM C4500-C6000\MPC4500_.inf"
$driverName = "RICOH IM C4500 PCL 6"

# Thông tin máy in và cổng
$printer1Name = "RICOH IM C4500 (may 1)"
$printer2Name = "RICOH IM C4500 (may 2)"
$port1Name = "IP_10.10.81.21"
$port2Name = "IP_10.10.80.21"
$port1Address = "10.10.81.21"
$port2Address = "10.10.80.21"
$paperSize = "A4"

# Cài đặt driver
pnputil /add-driver $driverPath /install
Add-PrinterDriver -Name $driverName

# Máy in 1
if (-not (Get-PrinterPort -Name $port1Name -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name $port1Name -PrinterHostAddress $port1Address
}
Add-Printer -Name $printer1Name -PortName $port1Name -DriverName $driverName
Set-PrintConfiguration -PrinterName $printer1Name -PaperSize $paperSize

# Máy in 2
if (-not (Get-PrinterPort -Name $port2Name -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name $port2Name -PrinterHostAddress $port2Address
}
Add-Printer -Name $printer2Name -PortName $port2Name -DriverName $driverName
Set-PrintConfiguration -PrinterName $printer2Name -PaperSize $paperSize

Pause
