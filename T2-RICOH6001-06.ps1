$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$driverPath = Join-Path $scriptDir "Aficio MP 6001-7001-8001-9001\OEMSETUP.INF"
$driverName = "RICOH Aficio MP 6001 PCL 6"
$portName = "IP_10.10.109.54"
$printerName = "T2-RICOH6001-06"
$printerIP = "10.10.109.54"
pnputil /add-driver $driverPath /install
Add-PrinterDriver -Name $driverName
if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name $portName -PrinterHostAddress $printerIP
}
Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
Set-PrintConfiguration -PrinterName $printerName -PaperSize "A4"
Invoke-Expression "rundll32 printui.dll,PrintUIEntry /e /n`"$printerName`""
Pause