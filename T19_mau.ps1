$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$driverPath = Join-Path $scriptDir "MP C4504ex - C6004ex series\oemsetup.inf"
$driverName = "RICOH MP C4504ex PCL 6"
$portName = "IP_10.10.84.20"
$printerName = "RICOH MP C4504ex (mau tang 19)"
$printerIP = "10.10.84.20"
pnputil /add-driver $driverPath /install
Add-PrinterDriver -Name $driverName
if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name $portName -PrinterHostAddress $printerIP
}
Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
Set-PrintConfiguration -PrinterName $printerName -PaperSize "A4"
Invoke-Expression "rundll32 printui.dll,PrintUIEntry /e /n`"$printerName`""
Pause