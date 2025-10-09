$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$driverPath = Join-Path $scriptDir "MP 2554-3054-3554-4054-5054-6054 series\MP_2554_.inf"
$driverName = "RICOH MP 3054 PCL 6"
$portName = "IP_10.10.108.52"
$printerName = "T2-RICOH3054-02"
$printerIP = "10.10.108.52"
pnputil /add-driver $driverPath /install
Add-PrinterDriver -Name $driverName
if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name $portName -PrinterHostAddress $printerIP
}
Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
Set-PrintConfiguration -PrinterName $printerName -PaperSize "A4"
Invoke-Expression "rundll32 printui.dll,PrintUIEntry /e /n`"$printerName`""
Pause