# Kiem tra xem script co dang chay voi quyen admin khong
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if (-not $IsAdmin) {
    Write-Host "Script Chua Chay Voi Quyen Admin. Dang Khoi Dong Lai..."
    # Lay duong dan script hien tai
    $scriptPath = $MyInvocation.MyCommand.Definition

    # Mo lai script voi quyen admin
    Start-Process powershell -ArgumentList "-File `"$scriptPath`"" -Verb RunAs

    # Thoat tien trinh hien tai
    exit
}

# Neu da co quyen admin, chay tiep phan code ben duoi
Write-Host "Script dang chay voi quyen admin!"

# Hien thi menu
function Show-Menu {
    Clear-Host
    Write-Host "========================================================="
    Write-Host "                      MENU LUA CHON                      "
    Write-Host "========================================================="
    Write-Host "1. Cai dat may in RICOH IM C4500 (SADORA MAY 1)"
    Write-Host "2. Cai dat may in RICOH Aficio MP MP 3054 (SADORA MAY 2)"
    Write-Host "3. Cai dat may in RICOH Aficio MP 9002 (SADORA MAY 3)"
    Write-Host "4. Cai dat may in RICOH Aficio MP 9002 (SADORA MAY 4)"
    Write-Host "5. Cai dat may in RICOH MP C4504 (SADORA MAY 5)"
    Write-Host "6. Cai dat may in RICOH Aficio MP 6001 (SADORA MAY 6)"
    Write-Host "7. Cai dat may in RICOH IM C4500 (THACO AGRI - IN MAU)"
    Write-Host "8. Cai dat may in RICOH MP 7503 (THACO AGRI - TRANG DEN)"
    Write-Host "9. Cai dat may in RICOH IM C4500 (TANG 18 MAY 1 VA MAY 2)"
    Write-Host "10. Cai dat may in RICOH Aficio MP 6002 (tai chinh)"
    Write-Host "11. Cai dat may in RICOH MP C4504ex (mau tang 19)"
    Write-Host "12. Cai dat may in RICOH IM C4500 (SADORA MAY 2)"
    Write-Host "13. Mo Device and Printers"
    Write-Host "0. Thoat chuong trinh"
    Write-Host "========================================================="
}

# Vong lap menu
do {
    Show-Menu
    $choice = Read-Host "Nhap lua chon (1-13 hoac 0 de thoat)"
    switch ($choice) {
        "1" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath1 = Join-Path $scriptDir "IM C4500-C6000\OEMSETUP.INF"
            $driverName1 = "RICOH IM C4500 PCL 6"
            $printer1Name = "RICOH IM C4500 (sadora may 1)"
            $port1Name = "IP_10.10.110.113"
            $port1Address = "10.10.110.113"
            $paperSize = "A4"
            pnputil /add-driver $driverPath1 /install
            Add-PrinterDriver -Name $driverName1
            if (-not (Get-PrinterPort -Name $port1Name -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $port1Name -PrinterHostAddress $port1Address
            }
            Add-Printer -Name $printer1Name -PortName $port1Name -DriverName $driverName1
            Set-PrintConfiguration -PrinterName $printer1Name -PaperSize $paperSize
            Pause
        }
        "2" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath2 = Join-Path $scriptDir "MP 2554-3054-3554-4054-5054-6054 series\disk1\MP_2554_.INF"
            $driverName2 = "RICOH MP 3054 PCL 6"
            $printer2Name = "RICOH MP 3054 (sadora may 2)"
            $port2Name = "IP_10.10.108.52"
            $port2Address = "10.10.108.52"
            $paperSize = "A4"
            pnputil /add-driver $driverPath2 /install
            Add-PrinterDriver -Name $driverName2
            if (-not (Get-PrinterPort -Name $port2Name -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $port2Name -PrinterHostAddress $port2Address
            }
            Add-Printer -Name $printer2Name -PortName $port2Name -DriverName $driverName2
            Set-PrintConfiguration -PrinterName $printer2Name -PaperSize $paperSize
            Pause
        }
        "3" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath3 = Join-Path $scriptDir "Aficio MP 6002-7502-9002\OEMSETUP.INF"
            $driverName3 = "RICOH Aficio MP 9002 PCL 6"
            $printer3Name = "RICOH Aficio MP 9002 (sadora may 3)"
            $port3Name = "IP_10.10.113.34"
            $port3Address = "10.10.113.34"
            $paperSize = "A4"
            pnputil /add-driver $driverPath3 /install
            Add-PrinterDriver -Name $driverName3
            if (-not (Get-PrinterPort -Name $port3Name -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $port3Name -PrinterHostAddress $port3Address
            }
            Add-Printer -Name $printer3Name -PortName $port3Name -DriverName $driverName3
            Set-PrintConfiguration -PrinterName $printer3Name -PaperSize $paperSize
            Pause
        }
        "4" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath4 = Join-Path $scriptDir "Aficio MP 6002-7502-9002\OEMSETUP.INF"
            $driverName4 = "RICOH Aficio MP 9002 PCL 6"
            $printer4Name = "RICOH Aficio MP 9002 (sadora may 4)"
            $port4Name = "IP_10.10.109.41"
            $port4Address = "10.10.109.41"
            $paperSize = "A4"
            pnputil /add-driver $driverPath4 /install
            Add-PrinterDriver -Name $driverName4
            if (-not (Get-PrinterPort -Name $port4Name -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $port4Name -PrinterHostAddress $port4Address
            }
            Add-Printer -Name $printer4Name -PortName $port4Name -DriverName $driverName4
            Set-PrintConfiguration -PrinterName $printer4Name -PaperSize $paperSize
            Pause
        }
        "5" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath5 = Join-Path $scriptDir "MP C4504-C6004 series\disk1\OEMSETUP.INF"
            $driverName5 = "RICOH MP C4504 PCL 6"
            $printer5Name = "RICOH MP C4504 (sadora may 5)"
            $port5Name = "IP_10.10.109.22"
            $port5Address = "10.10.109.22"
            $paperSize = "A4"
            pnputil /add-driver $driverPath5 /install
            Add-PrinterDriver -Name $driverName5
            if (-not (Get-PrinterPort -Name $port5Name -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $port5Name -PrinterHostAddress $port5Address
            }
            Add-Printer -Name $printer5Name -PortName $port5Name -DriverName $driverName5
            Set-PrintConfiguration -PrinterName $printer5Name -PaperSize $paperSize
            Pause
        }
        "6" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath6 = Join-Path $scriptDir "Aficio MP 6001-7001-8001-9001\OEMSETUP.INF"
            $driverName6 = "RICOH Aficio MP 6001 PCL 6"
            $printer6Name = "RICOH Aficio MP 6001 (sadora may 6)"
            $port6Name = "IP_10.10.109.54"
            $port6Address = "10.10.109.54"
            $paperSize = "A4"
            pnputil /add-driver $driverPath6 /install
            Add-PrinterDriver -Name $driverName6
            if (-not (Get-PrinterPort -Name $port6Name -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $port6Name -PrinterHostAddress $port6Address
            }
            Add-Printer -Name $printer6Name -PortName $port6Name -DriverName $driverName6
            Set-PrintConfiguration -PrinterName $printer6Name -PaperSize $paperSize
            Pause
        }
        "7" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath7 = Join-Path $scriptDir "IM C4500-C6000\OEMSETUP.INF"
            $driverName7 = "RICOH IM C4500 PCL 6"
            $printer7Name = "RICOH IM C4500 (THACO AGRI IN MAU)"
            $port7Name = "IP_10.10.110.105"
            $port7Address = "10.10.110.105"
            $paperSize = "A4"
            pnputil /add-driver $driverPath7 /install
            Add-PrinterDriver -Name $driverName7
            if (-not (Get-PrinterPort -Name $port7Name -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $port7Name -PrinterHostAddress $port7Address
            }
            Add-Printer -Name $printer7Name -PortName $port7Name -DriverName $driverName7
            Set-PrintConfiguration -PrinterName $printer7Name -PaperSize $paperSize
            Pause
        }
        "8" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath8 = Join-Path $scriptDir "MP 6503SP-7503SP-9003SP\OEMSETUP.INF"
            $driverName8 = "RICOH MP 7503 PCL 6"
            $printer8Name = "RICOH MP 7503 (THACO AGRI TRANG DEN)"
            $port8Name = "IP_10.10.110.108"
            $port8Address = "10.10.110.108"
            $paperSize = "A4"
            pnputil /add-driver $driverPath8 /install
            Add-PrinterDriver -Name $driverName8
            if (-not (Get-PrinterPort -Name $port8Name -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $port8Name -PrinterHostAddress $port8Address
            }
            Add-Printer -Name $printer8Name -PortName $port8Name -DriverName $driverName8
            Set-PrintConfiguration -PrinterName $printer8Name -PaperSize $paperSize
            Pause
        }
        "9" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath1 = Join-Path $scriptDir "IM C4500-C6000\MPC4500_.inf"
            $driverName1 = "RICOH IM C4500 PCL 6"
            $printer1Name = "RICOH IM C4500 (Tang 18 - may 1)"
            $printer2Name = "RICOH IM C4500 (Tang 18 - may 2)"
            $port1Name = "IP_10.10.81.21"
            $port2Name = "IP_10.10.80.21"
            $port1Address = "10.10.81.21"
            $port2Address = "10.10.80.21"
            $paperSize = "A4"
            pnputil /add-driver $driverPath1 /install
            Add-PrinterDriver -Name $driverName1
            if (-not (Get-PrinterPort -Name $port1Name -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $port1Name -PrinterHostAddress $port1Address
            }
            Add-Printer -Name $printer1Name -PortName $port1Name -DriverName $driverName1
            Set-PrintConfiguration -PrinterName $printer1Name -PaperSize $paperSize
            Add-PrinterDriver -Name $driverName2
            if (-not (Get-PrinterPort -Name $port2Name -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $port2Name -PrinterHostAddress $port2Address
            }
            Add-Printer -Name $printer2Name -PortName $port2Name -DriverName $driverName2
            Set-PrintConfiguration -PrinterName $printer2Name -PaperSize $paperSize
            Pause
        }
        "10" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath = Join-Path $scriptDir "Aficio MP 6002-7502-9002\oemsetup.inf"
            $driverName = "RICOH Aficio MP 6002 PCL 6"
            $portName = "IP_10.10.85.36"
            $printerName = "RICOH Aficio MP 6002 (Tang 19 - tai chinh)"
            $printerIP = "10.10.85.36"
            pnputil /add-driver $driverPath /install
            Add-PrinterDriver -Name $driverName
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $printerIP
            }
            Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            Set-PrintConfiguration -PrinterName $printerName -PaperSize "A4"
            Pause
        }
        "11" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath = Join-Path $scriptDir "MP C4504ex - C6004ex series\oemsetup.inf"
            $driverName = "RICOH MP C4504ex PCL 6"
            $portName = "IP_10.10.84.20"
            $printerName = "RICOH MP C4504ex (Tang 19 - in mau)"
            $printerIP = "10.10.84.20"
            pnputil /add-driver $driverPath /install
            Add-PrinterDriver -Name $driverName
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $printerIP
            }
            Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            Set-PrintConfiguration -PrinterName $printerName -PaperSize "A4"
            Pause
        }
        "12" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath = Join-Path $scriptDir "IM C4500-C6000\MPC4500_.inf"
            $driverName = "RICOH IM C4500 PCL 6"
            $printerName = "RICOH IM C4500 (sadora may 2)"
            $portName = "IP_10.10.108.41"
            $portAddress = "10.10.108.41"
            $paperSize = "A4"
            pnputil /add-driver $driverPath /install
            Add-PrinterDriver -Name $driverName
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
            }
            Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize
            Pause
        }

        "13" {
            Start-Process "control.exe" -ArgumentList "printers"
        }

        "0" {
            Write-Host "Thoat chuong trinh..."
        }
        default {
            Write-Host "Lua chon khong hop le, vui long thu lai."
            Pause
        }
    }
} while ($choice -ne "0")


