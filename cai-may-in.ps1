$OutputEncoding = [System.Text.Encoding]::UTF8
# Kiem tra xem script co dang chay voi quyen admin khong
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if (-not $IsAdmin) {
    Write-Host "Script chưa chạy với quyền admin. Đang khởi động lại..."
    
    $scriptPath = $MyInvocation.MyCommand.Definition
    
    Start-Process powershell -ArgumentList "-File `"$scriptPath`"" -Verb RunAs
    
    exit
}

# Hien thi menu
function Show-Menu {
    Clear-Host
    Write-Host "========================================================="
    Write-Host "                      MENU LUA CHON                      "
    Write-Host "========================================================="
    Write-Host "1. Cài đặt máy in RICOH IM C4500 (SADORA MAY 1 ID)"
    Write-Host "2. Cài đặt máy in RICOH Aficio MP MP 3054 (SADORA MAY 2)"
    Write-Host "3. Cài đặt máy in RICOH Aficio MP 9002 (SADORA MAY 3)"
    Write-Host "4. Cài đặt máy in RICOH Aficio MP 9002 (SADORA MAY 4)"
    Write-Host "5. Cài đặt máy in RICOH MP C4504 (SADORA MAY 5)"
    Write-Host "6. Cài đặt máy in RICOH Aficio MP 6001 (SADORA MAY 6)"
    Write-Host "7. Cài đặt máy in RICOH IM C4500 (THACO AGRI - IN MAU)"
    Write-Host "8. Cài đặt máy in RICOH MP 7503 (THACO AGRI - TRANG DEN)"
    Write-Host "9. Cài đặt máy in RICOH IM C4500 (TANG 18 MAY 1)"
    Write-Host "10. Cài đặt máy in RICOH IM C4500 (TANG 18 MAY 2)"
    Write-Host "11. Cài đặt máy in RICOH Aficio MP 6002 (tai chinh)"
    Write-Host "12. Cài đặt máy in RICOH MP C4504ex (mau tang 19)"
    Write-Host "13. Mở Device and Printers"
    Write-Host "14. Kiểm tra máy in mặc định"
    Write-Host "0. Thoát chương trình"
    Write-Host "========================================================="
}

# Vong lap menu
do {
    Show-Menu
    $choice = Read-Host "Nhập lựa chọn (1-14 hoặc 0 để thoát)"
    switch ($choice) {
        "1" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath = Join-Path $scriptDir "IM C4500-C6000\MPC4500_.inf"
            $driverName = "RICOH IM C4500 PCL 6"
            $printerName = "RICOH IM C4500 (sadora may 1 ID)"
            $portName = "IP_10.10.110.95"
            $portAddress = "10.10.110.95"
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
                Write-Host "Đang tạo cổng máy in $portName với địa chỉ $portAddress..."
                Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
            }
            if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                Write-Host "Đang cài đặt máy in $printerName..."
                Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            }
            Write-Host "Đang cấu hình máy in $printerName với khổ giấy $paperSize và trắng đen..."
            Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize -Color $false
            # Đặt máy in mặc định
            Write-Host "Đang đặt $printerName làm máy in mặc định..."
            $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            Write-Host "Đã đặt $printerName làm máy in mặc định."
            Write-Host "Đã thực hiện xong cài đặt và cấu hình cho $printerName."
            Write-Host "Đang mở thư mục Devices and Printers..."
            Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
            Pause
        }
        "2" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath = Join-Path $scriptDir "MP 2554-3054-3554-4054-5054-6054 series\disk1\MP_2554_.INF"
            $driverName = "RICOH MP 3054 PCL 6"
            $printerName = "RICOH MP 3054 (sadora may 2)"
            $portName = "IP_10.10.108.52"
            $portAddress = "10.10.108.52"
            $paperSize = "A4"
            # Kiem tra xem driver da ton tai chua
            $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
            if ($existingDriver) {
                Write-Host "Driver $driverName đã tồn tại, bỏ qua bước cài đặt."
            } else {
                Write-Host "Đang cài đặt driver $driverName..."
                pnputil /add-driver $driverPath /install
                Add-PrinterDriver -Name $driverName
            }
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
            }
            if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                Write-Host "Đang cài đặt máy in $printerName..."
                Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            }
            Write-Host "Đang cấu hình máy in $printerName với khổ giấy $paperSize..."
            Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize -Color $false
            # Đặt máy in mặc định
            Write-Host "Đang đặt $printerName làm máy in mặc định..."
            $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            Write-Host "Đã đặt $printerName làm máy in mặc định."
            Write-Host "Đã thực hiện xong cài đặt và cấu hình cho $printerName."
            Write-Host "Đang mở thư mục Devices and Printers..."
            Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
            Pause
        }
        "3" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath = Join-Path $scriptDir "Aficio MP 6002-7502-9002\OEMSETUP.INF"
            $driverName = "RICOH Aficio MP 9002 PCL 6"
            $printerName = "RICOH Aficio MP 9002 (sadora may 3)"
            $portName = "IP_10.10.113.34"
            $portAddress = "10.10.113.34"
            $paperSize = "A4"
            # Kiem tra xem driver da ton tai chua
            $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
            if ($existingDriver) {
                Write-Host "Driver $driverName đã tồn tại, bỏ qua bước cài đặt."
            } else {
                Write-Host "Đang cài đặt driver $driverName..."
                pnputil /add-driver $driverPath /install
                Add-PrinterDriver -Name $driverName
            }
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
            }
            if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                Write-Host "Đang cài đặt máy in $printerName..."
                Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            }
            Write-Host "Đang cấu hình máy in $printerName với khổ giấy $paperSize..."
            Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize -Color $false
            # Đặt máy in mặc định
            Write-Host "Đang đặt $printerName làm máy in mặc định..."
            $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            write-Host "Đã đặt $printerName làm máy in mặc định."
            Write-Host "Đã thực hiện xong cài đặt và cấu hình cho $printerName."
            Write-Host "Đang mở thư mục Devices and Printers..."
            Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
            Pause
        }
        "4" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath = Join-Path $scriptDir "Aficio MP 6002-7502-9002\OEMSETUP.INF"
            $driverName = "RICOH Aficio MP 9002 PCL 6"
            $printerName = "RICOH Aficio MP 9002 (sadora may 4)"
            $portName = "IP_10.10.109.41"
            $portAddress = "10.10.109.41"
            $paperSize = "A4"
            # Kiem tra xem driver da ton tai chua
            $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
            if ($existingDriver) {
                Write-Host "Driver $driverName đã tồn tại, bỏ qua bước cài đặt."
            } else {
                Write-Host "Đang cài đặt driver $driverName..."
                pnputil /add-driver $driverPath /install
                Add-PrinterDriver -Name $driverName
            }
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
            }
            if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                Write-Host "Đang cài đặt máy in $printerName..."
                Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            }
            Write-Host "Đang cấu hình máy in $printerName với khổ giấy $paperSize..."
            Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize -Color $false
            # Đặt máy in mặc định
            Write-Host "Đang đặt $printerName làm máy in mặc định..."
            $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            write-Host "Đã đặt $printerName làm máy in mặc định."
            Write-Host "Đã thực hiện xong cài đặt và cấu hình cho $printerName."
            Write-Host "Đang mở thư mục Devices and Printers..."
            Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
            Pause
        }
        "5" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath = Join-Path $scriptDir "MP C4504-C6004 series\MPC4504_.inf"
            $driverName = "RICOH MP C4504 PCL 6"
            $printerName = "RICOH MP C4504 (sadora may 5)"
            $portName = "IP_10.10.109.22"
            $portAddress = "10.10.109.22"
            $paperSize = "A4"
            # Kiem tra xem driver da ton tai chua
            $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
            if ($existingDriver) {
                Write-Host "Driver $driverName đã tồn tại, bỏ qua bước cài đặt."
            } else {
                Write-Host "Đang cài đặt driver $driverName..."
                pnputil /add-driver $driverPath /install
                Add-PrinterDriver -Name $driverName
            }
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
            }
            if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                Write-Host "Đang cài đặt máy in $printerName..."
                Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            }
            Write-Host "Đang cấu hình máy in $printerName với khổ giấy $paperSize..."
            Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize -Color $false
            # Đặt máy in mặc định
            Write-Host "Đang đặt $printerName làm máy in mặc định..."
            $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            write-Host "Đã đặt $printerName làm máy in mặc định."
            Write-Host "Đã thực hiện xong cài đặt và cấu hình cho $printerName."
            Write-Host "Đang mở thư mục Devices and Printers..."
            Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
            Pause
        }
        "6" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath = Join-Path $scriptDir "Aficio MP 6001-7001-8001-9001\OEMSETUP.INF"
            $driverName = "RICOH Aficio MP 6001 PCL 6"
            $printerName = "RICOH Aficio MP 6001 (sadora may 6)"
            $portName = "IP_10.10.109.54"
            $portAddress = "10.10.109.54"
            $paperSize = "A4"
            # Kiem tra xem driver da ton tai chua
            $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
            if ($existingDriver) {
                Write-Host "Driver $driverName đã tồn tại, bỏ qua bước cài đặt."
            } else {
                Write-Host "Đang cài đặt driver $driverName..."
                pnputil /add-driver $driverPath /install
                Add-PrinterDriver -Name $driverName
            }
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
            }
            if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                Write-Host "Đang cài đặt máy in $printerName..."
                Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            }
            Write-Host "Đang cấu hình máy in $printerName với khổ giấy $paperSize..."
            Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize -Color $false
            # Đặt máy in mặc định
            Write-Host "Đang đặt $printerName làm máy in mặc định..."
            $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            write-Host "Đã đặt $printerName làm máy in mặc định."
            Write-Host "Đã thực hiện xong cài đặt và cấu hình cho $printerName."
            Write-Host "Đang mở thư mục Devices and Printers..."
            Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
            Pause
        }
        "7" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath = Join-Path $scriptDir "IM C4500-C6000\OEMSETUP.INF"
            $driverName = "RICOH IM C4500 PCL 6"
            $printerName = "RICOH IM C4500 (THACO AGRI IN MAU)"
            $portName = "IP_10.10.110.105"
            $portAddress = "10.10.110.105"
            $paperSize = "A4"
            # Kiem tra xem driver da ton tai chua
            $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
            if ($existingDriver) {
                Write-Host "Driver $driverName đã tồn tại, bỏ qua bước cài đặt."
            } else {
                Write-Host "Đang cài đặt driver $driverName..."
                pnputil /add-driver $driverPath /install
                Add-PrinterDriver -Name $driverName
            }
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
            }
            if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                Write-Host "Đang cài đặt máy in $printerName..."
                Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            }
            
            Write-Host "Đang cấu hình máy in $printerName với khổ giấy $paperSize..."
            Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize -Color $false
            # Đặt máy in mặc định
            Write-Host "Đang đặt $printerName làm máy in mặc định..."
            $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            write-Host "Đã đặt $printerName làm máy in mặc định."
            Write-Host "Đã thực hiện xong cài đặt và cấu hình cho $printerName."
            Write-Host "Đang mở thư mục Devices and Printers..."
            Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
            Pause
        }
        "8" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath = Join-Path $scriptDir "MP 6503SP-7503SP-9003SP\OEMSETUP.INF"
            $driverName = "RICOH MP 7503 PCL 6"
            $printerName = "RICOH MP 7503 (THACO AGRI TRANG DEN)"
            $portName = "IP_10.10.110.108"
            $portAddress = "10.10.110.108"
            $paperSize = "A4"
            # Kiem tra xem driver da ton tai chua
            $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
            if ($existingDriver) {
                Write-Host "Driver $driverName đã tồn tại, bỏ qua bước cài đặt."
            } else {
                Write-Host "Đang cài đặt driver $driverName..."
                pnputil /add-driver $driverPath /install
                Add-PrinterDriver -Name $driverName
            }
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
            }
            if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                Write-Host "Đang cài đặt máy in $printerName..."
                Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            }
            
            Write-Host "Đang cấu hình máy in $printerName với khổ giấy $paperSize..."
            Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize -Color $false
            # Đặt máy in mặc định
            Write-Host "Đang đặt $printerName làm máy in mặc định..."
            $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            write-Host "Đã đặt $printerName làm máy in mặc định."
            Write-Host "Đã thực hiện xong cài đặt và cấu hình cho $printerName."
            Write-Host "Đang mở thư mục Devices and Printers..."
            Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
            Pause
        }
        "9" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath = Join-Path $scriptDir "IM C4500-C6000\MPC4500_.inf"
            $driverName = "RICOH IM C4500 PCL 6"
            $printerName = "RICOH IM C4500 (Tang 18 - may 1)"
            $portName = "IP_10.10.80.21"
            $portAddress = "10.10.80.21"
            $paperSize = "A4"
            # Kiem tra xem driver da ton tai chua
            $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
            if ($existingDriver) {
                Write-Host "Driver $driverName đã tồn tại, bỏ qua bước cài đặt."
            } else {
                Write-Host "Đang cài đặt driver $driverName..."
                pnputil /add-driver $driverPath /install
                Add-PrinterDriver -Name $driverName
            }
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
            }
            if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                Write-Host "Đang cài đặt máy in $printerName..."
                Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            }
            Write-Host "Đang cấu hình máy in $printerName với khổ giấy $paperSize..."
            Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize -Color $false
            # Đặt máy in mặc định
            Write-Host "Đang đặt $printerName làm máy in mặc định..."
            $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            write-Host "Đã đặt $printerName làm máy in mặc định."
            Write-Host "Đã thực hiện xong cài đặt và cấu hình cho $printerName."
            Write-Host "Đang mở thư mục Devices and Printers..."
            Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
            Pause
        }
        "10" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath = Join-Path $scriptDir "IM C4500-C6000\MPC4500_.inf"
            $driverName = "RICOH IM C4500 PCL 6"
            $printerName = "RICOH IM C4500 (Tang 18 - may 2)"
            $portName = "IP_10.10.81.21"
            $portAddress = "10.10.81.21"
            $paperSize = "A4"
            # Kiem tra xem driver da ton tai chua
            $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
            if ($existingDriver) {
                Write-Host "Driver $driverName đã tồn tại, bỏ qua bước cài đặt."
            } else {
                Write-Host "Đang cài đặt driver $driverName..."
                pnputil /add-driver $driverPath /install
                Add-PrinterDriver -Name $driverName
            }
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
            }
            if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                Write-Host "Đang cài đặt máy in $printerName..."
                Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            }
            # Đang cấu hình máy in $printerName với khổ giấy $paperSize và in trắng đen...
            Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize -Color $false
            # Đặt máy in mặc định
            Write-Host "Đang đặt $printerName làm máy in mặc định..."
            $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            write-Host "Đã đặt $printerName làm máy in mặc định."
            Write-Host "Đã thực hiện xong cài đặt và cấu hình cho $printerName."
            Write-Host "Đang mở thư mục Devices and Printers..."
            Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
            Pause
        }
        "11" {
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
            if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                Write-Host "Đang cài đặt máy in $printerName..."
                Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
            }
            Write-Host "Đang cấu hình máy in $printerName với khổ giấy $paperSize..."
            Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize -Color $false
            # Đặt máy in mặc định
            Write-Host "Đang đặt $printerName làm máy in mặc định..."
            $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            write-Host "Đã đặt $printerName làm máy in mặc định."
            Write-Host "Đã thực hiện xong cài đặt và cấu hình cho $printerName."
            Write-Host "Đang mở thư mục Devices and Printers..."
            Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
            Pause
        }
        "12" {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $driverPath11 = Join-Path $scriptDir "MP C4504ex - C6004ex series\oemsetup.inf"
            $driverName11 = "RICOH MP C4504ex PCL 6"
            $portName = "IP_10.10.84.20"
            $printerName = "RICOH MP C4504ex (Tang 19 - in mau)"
            $printerIP = "10.10.84.20"
            # Kiem tra xem driver da ton tai chua
            $existingDriver = Get-PrinterDriver -Name $driverName11 -ErrorAction SilentlyContinue
            if ($existingDriver) {
                Write-Host "Driver $driverName11 đã tồn tại, bỏ qua bước cài đặt."
            } else {
                Write-Host "Đang cài đặt driver $driverName11..."
                pnputil /add-driver $driverPath11 /install
                Add-PrinterDriver -Name $driverName11
            }
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $printerIP
            }
            if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                Write-Host "Đang cài đặt máy in $printerName..."
                Add-Printer -Name $printerName -PortName $portName -DriverName $driverName11
            }
            Write-Host "Đang cấu hình máy in $printerName với khổ giấy $paperSize..."
            Set-PrintConfiguration -PrinterName $printerName -PaperSize $paperSize -Color $false
            # Đặt máy in mặc định
            Write-Host "Đang đặt $printerName làm máy in mặc định..."
            $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'"
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            write-Host "Đã đặt $printerName làm máy in mặc định."
            Write-Host "Đã thực hiện xong cài đặt và cấu hình cho $printerName."
            Write-Host "Đang mở thư mục Devices and Printers..."
            Start-Process "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}"
            Pause
        }

        "13" {
            Start-Process "control.exe" -ArgumentList "printers"
        }
        "14" {
            $defaultPrinter = Get-CimInstance -ClassName Win32_Printer -Filter "Default=$true"
            if ($defaultPrinter) {
                Write-Host "Máy in mặc định hiện tại là: $($defaultPrinter.Name)"
            } else {
                Write-Host "Không tìm thấy máy in mặc định."
            }
            Pause
        }
        "0" {
            Write-Host "Thoát chương trình..."
        }
        default {
            Write-Host "Lựa chọn không hợp lệ, vui lòng thử lại."
            Pause
        }
    }
} while ($choice -ne "0")


