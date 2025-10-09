# ========================
# 1. Nhập thông tin từ người dùng
# ========================
$Email = Read-Host "Nhập địa chỉ email (ví dụ: nguyenvanluan4@thaco.com.vn)"
if (-not $Email.Contains("@")) {
    Write-Host "❌ Email không hợp lệ. Thoát script." -ForegroundColor Red
    exit
}

# Tách thông tin từ email
$UserName = $Email.Split("@")[0]
$DisplayName = $UserName -replace "\.", " "  # Gợi ý: Tách nếu bạn dùng kiểu nguyen.van.luan
$PstFileName = "$Email.pst"

# ========================
# 2. Tạo thư mục lưu PST
# ========================
$pstFolder = "D:\Mail"
if (-Not (Test-Path $pstFolder)) {
    New-Item -ItemType Directory -Path $pstFolder -Force | Out-Null
    Write-Host "Đã tạo thư mục $pstFolder"
}

# ========================
# 3. Thiết lập ForcePSTPath
# ========================
$outlookVer = "16.0"  # Outlook 2016/2019/365
$regPath = "HKCU:\Software\Microsoft\Office\$outlookVer\Outlook"
Set-ItemProperty -Path $regPath -Name "ForcePSTPath" -Value $pstFolder -Force
Write-Host "Đã đặt ForcePSTPath = $pstFolder"

# ========================
# 4. Cấu hình tài khoản POP3
# ========================
$IncomingSrv  = "mail.thaco.com.vn"
$IncomingPort = 995
$OutgoingSrv  = "mail.thaco.com.vn"
$OutgoingPort = 587
$ProfileName  = "Outlook"

# Tạo khóa profile nếu chưa có
$baseReg = "HKCU:\Software\Microsoft\Office\$outlookVer\Outlook\Profiles\$ProfileName\9375CFF0413111d3B88A00104B2A6676"
if (-Not (Test-Path $baseReg)) {
    New-Item -Path $baseReg -Force | Out-Null
}

# Tạo GUID đại diện tài khoản
$guid = [guid]::NewGuid().ToString("N").ToUpper()
$accountKey = "$baseReg\$guid"
New-Item -Path $accountKey -Force | Out-Null

# Thiết lập giá trị tài khoản
Set-ItemProperty -Path $accountKey -Name "Account Name"         -Value $Email
Set-ItemProperty -Path $accountKey -Name "Email"                -Value $Email
Set-ItemProperty -Path $accountKey -Name "User Name"            -Value $UserName
Set-ItemProperty -Path $accountKey -Name "Display Name"         -Value $DisplayName
Set-ItemProperty -Path $accountKey -Name "Server"               -Value $IncomingSrv
Set-ItemProperty -Path $accountKey -Name "POP3 Port"            -Value $IncomingPort
Set-ItemProperty -Path $accountKey -Name "SMTP Server"          -Value $OutgoingSrv
Set-ItemProperty -Path $accountKey -Name "SMTP Port"            -Value $OutgoingPort
Set-ItemProperty -Path $accountKey -Name "Use SSL"              -Value 1
Set-ItemProperty -Path $accountKey -Name "SMTP Use SSL"         -Value 1
Set-ItemProperty -Path $accountKey -Name "Leave Mail On Server" -Value 1
Set-ItemProperty -Path $accountKey -Name "Remove After Days"    -Value 4

# ========================
# 5. Thông báo kết quả
# ========================
Write-Host "`n✅ Đã thêm tài khoản POP3 cho: $Email"
Write-Host "📁 PST sẽ được lưu tại: $pstFolder\$PstFileName"
Write-Host "👉 Vui lòng mở Outlook để kiểm tra cấu hình."
