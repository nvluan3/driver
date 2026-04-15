powershell -Command "Get-ExecutionPolicy -List"
powershell -Command "Start-Process PowerShell -ArgumentList 'Set-ExecutionPolicy Unrestricted -Force' -Verb RunAs"
powershell -Command "Get-ExecutionPolicy -List"
pause