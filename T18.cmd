set "usbPath=%~dp0\T18.ps1"
powershell -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%usbPath%\"' -Verb RunAs"
pause