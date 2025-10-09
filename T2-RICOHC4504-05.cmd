set "usbPath=%~dp0\T2-RICOH6001-05.ps1"
powershell -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%usbPath%\"' -Verb RunAs"