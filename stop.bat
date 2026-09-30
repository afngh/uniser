@echo off
powershell -NoProfile -ExecutionPolicy Bypass -Command "$conn = Get-NetTCPConnection -LocalPort 8000 -State Listen -ErrorAction SilentlyContinue; if ($conn) { Stop-Process -Id $conn.OwningProcess -Force; Write-Host 'Stopped PIN Lookup server process (PID:' $conn.OwningProcess ')' -ForegroundColor Green } else { Write-Host 'PIN Lookup server is not currently running.' -ForegroundColor Yellow }"
echo.
pause
