@echo off
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo.
  echo [LOI] Chua tim thay Node.js tren may nay.
  echo Vui long cai Node.js truoc: https://nodejs.org (chon ban LTS), roi chay lai file nay.
  echo.
  pause
  exit /b 1
)

if not exist "node_modules" (
  echo Dang cai dat lan dau, vui long doi mot chut...
  call npm install
)

echo.
echo Dang khoi dong K^&H Bank Tracker...
echo Sau khi thay dong chu "dang chay tai http://localhost:3000", mo trinh duyet vao dia chi do.
echo.

start "" http://localhost:3000
call npm start

pause
