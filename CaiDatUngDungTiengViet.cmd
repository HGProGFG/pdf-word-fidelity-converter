@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_vietnamese_app.ps1"
if errorlevel 1 (
    echo.
    echo Cai dat chua hoan tat. Xem thong bao loi o tren.
    pause
    exit /b 1
)
echo.
echo Cai dat thanh cong. Ban co the mo ung dung tu Start Menu.
pause
