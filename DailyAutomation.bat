@echo off
setlocal
cd /d "C:\Users\yxu\OneDrive - Reconnect Community Health Services\Desktop\Report Pulling App"

if exist "main.py" (
    py "main.py"
    if errorlevel 1 python "main.py"
) else (
    echo Could not find main.py in:
    echo %cd%
)

echo.
echo Press any key to close...
pause >nul