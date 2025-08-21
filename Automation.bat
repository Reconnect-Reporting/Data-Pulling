@echo off
setlocal
set "APP_DIR=%~dp0"
set "MAIN=%APP_DIR%main.py"

REM ---- 0) (Optional) bootstrap a venv the first time ----
REM Uncomment this block if you want auto-setup.
REM if not exist "%APP_DIR%.venv" (
REM   echo Creating virtual environment...
REM   py -3 -m venv "%APP_DIR%.venv" || goto NOPY
REM   "%APP_DIR%.venv\Scripts\python.exe" -m pip install -r "%APP_DIR%requirements.txt"
REM )

REM ---- 1) Use venv if present ----
if exist "%APP_DIR%.venv\Scripts\pythonw.exe" (
  "%APP_DIR%.venv\Scripts\pythonw.exe" "%MAIN%"
  goto :eof
)

REM ---- 2) Try Python Launcher (most reliable; ships with Python on Windows) ----
where pyw >nul 2>&1
if %ERRORLEVEL%==0 (
  pyw -3 "%MAIN%"
  goto :eof
)

REM ---- 3) Try pythonw on PATH ----
where pythonw >nul 2>&1
if %ERRORLEVEL%==0 (
  pythonw "%MAIN%"
  goto :eof
)

echo.
echo Could not find Python.
echo Install from https://www.python.org/downloads/ and check "Add Python to PATH".
echo Or create a venv:  py -3 -m venv .venv
pause
goto :eof

:NOPY
echo.
echo Python launcher (py) not found. Install Python 3 first.
pause
