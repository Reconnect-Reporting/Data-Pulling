@echo off
setlocal

REM ----- Paths -----
set "APP_DIR=%~dp0"
set "VENV_DIR=%APP_DIR%.venv"
set "PY_VENV=%VENV_DIR%\Scripts\python.exe"
set "PYW_VENV=%VENV_DIR%\Scripts\pythonw.exe"
set "MAIN=%APP_DIR%main.py"
set "REQ=%APP_DIR%requirements.txt"

REM ----- Ensure Python (launcher) exists or a system python is available -----
where py >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
  where python >nul 2>&1
  if %ERRORLEVEL% NEQ 0 (
    echo Python not found. Please install Python 3 from https://www.python.org/downloads/ and check "Add to PATH".
    pause
    exit /b 1
  )
)

REM ----- Create venv and install deps ONLY on first run -----
if not exist "%PY_VENV%" (
  echo Creating virtual environment...
  where py >nul 2>&1
  if %ERRORLEVEL%==0 (
    py -3 -m venv "%VENV_DIR%" || (
      echo Failed to create venv. Aborting.
      pause
      exit /b 1
    )
  ) else (
    python -m venv "%VENV_DIR%" || (
      echo Failed to create venv. Aborting.
      pause
      exit /b 1
    )
  )

  echo Upgrading pip and installing requirements...
  "%PY_VENV%" -m pip install --upgrade pip setuptools wheel
  if exist "%REQ%" (
    "%PY_VENV%" -m pip install -r "%REQ%"
  ) else (
    echo No requirements.txt found. Skipping dependency install.
  )
)

REM ----- Launch (no console if GUI; fallback to console python if pythonw missing) -----
if exist "%PYW_VENV%" (
  "%PYW_VENV%" "%MAIN%"
) else (
  "%PY_VENV%" "%MAIN%"
)

endlocal
exit /b 0
