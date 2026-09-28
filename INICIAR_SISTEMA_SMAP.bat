@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>&1
if %errorlevel%==0 (set PY=py) else (set PY=python)
%PY% -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
  echo.
  echo No se encontro openpyxl. Instalando dependencia...
  %PY% -m pip install openpyxl
)
start "SISTEMA SMAP 2026" http://127.0.0.1:8876
%PY% server.py
pause
