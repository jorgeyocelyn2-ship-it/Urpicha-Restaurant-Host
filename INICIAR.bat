@echo off
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel% neq 0 (
  echo No se encontro Python. Instale Python 3 desde python.org y marque "Add Python to PATH".
  pause
  exit /b 1
)
start "" http://localhost:8080/admin
python app.py
pause
