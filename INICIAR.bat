@echo off
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel% neq 0 (
  echo No se encontro Python. Instale Python 3 desde python.org y marque "Add Python to PATH".
  pause
  exit /b 1
)
where tesseract >nul 2>nul
if %errorlevel% neq 0 (
  echo.
  echo AVISO: Tesseract OCR no esta instalado o no esta en PATH.
  echo El panel de pedidos abrira, pero los scanners TALMA y POLICIA no podran leer imagenes.
  echo Ejecuta INSTALAR_OCR_WINDOWS.bat para ver el instalador recomendado.
  echo.
)
python -m pip install -r requirements.txt
start "" http://localhost:8080/admin
python app.py
pause
