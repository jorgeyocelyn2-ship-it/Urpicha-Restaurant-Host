@echo off
setlocal
cd /d "%~dp0"
title Cuponera Empresa v10

if not exist "uploads" mkdir "uploads"
if not exist "outputs" mkdir "outputs"

echo Actualizando instalador...
py -m pip install --upgrade pip setuptools wheel
if errorlevel 1 python -m pip install --upgrade pip setuptools wheel

echo Instalando dependencias...
py -m pip install --upgrade -r requirements.txt
if errorlevel 1 python -m pip install --upgrade -r requirements.txt

if errorlevel 1 (
 echo ERROR: No se pudieron instalar las dependencias.
 pause
 exit /b 1
)

where tesseract >nul 2>nul
if errorlevel 1 (
 if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" goto ok
 if exist "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" goto ok
 echo Tesseract no esta instalado. Ejecuta instalar_ocr.bat.
 pause
 exit /b 1
)

:ok
start "" http://127.0.0.1:5050
py app.py
if errorlevel 1 python app.py
pause
