@echo off
title Conciliacion Telefonica ENTEL - PRODAC
cd /d "%~dp0"
echo =============================================================
echo   Iniciando Modulo Entel - Conciliacion Telefonica PRODAC
echo =============================================================
echo Verificando dependencias de Python...
python -c "import openpyxl, pypdf" 2>nul
if %errorlevel% neq 0 (
    echo Instalando dependencias requeridas...
    pip install openpyxl pypdf msoffcrypto-tool
)
echo Abriendo aplicacion en http://localhost:8080/entel.html ...
start http://localhost:8080/entel.html
python server.py
pause
