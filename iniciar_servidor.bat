@echo off
title Facturacion Telefonica - PRODAC
cd /d "%~dp0"
echo =============================================================
echo   Iniciando Dashboard de Facturacion Telefonica - PRODAC
echo =============================================================
echo Verificando dependencias de Python...
python -c "import msoffcrypto, openpyxl, pypdf" 2>nul
if %errorlevel% neq 0 (
    echo Instalando dependencias requeridas...
    pip install msoffcrypto-tool openpyxl pypdf
)
echo Abriendo aplicacion en http://localhost:8080/movistar.html ...
start http://localhost:8080/movistar.html
python server.py
pause
