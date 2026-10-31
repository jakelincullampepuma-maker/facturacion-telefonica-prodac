@echo off
title Facturacion Telefonica - PRODAC
cd /d "%~dp0"
echo =============================================================
echo   Iniciando Servidor de Conciliacion Telefonica - PRODAC
echo =============================================================
echo Verificando dependencias de Python...
python -c "import msoffcrypto, openpyxl, pypdf" 2>nul
if %errorlevel% neq 0 (
    echo Instalando dependencias requeridas...
    pip install msoffcrypto-tool openpyxl pypdf
)
echo.
echo Modulos disponibles en el servidor:
echo   - Entel:    http://localhost:8080/entel.html
echo   - Movistar: http://localhost:8080/movistar.html
echo   - Claro:    http://localhost:8080/claro.html
echo.
echo Abriendo aplicacion en http://localhost:8080/entel.html ...
start http://localhost:8080/entel.html
python server.py
pause
