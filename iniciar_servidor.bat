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
    python -m pip install -r requirements.txt
)
echo.
echo Modulos disponibles en el servidor:
echo   - Entel:    http://localhost:8080/entel
echo   - Movistar: http://localhost:8080/movistar
echo   - Claro:    http://localhost:8080/claro
echo.
echo Abriendo aplicacion en http://localhost:8080/movistar ...
start http://localhost:8080/movistar
python server.py
pause
