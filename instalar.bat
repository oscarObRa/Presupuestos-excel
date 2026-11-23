@echo off
chcp 65001 >nul
setlocal

REM ============================================================
REM  Presupuestador PONTEM - Instalacion inicial
REM  Crea el entorno virtual, instala dependencias y la plantilla
REM ============================================================

cd /d "%~dp0"

echo ============================================================
echo   INSTALACION - Presupuestador PONTEM
echo ============================================================
echo.

if not exist "venv\Scripts\activate.bat" (
    echo [1/3] Creando entorno virtual...
    python -m venv venv
    if errorlevel 1 (
        echo [ERROR] No se pudo crear el entorno virtual. Verifica que Python este instalado.
        pause
        exit /b 1
    )
) else (
    echo [1/3] El entorno virtual ya existe.
)

echo [2/3] Instalando dependencias...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Fallo la instalacion de dependencias.
    pause
    exit /b 1
)

echo [3/3] Generando plantilla base...
python generar_plantilla.py

echo.
echo ============================================================
echo   INSTALACION COMPLETA
echo ============================================================
echo   - Interfaz web:  iniciar_streamlit.bat
echo   - Linea de comandos:  presupuestar.bat
echo.
pause
endlocal
