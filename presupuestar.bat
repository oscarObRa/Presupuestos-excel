@echo off
chcp 65001 >nul
setlocal

REM ============================================================
REM  Presupuestador PONTEM - Linea de comandos (CLI)
REM  Uso:  presupuestar.bat [ruta_al_archivo.xlsx]
REM ============================================================

cd /d "%~dp0"

if not exist "venv\Scripts\activate.bat" (
    echo [ERROR] No se encontro el entorno virtual en "venv".
    echo Ejecuta primero:  instalar.bat
    pause
    exit /b 1
)

if not exist "plantillas\machote.xlsx" (
    echo [INFO] Generando plantilla base...
    call venv\Scripts\activate.bat
    python generar_plantilla.py
)

call venv\Scripts\activate.bat
python presupuestar.py %*

echo.
pause
endlocal
