@echo off
chcp 65001 >nul
setlocal

REM ============================================================
REM  Presupuestador PONTEM - Interfaz web (Streamlit)
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

echo.
echo Iniciando Presupuestador PONTEM (web)...
echo Se abrira en tu navegador en http://localhost:8501
echo Para detener, cierra esta ventana o presiona Ctrl+C.
echo.

call venv\Scripts\activate.bat
streamlit run app.py

endlocal
