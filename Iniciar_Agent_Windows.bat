@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo ERROR: no s'ha trobat l'entorn virtual .venv.
    echo Executa primer instalar_windows.bat abans de continuar.
    echo.
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: no s'ha pogut activar l'entorn virtual .venv.
    pause
    exit /b 1
)

streamlit run app.py
pause
