@echo off
REM Requereix Python i entorn instal·lats en aquest PC; veure README.
REM Pendent de provar en Windows real.
cd /d "%~dp0"
call .venv\Scripts\activate.bat
streamlit run app.py
pause
