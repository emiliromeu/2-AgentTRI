@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"

echo ============================================
echo   Agent TRIMESTRE -- Instal.lacio (Windows)
echo ============================================
echo.
echo Aquest script es fa NOMES UNA VEGADA per PC.
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: no s'ha trobat Python en aquest PC.
    echo.
    echo Cal instal.lar Python des de https://python.org
    echo IMPORTANT: durant la instal.lacio, marca la casella
    echo "Add Python to PATH" abans de prémer "Install Now".
    echo.
    echo Quan estigui instal.lat, torna a executar aquest arxiu
    echo ^(instalar_windows.bat^).
    echo.
    pause
    exit /b 1
)

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: s'ha trobat un "python" en aquest PC pero no respon
    echo correctament ^(pot ser un enllaç trencat de la Microsoft Store^).
    echo.
    echo Cal instal.lar Python des de https://python.org
    echo IMPORTANT: marca "Add Python to PATH" durant la instal.lacio.
    echo.
    pause
    exit /b 1
)

echo Python trobat:
python --version
echo.

if not exist ".venv" (
    echo Creant l'entorn virtual .venv ...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: no s'ha pogut crear l'entorn virtual .venv.
        echo Revisa el missatge d'error de dalt.
        pause
        exit /b 1
    )
    echo Entorn virtual creat.
) else (
    echo L'entorn virtual .venv ja existeix -- no es torna a crear.
)
echo.

call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: no s'ha pogut activar l'entorn virtual .venv.
    pause
    exit /b 1
)

echo Actualitzant pip ...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo ERROR: no s'ha pogut actualitzar pip.
    echo Revisa el missatge d'error de dalt i la connexio a internet.
    pause
    exit /b 1
)
echo.

echo Instal.lant les dependencies ^(requirements.txt^) -- pot trigar
echo uns minuts la primera vegada ...
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: la instal.lacio de dependencies ha fallat.
    echo Revisa el missatge d'error de dalt i la connexio a internet.
    pause
    exit /b 1
)
echo.

if not exist ".env" (
    echo Creant una plantilla .env ...
    (
        echo ANTHROPIC_API_KEY=
    ) > .env
    echo.
    echo S'ha creat un arxiu .env buit a la carpeta del programa.
    echo IMPORTANT: obre .env amb el Bloc de notes i enganxa la clau
    echo API d'AQUEST ordinador despres del signe "=", sense espais
    echo ni cometes. Mai copiis la clau d'un altre ordinador.
) else (
    echo Ja existeix un arxiu .env -- no es toca.
)
echo.

echo ============================================
echo   Instal.lacio completada correctament.
echo.
echo   Segueix aquests passos abans d'arrencar:
echo   1. Si no ho has fet, edita .env amb la clau API.
echo   2. Fes doble clic a Iniciar_Agent_Windows.bat
echo ============================================
echo.
pause
