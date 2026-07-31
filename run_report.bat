@echo off
REM ============================================================================
REM  "Il pulsante" — Windows.
REM
REM  Uso:  trascinare qui nulla, basta un doppio clic. Chiede la settimana.
REM        Oppure da riga di comando:  run_report.bat 31
REM
REM  Prima volta, in questa cartella:
REM      python -m pip install -e ".[excel]"
REM  Serve Excel desktop: il ricalcolo delle formule ad array 365 e del VBA non
REM  e' replicabile altrove.
REM ============================================================================
setlocal
cd /d "%~dp0"

set WEEK=%1
if "%WEEK%"=="" set /p WEEK=Numero settimana (es. 31):
if "%WEEK%"=="" (
    echo Nessuna settimana indicata. Esco.
    pause
    exit /b 2
)

echo.
echo === Controllo dei CSV in input\ ===
python -m fasterreports.omni.cli preflight --week %WEEK%
if errorlevel 1 (
    echo.
    echo I CSV non sono a posto: vedi i punti BLOCCATO qui sopra.
    echo Nessun workbook e' stato prodotto.
    pause
    exit /b 1
)

echo.
echo === Generazione del report ===
python -m fasterreports.omni.cli build --week %WEEK%
if errorlevel 1 (
    echo.
    echo Generazione fallita.
    pause
    exit /b 1
)

echo.
echo Fatto. Il file e' in output\
pause
