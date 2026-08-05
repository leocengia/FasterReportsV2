@echo off
REM ============================================================================
REM  "Il pulsante" — Windows.
REM
REM  Uso:  trascinare qui nulla, basta un doppio clic. Chiede la settimana.
REM        Oppure da riga di comando:  run_report.bat 31
REM
REM  Prima volta, in questa cartella: doppio clic su installa.bat.
REM  Serve Excel desktop: il ricalcolo delle formule ad array 365 e del VBA non
REM  e' replicabile altrove.
REM ============================================================================
setlocal
cd /d "%~dp0"

REM --- Python e le librerie ci sono? -----------------------------------------
REM Stesso controllo di installa.bat: "py" non viene mai sostituito dallo stub
REM del Microsoft Store, quindi e' il modo affidabile per sapere se Python c'e'
REM davvero. Senza questo controllo, chi non ha installato nulla vede il
REM messaggio criptico di Windows sull'"alias di esecuzione dell'app" invece di
REM sapere semplicemente cosa fare.
set "PYEXE="
py -3 -c "import fasterreports" >nul 2>&1
if not errorlevel 1 set "PYEXE=py -3"
if "%PYEXE%"=="" (
    python -c "import fasterreports" >nul 2>&1
    if not errorlevel 1 set "PYEXE=python"
)
if "%PYEXE%"=="" (
    echo.
    echo Il programma non e' ancora installato su questo PC.
    echo Fai doppio clic su installa.bat, poi rilancia questo file.
    echo.
    pause
    exit /b 1
)

set WEEK=%1
if "%WEEK%"=="" set /p WEEK=Numero settimana (es. 31):
if "%WEEK%"=="" (
    echo Nessuna settimana indicata. Esco.
    pause
    exit /b 2
)

echo.
echo === Controllo dei CSV in input\ ===
%PYEXE% -m fasterreports.omni.cli preflight --week %WEEK%
if errorlevel 1 (
    echo.
    echo I CSV non sono a posto: vedi i punti BLOCCATO qui sopra.
    echo Nessun workbook e' stato prodotto.
    pause
    exit /b 1
)

echo.
echo === Generazione del report ===
%PYEXE% -m fasterreports.omni.cli build --week %WEEK%
if errorlevel 1 (
    echo.
    echo Generazione fallita.
    pause
    exit /b 1
)

echo.
echo Fatto. Il file e' in output\
pause
