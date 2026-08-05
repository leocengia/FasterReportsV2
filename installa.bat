@echo off
REM ============================================================================
REM  Installazione di Omni Report — Windows.
REM
REM  Uso: doppio clic. Va fatto UNA VOLTA per PC, prima di usare run_report.bat.
REM
REM  Cosa fa:
REM    1. Controlla se c'e' un Python vero installato.
REM    2. Se non c'e', prova a installarlo da solo (con winget).
REM    3. Installa le librerie che il programma usa (pyyaml, xlwings).
REM
REM  Se qualcosa non va, questa finestra lo dice in chiaro: leggi il messaggio,
REM  non chiuderla subito. Se resti bloccato, manda uno screenshot di questa
REM  finestra a chi ti ha dato il programma.
REM ============================================================================
setlocal
cd /d "%~dp0"

echo ============================================================
echo   Installazione Omni Report
echo ============================================================
echo.

REM --- 1. C'e' gia' un Python vero? -----------------------------------------
REM Il launcher "py" e' quello messo dall'installer ufficiale di python.org.
REM Windows NON lo sostituisce mai con lo stub del Microsoft Store (quello
REM sostituisce solo "python.exe" e "python3.exe"), quindi e' il modo piu'
REM sicuro per capire se Python c'e' davvero o se "python" e' solo l'alias
REM vuoto che manda al Microsoft Store.
set "PYEXE="

py -3 --version >nul 2>&1
if not errorlevel 1 (
    set "PYEXE=py -3"
)

if "%PYEXE%"=="" (
    python --version 2>nul | findstr /r "^Python 3" >nul
    if not errorlevel 1 (
        set "PYEXE=python"
    )
)

if not "%PYEXE%"=="" goto python_ok

REM --- 2. Python non c'e': si prova a installarlo da solo -------------------
echo Su questo PC non c'e' un Python vero (solo il rimando al Microsoft Store).
echo.

where winget >nul 2>&1
if errorlevel 1 (
    echo Questo PC non ha "winget", quindi non posso installare Python da solo.
    echo.
    echo Vai su https://www.python.org/downloads/ e scarica Python.
    echo IMPORTANTE: nella prima schermata dell'installer, spunta la casella
    echo   "Add python.exe to PATH"
    echo prima di premere Install. Senza quella spunta il problema si ripete.
    echo.
    echo Dopo l'installazione, chiudi questa finestra e rilancia installa.bat.
    pause
    exit /b 1
)

echo Provo a installare Python con winget. Si apriranno delle finestre:
echo accetta le richieste che compaiono. Puo' volerci un minuto o due.
echo.
winget install -e --id Python.Python.3.12 --source winget --scope machine --accept-source-agreements --accept-package-agreements
if errorlevel 1 (
    echo.
    echo L'installazione con i permessi di amministratore non e' riuscita.
    echo Riprovo senza quei permessi...
    winget install -e --id Python.Python.3.12 --source winget --scope user --accept-source-agreements --accept-package-agreements
)
if errorlevel 1 (
    echo.
    echo L'installazione automatica non e' riuscita.
    echo Vai su https://www.python.org/downloads/ e installalo a mano.
    echo IMPORTANTE: spunta "Add python.exe to PATH" nell'installer.
    echo Poi rilancia installa.bat.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo Python e' stato installato.
echo IMPORTANTE: chiudi questa finestra e rilancia installa.bat
echo (Windows deve aggiornare il PATH: nella stessa finestra non si vede).
echo ============================================================
pause
exit /b 0

:python_ok
echo Trovato un Python funzionante:
%PYEXE% --version
echo.

REM --- 3. Le librerie del progetto -------------------------------------------
echo Installo le librerie necessarie...
echo.
%PYEXE% -m pip install --upgrade pip
%PYEXE% -m pip install -e ".[excel]"
if errorlevel 1 (
    echo.
    echo ============================================================
    echo L'installazione delle librerie e' FALLITA.
    echo Copia tutto il testo di questa finestra (tasto destro sulla barra
    echo del titolo, o seleziona con il mouse e Ctrl+C) e mandalo a chi ti
    echo ha dato il programma.
    echo ============================================================
    pause
    exit /b 1
)

echo.
echo === Verifica finale ===
%PYEXE% -m fasterreports.omni.cli check --no-excel
echo.
echo ============================================================
echo Fatto. Da ora in poi usa run_report.bat per generare il report.
echo ============================================================
pause
