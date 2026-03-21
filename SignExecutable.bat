@echo off
REM Script batch per firmare FatturaView.exe
REM Bypassa automaticamente l'Execution Policy di PowerShell

echo ========================================
echo   FatturaView - Firma Eseguibile
echo   Wrapper Batch (Bypass Policy)
echo ========================================
echo.

REM Verifica che lo script PowerShell esista
if not exist "%~dp0SignExecutable.ps1" (
    echo ERRORE: SignExecutable.ps1 non trovato!
    echo Percorso atteso: %~dp0SignExecutable.ps1
    echo.
    pause
    exit /b 1
)

echo Esecuzione script PowerShell con Bypass policy...
echo.

REM Esegui lo script PowerShell con bypass della execution policy
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0SignExecutable.ps1" %*

REM Verifica il risultato
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo   OPERAZIONE COMPLETATA CON SUCCESSO
    echo ========================================
) else (
    echo.
    echo ========================================
    echo   ERRORE - Codice: %ERRORLEVEL%
    echo ========================================
    echo.
    echo Controlla i messaggi sopra per dettagli.
)

echo.
pause
