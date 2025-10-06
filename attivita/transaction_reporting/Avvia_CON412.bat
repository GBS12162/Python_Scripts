@echo off
title CON-412 Transaction Reporting System
echo.
echo ============================================================
echo   CON-412 TRANSACTION REPORTING - SISTEMA AUTOMATICO
echo ============================================================
echo   Versione: 1.0
echo   Sistema di validazione ESMA per ordini CON-412
echo ============================================================
echo.
echo Avvio del sistema in corso...
echo.

REM Cambio alla directory dell'eseguibile
cd /d "%~dp0dist"

REM Avvio dell'eseguibile
CON412_TransactionReporting.exe

REM Pausa per vedere eventuali errori
if errorlevel 1 (
    echo.
    echo ============================================================
    echo   ERRORE: Il sistema si è chiuso con errori
    echo   Controllare i log nella cartella 'log' per dettagli
    echo ============================================================
    pause
) else (
    echo.
    echo ============================================================
    echo   Sistema completato con successo
    echo   Controllare i risultati nella cartella 'output'
    echo ============================================================
    pause
)