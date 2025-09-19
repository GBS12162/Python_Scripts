@echo off
title CON-412 Transaction Reporting - Sistema Interattivo
echo ================================================
echo CON-412 TRANSACTION REPORTING - REJECTING MENSILE
echo Sistema Interattivo con Configurazione Dinamica
echo ================================================
echo.

cd /d "%~dp0"
py main_con412.py

echo.
echo Premi un tasto per chiudere...
pause > nul