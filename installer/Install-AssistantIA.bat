@echo off
chcp 65001 >nul
title Installation AfriExcel Copilot - Microsoft Excel
echo.
echo ============================================================
echo    INSTALLATEUR RAPIDE AFRIEXCEL (SYSCOHADA & FINANCE)
echo ============================================================
echo.
echo Configuration automatique dans Microsoft Excel en cours...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/ralph-pagui/afriexcel/main/addin-frontend/public/install.ps1 | iex"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [i] Tentative avec protocole de secours...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = 3072 -bor 768; iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/ralph-pagui/afriexcel/main/addin-frontend/public/install.ps1'))"
)

echo.
pause
