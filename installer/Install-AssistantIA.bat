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

powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = 3072 -bor 12288 -bor 768 -bor 192; iex ((New-Object System.Net.WebClient).DownloadString('https://afriexcelai.netlify.app/install.ps1'))"

echo.
pause
