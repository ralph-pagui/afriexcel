@echo off
chcp 65001 > nul
title Installation Assistant IA Excel - Afrique Francophone
echo ========================================================
echo   Installation directe de l'Assistant IA dans Excel...
echo ========================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0installer_excel.ps1"
pause
