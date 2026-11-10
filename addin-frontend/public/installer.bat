@echo off
chcp 65001 >nul
title Installation Assistant IA Excel - Afrique

:: Vérifier et demander les droits administrateur automatiquement
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo ========================================================
    echo   Demande d'elevation administrateur en cours...
    echo ========================================================
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process cmd.exe -ArgumentList '/c \"\"%~f0\"\"' -Verb RunAs"
    exit /b
)

:: Exécuter le script PowerShell d'installation
if exist "%~dp0install.ps1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://afriexcelai.netlify.app/install.ps1 | iex"
)

pause
