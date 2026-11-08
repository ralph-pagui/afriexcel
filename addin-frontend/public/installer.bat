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
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-AssistantIA.ps1"

pause
