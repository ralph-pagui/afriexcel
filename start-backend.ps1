# Script de lancement du backend Python AfriExcel
$ErrorActionPreference = "Stop"

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "  Démarrage du Serveur Python Copilote IA Business Plan   " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Green

Set-Location -Path "$PSScriptRoot\backend"
python run_server.py
