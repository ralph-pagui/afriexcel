# ==============================================================================
# Script d'installation automatique - Assistant IA Excel (Afrique & SYSCOHADA)
# Compatible : Excel 2016, 2019, 2021 et Microsoft 365 (Windows)
# ==============================================================================

Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host "  INSTALLATION DE L'ASSISTANT IA EXCEL (AFRIQUE & SYSCOHADA)  " -ForegroundColor Yellow
Write-Host "==============================================================" -ForegroundColor Cyan

# 1. Fermer toutes les instances d'Excel pour libérer les caches
Write-Host "`n[1/5] Fermeture d'Excel si ouvert..." -ForegroundColor White
Stop-Process -Name excel -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

# 2. Définir le dossier permanent pour le complément
$targetFolder = "$env:APPDATA\AssistantIAExcel"
if (!(Test-Path $targetFolder)) {
    New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
}

# 3. Copier le fichier manifest.xml vers l'emplacement permanent
$sourceManifest = "F:\cop\manifest.xml"
if (!(Test-Path $sourceManifest)) {
    # Si le script est exécuté depuis un autre dossier
    $sourceManifest = Join-Path $PSScriptRoot "..\manifest.xml"
}
$destManifest = Join-Path $targetFolder "manifest.xml"
Copy-Item -Path $sourceManifest -Destination $destManifest -Force
Write-Host "[2/5] Manifest copié dans : $destManifest" -ForegroundColor Green

# 4. Enregistrement direct dans le registre WEF Developer de Microsoft Office
# Ce registre est le mécanisme officiel de sideloading reconnu par Excel 2016, 2019, 2021 et 365
Write-Host "[3/5] Enregistrement dans le Registre Office (WEF Developer)..." -ForegroundColor White
$devRegKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
if (!(Test-Path $devRegKey)) {
    New-Item -Path $devRegKey -Force | Out-Null
}

# Enregistrement par GUID et par nom
Set-ItemProperty -Path $devRegKey -Name "b29d4c78-5e4a-4b91-912a-8457fce2a912" -Value $destManifest
Set-ItemProperty -Path $devRegKey -Name "{b29d4c78-5e4a-4b91-912a-8457fce2a912}" -Value $destManifest
Set-ItemProperty -Path $devRegKey -Name "AssistantIAExcel" -Value $destManifest

# 5. Enregistrement du catalogue approuvé (Trusted Catalog) pour l'onglet Dossier partagé
Write-Host "[4/5] Configuration du catalogue approuvé (TrustedCatalogs)..." -ForegroundColor White
$catRegKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{b29d4c78-5e4a-4b91-912a-8457fce2a912}"
if (!(Test-Path $catRegKey)) {
    New-Item -Path $catRegKey -Force | Out-Null
}
Set-ItemProperty -Path $catRegKey -Name "Id" -Value "{b29d4c78-5e4a-4b91-912a-8457fce2a912}"
Set-ItemProperty -Path $catRegKey -Name "Url" -Value "$targetFolder\"
Set-ItemProperty -Path $catRegKey -Name "Flags" -Value 1 -Type DWord

# 6. Nettoyage du cache WEF d'Office pour forcer le rechargement immédiat
Write-Host "[5/5] Réinitialisation du cache Office..." -ForegroundColor White
$cacheFolders = @(
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\WEF\AggregatedCache",
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\WEF\AddinInfo",
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\WEF\AppCommands"
)
foreach ($folder in $cacheFolders) {
    if (Test-Path $folder) {
        Remove-Item -Path "$folder\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "`n==============================================================" -ForegroundColor Green
Write-Host "  INSTALLATION REUSSIE AVEC SUCCES !                         " -ForegroundColor Green
Write-Host "==============================================================" -ForegroundColor Green
Write-Host "Lancement d'Excel..." -ForegroundColor Yellow
Start-Process excel.exe
