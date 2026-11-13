# Script d'installation automatique 1-Clic pour l'Assistant IA Excel
# Compatible avec Excel 2016, Excel 2019, Excel 2021 et Microsoft 365 sous Windows
[CmdletBinding()]
param (
    [string]$NetlifyUrl = ""
)

Write-Host "=====================================================" -ForegroundColor Green
Write-Host "   INSTALLATION ASSISTANT IA EXCEL (AFRIQUE / OHADA)  " -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Green
Write-Host ""

# 1. Dossier de destination local
$TargetDir = "$env:LOCALAPPDATA\AssistantIAExcel"
if (!(Test-Path $TargetDir)) {
    New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null
}

Write-Host "[1/4] Préparation du dossier local : $TargetDir" -ForegroundColor White

# 2. Copie et adaptation du fichier Manifest XML
$SourceManifest = Join-Path $PSScriptRoot "..\manifest.xml"
$TargetManifest = Join-Path $TargetDir "manifest.xml"

if (Test-Path $SourceManifest) {
    $manifestContent = Get-Content -Path $SourceManifest -Raw -Encoding UTF8
    
    # Si une URL Netlify personnalisée est passée en paramètre, on l'injecte
    if ($NetlifyUrl -ne "") {
        Write-Host "[*] Injection de l'URL Netlify personnalisée : $NetlifyUrl" -ForegroundColor Yellow
        $cleanUrl = $NetlifyUrl.TrimEnd('/')
        $manifestContent = $manifestContent -replace "https://mon-assistant-excel.netlify.app", $cleanUrl
    }
    
    Set-Content -Path $TargetManifest -Value $manifestContent -Encoding UTF8
    Write-Host "[2/4] Fichier Manifest enregistré avec succès." -ForegroundColor White
} else {
    Write-Host "[!] Erreur : Fichier manifest.xml introuvable dans le dossier parent." -ForegroundColor Red
    Exit 1
}

# 3. Enregistrement dans le Registre Windows pour Office 2019
# Office 2019 exige un chemin réseau UNC (\\localhost\c$\...) et refuse les chemins locaux C:\
$uncLocalhost = "\\localhost\c$" + $TargetDir.Substring(2)
$uncComputer = "\\$env:COMPUTERNAME\c$" + $TargetDir.Substring(2)

Write-Host "[3/4] Configuration du registre Windows (Office 2019 UNC: $uncLocalhost)..." -ForegroundColor White

# Clé 1 : Catalog UNC Localhost
$CatalogGuid1 = "{b29d4c78-5e4a-4b91-912a-8457fce2a912}"
$RegistryPath1 = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogGuid1"
if (!(Test-Path $RegistryPath1)) { New-Item -Path $RegistryPath1 -Force | Out-Null }
Set-ItemProperty -Path $RegistryPath1 -Name "Url" -Value $uncLocalhost -Type String
Set-ItemProperty -Path $RegistryPath1 -Name "Flags" -Value 1 -Type DWord

# Clé 2 : Catalog UNC ComputerName
$CatalogGuid2 = "{c38e5d89-6f5b-4c02-823b-9568gdf3b023}"
$RegistryPath2 = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogGuid2"
if (!(Test-Path $RegistryPath2)) { New-Item -Path $RegistryPath2 -Force | Out-Null }
Set-ItemProperty -Path $RegistryPath2 -Name "Url" -Value $uncComputer -Type String
Set-ItemProperty -Path $RegistryPath2 -Name "Flags" -Value 1 -Type DWord

# Clé 3 : Sideloading direct Developer Manifest
$DevPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer\$CatalogGuid1"
if (!(Test-Path $DevPath)) { New-Item -Path $DevPath -Force | Out-Null }
Set-ItemProperty -Path $DevPath -Name "ManifestPath" -Value $TargetManifest -Type String

# 4. Vidage du cache Office WEF pour forcer la prise en compte immédiate
$WefCacheDir = "$env:LOCALAPPDATA\Microsoft\Office\16.0\WEF"
if (Test-Path $WefCacheDir) {
    try {
        Remove-Item "$WefCacheDir\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "[4/4] Cache Office actualisé." -ForegroundColor White
    } catch {
        # Si un fichier est verrouillé par Excel en cours d'exécution
    }
}

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Green
Write-Host "   SUCCÈS : L'ASSISTANT IA EST INSTALLÉ DANS EXCEL !  " -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Pour l'utiliser dès maintenant :" -ForegroundColor Yellow
Write-Host "1. Ouvrez Microsoft Excel (2016, 2019, 2021 ou 365)" -ForegroundColor White
Write-Host "2. Allez dans le menu : Insertion > Mes compléments > Dossier partagé" -ForegroundColor White
Write-Host "3. Sélectionnez 'Assistant IA Excel - Afrique' et cliquez sur 'Ajouter'" -ForegroundColor White
Write-Host "   (Le bouton 'Assistant IA' apparaîtra ensuite directement dans votre ruban !)" -ForegroundColor Cyan
Write-Host ""
Write-Host "Installation terminée avec succès !" -ForegroundColor Green
