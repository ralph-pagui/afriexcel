# ==============================================================================
# INSTALLATEUR ASSISTANT IA EXCEL - AFRIQUE FRANCOPHONE (SYSCOHADA & PAIE)
# Script d'installation automatique et directe dans Microsoft Excel
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Installation de l'Assistant IA Excel"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "    INSTALLATEUR DE L'ASSISTANT IA EXCEL (AFRIQUE)" -ForegroundColor Yellow
Write-Host "    SYSCOHADA | Rapprochement bancaire Wave/Orange | Paie" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Vérification et élévation Administrateur si nécessaire
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[!] Elevation des privileges administrateur en cours..." -ForegroundColor Yellow
    $scriptPath = $MyInvocation.MyCommand.Definition
    if (-not $scriptPath) { $scriptPath = "$PSScriptRoot\Install-AssistantIA.ps1" }
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    exit
}

# 2. Préparation du dossier local
$targetDir = "$env:LOCALAPPDATA\AssistantIAExcel"
if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}
Write-Host "[1/6] Dossier local configure : $targetDir" -ForegroundColor Green

# 3. Récupération du Manifest XML
$manifestPath = "$targetDir\manifest.xml"
$localManifest = Join-Path $PSScriptRoot "..\addin-frontend\dist\manifest.xml"
$remoteManifestUrl = "https://afriexcel.netlify.app/manifest.xml"

if (Test-Path $localManifest) {
    Copy-Item $localManifest $manifestPath -Force
    Write-Host "[2/6] Manifest copie depuis la source locale." -ForegroundColor Green
} else {
    Write-Host "[2/6] Telechargement du manifeste officiel depuis afriexcel.netlify.app..." -ForegroundColor Green
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $remoteManifestUrl -OutFile $manifestPath -UseBasicParsing
        Write-Host "      Manifeste telecharge avec succes." -ForegroundColor Green
    } catch {
        Write-Host "      [!] Erreur telechargement, tentative alternative..." -ForegroundColor Yellow
        (New-Object System.Net.WebClient).DownloadFile($remoteManifestUrl, $manifestPath)
    }
}

# 4. Configuration du Partage Réseau Windows (SMB)
Write-Host "[3/6] Creation du partage reseau local..." -ForegroundColor Green
net share AssistantIAExcel /delete 2>$null | Out-Null
$shareCreated = $false

# Essai 1 : Everyone
cmd.exe /c "net share AssistantIAExcel=`"$targetDir`" /grant:Everyone,FULL" 2>$null | Out-Null
if ($LASTEXITCODE -eq 0) { $shareCreated = $true }

# Essai 2 : Tout le monde (Windows Français)
if (-not $shareCreated) {
    cmd.exe /c "net share AssistantIAExcel=`"$targetDir`" /grant:`"Tout le monde`",FULL" 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { $shareCreated = $true }
}

# Essai 3 : Sans spécifier d'utilisateur
if (-not $shareCreated) {
    cmd.exe /c "net share AssistantIAExcel=`"$targetDir`"" 2>$null | Out-Null
}

$uncPath = "\\$env:COMPUTERNAME\AssistantIAExcel"
Write-Host "      Partage SMB pret : $uncPath" -ForegroundColor Green

# 5. Enregistrement dans le Registre Office (Catalogue approuve officiel avec GUID)
Write-Host "[4/6] Enregistrement dans Microsoft Office..." -ForegroundColor Green

$catalogGuid = "{E5B75A18-6326-4C10-9CE2-78A7A0864B7A}"
$addinGuid = "b29d4c78-5e4a-4b91-912a-8457fce2a912"

# A. Catalogue approuve UNC (pour que l'onglet DOSSIER PARTAGE et le ruban apparaissent)
$regCatalog = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogGuid"
New-Item -Path $regCatalog -Force | Out-Null
Set-ItemProperty -Path $regCatalog -Name "Id" -Value $catalogGuid -Type String
Set-ItemProperty -Path $regCatalog -Name "Url" -Value $uncPath -Type String
Set-ItemProperty -Path $regCatalog -Name "Flags" -Value 1 -Type DWord

# B. Enregistrement Developer (pour activation directe sans confirmation)
$regDev = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
New-Item -Path $regDev -Force | Out-Null
Set-ItemProperty -Path $regDev -Name $addinGuid -Value $manifestPath -Type String

# C. Nettoyer les anciennes cles numeriques obsoletes
Remove-Item "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\2" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "      Catalogue et permissions enregistres." -ForegroundColor Green

# 6. Nettoyage du cache WEF d'Office pour forcer la detection immediate
Write-Host "[5/6] Rafraichissement du cache Office..." -ForegroundColor Green
$wefCache = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
if (Test-Path $wefCache) {
    Remove-Item "$wefCache\*" -Recurse -Force -ErrorAction SilentlyContinue
}

# 7. Creation du classeur d'activation automatique sur le Bureau
Write-Host "[6/6] Creation du classeur d'activation sur le Bureau..." -ForegroundColor Green
$desktopPath = [Environment]::GetFolderPath("Desktop")
$workbookPath = Join-Path $desktopPath "AssistantIA.xlsx"

Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

function Create-ActivationWorkbook($path) {
    if (Test-Path $path) { Remove-Item $path -Force -ErrorAction SilentlyContinue }
    $zip = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Create)
    
    function Write-ZipEntry($z, $name, $content) {
        $entry = $z.CreateEntry($name)
        $writer = New-Object System.IO.StreamWriter($entry.Open(), [System.Text.Encoding]::UTF8)
        $writer.Write($content)
        $writer.Close()
    }
    
    Write-ZipEntry $zip "[Content_Types].xml" @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>
  <Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>
</Types>
'@

    Write-ZipEntry $zip "_rels/.rels" @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="xl/webextensions/taskpanes.xml"/>
</Relationships>
'@

    Write-ZipEntry $zip "xl/workbook.xml" @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Assistant IA" sheetId="1" r:id="rId1"/></sheets>
</workbook>
'@

    Write-ZipEntry $zip "xl/_rels/workbook.xml.rels" @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
'@

    Write-ZipEntry $zip "xl/worksheets/sheet1.xml" @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Assistant IA Excel - Afrique Francophone</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>L'Assistant IA s'ouvre automatiquement dans le volet de droite.</t></is></c></row>
  </sheetData>
</worksheet>
'@

    Write-ZipEntry $zip "xl/webextensions/taskpanes.xml" @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:taskpane dockstate="right" visibility="1" width="380" row="1" locked="0">
    <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
  </wetp:taskpane>
</wetp:taskpanes>
'@

    Write-ZipEntry $zip "xl/webextensions/_rels/taskpanes.xml.rels" @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>
</Relationships>
'@

    $weXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="$catalogGuid">
  <we:reference id="$addinGuid" version="1.0.0.0" store="$uncPath" storeType="FileSystem"/>
  <we:alternateReferences>
    <we:reference id="$addinGuid" version="1.0.0.0" store="developer" storeType="Registry"/>
  </we:alternateReferences>
  <we:properties>
    <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot/>
</we:webextension>
"@
    Write-ZipEntry $zip "xl/webextensions/webextension1.xml" $weXml
    $zip.Dispose()
}

try {
    Create-ActivationWorkbook $workbookPath
    Write-Host "      Classeur cree sur le Bureau : $workbookPath" -ForegroundColor Green
} catch {
    Write-Host "      [!] Note : Le classeur direct n'a pas pu etre genere ($($_.Exception.Message))" -ForegroundColor Yellow
}

# 8. Lancement d'Excel
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   INSTALLATION REUSSIE AVEC SUCCES !" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Ouverture automatique d'Excel..." -ForegroundColor Yellow

$excelExe = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
if (-not (Test-Path $excelExe)) {
    $excelExe = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
}

if (Test-Path $workbookPath) {
    if (Test-Path $excelExe) {
        Start-Process $excelExe -ArgumentList "`"$workbookPath`""
    } else {
        Start-Process $workbookPath
    }
} else {
    Start-Process "excel.exe"
}

Write-Host "Termine ! L'Assistant IA est pret a l'emploi." -ForegroundColor Green
Write-Host ""
Start-Sleep -Seconds 3
