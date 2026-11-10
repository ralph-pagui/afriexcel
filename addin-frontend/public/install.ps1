# ==============================================================================
# INSTALLATEUR ASSISTANT IA EXCEL - AFRIQUE FRANCOPHONE (SYSCOHADA & PAIE)
# Script d'installation automatique et directe dans Microsoft Excel
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Installation de l'Assistant IA Excel"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "    INSTALLATEUR DE L'ASSISTANT IA EXCEL (AFRIEXCEL)" -ForegroundColor Yellow
Write-Host "    SYSCOHADA | 20 Feuilles Bancables | KPay | Finance" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Verification des privileges
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
    Write-Host "[+] Droits Administrateur detectes : activation et partage SMB complets." -ForegroundColor Green
} else {
    Write-Host "[i] Droits Utilisateur Standard : enregistrement direct via le registre Office." -ForegroundColor Yellow
}

# 2. Preparation du dossier local
$targetDir = "$env:LOCALAPPDATA\AssistantIAExcel"
if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}
Write-Host "[1/5] Dossier local configure : $targetDir" -ForegroundColor Green

# 3. Telechargement du Manifeste XML Officiel
$manifestPath = "$targetDir\manifest.xml"
$remoteManifestUrl = "https://afriexcelai.netlify.app/manifest.xml"

Write-Host "[2/5] Telechargement du manifeste officiel (https://afriexcelai.netlify.app)..." -ForegroundColor Green
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $remoteManifestUrl -OutFile $manifestPath -UseBasicParsing
    Write-Host "      Manifeste telecharge avec succes." -ForegroundColor Green
} catch {
    try {
        (New-Object System.Net.WebClient).DownloadFile($remoteManifestUrl, $manifestPath)
        Write-Host "      Manifeste telecharge via WebClient." -ForegroundColor Green
    } catch {
        Write-Host "      [!] Erreur lors du telechargement du manifeste : $($_.Exception.Message)" -ForegroundColor Red
    }
}

# 4. Partage Reseau SMB (si administrateur)
$catalogGuid = "{E5B75A18-6326-4C10-9CE2-78A7A0864B7A}"
$addinGuid = "b29d4c78-5e4a-4b91-912a-8457fce2a912"
$uncPath = "\\$env:COMPUTERNAME\AssistantIAExcel"

if ($isAdmin) {
    try {
        net share AssistantIAExcel /delete 2>$null | Out-Null
        cmd.exe /c "net share AssistantIAExcel=`"$targetDir`" /grant:Everyone,FULL" 2>$null | Out-Null
        if ($LASTEXITCODE -ne 0) {
            cmd.exe /c "net share AssistantIAExcel=`"$targetDir`" /grant:`"Tout le monde`",FULL" 2>$null | Out-Null
        }
        Write-Host "      Partage SMB configure : $uncPath" -ForegroundColor Green
        
        $regCatalog = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogGuid"
        New-Item -Path $regCatalog -Force | Out-Null
        Set-ItemProperty -Path $regCatalog -Name "Id" -Value $catalogGuid -Type String
        Set-ItemProperty -Path $regCatalog -Name "Url" -Value $uncPath -Type String
        Set-ItemProperty -Path $regCatalog -Name "Flags" -Value 1 -Type DWord
    } catch {
        Write-Host "      [i] Partage SMB ignore : enregistrement direct Developer actif." -ForegroundColor Gray
    }
}

# 5. Enregistrement direct Developer dans le Registre Office (fonctionne sans admin)
Write-Host "[3/5] Enregistrement dans Microsoft Office..." -ForegroundColor Green
$regDev = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
New-Item -Path $regDev -Force | Out-Null
Set-ItemProperty -Path $regDev -Name $addinGuid -Value $manifestPath -Type String
Write-Host "      Cle Developer enregistree ($addinGuid)." -ForegroundColor Green

# 6. Nettoyage du cache WEF d'Office
Write-Host "[4/5] Rafraichissement du cache Office..." -ForegroundColor Green
$wefCache = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
if (Test-Path $wefCache) {
    Remove-Item "$wefCache\*" -Recurse -Force -ErrorAction SilentlyContinue
}

# 7. Creation du classeur d'activation automatique sur le Bureau
Write-Host "[5/5] Creation du classeur d'activation sur le Bureau..." -ForegroundColor Green
$desktopPath = [Environment]::GetFolderPath("Desktop")
$workbookPath = Join-Path $desktopPath "AssistantIA_AfriExcel.xlsx"

Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
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
  <sheets><sheet name="AfriExcel Copilot" sheetId="1" r:id="rId1"/></sheets>
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
    <row r="1"><c r="A1" t="inlineStr"><is><t>AfriExcel Copilot - Business Plan OHADA &amp; Finance</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>L'Assistant IA s'ouvre automatiquement dans le volet a droite.</t></is></c></row>
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

    $storeLocation = if ($isAdmin) { $uncPath } else { "developer" }
    $storeType = if ($isAdmin) { "FileSystem" } else { "Registry" }

    $weXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="$catalogGuid">
  <we:reference id="$addinGuid" version="2.0.0.0" store="$storeLocation" storeType="$storeType"/>
  <we:alternateReferences>
    <we:reference id="$addinGuid" version="2.0.0.0" store="developer" storeType="Registry"/>
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
    Write-Host "      Classeur direct pret sur le Bureau : AssistantIA_AfriExcel.xlsx" -ForegroundColor Green
} catch {
    Write-Host "      [i] Creation du classeur d'activation optionnel ignoree." -ForegroundColor Gray
}

# 8. Notification Excel
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "    INSTALLATION REUSSIE AVEC SUCCES !" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Vous pouvez desormais utiliser l'Assistant IA :" -ForegroundColor White
Write-Host "  1. Ouvrez Microsoft Excel." -ForegroundColor White
Write-Host "  2. Un bouton 'Copilote OHADA' est present dans l'onglet 'Accueil' du ruban." -ForegroundColor White
Write-Host "  3. Ou double-cliquez directement sur 'AssistantIA_AfriExcel.xlsx' sur votre Bureau." -ForegroundColor White
Write-Host ""

$excelProcesses = Get-Process excel -ErrorAction SilentlyContinue
if ($excelProcesses) {
    Write-Host "[!] Note : Excel etait deja ouvert. Veuillez le fermer puis le rouvrir pour actualiser le ruban." -ForegroundColor Yellow
} else {
    Write-Host "Lancement automatique d'Excel..." -ForegroundColor Yellow
    if (Test-Path $workbookPath) {
        Start-Process $workbookPath
    } else {
        Start-Process "excel.exe"
    }
}

Start-Sleep -Seconds 2
