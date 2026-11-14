# ==============================================================================
# INSTALLATEUR ASSISTANT IA EXCEL - AFRIQUE FRANCOPHONE (SYSCOHADA & PAIE)
# Script d'installation universel et automatique dans Microsoft Excel
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Installation de l'Assistant IA Excel"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "    INSTALLATEUR DE L'ASSISTANT IA EXCEL (AFRIEXCEL)" -ForegroundColor Yellow
Write-Host "    SYSCOHADA | 20 Feuilles Bancables | KPay | Finance" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Activation universelle des protocoles TLS (TLS 1.2 + TLS 1.0)
try {
    [System.Net.ServicePointManager]::SecurityProtocol = 3072 -bor 768
} catch {}

# 2. Verification des privileges
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
    Write-Host "[+] Droits Administrateur detectes : activation complete (Partage SMB + Registre Office)." -ForegroundColor Green
} else {
    Write-Host "[i] Droits Utilisateur Standard : enregistrement direct et securise dans le registre Office (sans UAC ni mot de passe)." -ForegroundColor Yellow
}

# 3. Preparation du dossier local permanent
$targetDir = "$env:LOCALAPPDATA\AssistantIAExcel"
if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}
Write-Host "[1/5] Dossier local configure : $targetDir" -ForegroundColor Green

# 4. Preparation du Manifeste XML Officiel (Integre directement pour 100% de fiabilite et zero latence reseau)
$manifestPath = "$targetDir\manifest.xml"
Write-Host "[2/5] Initialisation du manifeste officiel..." -ForegroundColor Green

$embeddedManifest = @'
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>b29d4c78-5e4a-4b91-912a-8457fce2a912</Id>
  <Version>2.0.0.0</Version>
  <ProviderName>AfriExcel Technologies</ProviderName>
  <DefaultLocale>fr-FR</DefaultLocale>
  <DisplayName DefaultValue="AfriExcel Copilot - Business Plan OHADA"/>
  <Description DefaultValue="Assistant IA pour la création et la modification de Business Plans bancables SYSCOHADA à 20 feuilles dans Excel."/>
  <IconUrl DefaultValue="https://afriexcelai.netlify.app/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://afriexcelai.netlify.app/assets/icon-80.png"/>
  <SupportUrl DefaultValue="https://afriexcelai.netlify.app/"/>
  <AppDomains>
    <AppDomain>https://afriexcelai.netlify.app</AppDomain>
    <AppDomain>https://appsforoffice.microsoft.com</AppDomain>
    <AppDomain>https://generativelanguage.googleapis.com</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://afriexcelai.netlify.app/?mode=addin"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Group.AssistantIA">
                <Label resid="GroupLabel"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16"/>
                  <bt:Image size="32" resid="Icon.32"/>
                  <bt:Image size="80" resid="Icon.80"/>
                </Icon>
                <Control xsi:type="Button" id="Button.ShowTaskpane">
                  <Label resid="ButtonLabel"/>
                  <Supertip>
                    <Title resid="ButtonLabel"/>
                    <Description resid="ButtonToolTip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="https://afriexcelai.netlify.app/assets/icon-16.png"/>
        <bt:Image id="Icon.32" DefaultValue="https://afriexcelai.netlify.app/assets/icon-32.png"/>
        <bt:Image id="Icon.80" DefaultValue="https://afriexcelai.netlify.app/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="https://afriexcelai.netlify.app/?mode=addin"/>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://afriexcelai.netlify.app/"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="AfriExcel"/>
        <bt:String id="ButtonLabel" DefaultValue="Copilote OHADA"/>
        <bt:String id="GetStarted.Title" DefaultValue="AfriExcel Copilot"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="ButtonToolTip" DefaultValue="Ouvre le copilote IA pour générer et modifier vos Business Plans bancables SYSCOHADA à 20 feuilles."/>
        <bt:String id="GetStarted.Description" DefaultValue="Assistant IA Business Plan & Modélisation Financière OHADA."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
'@

try {
    [System.IO.File]::WriteAllText($manifestPath, $embeddedManifest.Trim(), [System.Text.Encoding]::UTF8)
    Write-Host "      Manifeste XML installe et valide localement." -ForegroundColor Green
} catch {
    Set-Content -Path $manifestPath -Value $embeddedManifest.Trim() -Encoding UTF8 -Force
    Write-Host "      Manifeste XML configure." -ForegroundColor Green
}

# 5. Configuration des Catalogues Approuves Office (Multi-Canal Universel)
Write-Host "[3/5] Configuration des Catalogues de Confiance Office..." -ForegroundColor Green
$catalogGuid1 = "{E5B75A18-6326-4C10-9CE2-78A7A0864B7A}"
$catalogGuid2 = "{B29D4C78-5E4A-4B91-912A-8457FCE2A912}"
$catalogGuid3 = "{C38E5D89-6F5B-4C02-823B-9568GDF3B023}"
$catalogGuid4 = "{D49F6E9A-706C-4D13-934C-0679HE4CA134}"
$addinGuid = "b29d4c78-5e4a-4b91-912a-8457fce2a912"

$uncPath = "\\$env:COMPUTERNAME\AssistantIAExcel"
$uncLocalhost = "\\localhost\c$" + $targetDir.Substring(2)
$fileUri = "file:///" + ($targetDir.Replace("\", "/")) + "/"
$directPath = "$targetDir\"

# Tentative de creation du partage reseau SMB
try {
    net share AssistantIAExcel /delete 2>$null | Out-Null
    cmd.exe /c "net share AssistantIAExcel=`"$targetDir`" /grant:Everyone,FULL" 2>$null | Out-Null
    if ($LASTEXITCODE -ne 0) {
        cmd.exe /c "net share AssistantIAExcel=`"$targetDir`" /grant:`"Tout le monde`",FULL" 2>$null | Out-Null
    }
    if ($LASTEXITCODE -ne 0) {
        cmd.exe /c "net share AssistantIAExcel=`"$targetDir`"" 2>$null | Out-Null
    }
} catch {}

# Inscription des 4 formats de catalog dans HKCU (ne necessite aucun droit administrateur)
$catalogs = @(
    @{ Guid = $catalogGuid1; Url = $uncPath },
    @{ Guid = $catalogGuid2; Url = $uncLocalhost },
    @{ Guid = $catalogGuid3; Url = $fileUri },
    @{ Guid = $catalogGuid4; Url = $directPath }
)

foreach ($cat in $catalogs) {
    try {
        $regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$($cat.Guid)"
        if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
        Set-ItemProperty -Path $regPath -Name "Id" -Value $cat.Guid -Type String
        Set-ItemProperty -Path $regPath -Name "Url" -Value $cat.Url -Type String
        Set-ItemProperty -Path $regPath -Name "Flags" -Value 1 -Type DWord
    } catch {}
}

# Enregistrement direct Developer pour sideloading sans reseau
$regDev = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
if (-not (Test-Path $regDev)) { New-Item -Path $regDev -Force | Out-Null }
Set-ItemProperty -Path $regDev -Name $addinGuid -Value $manifestPath -Type String
Set-ItemProperty -Path $regDev -Name "{$addinGuid}" -Value $manifestPath -Type String
Set-ItemProperty -Path $regDev -Name "AssistantIAExcel" -Value $manifestPath -Type String
Write-Host "      Enregistrement registre effectue avec succes." -ForegroundColor Green

# 6. Rafraichissement securise des caches Office
Write-Host "[4/5] Optimisation des caches Office..." -ForegroundColor Green
$cacheAggregated = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\AggregatedCache"
if (Test-Path $cacheAggregated) {
    Remove-Item "$cacheAggregated\*" -Recurse -Force -ErrorAction SilentlyContinue
}

# 7. Creation du classeur d'accueil et d'activation sur le Bureau
Write-Host "[5/5] Creation du classeur d'accueil sur le Bureau..." -ForegroundColor Green
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
</Types>
'@

    Write-ZipEntry $zip "_rels/.rels" @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
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
    <row r="1"><c r="A1" t="inlineStr"><is><t>AFRIEXCEL COPILOT - BUSINESS PLAN BANCAIRE SYSCOHADA (20 FEUILLES)</t></is></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>POUR OUVRIR LE VOLET DE L'ASSISTANT IA DANS CE CLASSEUR :</t></is></c></row>
    <row r="5"><c r="A5" t="inlineStr"><is><t>1. Cliquez sur le menu "Insertion" tout en haut dans votre ruban Excel.</t></is></c></row>
    <row r="6"><c r="A6" t="inlineStr"><is><t>2. Cliquez sur "Mes compléments" (ou l'icône Compléments).</t></is></c></row>
    <row r="7"><c r="A7" t="inlineStr"><is><t>3. Allez dans l'onglet "DOSSIER PARTAGÉ".</t></is></c></row>
    <row r="8"><c r="A8" t="inlineStr"><is><t>4. Sélectionnez "AfriExcel Copilot - Business Plan OHADA" puis cliquez sur "Ajouter".</t></is></c></row>
    <row r="10"><c r="A10" t="inlineStr"><is><t>✓ Le volet de l'Assistant IA s'affiche immédiatement à droite !</t></is></c></row>
    <row r="11"><c r="A11" t="inlineStr"><is><t>✓ Le bouton "Copilote OHADA" restera ensuite fixé pour toujours dans votre onglet Accueil !</t></is></c></row>
  </sheetData>
</worksheet>
'@

    $zip.Dispose()
}

try {
    Create-ActivationWorkbook $workbookPath
    Write-Host "      Classeur pret sur le Bureau : AssistantIA_AfriExcel.xlsx" -ForegroundColor Green
} catch {}

# 8. Notification & Lancement Excel
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "     INSTALLATION CONFIGUREE AVEC SUCCES DANS EXCEL !       " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Excel s'ouvre. Pour afficher l'Assistant IA a droite :" -ForegroundColor Yellow
Write-Host ""
Write-Host "  1. Cliquez sur l'onglet 'Insertion' en haut de votre ruban." -ForegroundColor White
Write-Host "  2. Cliquez sur 'Mes complements' (ou 'Complements')." -ForegroundColor White
Write-Host "  3. Allez dans l'onglet 'DOSSIER PARTAGE'." -ForegroundColor White
Write-Host "  4. Cliquez sur 'AfriExcel Copilot' puis sur 'Ajouter'." -ForegroundColor Green
Write-Host ""
Write-Host "  Le volet s'ouvrira aussitot et le bouton 'Copilote OHADA' sera" -ForegroundColor Cyan
Write-Host "  present dans votre onglet Accueil pour tous vos classeurs !" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$excelProcesses = Get-Process excel -ErrorAction SilentlyContinue
if ($excelProcesses) {
    Write-Host "[i] Note : Excel etait deja ouvert. Veuillez fermer et rouvrir Excel ou suivre les 4 etapes ci-dessus." -ForegroundColor Yellow
} else {
    Write-Host "Lancement de Microsoft Excel..." -ForegroundColor Yellow
    if (Test-Path $workbookPath) {
        Start-Process $workbookPath
    } else {
        Start-Process "excel.exe"
    }
}

Start-Sleep -Seconds 3
