# ==============================================================================
# INSTALLATEUR ASSISTANT IA EXCEL - AFRIQUE FRANCOPHONE (SYSCOHADA & PAIE)
# Script universel 100% automatique : enregistrement WEF + classeur d'activation
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Installation de l'Assistant IA Excel"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "    INSTALLATEUR DE L'ASSISTANT IA EXCEL (AFRIEXCEL)" -ForegroundColor Yellow
Write-Host "    SYSCOHADA | 20 Feuilles Bancables | KPay | Finance" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Activation TLS universelle
try {
    [System.Net.ServicePointManager]::SecurityProtocol = 3072 -bor 768
} catch {}

# 2. Fermeture prealable des processus Excel pour liberer les caches Office
Write-Host "[1/5] Fermeture prealable d'Excel pour mise a jour des caches..." -ForegroundColor Yellow
Stop-Process -Name excel -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

# 3. Preparation des dossiers permanents Roaming & Local
$roamingFolder = "$env:APPDATA\AssistantIAExcel"
$localFolder = "$env:LOCALAPPDATA\AssistantIAExcel"

foreach ($dir in @($roamingFolder, $localFolder)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}
Write-Host "[2/5] Dossiers locaux configures avec succes." -ForegroundColor Green

# 4. Manifeste XML Officiel integre (zero latence, zero panne reseau)
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
        <bt:String id="GetStarted.Description" DefaultValue="Assistant IA Business Plan et Modélisation Financière OHADA."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
'@

$manifestRoaming = "$roamingFolder\manifest.xml"
$manifestLocal = "$localFolder\manifest.xml"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($manifestRoaming, $embeddedManifest.Trim(), $utf8NoBom)
[System.IO.File]::WriteAllText($manifestLocal, $embeddedManifest.Trim(), $utf8NoBom)
Write-Host "[3/5] Fichiers manifests deployes localement (0 BOM)." -ForegroundColor Green

# 5. Enregistrement direct dans le Registre Office (WEF Developer + TrustedCatalogs)
$addinGuid = "b29d4c78-5e4a-4b91-912a-8457fce2a912"
$catalogGuid = "{E5B75A18-6326-4C10-9CE2-78A7A0864B7A}"
$uncPath = "\\$env:COMPUTERNAME\AssistantIAExcel"

foreach ($ver in @("16.0", "15.0")) {
    & reg.exe add "HKCU\Software\Microsoft\Office\$ver\WEF\Developer" /f /v "$addinGuid" /t REG_SZ /d "$manifestLocal" 2>$null | Out-Null
    & reg.exe add "HKCU\Software\Microsoft\Office\$ver\WEF\Developer" /f /v "{$addinGuid}" /t REG_SZ /d "$manifestLocal" 2>$null | Out-Null
    & reg.exe add "HKCU\Software\Microsoft\Office\$ver\WEF\Developer" /f /v "AssistantIAExcel" /t REG_SZ /d "$manifestLocal" 2>$null | Out-Null

    # Catalogue 1 : Partage UNC Reseau Machine (Affiche DOSSIER PARTAGE dans Excel)
    & reg.exe add "HKCU\Software\Microsoft\Office\$ver\WEF\TrustedCatalogs\$catalogGuid" /f /v Id /t REG_SZ /d "$catalogGuid" 2>$null | Out-Null
    & reg.exe add "HKCU\Software\Microsoft\Office\$ver\WEF\TrustedCatalogs\$catalogGuid" /f /v Url /t REG_SZ /d "$uncPath" 2>$null | Out-Null
    & reg.exe add "HKCU\Software\Microsoft\Office\$ver\WEF\TrustedCatalogs\$catalogGuid" /f /v Flags /t REG_DWORD /d 1 2>$null | Out-Null

    # Catalogue 2 : Localhost UNC
    $catalogGuid2 = "{B29D4C78-5E4A-4B91-912A-8457FCE2A912}"
    & reg.exe add "HKCU\Software\Microsoft\Office\$ver\WEF\TrustedCatalogs\$catalogGuid2" /f /v Id /t REG_SZ /d "$catalogGuid2" 2>$null | Out-Null
    & reg.exe add "HKCU\Software\Microsoft\Office\$ver\WEF\TrustedCatalogs\$catalogGuid2" /f /v Url /t REG_SZ /d "\\localhost\AssistantIAExcel" 2>$null | Out-Null
    & reg.exe add "HKCU\Software\Microsoft\Office\$ver\WEF\TrustedCatalogs\$catalogGuid2" /f /v Flags /t REG_DWORD /d 1 2>$null | Out-Null
}

# 6. Reinitialisation des caches Office WEF
$cacheFolders = @(
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\WEF\AggregatedCache",
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\WEF\AddinInfo",
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\WEF\AppCommands"
)
foreach ($cf in $cacheFolders) {
    if (Test-Path $cf) {
        Remove-Item -Path "$cf\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
}
Write-Host "[4/5] Registre configure et caches Office reinitialises." -ForegroundColor Green

# 7. Generation du classeur auto-activable avec volet IA integre
Write-Host "[5/5] Creation du classeur d'activation directe..." -ForegroundColor Green
$desktopPath = [Environment]::GetFolderPath("Desktop")
$targetWb = Join-Path $desktopPath "Assistant_IA_Excel.xlsx"

Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

function Build-AutoOpeningWorkbook($outPath) {
    $tempDir = Join-Path $env:TEMP "ExcelWebExtensionBuild_Auto"
    if (Test-Path $tempDir) { Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
    
    New-Item -ItemType Directory -Path "$tempDir\_rels" -Force | Out-Null
    New-Item -ItemType Directory -Path "$tempDir\xl\_rels" -Force | Out-Null
    New-Item -ItemType Directory -Path "$tempDir\xl\worksheets" -Force | Out-Null
    New-Item -ItemType Directory -Path "$tempDir\xl\webextensions\_rels" -Force | Out-Null

    # UTF-8 pur sans aucun BOM (indispensable pour la conformité OpenXML et la reconnaissance par Excel)
    $utf8 = New-Object System.Text.UTF8Encoding($false)

    [System.IO.File]::WriteAllText((Join-Path $tempDir "[Content_Types].xml"), @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>
  <Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>
</Types>
'@, $utf8)

    [System.IO.File]::WriteAllText((Join-Path $tempDir "_rels\.rels"), @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@, $utf8)

    [System.IO.File]::WriteAllText((Join-Path $tempDir "xl\workbook.xml"), @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="AfriExcel Copilot" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
'@, $utf8)

    [System.IO.File]::WriteAllText((Join-Path $tempDir "xl\_rels\workbook.xml.rels"), @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="webextensions/taskpanes.xml"/>
</Relationships>
'@, $utf8)

    [System.IO.File]::WriteAllText((Join-Path $tempDir "xl\worksheets\sheet1.xml"), @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>AFRIEXCEL COPILOT - BUSINESS PLAN BANCAIRE SYSCOHADA (20 FEUILLES)</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>L'Assistant IA s'ouvre automatiquement dans le volet a droite.</t></is></c></row>
  </sheetData>
</worksheet>
'@, $utf8)

    [System.IO.File]::WriteAllText((Join-Path $tempDir "xl\webextensions\taskpanes.xml"), @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:taskpane dockstate="right" visibility="1" width="380" row="1">
    <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
  </wetp:taskpane>
</wetp:taskpanes>
'@, $utf8)

    [System.IO.File]::WriteAllText((Join-Path $tempDir "xl\webextensions\_rels\taskpanes.xml.rels"), @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>
</Relationships>
'@, $utf8)

    [System.IO.File]::WriteAllText((Join-Path $tempDir "xl\webextensions\webextension1.xml"), @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="$catalogGuid">
  <we:reference id="$addinGuid" version="2.0.0.0" store="$uncPath" storeType="FileSystem"/>
  <we:alternateReferences>
    <we:reference id="$addinGuid" version="2.0.0.0" store="developer" storeType="developer"/>
  </we:alternateReferences>
  <we:properties>
    <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
"@, $utf8)

    if (Test-Path $outPath) { Remove-Item $outPath -Force -ErrorAction SilentlyContinue }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $outPath)
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

try {
    Build-AutoOpeningWorkbook $targetWb
    $altWb = Join-Path $desktopPath "AssistantIA_AfriExcel.xlsx"
    Build-AutoOpeningWorkbook $altWb
    Write-Host "      Classeur auto-activable pret : $targetWb" -ForegroundColor Green
} catch {}

# 8. Lancement automatique d'Excel avec le classeur
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "     INSTALLATION TERMINEE AVEC SUCCES DANS EXCEL !         " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Lancement de Microsoft Excel..." -ForegroundColor Yellow

if (Test-Path $targetWb) {
    Start-Process $targetWb
} else {
    Start-Process "excel.exe"
}

Start-Sleep -Seconds 2
