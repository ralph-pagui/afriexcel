$tempDir = Join-Path $env:TEMP "ExcelWebExtensionBuild"
if (Test-Path $tempDir) { Remove-Item -Path $tempDir -Recurse -Force }
New-Item -ItemType Directory -Path "$tempDir\_rels" -Force | Out-Null
New-Item -ItemType Directory -Path "$tempDir\xl\_rels" -Force | Out-Null
New-Item -ItemType Directory -Path "$tempDir\xl\worksheets" -Force | Out-Null
New-Item -ItemType Directory -Path "$tempDir\xl\webextensions\_rels" -Force | Out-Null

$utf8 = [System.Text.Encoding]::UTF8

# 1. [Content_Types].xml
$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>
  <Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>
</Types>
'@
[System.IO.File]::WriteAllText((Join-Path $tempDir "[Content_Types].xml"), $contentTypes, $utf8)

# 2. _rels/.rels
$dotRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $tempDir "_rels\.rels"), $dotRels, $utf8)

# 3. xl/workbook.xml
$wbXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Feuille1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
'@
[System.IO.File]::WriteAllText((Join-Path $tempDir "xl\workbook.xml"), $wbXml, $utf8)

# 4. xl/_rels/workbook.xml.rels
$wbRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="webextensions/taskpanes.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $tempDir "xl\_rels\workbook.xml.rels"), $wbRels, $utf8)

# 5. xl/worksheets/sheet1.xml
$sheet1 = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
'@
[System.IO.File]::WriteAllText((Join-Path $tempDir "xl\worksheets\sheet1.xml"), $sheet1, $utf8)

# 6. xl/webextensions/taskpanes.xml
$tpXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:taskpane dockstate="right" visibility="1" width="350" row="1">
    <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
  </wetp:taskpane>
</wetp:taskpanes>
'@
[System.IO.File]::WriteAllText((Join-Path $tempDir "xl\webextensions\taskpanes.xml"), $tpXml, $utf8)

# 7. xl/webextensions/_rels/taskpanes.xml.rels
$tpRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $tempDir "xl\webextensions\_rels\taskpanes.xml.rels"), $tpRels, $utf8)

# 8. xl/webextensions/webextension1.xml
$weXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{b29d4c78-5e4a-4b91-912a-8457fce2a912}">
  <we:reference id="b29d4c78-5e4a-4b91-912a-8457fce2a912" version="1.0.0.0" store="developer" storeType="developer"/>
  <we:alternateReferences/>
  <we:properties>
    <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
'@
[System.IO.File]::WriteAllText((Join-Path $tempDir "xl\webextensions\webextension1.xml"), $weXml, $utf8)

# Créer le fichier .xlsx final sur le bureau de l'utilisateur et dans le projet
$destFiles = @(
    "F:\cop\Assistant_IA_Excel.xlsx",
    "$([Environment]::GetFolderPath('Desktop'))\Assistant_IA_Excel.xlsx"
)

foreach ($dest in $destFiles) {
    if (Test-Path $dest) { Remove-Item $dest -Force }
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, "F:\cop\Assistant_IA_Excel.xlsx")
Copy-Item "F:\cop\Assistant_IA_Excel.xlsx" -Destination "$([Environment]::GetFolderPath('Desktop'))\Assistant_IA_Excel.xlsx" -Force

Write-Host "Classeur généré avec succès : F:\cop\Assistant_IA_Excel.xlsx" -ForegroundColor Green
Write-Host "Copie placée sur le Bureau : $([Environment]::GetFolderPath('Desktop'))\Assistant_IA_Excel.xlsx" -ForegroundColor Green
