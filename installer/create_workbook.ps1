# Creer un fichier .xlsx avec WebExtension embarquee
# qui force le chargement de l'add-in a l'ouverture

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$outputPath = Join-Path $env:USERPROFILE 'Desktop\AssistantIA.xlsx'

# Supprimer s'il existe
if (Test-Path $outputPath) { Remove-Item $outputPath -Force }

# D'abord creer un xlsx vide via Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Cells.Item(1,1).Value2 = "Assistant IA Excel - Afrique Francophone"
$ws.Cells.Item(2,1).Value2 = "Ce classeur installe automatiquement l'add-in Assistant IA."
$ws.Cells.Item(3,1).Value2 = "Si le volet ne s'ouvre pas, allez dans Insert > My Add-ins."
$ws.Cells.Item(1,1).Font.Bold = $true
$ws.Cells.Item(1,1).Font.Size = 14
$wb.SaveAs($outputPath, 51) # xlOpenXMLWorkbook
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Start-Sleep -Seconds 2

Write-Host "Base xlsx created: $outputPath"

# Maintenant ouvrir le xlsx comme archive ZIP et ajouter les fichiers WebExtension
$zip = [System.IO.Compression.ZipFile]::Open($outputPath, 'Update')

# === 1. webextension1.xml ===
$weXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
                 id="b29d4c78-5e4a-4b91-912a-8457fce2a912">
  <we:reference id="b29d4c78-5e4a-4b91-912a-8457fce2a912"
                version="1.0.0.0"
                store="https://afriexcel.netlify.app/"
                storeType="SPCatalog"/>
  <we:alternateReferences>
    <we:reference id="b29d4c78-5e4a-4b91-912a-8457fce2a912"
                  version="1.0.0.0"
                  store="EN-US"
                  storeType="OMEX"/>
    <we:reference id="b29d4c78-5e4a-4b91-912a-8457fce2a912"
                  version="1.0.0.0"
                  store=""
                  storeType="Registry"/>
  </we:alternateReferences>
  <we:properties>
    <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot/>
</we:webextension>
'@

$weEntry = $zip.CreateEntry('xl/webextensions/webextension1.xml')
$weWriter = New-Object System.IO.StreamWriter($weEntry.Open())
$weWriter.Write($weXml)
$weWriter.Dispose()
Write-Host "Added webextension1.xml"

# === 2. taskpanes.xml ===
$tpXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <wetp:taskpane dockstate="right" visibility="1" width="350" row="0">
    <wetp:webextensionref r:id="rId1"/>
  </wetp:taskpane>
</wetp:taskpanes>
'@

$tpEntry = $zip.CreateEntry('xl/webextensions/taskpanes.xml')
$tpWriter = New-Object System.IO.StreamWriter($tpEntry.Open())
$tpWriter.Write($tpXml)
$tpWriter.Dispose()
Write-Host "Added taskpanes.xml"

$zip.Dispose()

# === 3. Maintenant il faut modifier [Content_Types].xml et xl/_rels/workbook.xml.rels ===
# Reouvrir le zip
Start-Sleep -Seconds 1
$zip = [System.IO.Compression.ZipFile]::Open($outputPath, 'Update')

# Lire le Content_Types actuel
$ctEntry = $zip.GetEntry('[Content_Types].xml')
$ctReader = New-Object System.IO.StreamReader($ctEntry.Open())
$ctXml = $ctReader.ReadToEnd()
$ctReader.Dispose()

# Ajouter les types pour webextensions
$insertBefore = '</Types>'
$newTypes = @'
  <Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>
  <Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>
</Types>
'@
$ctXml = $ctXml.Replace($insertBefore, $newTypes)

# Supprimer et recreer
$ctEntry.Delete()
$newCtEntry = $zip.CreateEntry('[Content_Types].xml')
$ctWriter = New-Object System.IO.StreamWriter($newCtEntry.Open())
$ctWriter.Write($ctXml)
$ctWriter.Dispose()
Write-Host "Updated [Content_Types].xml"

# Lire le workbook.xml.rels
$relsEntry = $zip.GetEntry('xl/_rels/workbook.xml.rels')
$relsReader = New-Object System.IO.StreamReader($relsEntry.Open())
$relsXml = $relsReader.ReadToEnd()
$relsReader.Dispose()

# Ajouter la relation vers taskpanes.xml
$insertBefore2 = '</Relationships>'
$newRel = @'
  <Relationship Id="rId_wp1" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="webextensions/taskpanes.xml"/>
</Relationships>
'@
$relsXml = $relsXml.Replace($insertBefore2, $newRel)

$relsEntry.Delete()
$newRelsEntry = $zip.CreateEntry('xl/_rels/workbook.xml.rels')
$relsWriter = New-Object System.IO.StreamWriter($newRelsEntry.Open())
$relsWriter.Write($relsXml)
$relsWriter.Dispose()
Write-Host "Updated workbook.xml.rels"

# === 4. Ajouter la relation webextension dans taskpanes.xml.rels ===
$tpRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>
</Relationships>
'@

$tpRelsEntry = $zip.CreateEntry('xl/webextensions/_rels/taskpanes.xml.rels')
$tpRelsWriter = New-Object System.IO.StreamWriter($tpRelsEntry.Open())
$tpRelsWriter.Write($tpRelsXml)
$tpRelsWriter.Dispose()
Write-Host "Added taskpanes.xml.rels"

$zip.Dispose()

Write-Host "`n=== WORKBOOK CREATED ==="
Write-Host "File: $outputPath"
Write-Host "Size: $((Get-Item $outputPath).Length) bytes"
Write-Host "`nOuvrez ce fichier dans Excel."
Write-Host "Le volet Assistant IA devrait s'ouvrir automatiquement!"
