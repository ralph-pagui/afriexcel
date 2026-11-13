# Approche FINALE: Utiliser Excel COM pour installer l'add-in directement
# comme le fait le bouton "Insert > My Add-ins > SHARED FOLDER > Add"

# Fermer Excel d'abord
Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3

# Version Office
$c2r = Get-ItemProperty 'HKLM:\Software\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
if ($c2r) {
    Write-Host "Office Version: $($c2r.ClientVersionToReport)"
    Write-Host "Platform: $($c2r.Platform)"
    Write-Host "Channel: $($c2r.CDNBaseUrl)"
}

# Lancer Excel via COM
Write-Host "`nCreation de l'instance Excel COM..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
Write-Host "Excel version: $($excel.Version)"
Write-Host "Excel build: $($excel.Build)"

# Ouvrir un classeur vide
$wb = $excel.Workbooks.Add()
Start-Sleep -Seconds 2

# Lister les add-ins web existants
Write-Host "`n=== COMAddIns ==="
foreach ($addin in $excel.COMAddIns) {
    Write-Host "  COM: $($addin.Description) - $($addin.ProgId) - Connected: $($addin.Connect)"
}

# Verifier s'il y a une methode pour installer les web add-ins
Write-Host "`n=== Excel Application Members (Web*) ==="
$excel | Get-Member -Name "*Web*","*Addin*","*AddIn*","*Extension*" | ForEach-Object {
    Write-Host "  $($_.MemberType): $($_.Name)"
}

# Essayer d'acceder aux DefaultWebOptions ou similaire
Write-Host "`n=== Workbook WebExtensions check ==="
try {
    $extensions = $wb.ContentTypeProperties
    Write-Host "ContentTypeProperties count: $($extensions.Count)"
} catch {
    Write-Host "No ContentTypeProperties"
}

# Ne pas fermer Excel - laisser ouvert pour debug
Write-Host "`nExcel est ouvert. NE PAS FERMER."
Write-Host "Verifiez maintenant Insert > My Add-ins"
