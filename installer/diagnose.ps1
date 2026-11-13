# === Etape 1: Runtime logging pour diagnostiquer ===
$logPath = Join-Path $env:TEMP 'OfficeAddins.log'
$regPath = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
if (!(Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }

# Verifier RuntimeLogging subkey
$rtPath = Join-Path $regPath 'RuntimeLogging'
if (!(Test-Path $rtPath)) { New-Item -Path $rtPath -Force | Out-Null }
Set-ItemProperty -Path $rtPath -Name '(Default)' -Value $logPath -Force
Write-Host "Runtime log: $logPath"

# === Etape 2: S'assurer que le manifest local est correct ===
$manifestDst = Join-Path $env:LOCALAPPDATA 'AssistantIAExcel\manifest.xml'
Write-Host "Manifest local: $manifestDst"
Write-Host "Manifest existe: $(Test-Path $manifestDst)"

# === Etape 3: Verifier le contenu du manifest ===
if (Test-Path $manifestDst) {
    $xml = [xml](Get-Content $manifestDst -Encoding UTF8)
    Write-Host "Type: $($xml.OfficeApp.GetAttribute('xsi:type'))"
    Write-Host "ID: $($xml.OfficeApp.Id)"
    Write-Host "Name: $($xml.OfficeApp.DisplayName.DefaultValue)"
    $hosts = $xml.OfficeApp.Hosts.Host
    Write-Host "Host: $($hosts.Name)"
}

# === Etape 4: Verifier le registre WEF\Developer ===
Write-Host "`n=== WEF Developer entries ==="
Get-ItemProperty -Path $regPath | Format-List

# === Etape 5: Fermer Excel ===
Write-Host "`nFermeture d'Excel..."
Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3

# === Etape 6: Supprimer le log precedent ===
if (Test-Path $logPath) { Remove-Item $logPath -Force }

# === Etape 7: Lancer Excel ===
Write-Host "Lancement d'Excel..."
Start-Process excel.exe
Start-Sleep -Seconds 8

# === Etape 8: Lire le runtime log ===
Write-Host "`n=== RUNTIME LOG ==="
if (Test-Path $logPath) {
    Get-Content $logPath
} else {
    Write-Host "Pas de log genere - le runtime logging ne fonctionne peut-etre pas"
    Write-Host "Tentative de lecture d'autres emplacements..."
    $altLog = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\OTele\officelog.etl'
    if (Test-Path $altLog) {
        Write-Host "OTele log found: $altLog"
    }
}

Write-Host "`n=== FIN DU DIAGNOSTIC ==="
