# === APPROCHE DIRECTE: Sideload via COM + manipulation fichier Excel ===

$manifestPath = Join-Path $env:LOCALAPPDATA 'AssistantIAExcel\manifest.xml'
Write-Host "Manifest: $manifestPath"
Write-Host "Exists: $(Test-Path $manifestPath)"

# === Fix RuntimeLogging - format correct ===
$rtKey = 'HKCU:\Software\Microsoft\Office\16.0\Wef'
$logFile = [System.IO.Path]::Combine($env:USERPROFILE, 'Desktop', 'OfficeAddins.log')
if (!(Test-Path $rtKey)) { New-Item -Path $rtKey -Force | Out-Null }
New-ItemProperty -Path $rtKey -Name 'RuntimeLogging' -Value $logFile -PropertyType String -Force | Out-Null
Write-Host "RuntimeLogging set to: $logFile"

# === Fermer Excel ===
Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3

# === Nettoyer completement le cache WEF ===
$wefPaths = @(
    (Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'),
    (Join-Path $env:LOCALAPPDATA 'Microsoft\Office\Wef')
)
foreach ($p in $wefPaths) {
    if (Test-Path $p) {
        Remove-Item "$p\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Cleaned: $p"
    }
}

# === Lancer Excel ===
Write-Host "`nLancement d'Excel..."
Start-Process excel.exe
Start-Sleep -Seconds 10

# === Verifier le runtime log ===
Write-Host "`n=== RUNTIME LOG ==="
if (Test-Path $logFile) {
    $content = Get-Content $logFile -ErrorAction SilentlyContinue
    if ($content) {
        $content | ForEach-Object { Write-Host $_ }
    } else {
        Write-Host "(log vide)"
    }
} else {
    Write-Host "(log non cree)"
    
    # Chercher le log dans d'autres emplacements possibles
    $altPaths = @(
        (Join-Path $env:TEMP 'OfficeAddins.log'),
        (Join-Path $env:USERPROFILE 'OfficeAddins.log'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Office\OfficeAddins.log')
    )
    foreach ($alt in $altPaths) {
        if (Test-Path $alt) {
            Write-Host "Found log at: $alt"
            Get-Content $alt | ForEach-Object { Write-Host $_ }
        }
    }
}

# === Lister les WEF entries ===
Write-Host "`n=== ALL WEF REGISTRY ==="
$wefKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF'
if (Test-Path $wefKey) {
    Get-ChildItem $wefKey -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        Write-Host "`n[$($_.PSChildName)]"
        $_ | Get-ItemProperty -ErrorAction SilentlyContinue | Format-List *
    }
}

Write-Host "`n=== Office version info ==="
$clickToRun = 'HKLM:\Software\Microsoft\Office\ClickToRun\Configuration'
if (Test-Path $clickToRun) {
    $ver = Get-ItemProperty $clickToRun -ErrorAction SilentlyContinue
    Write-Host "Platform: $($ver.Platform)"
    Write-Host "ClientVersionToReport: $($ver.ClientVersionToReport)"
    Write-Host "AudienceId: $($ver.AudienceId)"
}
