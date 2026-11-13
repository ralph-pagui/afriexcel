$dst = Join-Path $env:LOCALAPPDATA 'AssistantIAExcel'
Write-Host "Destination: $dst"
if (!(Test-Path $dst)) {
    New-Item -Path $dst -ItemType Directory -Force | Out-Null
    Write-Host "Directory created"
}
Copy-Item -Path 'f:\cop\addin-frontend\dist\manifest.xml' -Destination (Join-Path $dst 'manifest.xml') -Force
$exists = Test-Path (Join-Path $dst 'manifest.xml')
Write-Host "Manifest exists: $exists"
if ($exists) {
    $content = Get-Content (Join-Path $dst 'manifest.xml') -Head 3
    Write-Host $content
}
