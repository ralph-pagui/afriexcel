# APPROCHE: Serveur HTTP local avec Node.js + WebDAV minimal
# Sert le manifest.xml sur localhost avec les bons headers

$serverScript = @'
const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = 8765;
const MANIFEST_DIR = process.argv[2] || ".";

const server = http.createServer((req, res) => {
  console.log(`${req.method} ${req.url}`);
  
  // Handle PROPFIND (WebDAV directory listing)
  if (req.method === "PROPFIND" || (req.method === "GET" && req.url === "/")) {
    // Return XML listing of manifest files
    const files = fs.readdirSync(MANIFEST_DIR).filter(f => f.endsWith(".xml"));
    
    if (req.method === "PROPFIND") {
      let responses = files.map(f => `
        <D:response>
          <D:href>http://localhost:${PORT}/${f}</D:href>
          <D:propstat>
            <D:prop>
              <D:displayname>${f}</D:displayname>
              <D:getcontenttype>application/xml</D:getcontenttype>
            </D:prop>
            <D:status>HTTP/1.1 200 OK</D:status>
          </D:propstat>
        </D:response>`).join("");
      
      res.writeHead(207, {
        "Content-Type": "application/xml; charset=utf-8",
        "Access-Control-Allow-Origin": "*"
      });
      res.end(`<?xml version="1.0" encoding="utf-8"?>
<D:multistatus xmlns:D="DAV:">
  <D:response>
    <D:href>http://localhost:${PORT}/</D:href>
    <D:propstat>
      <D:prop>
        <D:resourcetype><D:collection/></D:resourcetype>
      </D:prop>
      <D:status>HTTP/1.1 200 OK</D:status>
    </D:propstat>
  </D:response>${responses}
</D:multistatus>`);
      return;
    }
    
    // GET / - simple HTML listing
    let html = files.map(f => `<a href="/${f}">${f}</a><br>`).join("");
    res.writeHead(200, {"Content-Type": "text/html"});
    res.end(`<html><body><h1>Manifest Files</h1>${html}</body></html>`);
    return;
  }
  
  // Handle OPTIONS (CORS preflight)
  if (req.method === "OPTIONS") {
    res.writeHead(200, {
      "Allow": "GET, PROPFIND, OPTIONS",
      "DAV": "1",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, PROPFIND, OPTIONS",
      "Access-Control-Allow-Headers": "Depth, Content-Type"
    });
    res.end();
    return;
  }
  
  // Serve manifest files
  const filePath = path.join(MANIFEST_DIR, req.url.replace(/^\//, ""));
  if (fs.existsSync(filePath) && filePath.endsWith(".xml")) {
    const content = fs.readFileSync(filePath, "utf8");
    res.writeHead(200, {
      "Content-Type": "application/xml; charset=utf-8",
      "Access-Control-Allow-Origin": "*"
    });
    res.end(content);
  } else {
    res.writeHead(404);
    res.end("Not found");
  }
});

server.listen(PORT, () => {
  console.log(`WebDAV catalog server running at http://localhost:${PORT}/`);
  console.log(`Serving manifests from: ${MANIFEST_DIR}`);
});
'@

# Ecrire le serveur
$serverPath = Join-Path $env:LOCALAPPDATA 'AssistantIAExcel\server.js'
[System.IO.File]::WriteAllText($serverPath, $serverScript, [System.Text.Encoding]::UTF8)
Write-Host "Server script written to: $serverPath"

# Lancer le serveur en arriere-plan
$manifestDir = Join-Path $env:LOCALAPPDATA 'AssistantIAExcel'
Write-Host "Starting local WebDAV server..."
$proc = Start-Process -FilePath "node" -ArgumentList "`"$serverPath`" `"$manifestDir`"" -PassThru -WindowStyle Hidden
Write-Host "Server PID: $($proc.Id)"
Start-Sleep -Seconds 2

# Verifier que le serveur tourne
try {
    $response = Invoke-WebRequest -Uri "http://localhost:8765/manifest.xml" -UseBasicParsing -TimeoutSec 5
    Write-Host "Server OK! Status: $($response.StatusCode)"
    Write-Host "Content preview: $($response.Content.Substring(0, [Math]::Min(200, $response.Content.Length)))"
} catch {
    Write-Host "Server check failed: $($_.Exception.Message)"
}

# Enregistrer http://localhost:8765/ comme TrustedCatalog
Write-Host "`nRegistering localhost catalog..."
$regPath = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\3'
if (!(Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
Set-ItemProperty -Path $regPath -Name 'Url' -Value 'http://localhost:8765/' -Type String
Set-ItemProperty -Path $regPath -Name 'Flags' -Value 1 -Type DWord
Set-ItemProperty -Path $regPath -Name 'Name' -Value 'AssistantIA-Local' -Type String
Write-Host "Catalog registered!"

# Fermer et relancer Excel
Write-Host "`nRestarting Excel..."
Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3
Start-Process excel.exe
Start-Sleep -Seconds 5

Write-Host "`n=== DONE ==="
Write-Host "Le serveur local tourne sur http://localhost:8765/"
Write-Host "Allez dans Insert > My Add-ins > SHARED FOLDER"
Write-Host "IMPORTANT: Ne fermez pas cette fenetre (le serveur s'arreterait)"
