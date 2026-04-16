# ═══════════════════════════════════════════════════════════════════
#  RDFA Dashboard — PowerPoint Add-in Installer
#  Roger D. Fields Associates, Inc.
#
#  Run this one line in PowerShell (as any user, no admin needed):
#  irm https://bobsledohio.github.io/rdfa-dashboard/install.ps1 | iex
# ═══════════════════════════════════════════════════════════════════

$ErrorActionPreference = 'Stop'

Write-Host ""
Write-Host "  ┌─────────────────────────────────────────────┐" -ForegroundColor DarkRed
Write-Host "  │   RDFA Dashboard — PowerPoint Add-in        │" -ForegroundColor DarkRed
Write-Host "  │   Roger D. Fields Associates, Inc.          │" -ForegroundColor DarkRed
Write-Host "  └─────────────────────────────────────────────┘" -ForegroundColor DarkRed
Write-Host ""

# ── 1. Create local folder ───────────────────────────────────────
$dest = "$env:APPDATA\RDFA"
if (-not (Test-Path $dest)) {
    New-Item -ItemType Directory -Path $dest | Out-Null
}

# ── 2. Download manifest from GitHub Pages ───────────────────────
Write-Host "  Downloading manifest..." -NoNewline
$manifestUrl  = "https://bobsledohio.github.io/rdfa-dashboard/manifest.xml"
$manifestPath = "$dest\manifest.xml"
try {
    Invoke-WebRequest -Uri $manifestUrl -OutFile $manifestPath -UseBasicParsing
    Write-Host " Done." -ForegroundColor Green
} catch {
    Write-Host " FAILED." -ForegroundColor Red
    Write-Host "  Could not reach $manifestUrl" -ForegroundColor Red
    Write-Host "  Make sure you have internet access and try again." -ForegroundColor Yellow
    exit 1
}

# ── 3. Register add-in with Office ───────────────────────────────
Write-Host "  Registering add-in with Office..." -NoNewline
$regPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer"
if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}
Set-ItemProperty -Path $regPath -Name "RDFA Dashboard" -Value $manifestPath
Write-Host " Done." -ForegroundColor Green

# ── 4. Done ──────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ✔  Installation complete!" -ForegroundColor Green
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor White
Write-Host "  1. Open (or restart) PowerPoint" -ForegroundColor Gray
Write-Host "  2. Insert  →  My Add-ins  →  Developer Add-ins tab" -ForegroundColor Gray
Write-Host "  3. Select  RDFA Dashboard  →  Insert" -ForegroundColor Gray
Write-Host "  4. Resize the panel to fill your slide" -ForegroundColor Gray
Write-Host ""
Write-Host "  Dashboard loads live from rdfa.com — internet required." -ForegroundColor DarkGray
Write-Host ""
