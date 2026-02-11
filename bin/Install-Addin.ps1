# Install Excel VBA Extension Addin
# Lifecycle Hook: "postinstall"
# Executed: During 'npm install' (development environment only)
# Purpose: Copy excel-vba-addin.xlam to %APPDATA%\Microsoft\AddIns

$AddinName = "excel-vba-addin.xlam"
$SourceAddinPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "excel\addin\$AddinName")
$AddinFolder = Join-Path $env:APPDATA "Microsoft\AddIns"
$DestAddinPath = Join-Path $AddinFolder $AddinName

if (-Not (Test-Path $SourceAddinPath)) {
    Write-Host "[WARNING] Addin source not found: $SourceAddinPath"
    exit 0
}

if (-Not (Test-Path $AddinFolder)) {
    New-Item -ItemType Directory -Path $AddinFolder -Force | Out-Null
}

try {
    Copy-Item $SourceAddinPath $DestAddinPath -Force
    Write-Host "[INFO] Addin installed: $DestAddinPath"
    exit 0
}
catch {
    Write-Host "[WARNING] Failed to install addin: $_"
    exit 0  # Don't fail npm install if addin setup fails
}
