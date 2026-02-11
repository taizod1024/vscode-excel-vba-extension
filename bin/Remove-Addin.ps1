# Remove Excel VBA Extension Addin
# Lifecycle Hook: "preuninstall"
# Executed: During 'npm uninstall' (development environment cleanup)
# Purpose: Remove excel-vba-addin.xlam from %APPDATA%\Microsoft\AddIns
# Note: Not used for Marketplace installations (vscode:uninstall handles cleanup)

$AddinName = "excel-vba-addin.xlam"
$AddinFolder = Join-Path $env:APPDATA "Microsoft\AddIns"
$AddinPath = Join-Path $AddinFolder $AddinName

if (Test-Path $AddinPath) {
    try {
        Remove-Item $AddinPath -Force
        Write-Host "[INFO] Addin removed: $AddinPath"
        exit 0
    }
    catch {
        Write-Host "[ERROR] Failed to remove addin: $_"
        exit 0  # Don't fail uninstall if addin removal fails
    }
}
else {
    Write-Host "[INFO] Addin not found: $AddinPath"
    exit 0
}
