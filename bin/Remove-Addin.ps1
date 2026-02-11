# Remove Excel VBA Extension Addin
# Lifecycle Hook: "vscode:uninstall"
# Executed: By VS Code after extension uninstall
# Purpose: Clean up excel-vba-addin.xlam from %APPDATA%\Microsoft\AddIns

$AddinName = "excel-vba-addin.xlam"
$AddinFolder = Join-Path $env:APPDATA "Microsoft\AddIns"
$AddinPath = Join-Path $AddinFolder $AddinName
$LogPath = Join-Path $env:APPDATA "Microsoft\AddIns\remove-addin.log"

function Log {
    param([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] $Message"
    Write-Host $LogMessage
    Add-Content -Path $LogPath -Value $LogMessage -Force
}

try {
    Log "[START] Remove-Addin.ps1 execution started"
    Log "Addin path: $AddinPath"
    
    if (Test-Path $AddinPath) {
        Remove-Item $AddinPath -Force
        Log "[SUCCESS] Addin removed: $AddinPath"
        exit 0
    }
    else {
        Log "[INFO] Addin not found: $AddinPath"
        exit 0
    }
}
catch {
    Log "[ERROR] Failed to remove addin: $_"
    exit 0  # Don't fail the uninstall process
}
