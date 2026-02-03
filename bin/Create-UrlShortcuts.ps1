# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $workspacePath
)

# Initialize error handling and encoding
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host -ForegroundColor Yellow "Create-UrlShortcuts.ps1:"

try {
    Write-Host -ForegroundColor Green "- workspacePath: $workspacePath"
    
    # Get active Excel workbook
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $workbook = $excel.ActiveWorkbook
    $fileName = $workbook.Name
    
    Write-Host -ForegroundColor Green "- fileName: $fileName"
    
    # Remove extension from filename
    $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    
    # Create shortcut path
    $shortcutPath = Join-Path $workspacePath "$fileNameWithoutExt.url"
    
    # Create empty 0-byte file
    $null | Out-File -LiteralPath $shortcutPath -Encoding UTF8 -Force
    
    Write-Host -ForegroundColor Green "  - Created: $shortcutPath"
    Write-Host -ForegroundColor Green "[SUCCESS] Shortcut file created successfully"
}
catch {
    Write-Host -ForegroundColor Red "[ERROR] $_"
    exit 1
}
