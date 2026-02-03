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
    
    # Get all open Excel workbooks
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $workbookCount = 0
    
    # Process all open workbooks (excluding local files)
    foreach ($workbook in $excel.Workbooks) {
        $fileUrl = $workbook.FullName
        
        # Skip local files - only process OneDrive/SharePoint URLs
        if ($fileUrl -notmatch '^https?://') {
            Write-Host -ForegroundColor Yellow "  - Skipping local file: $fileUrl"
            continue
        }
        
        $fileName = $workbook.Name
        
        # Remove extension from filename
        $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        
        # Create shortcut path
        $shortcutPath = Join-Path $workspacePath "$fileNameWithoutExt.url"
        
        # Create Internet Shortcut content
        $content = @"
[InternetShortcut]
URL=$fileUrl
"@
        
        # Create .url file with Internet Shortcut format
        $content | Out-File -LiteralPath $shortcutPath -Encoding UTF8 -Force
        
        Write-Host -ForegroundColor Green "  - Created: $shortcutPath"
        $workbookCount++
    }
    
    if ($workbookCount -eq 0) {
        Write-Host -ForegroundColor Yellow "[INFO] No OneDrive/SharePoint workbooks found"
    }
    else {
        Write-Host -ForegroundColor Green "[SUCCESS] Internet Shortcut files created for $workbookCount workbook(s)"
    }
}
catch {
    Write-Host -ForegroundColor Red "[ERROR] $_"
    exit 1
}
