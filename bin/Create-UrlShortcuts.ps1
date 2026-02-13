
param(
    [Parameter(Mandatory = $true)] [string] $workspacePath
)

# Initialize error handling and encoding
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "Create-UrlShortcuts.ps1:"

try {
    Write-Host "- workspacePath: $workspacePath"
    
    # Get all open Excel workbooks
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $workbookCount = 0
    
    # Process all open workbooks (excluding local files)
    foreach ($workbook in $excel.Workbooks) {
        $fileUrl = $workbook.FullName
        
        # Skip local files - only process OneDrive/SharePoint URLs
        if ($fileUrl -notmatch '^https?://') {
            Write-Host "  - Skipping local file: $fileUrl"
            continue
        }
        
        $fileName = $workbook.Name
        
        # Create shortcut path with full filename + .url extension
        $shortcutPath = Join-Path $workspacePath "$fileName.url"
        
        # Create Internet Shortcut content
        $content = @"
[InternetShortcut]
URL=$fileUrl
"@
        
        # Create .url file with Internet Shortcut format (UTF-8 encoding)
        $content | Out-File -LiteralPath $shortcutPath -Encoding UTF8 -Force
        
        Write-Host "  - Created: $shortcutPath"
        $workbookCount++
    }
    
    if ($workbookCount -eq 0) {
        Write-Host "[INFO] No OneDrive/SharePoint workbooks found"
    }
    else {
        Write-Host "[SUCCESS] Internet Shortcut files created for $workbookCount workbook(s)"
    }
}
catch {
    Write-Host "[ERROR] $_"
    exit 1
}
