# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $bookPath,
    [Parameter(Mandatory = $true)] [string] $imageOutputPath
)

# Import common functions
. (Join-Path $PSScriptRoot "Common.ps1")

try {
    # Initialize
    Initialize-Script "Export-SheetAsPng.ps1" | Out-Null
    Write-Host "- bookPath: $($bookPath)"
    Write-Host "- imageOutputPath: $($imageOutputPath)"

    # Get running Excel instance
    $excel = Get-ExcelInstance

    # Get workbook info and find the workbook
    $macroInfo = Get-BookInfo $bookPath
    $result = Find-VBProject $excel $macroInfo.ResolvedPath $macroInfo.IsAddIn
    $workbook = $result.Workbook
    if ($null -eq $workbook) {
        throw "No workbook open."
    }

    # Clean output directory
    if (Test-Path $imageOutputPath) {
        Remove-Item $imageOutputPath -Recurse -Force
    }

    # Create output directory
    New-Item -ItemType Directory -Force -Path $imageOutputPath | Out-Null

    # Activate Excel window
    $shell = New-Object -ComObject WScript.Shell
    $shell.AppActivate($excel.Caption) | Out-Null

    # Disable user interaction during processing
    $excel.Interactive = $false

    try {
        # Get all sheets
        $sheetCount = $workbook.Sheets.Count
        Write-Host "- Total sheets: $sheetCount"

        # Count sheets that end with .png
        $sheetsToExportCount = 0
        for ($i = 1; $i -le $sheetCount; $i++) {
            $sheet = $workbook.Sheets.Item($i)
            $sheetName = $sheet.Name
            if ($sheetName -match '.*\.png$') {
                $sheetsToExportCount++
            }
        }
        Write-Host "- Sheets ending with .png: $sheetsToExportCount"

        # Process sheets that end with .png
        for ($i = 1; $i -le $sheetCount; $i++) {
            $sheet = $workbook.Sheets.Item($i)
            $sheetName = $sheet.Name

            # Check if sheet name ends with .png
            if ($sheetName -match '.*\.png$') {
                Write-Host "  - Exporting: $sheetName"
                
                # Try to select sheet, but proceed if Select() is not supported
                try {
                    $sheet.Select() | Out-Null
                } catch {
                    # Some sheet types (e.g., chart sheets) may not support Select()
                    # This is not critical for exporting
                }

                # Get print area
                $printArea = $sheet.PageSetup.PrintArea
                Write-Host "    - Print area: $printArea"

                # Determine what range to export
                $rangeToExport = $null
                if (-not [string]::IsNullOrEmpty($printArea)) {
                    # Use print area if defined
                    $rangeToExport = $sheet.Range($printArea)
                    Write-Host "    - Using print area"
                }
                else {
                    # Use UsedRange if print area is not defined
                    $usedRange = $sheet.UsedRange
                    if ($null -ne $usedRange) {
                        $rangeToExport = $usedRange
                        Write-Host "    - Using UsedRange"
                    }
                }

                if ($null -eq $rangeToExport) {
                    throw "Sheet '$sheetName' has no print area or used range."
                }

                # Get the range and copy to clipboard
                $range = $rangeToExport
                $range.Copy() | Out-Null
                Write-Host "    - Copied to clipboard"

                # Create GDI+ image from clipboard
                [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
                [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
                
                # Get image from clipboard
                $image = [System.Windows.Forms.Clipboard]::GetImage()
                
                if ($null -ne $image) {
                    $outputFile = Join-Path $imageOutputPath "$sheetName"
                    $image.Save($outputFile, [System.Drawing.Imaging.ImageFormat]::Png)
                    $image.Dispose()
                    Write-Host "    - Saved to: $outputFile"
                }
                else {
                    throw "Failed to get image from clipboard"
                }
            }
        }

        Write-Host "- Export complete"
    }
    finally {
        # Re-enable user interaction
        $excel.Interactive = $true
    }

}
catch {
    Write-Host "ERROR: $_"
    exit 1
}