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
                }
                catch {
                    # Some sheet types (e.g., chart sheets) may not support Select()
                    # This is not critical for exporting
                }

                # Get page break preview range
                # Priority: PrintArea > Page breaks > UsedRange
                
                $rangeToExport = $null
                
                # Check PageSetup print area first - this is the page break preview blue border
                $printArea = $sheet.PageSetup.PrintArea
                if (-not [string]::IsNullOrEmpty($printArea)) {
                    Write-Host "    - Print area defined: $printArea"
                    try {
                        $rangeToExport = $sheet.Range($printArea)
                        $printMaxRow = $rangeToExport.Row + $rangeToExport.Rows.Count - 1
                        $printMaxCol = $rangeToExport.Column + $rangeToExport.Columns.Count - 1
                        Write-Host "    - Using print area: rows 1-$printMaxRow, columns 1-$printMaxCol"
                    }
                    catch {
                        Write-Host "    - Warning: Could not parse print area"
                        $rangeToExport = $null
                    }
                }
                
                # If no print area, fall back to page breaks or UsedRange
                if ($null -eq $rangeToExport) {
                    $hPageBreaks = $sheet.HPageBreaks.Count
                    $vPageBreaks = $sheet.VPageBreaks.Count
                    Write-Host "    - Horizontal page breaks: $hPageBreaks, Vertical page breaks: $vPageBreaks"

                    $maxRow = 1
                    $maxCol = 1

                    # Get all horizontal page break positions
                    if ($hPageBreaks -gt 0) {
                        for ($hpIdx = 1; $hpIdx -le $hPageBreaks; $hpIdx++) {
                            $pageBreak = $sheet.HPageBreaks($hpIdx)
                            $breakRow = $pageBreak.Location.Row
                            Write-Host "    - Horizontal page break $hpIdx at row: $breakRow"
                            if ($breakRow -gt $maxRow) {
                                $maxRow = $breakRow
                            }
                        }
                        Write-Host "    - Using maximum horizontal page break row: $maxRow"
                    }

                    # Get all vertical page break positions
                    if ($vPageBreaks -gt 0) {
                        for ($vpIdx = 1; $vpIdx -le $vPageBreaks; $vpIdx++) {
                            $pageBreak = $sheet.VPageBreaks($vpIdx)
                            $breakCol = $pageBreak.Location.Column
                            Write-Host "    - Vertical page break $vpIdx at column: $breakCol"
                            if ($breakCol -gt $maxCol) {
                                $maxCol = $breakCol
                            }
                        }
                        Write-Host "    - Using maximum vertical page break column: $maxCol"
                    }

                    # Check shapes extent
                    if ($sheet.Shapes.Count -gt 0) {
                        Write-Host "    - Sheet has $($sheet.Shapes.Count) shapes"
                        foreach ($shape in $sheet.Shapes) {
                            try {
                                $shapeBottom = $shape.Top + $shape.Height
                                $shapeRight = $shape.Left + $shape.Width
                                
                                # Find the row at the bottom of the shape
                                for ($r = 1; $r -le 1000; $r++) {
                                    $cell = $sheet.Cells.Item($r, 1)
                                    if ($cell.Top -gt $shapeBottom) {
                                        $shapeRow = $r - 1
                                        if ($shapeRow -lt 1) { $shapeRow = 1 }
                                        if ($shapeRow -gt $maxRow) { $maxRow = $shapeRow }
                                        break
                                    }
                                }
                                
                                # Find the column at the right of the shape
                                for ($c = 1; $c -le 256; $c++) {
                                    $cell = $sheet.Cells.Item(1, $c)
                                    if ($cell.Left -gt $shapeRight) {
                                        $shapeCol = $c - 1
                                        if ($shapeCol -lt 1) { $shapeCol = 1 }
                                        if ($shapeCol -gt $maxCol) { $maxCol = $shapeCol }
                                        break
                                    }
                                }
                            }
                            catch {
                                # Skip shapes that can't be processed
                            }
                        }
                        Write-Host "    - Shapes extend to row: $maxRow, column: $maxCol"
                    }

                    # Fall back to UsedRange if no page breaks and no shapes
                    $usedRange = $sheet.UsedRange
                    if ($null -ne $usedRange) {
                        $usedMaxRow = $usedRange.Row + $usedRange.Rows.Count - 1
                        $usedMaxCol = $usedRange.Column + $usedRange.Columns.Count - 1
                        Write-Host "    - UsedRange: rows 1-$usedMaxRow, columns 1-$usedMaxCol"
                        
                        if ($hPageBreaks -eq 0 -and $usedMaxRow -gt $maxRow) {
                            $maxRow = $usedMaxRow
                        }
                        if ($vPageBreaks -eq 0 -and $usedMaxCol -gt $maxCol) {
                            $maxCol = $usedMaxCol
                        }
                    }

                    # Default minimum
                    if ($maxRow -lt 1) { $maxRow = 50 }
                    if ($maxCol -lt 1) { $maxCol = 10 }

                    Write-Host "    - Final export range: rows 1-$maxRow, columns 1-$maxCol"
                    $rangeToExport = $sheet.Range($sheet.Cells(1, 1), $sheet.Cells($maxRow, $maxCol))
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
