param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelFilePath,
    
    [Parameter(Mandatory = $true)]
    [string]$CsvOutputPath
)

# Import common functions
. (Join-Path $PSScriptRoot "Common.ps1")

# Initialize
Initialize-Script $MyInvocation.MyCommand.Name | Out-Null
Write-Host -ForegroundColor Green "- ExcelFilePath: $($ExcelFilePath)"
Write-Host -ForegroundColor Green "- CsvOutputPath: $($CsvOutputPath)"

# Get running Excel instance
$excel = Get-ExcelInstance

try {
    # Check if Excel file exists
    if (-not (Test-Path $ExcelFilePath)) {
        throw "EXCEL FILE NOT FOUND: $($ExcelFilePath)"
    }
    
    # Check if the workbook is open in Excel
    $fullPath = [System.IO.Path]::GetFullPath($ExcelFilePath)
    
    $workbook = $null
    foreach ($openWorkbook in $excel.Workbooks) {
        if ($openWorkbook.FullName -eq $fullPath) {
            $workbook = $openWorkbook
            break
        }
    }
    
    if ($null -eq $workbook) {
        throw "EXCEL WORKBOOK NOT OPEN: $($fullPath) is not currently open in Excel"
    }

    # Clean output directory
    if (Test-Path $CsvOutputPath) {
        Remove-Item $CsvOutputPath -Recurse -Force
    }

    # Create output directory
    New-Item -ItemType Directory -Force -Path $CsvOutputPath | Out-Null
    
    # Activate Excel window
    $shell = New-Object -ComObject WScript.Shell
    $shell.AppActivate($excel.Caption)
    
    # Disable user interaction during processing
    $excel.Interactive = $false
    
    # Get all sheets
    $sheetCount = $workbook.Sheets.Count
    
    # Count sheets that end with .csv
    $sheetsToExportCount = 0
    for ($i = 1; $i -le $sheetCount; $i++) {
        $sheet = $workbook.Sheets.Item($i)
        if ($sheet.Name -match "\.csv$") {
            $sheetsToExportCount++
        }
    }
    
    # Export all sheets that end with .csv
    $currentIndex = 0
    for ($i = 1; $i -le $sheetCount; $i++) {
        $sheet = $workbook.Sheets.Item($i)
        $sheetName = $sheet.Name
        
        # Only process sheets that end with .csv
        if ($sheetName -match "\.csv$") {
            $currentIndex++
            Write-Host "Exporting sheet: $sheetName"
            
            # Activate the sheet
            $sheet.Activate()
            
            # Update status bar
            $excel.StatusBar = "Loading sheet ${currentIndex} of ${sheetsToExportCount}: $sheetName"
            
            # Get the used range from the source sheet
            $usedRange = $sheet.UsedRange
            if ($usedRange -and $usedRange.Cells.Count -gt 0) {
                # Get dimensions
                $rows = $usedRange.Rows.Count
                $cols = $usedRange.Columns.Count
                
                # Get all values at once using Value2
                $allValues = $usedRange.Value2
                
                # Create CSV content
                $csvLines = @()
                
                # Handle different array dimensions
                if ($rows -eq 1 -and $cols -eq 1) {
                    # Single cell
                    $value = if ( $null -eq $allValues) { "" } else { $allValues }
                    $value = $value.ToString()
                    if ($value -match '[",\r\n]') {
                        $value = '"' + ($value -replace '"', '""') + '"'
                    }
                    $csvLines += $value
                }
                elseif ($rows -eq 1) {
                    # Single row - allValues is 1D array
                    $line = @()
                    for ($c = 0; $c -lt $cols; $c++) {
                        $value = $allValues[$c]
                        if ($null -eq $value) {
                            $value = ""
                        }
                        $value = $value.ToString()
                        if ($value -match '[",\r\n]') {
                            $value = '"' + ($value -replace '"', '""') + '"'
                        }
                        $line += $value
                    }
                    $csvLines += $line -join ","
                }
                elseif ($cols -eq 1) {
                    # Single column - allValues is 1D array
                    for ($r = 0; $r -lt $rows; $r++) {
                        $value = $allValues[$r]
                        if ($null -eq $value) {
                            $value = ""
                        }
                        $value = $value.ToString()
                        if ($value -match '[",\r\n]') {
                            $value = '"' + ($value -replace '"', '""') + '"'
                        }
                        $csvLines += $value
                    }
                }
                else {
                    # Multiple rows and columns - allValues is 2D array (1-based)
                    for ($r = 1; $r -le $rows; $r++) {
                        $line = @()
                        for ($c = 1; $c -le $cols; $c++) {
                            $value = $allValues[$r, $c]
                            if ($null -eq $value) {
                                $value = ""
                            }
                            $value = $value.ToString()
                            if ($value -match '[",\r\n]') {
                                $value = '"' + ($value -replace '"', '""') + '"'
                            }
                            $line += $value
                        }
                        $csvLines += $line -join ","
                    }
                }
                
                # Save as UTF-8 CSV using ADODB.Stream
                $csvFileName = $sheetName
                $csvFilePath = Join-Path $CsvOutputPath $csvFileName
                
                $stream = New-Object -ComObject ADODB.Stream
                $stream.Type = 2  # Text
                $stream.Charset = "UTF-8"
                $stream.Open()
                $stream.WriteText(($csvLines -join "`r`n"), 1)  # 1 means add line terminator
                $stream.SaveToFile($csvFilePath, 2)  # 2 means overwrite
                $stream.Close()
                
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($stream) | Out-Null
                
                Write-Host "Exported: $csvFileName ($($csvLines.Count) rows)"
            }
            else {
                Write-Host "Sheet is empty: $sheetName"
            }
        }
    }
    
    # Clear status bar
    $excel.StatusBar = $false
    
    # Re-enable user interaction
    $excel.Interactive = $true
    
    Write-Host "Export completed successfully"
}
catch {
    [Console]::Error.WriteLine("$($_)")
    exit 1
}
finally {
    # Clean up COM objects (but do not quit Excel to keep it open)
    if ($workbook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
