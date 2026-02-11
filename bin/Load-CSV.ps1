param(
    [Parameter(Mandatory = $true)]
    [string]$excelFilePath,
    
    [Parameter(Mandatory = $true)]
    [string]$csvOutputPath
)

# Import common functions
. (Join-Path $PSScriptRoot "Common.ps1")

# Initialize
Initialize-Script $MyInvocation.MyCommand.Name | Out-Null
Write-Host -ForegroundColor Green "- excelFilePath: $($excelFilePath)"
Write-Host -ForegroundColor Green "- csvOutputPath: $($csvOutputPath)"

# Get running Excel instance
$excel = Get-ExcelInstance

try {
    # Check if the file is a .url marker file
    $isUrlFile = [System.IO.Path]::GetExtension($excelFilePath).ToLower() -eq ".url"
    
    if ($isUrlFile) {
        # For .url files, try to find the corresponding Excel file in the same directory as CsvOutputPath
        $csvDir = Split-Path $csvOutputPath -Parent
        $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($CsvOutputPath)
        
        # Look for Excel files matching the CSV folder name (without _csv suffix)
        $possibleFiles = @()
        foreach ($ext in @(".xlsm", ".xlsx", ".xlam")) {
            $possiblePath = Join-Path $csvDir ($baseFileName + $ext)
            if (Test-Path $possiblePath) {
                $possibleFiles += $possiblePath
            }
        }
        
        if ($possibleFiles.Count -eq 0) {
            throw ".URL FILE DETECTED: Cannot find corresponding Excel file in $(Split-Path $CsvOutputPath -Parent). Expected file: $($baseFileName).xlsx|.xlsm|.xlam"
        }
        
        if ($possibleFiles.Count -gt 1) {
            throw ".URL FILE DETECTED: Found multiple Excel files matching $($baseFileName). Please specify the exact file path."
        }
        
        $ExcelFilePath = $possibleFiles[0]
        $fullPath = [System.IO.Path]::GetFullPath($ExcelFilePath)
    }
    else {
        $fullPath = [System.IO.Path]::GetFullPath($excelFilePath)
        
        if (-not (Test-Path $fullPath)) {
            throw "EXCEL FILE NOT FOUND: $($fullPath)"
        }
    }
    
    # Find the workbook in open workbooks
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
    if (Test-Path $csvOutputPath) {
        Remove-Item $csvOutputPath -Recurse -Force
    }

    # Create output directory
    New-Item -ItemType Directory -Force -Path $csvOutputPath | Out-Null
    
    # Activate Excel window
    $shell = New-Object -ComObject WScript.Shell
    $shell.AppActivate($excel.Caption) | Out-Null
    
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
            $sheet.Activate() | Out-Null
            
            # Update status bar
            $excel.StatusBar = "Loading sheet ${currentIndex} of ${sheetsToExportCount}: $sheetName"
            
            # Get the used range from the source sheet
            try {
                $usedRange = $sheet.UsedRange
                if (-not $usedRange) {
                    $usedRange = $null
                }
            }
            catch {
                Write-Host "Warning: Could not get UsedRange from $sheetName"
                $usedRange = $null
            }
            
            # Check if sheet has content
            if ($usedRange -and $usedRange.Cells.Count -gt 0) {
                try {
                    # Get dimensions
                    $rows = $usedRange.Rows.Count
                    $cols = $usedRange.Columns.Count
                    
                    # Skip if empty
                    if ($rows -le 0 -or $cols -le 0) {
                        Write-Host "Sheet is empty: $sheetName"
                        continue
                    }
                    
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
                        # Single row - handle 2D array from Excel
                        $line = @()
                        for ($c = 1; $c -le $cols; $c++) {
                            $value = $allValues[1, $c]
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
                        # Single column - handle 2D array from Excel
                        for ($r = 1; $r -le $rows; $r++) {
                            $value = $allValues[$r, 1]
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
                    $csvFilePath = Join-Path $csvOutputPath $csvFileName
                    
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
                catch {
                    Write-Host "Error processing sheet $sheetName : $_"
                }
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
