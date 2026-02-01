param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelFilePath,
    
    [Parameter(Mandatory = $true)]
    [string]$CsvInputPath
)

# set error action
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Get running Excel instance
$excel = $null
try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
}
catch {
    throw "NO EXCEL FOUND. Please Open Excel first."
}

try {
    # Check if CSV input path exists
    if (-not (Test-Path $CsvInputPath)) {
        throw "CSV FOLDER NOT FOUND: $($CsvInputPath)"
    }
    
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
    
    # Get all CSV files from input path
    $csvFiles = Get-ChildItem -Path $CsvInputPath -Filter "*.csv" -File

    if ($csvFiles.Count -eq 0) {
        Write-Host "No CSV files found in: $CsvInputPath"
        exit 0
    }

    # Disable screen updating for performance
    $excel.ScreenUpdating = $false
    
    # Process each CSV file
    foreach ($csvFile in $csvFiles) {
        $sheetName = $csvFile.BaseName
        Write-Host "Importing sheet: $sheetName"
        
        # Check if sheet with same name exists and delete it
        $existingSheet = $null
        try {
            $existingSheet = $workbook.Sheets.Item($sheetName)
        }
        catch {
            # Sheet doesn't exist, which is fine
        }
        
        if ($existingSheet) {
            # Disable alerts to prevent confirmation dialog
            $excel.DisplayAlerts = $false
            $existingSheet.Delete()
            $excel.DisplayAlerts = $true
        }
        
        # Create new sheet
        $newSheet = $workbook.Sheets.Add([System.Type]::Missing, $workbook.Sheets(1))
        $newSheet.Name = $sheetName
        
        # Read CSV file as UTF-8
        $stream = New-Object -ComObject ADODB.Stream
        $stream.Type = 2  # Text
        $stream.Charset = "UTF-8"
        $stream.Open()
        $stream.LoadFromFile([System.IO.Path]::GetFullPath($csvFile.FullName))
        $csvContent = $stream.ReadText(-1)  # -1 means read all
        $stream.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($stream) | Out-Null
        $stream = $null
        
        # Parse CSV content
        $lines = $csvContent -split "`r`n" | Where-Object { $_.Length -gt 0 }
        
        # Store all data in a 2D array
        $data = @()
        
        foreach ($line in $lines) {
            # Simple CSV parsing (handles quoted fields with commas)
            $fields = @()
            $current = ""
            $inQuotes = $false
            
            for ($i = 0; $i -lt $line.Length; $i++) {
                $char = $line[$i]
                
                if ($char -eq '"') {
                    if ($inQuotes -and $i + 1 -lt $line.Length -and $line[$i + 1] -eq '"') {
                        # Escaped quote
                        $current += '"'
                        $i++
                    }
                    else {
                        # Toggle quote state
                        $inQuotes = -not $inQuotes
                    }
                }
                elseif ($char -eq ',' -and -not $inQuotes) {
                    $fields += $current
                    $current = ""
                }
                else {
                    $current += $char
                }
            }
            $fields += $current
            
            # Clean up field values (remove surrounding quotes if present)
            $cleanedFields = @()
            foreach ($field in $fields) {
                if ($field -match '^"(.*)"$') {
                    $cleanedFields += $matches[1]
                }
                else {
                    $cleanedFields += $field
                }
            }
            
            $data += , $cleanedFields
        }
        
        # Write all data to sheet at once using array formula
        if ($data.Count -gt 0) {
            $maxCols = ($data | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
            $rowCount = $data.Count
            
            # Get range and populate with data
            $range = $newSheet.Range("A1").Resize($rowCount, $maxCols)
            
            # Create a COM array for Excel
            $excelArray = New-Object 'object[,]' $rowCount, $maxCols
            for ($r = 0; $r -lt $rowCount; $r++) {
                for ($c = 0; $c -lt $maxCols; $c++) {
                    if ($c -lt $data[$r].Count) {
                        $excelArray[$r, $c] = $data[$r][$c]
                    }
                    else {
                        $excelArray[$r, $c] = ""
                    }
                }
            }
            
            # Set all values at once
            $range.Value2 = $excelArray
        }
        
        Write-Host "Imported: $sheetName"
    }
    
    # Save the workbook
    $workbook.Save()
    
    # Re-enable screen updating
    $excel.ScreenUpdating = $true
    
    Write-Host "Import completed successfully"
}
catch {
    Write-Error "Error during import: $_"
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
