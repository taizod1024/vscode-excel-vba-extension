param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelFilePath,
    
    [Parameter(Mandatory = $true)]
    [string]$CsvInputPath
)

# Create Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # Open the workbook
    $workbook = $excel.Workbooks.Open([System.IO.Path]::GetFullPath($ExcelFilePath))
    
    # Get all CSV files from input path
    $csvFiles = Get-ChildItem -Path $CsvInputPath -Filter "*.csv" -File
    
    if ($csvFiles.Count -eq 0) {
        Write-Host "No CSV files found in: $CsvInputPath"
        $workbook.Close($false)
        exit 0
    }
    
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
            Write-Host "Deleting existing sheet: $sheetName"
            $existingSheet.Delete()
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
        
        # Parse CSV content
        $lines = $csvContent -split "`r`n" | Where-Object { $_.Length -gt 0 }
        
        $rowNum = 1
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
            
            # Write to sheet
            $colNum = 1
            foreach ($field in $fields) {
                # Remove surrounding quotes if present
                if ($field -match '^"(.*)"$') {
                    $field = $matches[1]
                }
                $newSheet.Cells.Item($rowNum, $colNum).Value = $field
                $colNum++
            }
            
            $rowNum++
        }
        
        Write-Host "Imported: $sheetName"
    }
    
    # Save the workbook
    $workbook.Save()
    $workbook.Close($false)
    Write-Host "Import completed successfully"
}
catch {
    Write-Error "Error during import: $_"
    exit 1
}
finally {
    # Clean up COM objects
    if ($workbook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($stream) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($stream) | Out-Null
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
