param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelFilePath,
    
    [Parameter(Mandatory = $true)]
    [string]$CsvInputPath
)

# Import common functions
. (Join-Path $PSScriptRoot "Common.ps1")

# Initialize
Initialize-Script $MyInvocation.MyCommand.Name | Out-Null
Write-Host -ForegroundColor Green "- ExcelFilePath: $($ExcelFilePath)"
Write-Host -ForegroundColor Green "- CsvInputPath: $($CsvInputPath)"

# Function to read and parse CSV file
function Read-CsvFile {
    param(
        [string]$CsvFilePath
    )
    
    # Read CSV file as UTF-8
    $stream = New-Object -ComObject ADODB.Stream
    $stream.Type = 2  # Text
    $stream.Charset = "UTF-8"
    $stream.Open()
    $stream.LoadFromFile([System.IO.Path]::GetFullPath($CsvFilePath))
    $csvContent = $stream.ReadText(-1)  # -1 means read all
    $stream.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($stream) | Out-Null
    $stream = $null
    
    # Parse CSV content
    $lines = @($csvContent -split "`r`n" | Where-Object { $_.Length -gt 0 })
    
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
        
        $data += , @($cleanedFields)
    }
    
    return , $data
}

# Function to populate sheet with data
function Update-SheetData {
    param(
        [object]$Sheet,
        [object]$Data
    )
    
    if ($null -eq $Data) { return }
    
    $dataArray = @($Data)
    $rowCount = @($dataArray).Count
    if ($rowCount -eq 0) { return }
    
    # Get max columns
    $maxCols = 0
    foreach ($row in $dataArray) {
        if ($null -ne $row) {
            $colCount = @($row).Count
            if ($colCount -gt $maxCols) { $maxCols = $colCount }
        }
    }
    
    if ($maxCols -eq 0) { return }
    
    # Clear all cells and create COM array
    $Sheet.Cells.Clear() | Out-Null
    $excelArray = New-Object 'object[,]' $rowCount, $maxCols
    
    # Populate array
    for ($r = 0; $r -lt $rowCount; $r++) {
        $fields = @($dataArray[$r])
        for ($c = 0; $c -lt $maxCols; $c++) {
            $excelArray[$r, $c] = if ($c -lt @($fields).Count) { $fields[$c] } else { "" }
        }
    }
    
    # Set values in Excel with text format
    $range = $Sheet.Range("A1").Resize($rowCount, $maxCols)
    $range.NumberFormat = "@"  # Set format to text
    $range.Value2 = $excelArray
}

# Get running Excel instance
$excel = Get-ExcelInstance

try {
    # Check if CSV input path exists
    if (-not (Test-Path $CsvInputPath)) {
        throw "CSV FOLDER NOT FOUND: $($CsvInputPath)"
    }
    
    # Check if the file is a .url marker file
    $isUrlFile = [System.IO.Path]::GetExtension($ExcelFilePath).ToLower() -eq ".url"
    
    if ($isUrlFile) {
        # For .url files, try to find the corresponding Excel file in the same directory
        $csvDir = Split-Path $CsvInputPath -Parent
        $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($CsvInputPath)
        
        # Look for Excel files matching the CSV folder name (without _csv suffix)
        $possibleFiles = @()
        foreach ($ext in @(".xlsm", ".xlsx", ".xlam")) {
            $possiblePath = Join-Path $csvDir ($baseFileName + $ext)
            if (Test-Path $possiblePath) {
                $possibleFiles += $possiblePath
            }
        }
        
        if ($possibleFiles.Count -eq 0) {
            throw ".URL FILE DETECTED: Cannot find corresponding Excel file in $(Split-Path $CsvInputPath -Parent). Expected file: $($baseFileName).xlsx|.xlsm|.xlam"
        }
        
        if ($possibleFiles.Count -gt 1) {
            throw ".URL FILE DETECTED: Found multiple Excel files matching $($baseFileName). Please specify the exact file path."
        }
        
        $ExcelFilePath = $possibleFiles[0]
        $fullPath = [System.IO.Path]::GetFullPath($ExcelFilePath)
    }
    else {
        $fullPath = [System.IO.Path]::GetFullPath($ExcelFilePath)
        
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
    
    # Get all CSV files from input path
    $csvFiles = Get-ChildItem -Path $CsvInputPath -Filter "*.csv" -File | Sort-Object -Property BaseName

    if ($csvFiles.Count -eq 0) {
        Write-Host "No CSV files found in: $CsvInputPath"
        exit 0
    }

    # Activate Excel window
    $shell = New-Object -ComObject WScript.Shell
    $shell.AppActivate($excel.Caption) | Out-Null

    # Disable screen updating for performance
    $excel.ScreenUpdating = $false
    $excel.Interactive = $false
    $originalCalculation = $excel.Calculation
    $excel.Calculation = -4135  # xlCalculationManual
    # Get existing sheet names and their order
    $existingSheetNames = @()
    foreach ($sheet in $workbook.Sheets) {
        $existingSheetNames += $sheet.Name
    }
    
    # Create a hashtable to store CSV data (keyed by CSV filename with .csv extension)
    $csvData = @{}
    foreach ($csvFile in $csvFiles) {
        $csvData[$csvFile.Name] = $csvFile
    }
    
    # Count total sheets to process
    $sheetsToProcessCount = 0
    foreach ($existingSheetName in $existingSheetNames) {
        if ($csvData.ContainsKey($existingSheetName)) {
            $sheetsToProcessCount++
        }
    }
    
    # Add new sheets count
    $newSheetNames = @()
    foreach ($sheetName in $csvData.Keys) {
        if ($existingSheetNames -notcontains $sheetName) {
            $newSheetNames += $sheetName
        }
    }
    $newSheetNames = $newSheetNames | Sort-Object
    $sheetsToProcessCount += $newSheetNames.Count
    
    # Function to import sheet data
    function Import-SheetData {
        param(
            [object]$Sheet,
            [object]$CsvFile,
            [int]$CurrentIndex,
            [int]$TotalCount
        )
        
        $sheetName = $CsvFile.Name
        Write-Host "Importing sheet: $sheetName"
        
        # Temporarily enable screen updating to show sheet selection
        $excel.ScreenUpdating = $true
        $Sheet.Activate()
        $excel.ScreenUpdating = $false
        
        # Update status bar
        $excel.StatusBar = "Saving sheet ${CurrentIndex} of ${TotalCount}: $sheetName"
        
        # Read CSV file and populate sheet
        $data = Read-CsvFile -CsvFilePath $CsvFile.FullName
        Update-SheetData -Sheet $Sheet -Data $data
        
        # Set font for entire sheet to Meiryo UI 9pt and auto-fit row heights
        try {
            $allCells = $Sheet.Cells
            $allCells.Font.Name = "Meiryo UI"
            $allCells.Font.Size = 9
            
            # Auto-fit row heights
            $Sheet.Rows.AutoFit() | Out-Null
            
            Write-Host "Applied formatting: Meiryo UI 9pt and auto-fit row heights"
        }
        catch {
            Write-Host "Warning: Could not apply formatting to $sheetName : $_"
        }
        
        Write-Host "Imported: $sheetName ($($data.Count) rows)"
    }
    
    # Function to convert range to table
    function Convert-RangeToTable {
        param(
            [object]$Sheet
        )
        
        # Find the last used cell
        $usedRange = $Sheet.UsedRange
        if ($usedRange.Rows.Count -gt 0 -and $usedRange.Columns.Count -gt 0) {
            # Create table starting from A1
            $tableRange = $Sheet.Range("A1").Resize($usedRange.Rows.Count, $usedRange.Columns.Count)
            
            # Create table object (ListObject in Excel)
            [void]$Sheet.ListObjects.Add(1, $tableRange, $null, 1)
            
            # Set table style
            $Sheet.ListObjects(1).TableStyle = "TableStyleLight1"
            
            # Freeze first row and first column
            try {
                $Sheet.Activate()
                $Sheet.Range("B2").Select()
                $excel.ActiveWindow.FreezePanes = $true
            }
            catch {
                Write-Host -ForegroundColor Yellow "- Warning: Failed to set freeze panes"
            }
            
            Write-Host "Converted to table: $($Sheet.Name)"
        }
    }
    
    # Process sheets in the order they appear in the workbook
    # First, update existing sheets
    $currentIndex = 0
    foreach ($existingSheetName in $existingSheetNames) {
        if ($csvData.ContainsKey($existingSheetName)) {
            $currentIndex++
            $csvFile = $csvData[$existingSheetName]
            $existingSheet = $workbook.Sheets.Item($existingSheetName)
            
            # Reset freeze panes before clearing
            try {
                $existingSheet.Activate()
                $excel.ActiveWindow.FreezePanes = $false
                $excel.ActiveWindow.SplitRow = 0
                $excel.ActiveWindow.SplitColumn = 0
            }
            catch {
                Write-Host -ForegroundColor Yellow "- Warning: Failed to reset freeze panes"
            }
            
            # Clear existing data
            $existingSheet.Cells.Clear() | Out-Null
            
            # Import the sheet data
            Import-SheetData -Sheet $existingSheet -CsvFile $csvFile -CurrentIndex $currentIndex -TotalCount $sheetsToProcessCount | Out-Null
            
            # Convert range to table
            Convert-RangeToTable -Sheet $existingSheet | Out-Null
        }
    }
    
    # Add new sheets (CSV files that don't exist in the workbook) at the end
    foreach ($sheetName in $newSheetNames) {
        $currentIndex++
        $csvFile = $csvData[$sheetName]
        
        # Create new sheet at the end
        $newSheet = $workbook.Sheets.Add([System.Type]::Missing, $workbook.Sheets($workbook.Sheets.Count))
        $newSheet.Name = $sheetName
        
        # Reset freeze panes for new sheet
        try {
            $newSheet.Activate()
            $excel.ActiveWindow.FreezePanes = $false
            $excel.ActiveWindow.SplitRow = 0
            $excel.ActiveWindow.SplitColumn = 0
        }
        catch {
            Write-Host -ForegroundColor Yellow "- Warning: Failed to reset freeze panes"
        }
        
        # Import the sheet data
        Import-SheetData -Sheet $newSheet -CsvFile $csvFile -CurrentIndex $currentIndex -TotalCount $sheetsToProcessCount | Out-Null
        
        # Convert range to table
        Convert-RangeToTable -Sheet $newSheet | Out-Null
    }
    
    # Clear status bar
    $excel.StatusBar = $false
    
    # Save the workbook
    $workbook.Save()
    
    # Restore settings
    $excel.Calculation = $originalCalculation
    $excel.ScreenUpdating = $true
    $excel.Interactive = $true
    
    Write-Host "Import completed successfully"
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
