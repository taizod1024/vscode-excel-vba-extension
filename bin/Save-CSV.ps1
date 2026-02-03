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
    
    return $data
}

# Function to populate sheet with data
function Update-SheetData {
    param(
        [object]$Sheet,
        [object[]]$Data
    )
    
    if ($Data.Count -gt 0) {
        $maxCols = ($Data | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
        $rowCount = $Data.Count
        
        # Clear all cells first
        $Sheet.Cells.Clear() | Out-Null
        
        # Create a COM array for Excel
        $excelArray = New-Object 'object[,]' $rowCount, $maxCols
        for ($r = 0; $r -lt $rowCount; $r++) {
            for ($c = 0; $c -lt $maxCols; $c++) {
                if ($c -lt $Data[$r].Count) {
                    $excelArray[$r, $c] = $Data[$r][$c]
                }
                else {
                    $excelArray[$r, $c] = ""
                }
            }
        }
        
        # Set all values at once
        $range = $Sheet.Range("A1").Resize($rowCount, $maxCols)
        $range.Value2 = $excelArray
    }
}

# Get running Excel instance
$excel = Get-ExcelInstance

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
    $csvFiles = Get-ChildItem -Path $CsvInputPath -Filter "*.csv" -File | Sort-Object -Property BaseName

    if ($csvFiles.Count -eq 0) {
        Write-Host "No CSV files found in: $CsvInputPath"
        exit 0
    }

    # Activate Excel window
    $shell = New-Object -ComObject WScript.Shell
    $shell.AppActivate($excel.Caption)

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
            [int]$TotalCount,
            [object]$ExcelApp
        )
        
        $sheetName = $CsvFile.Name
        Write-Host "Importing sheet: $sheetName"
        
        # Temporarily enable screen updating to show sheet selection
        $ExcelApp.ScreenUpdating = $true
        $Sheet.Activate()
        $ExcelApp.ScreenUpdating = $false
        
        # Update status bar
        $ExcelApp.StatusBar = "Saving sheet ${CurrentIndex} of ${TotalCount}: $sheetName"
        
        # Read CSV file and populate sheet
        $data = Read-CsvFile -CsvFilePath $CsvFile.FullName
        Update-SheetData -Sheet $Sheet -Data $data
        
        # Set font for entire sheet to Meiryo UI 9pt and auto-fit row heights
        try {
            $allCells = $Sheet.Cells
            $allCells.Font.Name = "xxxMeiryo UIxx"
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
            [object]$Sheet,
            [object]$ExcelApp
        )
        
        # Find the last used cell
        $usedRange = $Sheet.UsedRange
        if ($usedRange.Rows.Count -gt 0 -and $usedRange.Columns.Count -gt 0) {
            # Create table starting from A1
            $tableRange = $Sheet.Range("A1").Resize($usedRange.Rows.Count, $usedRange.Columns.Count)
            
            # Create table object (ListObject in Excel)
            [void]$Sheet.ListObjects.Add(1, $tableRange, $null, 1)
            
            # Set table style
            $Sheet.ListObjects(1).TableStyle = "TableStyleLight2"
            
            # Freeze first row and first column
            $Sheet.Range("B2").Select()
            $ExcelApp.ActiveWindow.FreezePanes = $true
            
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
            
            # Clear existing data
            $existingSheet.Cells.Clear() | Out-Null
            
            # Import the sheet data
            Import-SheetData -Sheet $existingSheet -CsvFile $csvFile -CurrentIndex $currentIndex -TotalCount $sheetsToProcessCount -ExcelApp $excel | Out-Null
            
            # Convert range to table
            Convert-RangeToTable -Sheet $existingSheet -ExcelApp $excel | Out-Null
        }
    }
    
    # Add new sheets (CSV files that don't exist in the workbook) at the end
    foreach ($sheetName in $newSheetNames) {
        $currentIndex++
        $csvFile = $csvData[$sheetName]
        
        # Create new sheet at the end
        $newSheet = $workbook.Sheets.Add([System.Type]::Missing, $workbook.Sheets($workbook.Sheets.Count))
        $newSheet.Name = $sheetName
        
        # Import the sheet data
        Import-SheetData -Sheet $newSheet -CsvFile $csvFile -CurrentIndex $currentIndex -TotalCount $sheetsToProcessCount -ExcelApp $excel | Out-Null
        
        # Convert range to table
        Convert-RangeToTable -Sheet $newSheet -ExcelApp $excel | Out-Null
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
