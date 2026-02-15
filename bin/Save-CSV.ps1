param(
    [Parameter(Mandatory = $true)]
    [string]$bookPath,
    
    [Parameter(Mandatory = $true)]
    [string]$csvInputPath
)

# Import common functions
. (Join-Path $PSScriptRoot "Common.ps1")

# Initialize
Initialize-Script $MyInvocation.MyCommand.Name | Out-Null
Write-Host "- bookPath: $($bookPath)"
Write-Host "- csvInputPath: $($csvInputPath)"

# Constants
$DEFAULT_FONT_NAME = "Meiryo UI"
$DEFAULT_FONT_SIZE = 9

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
    if (-not (Test-Path $csvInputPath)) {
        throw "CSV folder not found: $csvInputPath"
    }
    
    # Check if the file is a .url marker file
    $isUrlFile = [System.IO.Path]::GetExtension($bookPath).ToLower() -eq ".url"
    
    if ($isUrlFile) {
        # Convert folder name (from csvInputPath) to Excel file name (あああ_xlsx -> あああ.xlsx)
        $csvDir = Split-Path $csvInputPath -Parent           # /path/to/あああ_xlsx
        $folderName = Split-Path $csvDir -Leaf                # あああ_xlsx
        
        # Match the folder name pattern
        $match = $folderName -imatch '^(.+?)_(xlsm|xlsx|xlam)$'
        if (-not $match) {
            throw "Invalid folder name format: $folderName"
        }
        
        $expectedFileName = "$($matches[1]).$($matches[2])"
        $fullPath = ""
    }
    else {
        $fullPath = [System.IO.Path]::GetFullPath($bookPath)
        $expectedFileName = [System.IO.Path]::GetFileNameWithoutExtension($bookPath)
        
        if (-not (Test-Path $fullPath)) {
            throw "Workbook file not found: $fullPath"
        }
    }
    
    # Find the workbook in open workbooks
    $workbook = $null
    foreach ($openWorkbook in $excel.Workbooks) {
        $openBookFullName = $openWorkbook.FullName
        $openBookName = $openWorkbook.Name
        
        if ($isUrlFile) {
            # For .url files, match by filename (without extension)
            $openBookBaseName = [System.IO.Path]::GetFileNameWithoutExtension($openBookName)
            if ($openBookBaseName -ieq [System.IO.Path]::GetFileNameWithoutExtension($expectedFileName)) {
                $workbook = $openWorkbook
                break
            }
        }
        else {
            # For regular files, match by full path
            if ($openBookFullName -eq $fullPath) {
                $workbook = $openWorkbook
                break
            }
        }
    }
    
    if ($null -eq $workbook) {
        if ($isUrlFile) {
            throw "Workbook not open: $expectedFileName"
        }
        else {
            throw "Workbook not open: $fullPath"
        }
    }
    
    # Get all CSV files from input path
    $csvFiles = Get-ChildItem -Path $csvInputPath -Filter "*.csv" -File | Sort-Object -Property BaseName

    if ($csvFiles.Count -eq 0) {
        Write-Host "No CSV files found in: $csvInputPath"
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
    $csvDataMap = @{}
    foreach ($csvFile in $csvFiles) {
        $csvDataMap[$csvFile.Name] = $csvFile
    }
    
    # Count total sheets to process
    $sheetsToProcessCount = 0
    foreach ($existingSheetName in $existingSheetNames) {
        if ($csvDataMap.ContainsKey($existingSheetName)) {
            $sheetsToProcessCount++
        }
    }
    
    # Add new sheets count
    $newSheetNames = @()
    foreach ($csvFileName in $csvDataMap.Keys) {
        if ($existingSheetNames -notcontains $csvFileName) {
            $newSheetNames += $csvFileName
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
        
        # Set font for entire sheet to $DEFAULT_FONT_NAME $($DEFAULT_FONT_SIZE)pt and auto-fit row heights
        try {
            $allCells = $Sheet.Cells
            $allCells.Font.Name = $DEFAULT_FONT_NAME
            $allCells.Font.Size = $DEFAULT_FONT_SIZE
            
            # Auto-fit row heights
            $Sheet.Rows.AutoFit() | Out-Null
            
            Write-Host "Applied formatting: $($DEFAULT_FONT_NAME) $($DEFAULT_FONT_SIZE)pt and auto-fit row heights"
        }
        catch {
            Write-Host "Warning: Could not apply formatting to $sheetName : $_"
        }
        
        Write-Host "Imported: $sheetName ($($data.Count) rows)"
    }
    
    # Function to reset freeze panes
    function Reset-FreezePanes {
        param(
            [object]$Sheet
        )
        
        try {
            $Sheet.Activate()
            $excel.ActiveWindow.FreezePanes = $false
            $excel.ActiveWindow.SplitRow = 0
            $excel.ActiveWindow.SplitColumn = 0
        }
        catch {
            Write-Host "- Warning: Failed to reset freeze panes"
        }
    }
    
    # Function to set freeze panes
    function Set-FreezePanes {
        param(
            [object]$Sheet
        )
        
        try {
            $Sheet.Activate()
            $Sheet.Range("B2").Select()
            $excel.ActiveWindow.FreezePanes = $true
        }
        catch {
            Write-Host "- Warning: Failed to set freeze panes"
        }
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
            try {
                # Use GetType().InvokeMember for more reliable COM invocation with optional parameters
                $listObjects = $Sheet.ListObjects
                [void]$listObjects.Add(1, $tableRange, $null, 1, $null)
            }
            catch {
                Write-Host "- Warning: Could not create table: $_"
                return
            }
            
            # Set table style
            $Sheet.ListObjects(1).TableStyle = "TableStyleLight4"
            
            # Set freeze panes
            Set-FreezePanes -Sheet $Sheet
            
            Write-Host "Converted to table: $($Sheet.Name)"
        }
    }
    
    # Function to import sheet data and convert to table
    function Invoke-SheetDataImport {
        param(
            [object]$Sheet,
            [object]$CsvFile,
            [int]$CurrentIndex,
            [int]$TotalCount
        )
        
        # Import the sheet data
        Import-SheetData -Sheet $Sheet -CsvFile $CsvFile -CurrentIndex $CurrentIndex -TotalCount $TotalCount | Out-Null
        
        # Convert range to table
        Convert-RangeToTable -Sheet $Sheet | Out-Null
    }
    
    # Process sheets in the order they appear in the workbook
    # First, update existing sheets
    $currentIndex = 0
    foreach ($existingSheetName in $existingSheetNames) {
        if ($csvDataMap.ContainsKey($existingSheetName)) {
            $currentIndex++
            $csvFile = $csvDataMap[$existingSheetName]
            $existingSheet = $workbook.Sheets.Item($existingSheetName)
            
            # Reset freeze panes before clearing
            Reset-FreezePanes -Sheet $existingSheet
            
            # Clear existing data
            $existingSheet.Cells.Clear() | Out-Null
            
            # Import sheet data and convert to table
            Invoke-SheetDataImport -Sheet $existingSheet -CsvFile $csvFile -CurrentIndex $currentIndex -TotalCount $sheetsToProcessCount
        }
    }
    
    # Add new sheets (CSV files that don't exist in the workbook) at the end
    foreach ($csvFileName in $newSheetNames) {
        $currentIndex++
        $csvFile = $csvDataMap[$csvFileName]
        
        # Create new sheet at the end
        $lastSheet = $workbook.Sheets($workbook.Sheets.Count)
        $newSheet = $workbook.Sheets.Add($null, $lastSheet)
        $newSheet.Name = $csvFileName
        
        # Reset freeze panes for new sheet
        Reset-FreezePanes -Sheet $newSheet
        
        # Import sheet data and convert to table
        Invoke-SheetDataImport -Sheet $newSheet -CsvFile $csvFile -CurrentIndex $currentIndex -TotalCount $sheetsToProcessCount
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
