param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelFilePath,
    
    [Parameter(Mandatory = $true)]
    [string]$CsvOutputPath
)

# Create output directory if it doesn't exist
if (-not (Test-Path $CsvOutputPath)) {
    New-Item -ItemType Directory -Force -Path $CsvOutputPath | Out-Null
}

# Create Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # Open the workbook
    $workbook = $excel.Workbooks.Open([System.IO.Path]::GetFullPath($ExcelFilePath))
    
    # Get all sheets
    $sheetCount = $workbook.Sheets.Count
    
    # Export all sheets except those starting with "Sheet"
    for ($i = 1; $i -le $sheetCount; $i++) {
        $sheet = $workbook.Sheets.Item($i)
        $sheetName = $sheet.Name
        
        # Skip sheets that match the pattern Sheet* (like Sheet, Sheet1, Sheet2, etc.)
        if ($sheetName -notmatch "^Sheet\d*$") {
            Write-Host "Exporting sheet: $sheetName"
            
            # Get the used range from the source sheet
            $usedRange = $sheet.UsedRange
            if ($usedRange -and $usedRange.Cells.Count -gt 0) {
                # Get dimensions
                $rows = $usedRange.Rows.Count
                $cols = $usedRange.Columns.Count
                
                # Create CSV content
                $csvLines = @()
                
                for ($r = 1; $r -le $rows; $r++) {
                    $line = @()
                    for ($c = 1; $c -le $cols; $c++) {
                        $cell = $usedRange.Cells.Item($r, $c)
                        $value = $cell.Value2
                        
                        # Handle null values
                        if ($value -eq $null) {
                            $value = ""
                        }
                        
                        # Escape quotes and wrap in quotes if needed
                        $value = $value.ToString()
                        if ($value -match '[",\r\n]') {
                            $value = '"' + ($value -replace '"', '""') + '"'
                        }
                        
                        $line += $value
                    }
                    $csvLines += $line -join ","
                }
                
                # Save as UTF-8 CSV using ADODB.Stream
                $csvFileName = "$sheetName.csv"
                $csvFilePath = Join-Path $CsvOutputPath $csvFileName
                
                $stream = New-Object -ComObject ADODB.Stream
                $stream.Type = 2  # Text
                $stream.Charset = "UTF-8"
                $stream.Open()
                $stream.WriteText(($csvLines -join "`r`n"), 1)  # 1 means add line terminator
                $stream.SaveToFile($csvFilePath, 2)  # 2 means overwrite
                $stream.Close()
                
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($stream) | Out-Null
                
                Write-Host "Exported: $csvFileName"
            }
            else {
                Write-Host "Sheet is empty: $sheetName"
            }
        }
    }
    
    $workbook.Close($false)
    Write-Host "Export completed successfully"
}
catch {
    Write-Error "Error during export: $_"
    exit 1
}
finally {
    # Clean up COM objects
    if ($workbook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
