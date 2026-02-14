
param(
    [string]$ExcelPath,
    [string]$SheetName
)

try {
    # Get Excel application object
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        # If Excel is not running, create a new instance
        $excel = New-Object -ComObject Excel.Application
    }

    $excel.Visible = $true
    $excel.ScreenUpdating = $false

    # Open the workbook
    Write-Output "Opening workbook: $ExcelPath"
    $workbook = $excel.Workbooks.Open($ExcelPath, $false, $false)
    Write-Output "Workbook opened successfully"
    
    # List all sheet names for debugging
    Write-Output "Attempting to list sheets..."
    $sheetNames = @()
    try {
        foreach ($ws in $workbook.Sheets) {
            Write-Output "Sheet found: $($ws.Name)"
            $sheetNames += $ws.Name
        }
    }
    catch {
        Write-Output "Error listing sheets: $_"
    }
    Write-Output "Available sheets: $($sheetNames -join ', ')"
    
    # Find and select the sheet by name
    $sheet = $null
    foreach ($ws in $workbook.Sheets) {
        if ($ws.Name -eq $SheetName) {
            $sheet = $ws
            Write-Output "Found sheet: $($ws.Name)"
            break
        }
    }
    
    if ($sheet) {
        Write-Output "Activating sheet: $SheetName"
        $sheet.Activate()
        Write-Output "Sheet activated"
        
        # Activate Excel window using WScript.Shell
        try {
            $shell = New-Object -ComObject WScript.Shell
            $shell.AppActivate($excel.Caption) | Out-Null
            Write-Output "Excel window activated"
        }
        catch {
            Write-Output "Warning: Could not activate Excel window: $_"
        }
        
        Write-Output "Sheet '$SheetName' selected successfully"
    }
    else {
        Write-Error "Sheet '$SheetName' not found in workbook. Available sheets: $($sheetNames -join ', ')"
        exit 1
    }
}
catch {
    Write-Error "Failed to open sheet: $_"
    exit 1
}
finally {
    $excel.ScreenUpdating = $true
}
