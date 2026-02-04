# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $filePath
)

# Import common functions
. (Join-Path $PSScriptRoot "Common.ps1")

try {
    # Initialize
    Initialize-Script $MyInvocation.MyCommand.Name | Out-Null
    Write-Host -ForegroundColor Green "- filePath: $($filePath)"

    # Get or create Excel instance
    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        Write-Host -ForegroundColor Green "- using existing Excel instance"
    }
    catch {
        Write-Host -ForegroundColor Green "- creating new Excel instance"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
    }

    # Create new workbook
    Write-Host -ForegroundColor Green "- creating new Excel workbook"
    $workbook = $excel.Workbooks.Add()

    # Save the workbook
    Write-Host -ForegroundColor Green "- saving workbook: $filePath"
    $workbook.SaveAs($filePath)

    # Open the created workbook
    Write-Host -ForegroundColor Green "- opening workbook in Excel"
    $excel.Workbooks.Open($filePath) | Out-Null

    # Bring Excel to foreground
    Write-Host -ForegroundColor Green "- bringing Excel to foreground"
    $excel.Visible = $true
    
    # Use Shell to activate Excel window
    $shell = New-Object -ComObject wscript.shell
    $shell.AppActivate($excel.Caption) | Out-Null

    Write-Host -ForegroundColor Green "[SUCCESS] New Excel file created and opened"
}
catch {
    Write-Host -ForegroundColor Red "[ERROR] $_"
    exit 1
}
