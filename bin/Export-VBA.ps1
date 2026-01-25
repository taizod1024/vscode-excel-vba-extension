# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $bookPath,
    [Parameter(Mandatory = $true)] [string] $tmpPath
)

# Configuration
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Function to safely remove a directory
function Remove-PathToLongDirectory {
    Param([string]$directory)
    # Use robocopy to recursively delete long paths
    $parent = [System.IO.Path]::GetTempPath()
    $tempDirectory = New-Item -ItemType Directory -Path (Join-Path $parent ([System.Guid]::NewGuid()))
    robocopy /MIR $tempDirectory.FullName $directory | Out-Null
    Remove-Item $directory -Force | Out-Null
    Remove-Item $tempDirectory -Force | Out-Null
}

try {
    
    # Display script name
    $scriptName = $MyInvocation.MyCommand.Name
    Write-Host -ForegroundColor Yellow "$($scriptName):"
    Write-Host -ForegroundColor Green "- bookPath: $bookPath"
    Write-Host -ForegroundColor Green "- tmpPath: $tmpPath"

    # Check if Excel is already running
    Write-Host -ForegroundColor Green "- checking Excel running"
    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        throw "NO EXCEL FOUND, START EXCEL"
    }
    $book = $null

    # Check if the workbook is already open in Excel
    Write-Host -ForegroundColor Green "- checking if workbook is open in Excel"
    $resolvedPath = (Resolve-Path $bookPath).Path
    $book = $null
    foreach ($wb in $excel.Workbooks) {
        if ($wb.FullName -eq $resolvedPath) {
            $book = $wb
            break
        }
    }
    
    if ($null -eq $book) {
        throw "NO OPENED WORKBOOK FOUND, OPEN WORKBOOK"
    }
    
    # Access VB Project
    Write-Host -ForegroundColor Green "- accessing VB Project"
    $vbProject = $book.VBProject
    Write-Host -ForegroundColor Green "- project name: $($vbProject.Name)"
    $componentCount = $vbProject.VBComponents.Count
    Write-Host -ForegroundColor Green "- found $componentCount component(s)"
    
    if ($componentCount -eq 0) {
        throw @"
NO VB COMPONENTS FOUND, ENABLE VBA PROJECT OBJECT MODEL ACCESS:
1. Open Excel
2. File > Options > Trust Center > Trust Center Settings
3. Macro Settings > Check 'Trust access to the VBA project object model'
4. Close Excel and re-open the workbook
"@
    }
    
    # Clean temporary directory
    Write-Host -ForegroundColor Green "- cleaning tmpPath"
    if (Test-Path $tmpPath) { 
        Remove-PathToLongDirectory $tmpPath
    }
    Write-Host -ForegroundColor Green "- creating tmpPath"
    New-Item $tmpPath -ItemType Directory | Out-Null
    
    # Track first difference for diff view
    $firstDiffFile = $null
    $firstDiffOldPath = $null

    # Export each component
    for ($i = 1; $i -le $componentCount; $i++) {
        $component = $vbProject.VBComponents.Item($i)
        $componentName = $component.Name
        $componentType = $component.Type
        
        Write-Host -ForegroundColor Green "- exporting component [$i/$componentCount] $componentName"
        
        # Skip Document Modules (cannot be exported)
        if ($componentType -eq 100) {
            Write-Host -ForegroundColor Yellow "  (skipped - Document Module)"
            continue
        }
        
        # Determine file extension based on component type
        $fileExt = switch ($componentType) {
            1 { ".bas" }      # Standard Module
            2 { ".cls" }      # Class Module
            3 { ".frm" }      # Form
            default { ".bas" }
        }
        
        $filePath = Join-Path $tmpPath "$componentName$fileExt"
        [void]$component.Export($filePath)
        Write-Host -ForegroundColor Cyan "  exported to $filePath"
    }
    
    Write-Host -ForegroundColor Green "- done"
    exit 0
}
catch {
    [Console]::Error.WriteLine("$($_)")
    exit 1
}
finally {
}
