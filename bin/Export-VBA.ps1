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
    Write-Host -ForegroundColor Yellow "Export-VBA.ps1"
    Write-Host -ForegroundColor Green "- bookPath: $bookPath"
    Write-Host -ForegroundColor Green "- tmpPath: $tmpPath"

    # Clean temporary directory
    Write-Host -ForegroundColor Green "- cleaning tmpPath"
    if (Test-Path $tmpPath) { 
        Remove-PathToLongDirectory $tmpPath
    }
    
    Write-Host -ForegroundColor Green "- creating tmpPath"
    New-Item $tmpPath -ItemType Directory | Out-Null
    Push-Location $tmpPath

    # Check if Excel is already running
    Write-Host -ForegroundColor Green "- checking excel running"
    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        throw "FIRST, START EXCEL"
    }
    $book = $null

    try {
        # Open workbook
        Write-Host -ForegroundColor Green "- opening workbook"
        $book = $excel.Workbooks.Open((Resolve-Path $bookPath).Path)
        
        # Access VB Project
        Write-Host -ForegroundColor Green "- accessing VB Project"
        $vbProject = $book.VBProject
        Write-Host -ForegroundColor Green "- project name: $($vbProject.Name)"

        # Get component count
        $componentCount = $vbProject.VBComponents.Count
        Write-Host -ForegroundColor Green "- found $componentCount component(s)"
        
        if ($componentCount -eq 0) {
            throw @"
No VB components found. Enable VBA Project Object Model access:
1. Open Excel
2. File > Options > Trust Center > Trust Center Settings
3. Macro Settings > Check 'Trust access to the VBA project object model'
4. Close Excel and re-open the workbook
"@
        }
        
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
            
            try {
                [void]$component.Export($filePath)
                Write-Host -ForegroundColor Cyan "  exported to $filePath"
            }
            catch {
                Write-Host -ForegroundColor Red "  ERROR: Failed to export"
                Write-Host -ForegroundColor Red "  Reason: $_"
                throw $_
            }
        }
    }
    catch {
        Write-Host -ForegroundColor Red "ERROR: $($_)"
        throw
    }
    finally {
        # Cleanup Excel resources
        if ($null -ne $book) {
            Write-Host -ForegroundColor Green "- closing workbook"
            try { $book.Close($false) } catch { }
        }
        if ($null -ne $excel) {
            Write-Host -ForegroundColor Green "- closing Excel"
            try { $excel.Quit() } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        Pop-Location
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
