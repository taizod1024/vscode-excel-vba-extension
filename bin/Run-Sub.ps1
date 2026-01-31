# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $macroPath,
    [Parameter(Mandatory = $true)] [string] $subName
)

# set error action
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {
    # Display script name
    $scriptName = $MyInvocation.MyCommand.Name
    Write-Host -ForegroundColor Yellow "$($scriptName):"
    Write-Host -ForegroundColor Green "- macroPath: $($macroPath)"
    Write-Host -ForegroundColor Green "- subName: $($subName)"

    # check if Excel is running
    Write-Host -ForegroundColor Green "- checking Excel running"
    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        throw "NO EXCEL FOUND. Please Open Excel Macro."
    }

    # Resolve the macro file path
    $resolvedPath = (Resolve-Path $macroPath).Path
    Write-Host -ForegroundColor Cyan "- resolvedPath: $resolvedPath"
    
    # Determine if this is an add-in (.xlam) or workbook (.xlsm/.xlsx)
    $fileExtension = [System.IO.Path]::GetExtension($resolvedPath).ToLower()
    Write-Host -ForegroundColor Cyan "- file extension: $fileExtension"
    
    $isAddIn = ($fileExtension -eq ".xlam")
    
    $vbProject = $null
    
    if ($isAddIn) {
        # For add-ins (.xlam), search through VBE.VBProjects
        Write-Host -ForegroundColor Cyan "- searching VBE.VBProjects (add-in):"
        try {
            $vbe = $excel.VBE
            if ($null -eq $vbe) {
                throw "Excel.VBE is null - VBA project object model access may not be enabled"
            }
            
            $vbProjects = $vbe.VBProjects
            if ($null -eq $vbProjects) {
                throw "Excel.VBE.VBProjects is null - VBA project object model access may not be enabled"
            }
            
            foreach ($vbProj in $vbProjects) {
                if ($vbProj.FileName -eq $resolvedPath) {
                    $vbProject = $vbProj
                    break
                }
            }
        }
        catch {
            throw "Failed to access VBA project object model. Please enable: Excel Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model"
        }
    }
    else {
        # For workbooks, get the active workbook or search by path
        Write-Host -ForegroundColor Cyan "- searching workbooks:"
        if ($null -ne $excel.ActiveWorkbook) {
            $activeWorkbookPath = $excel.ActiveWorkbook.FullName
            if ($activeWorkbookPath -eq $resolvedPath) {
                Write-Host -ForegroundColor Cyan "  using active workbook: $activeWorkbookPath"
                $vbProject = $excel.ActiveWorkbook.VBProject
            }
        }
        
        if ($null -eq $vbProject) {
            # Search through open workbooks
            foreach ($workbook in $excel.Workbooks) {
                if ($workbook.FullName -eq $resolvedPath) {
                    Write-Host -ForegroundColor Cyan "  found workbook: $($workbook.FullName)"
                    $vbProject = $workbook.VBProject
                    break
                }
            }
        }
    }
    
    if ($null -eq $vbProject) {
        throw "Macro file not found in Excel: $resolvedPath"
    }

    # Run the Sub
    Write-Host -ForegroundColor Green "- running Sub: $subName"
    
    # Build the module reference
    # Try to find the sub in any module
    $subFound = $false
    
    # Bring Excel window to foreground before running the sub
    try {
        $shell = New-Object -ComObject WScript.Shell
        $shell.AppActivate($excel.Caption)
    }
    catch {
        Write-Host -ForegroundColor Yellow "- Warning: Could not activate window: $_"
    }
    
    try {
        # First, try to run it directly using Application.Run
        # This is the most reliable method
        $excel.Run($subName)
        $subFound = $true
    }
    catch {
        # If direct run fails, try to find the module and run it
        Write-Host -ForegroundColor Yellow "- Direct run failed, searching in modules..."
        
        foreach ($vbModule in $vbProject.VBComponents) {
            $moduleCode = $vbModule.CodeModule
            $linesCount = $moduleCode.CountOfLines
            
            # Search for the Sub declaration
            for ($i = 1; $i -le $linesCount; $i++) {
                $line = $moduleCode.Lines($i, 1)
                if ($line -match "^\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+$subName\s*(?:\(|$)") {
                    Write-Host -ForegroundColor Cyan "  found in module: $($vbModule.Name)"
                    try {
                        $excel.Run("$($vbModule.Name).$subName")
                        $subFound = $true
                        break
                    }
                    catch {
                        throw "Failed to run Sub: $_"
                    }
                }
            }
            
            if ($subFound) {
                break
            }
        }
    }
    
    if (-not $subFound) {
        throw "Sub not found: $subName"
    }

    Write-Host -ForegroundColor Green "[SUCCESS] Sub executed: $subName"
}
catch {
    Write-Host -ForegroundColor Red "[ERROR] $($_.Exception.Message)"
    exit 1
}
