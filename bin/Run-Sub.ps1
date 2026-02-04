# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $macroPath,
    [Parameter(Mandatory = $true)] [string] $subName
)

# Import common functions
. (Join-Path $PSScriptRoot "Common.ps1")

try {
    # Initialize
    Initialize-Script $MyInvocation.MyCommand.Name | Out-Null
    Write-Host -ForegroundColor Green "- macroPath: $($macroPath)"
    Write-Host -ForegroundColor Green "- subName: $($subName)"

    # Get Excel instance
    $excel = Get-ExcelInstance
    
    # Get VB Project
    $macroInfo = Get-BookInfo $macroPath
    $result = Find-VBProject $excel $macroInfo.ResolvedPath $macroInfo.IsAddIn
    $vbProject = $result.VBProject

    # Run the Sub
    Write-Host -ForegroundColor Green "- running Sub: $subName"
    
    # Build the module reference
    # Try to find the sub in any module
    $subFound = $false
    
    # Bring Excel window to foreground before running the sub
    try {
        $shell = New-Object -ComObject WScript.Shell
        $shell.AppActivate($excel.Caption) | Out-Null
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
                        throw "FAILED TO RUN SUB: $_"
                    }
                }
            }
            
            if ($subFound) {
                break
            }
        }
    }
    
    if (-not $subFound) {
        throw "SUB NOT FOUND: $subName"
    }

    Write-Host -ForegroundColor Green "[SUCCESS] Sub executed: $subName"
}
catch {
    [Console]::Error.WriteLine("$($_)")
    exit 1
}
