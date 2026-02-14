# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $bookPath,
    [Parameter(Mandatory = $true)] [string] $subName
)

# Import common functions
. (Join-Path $PSScriptRoot "Common.ps1")

try {
    # Initialize
    Initialize-Script $MyInvocation.MyCommand.Name | Out-Null
    Write-Host "- bookPath: $($bookPath)"
    Write-Host "- subName: $($subName)"

    # Get Excel instance
    $excel = Get-ExcelInstance
    
    # Get VB Project
    $macroInfo = Get-BookInfo $bookPath
    $result = Find-VBProject $excel $macroInfo.ResolvedPath $macroInfo.IsAddIn
    $vbProject = $result.VBProject

    # Run the Sub
    Write-Host "- running Sub: $subName"
    
    # Build the module reference
    # Try to find the sub in any module
    $subFound = $false
    
    # Bring Excel window to foreground before running the sub
    try {
        $shell = New-Object -ComObject WScript.Shell
        $shell.AppActivate($excel.Caption) | Out-Null
    }
    catch {
        Write-Host "- Warning: Could not activate window: $_"
    }
    
    # Disable alerts and events during macro execution
    $originalDisplayAlerts = $excel.DisplayAlerts
    $originalEnableEvents = $excel.EnableEvents
    
    try {
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        
        try {
            # First, try to run it directly using Application.Run
            # This is the most reliable method
            $excel.Run($subName)
            $subFound = $true
        }
        catch {
            # If direct run fails, try to find the module and run it
            Write-Host "- Direct run failed, searching in modules..."
            
            foreach ($vbModule in $vbProject.VBComponents) {
                $moduleCode = $vbModule.CodeModule
                $linesCount = $moduleCode.CountOfLines
                
                # Search for the Sub declaration
                for ($i = 1; $i -le $linesCount; $i++) {
                    $line = $moduleCode.Lines($i, 1)
                    if ($line -match "^\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+$subName\s*(?:\(|$)") {
                        Write-Host "  found in module: $($vbModule.Name)"
                        try {
                            $excel.Run("$($vbModule.Name).$subName")
                            $subFound = $true
                            break
                        }
                        catch {
                            # Check if error is related to macro security
                            if ($_ -match "実行できません|cannot run|disabled|security") {
                                throw "Macro execution blocked by Excel security settings. Run the macro manually in Excel or check security settings."
                            }
                            else {
                                throw "Failed to run subroutine: $_"
                            }
                        }
                    }
                }
                
                if ($subFound) {
                    break
                }
            }
        }
    }
    finally {
        # Restore original settings
        $excel.DisplayAlerts = $originalDisplayAlerts
        $excel.EnableEvents = $originalEnableEvents
    }
    
    if (-not $subFound) {
        throw "SUB NOT FOUND: $subName"
    }

    Write-Host "Sub executed successfully"
}
catch {
    [Console]::Error.WriteLine("$($_)")
    exit 1
}
