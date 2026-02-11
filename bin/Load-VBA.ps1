# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $bookPath,
    [Parameter(Mandatory = $true)] [string] $tmpPath
)

# Import common functions
. (Join-Path $PSScriptRoot "Common.ps1")

try {
    # Initialize
    Initialize-Script $MyInvocation.MyCommand.Name | Out-Null
    Write-Host -ForegroundColor Green "- bookPath: $($bookPath)"
    Write-Host -ForegroundColor Green "- tmpPath: $($tmpPath)"

    # Get Excel instance
    $excel = Get-ExcelInstance
    
    # Get VB Project
    $macroInfo = Get-BookInfo $bookPath
    $result = Find-VBProject $excel $macroInfo.ResolvedPath $macroInfo.IsAddIn
    $vbProject = $result.VBProject
    
    # Access VB Project
    Write-Host -ForegroundColor Green "- accessing VB Project"
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
    
    # Load each component
    for ($i = 1; $i -le $componentCount; $i++) {
        $component = $vbProject.VBComponents.Item($i)
        $componentName = $component.Name
        $componentType = $component.Type
        
        Write-Host -ForegroundColor Green "- loading component [$i/$componentCount] $componentName"
        
        # Determine file extension based on component type
        $fileExt = switch ($componentType) {
            1 { ".bas" }      # Standard Module
            2 { ".cls" }      # Class Module
            3 { ".frm" }      # Form
            100 { ".cls" }    # Document Module
            default { ".bas" }
        }
        
        $filePath = Join-Path $tmpPath "$componentName$fileExt"
        [void]$component.Export($filePath)

        # Remove trailing whitespace and blank lines before import
        $content = [System.IO.File]::ReadAllText($filePath, [System.Text.Encoding]::GetEncoding('shift_jis'))
        
        # Remove blank lines before VBA code starts
        $content = Remove-BlankLinesBeforeVBACode $content
        
        $content = $content -replace '\s+$', ''
        [System.IO.File]::WriteAllText($filePath, $content, [System.Text.Encoding]::GetEncoding('shift_jis'))
        
        Write-Host -ForegroundColor Cyan "  loaded to $filePath"
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
