# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $macroPath,
    [Parameter(Mandatory = $true)] [string] $tmpPath
)

# set error action
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# https://stackoverflow.com/a/39587889
function Remove-PathToLongDirectory {
    Param(
        [string]$directory
    )

    # create a temporary (empty) directory
    $parent = [System.IO.Path]::GetTempPath()
    [string] $name = [System.Guid]::NewGuid()
    $tempDirectory = New-Item -ItemType Directory -Path (Join-Path $parent $name)

    robocopy /MIR $tempDirectory.FullName $directory | out-null
    Remove-Item $directory -Force | out-null
    Remove-Item $tempDirectory -Force | out-null
}

# Function to remove blank lines before VBA code starts
function Remove-BlankLinesBeforeVBACode {
    Param([string]$content)
    # Remove blank lines before the first code keyword (excluding Attribute, VERSION, Begin, End)
    # Only replace once using regex match
    $match = [regex]::Match($content, "[\r\n]+(?=\s*(Option|Sub|Function|Const|Private|Public|Dim|Type|Enum|Declare|Global|Static|Param|'))")
    if ($match.Success) {
        return $content.Remove($match.Index, $match.Length).Insert($match.Index, "`r`n")
    }
    return $content
}

try {
    
    # Display script name
    $scriptName = $MyInvocation.MyCommand.Name
    Write-Host -ForegroundColor Yellow "$($scriptName):"
    Write-Host -ForegroundColor Green "- macroPath: $($macroPath)"
    Write-Host -ForegroundColor Green "- tmpPath: $($tmpPath)"

    # check if Excel is running
    Write-Host -ForegroundColor Green "- checking Excel running"
    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        throw "NO EXCEL FOUND. Please Open Excel."
    }
    $macro = $null

    # Check if the workbook is already open in Excel
    Write-Host -ForegroundColor Green "- checking if workbook is open in Excel"
    $resolvedPath = (Resolve-Path $macroPath).Path
    $macro = $null
    foreach ($wb in $excel.workbooks) {
        if ($wb.FullName -eq $resolvedPath) {
            $macro = $wb
            break
        }
    }
    
    if ($null -eq $macro) {
        throw "NO OPENED WORKBOOK FOUND. Please Open Workbook."
    }
    
    # Access VB Project
    Write-Host -ForegroundColor Green "- accessing VB Project"
    $vbProject = $macro.VBProject
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
        
        # Skip Document Modules (cannot be loaded)
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
