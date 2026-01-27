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
    # https://learn.microsoft.com/en-us/answers/questions/4911760/excel-vba-bug-importing-a-form-adds-a-newline-at-t
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

    # Check if the workbook or add-in is already open in Excel
    Write-Host -ForegroundColor Green "- checking if workbook/add-in is open in Excel"
    $resolvedPath = (Resolve-Path $macroPath).Path
    Write-Host -ForegroundColor Cyan "  resolvedPath: $resolvedPath"
    $macro = $null
    $vbProject = $null
    
    # Determine if this is an add-in (.xlam) or workbook (.xlsm/.xlsx)
    $fileExtension = [System.IO.Path]::GetExtension($resolvedPath).ToLower()
    Write-Host -ForegroundColor Cyan "  file extension: $fileExtension"
    
    $isAddIn = ($fileExtension -eq ".xlam")
    
    if ($isAddIn) {
        # For add-ins (.xlam), search through VBE.VBProjects
        Write-Host -ForegroundColor Cyan "  searching VBE.VBProjects (add-in):"
        try {
            $vbe = $excel.VBE
            if ($null -eq $vbe) {
                throw "Excel.VBE is null - VBA project object model access may not be enabled"
            }
            
            $vbProjects = $vbe.VBProjects
            if ($null -eq $vbProjects) {
                throw "Excel.VBE.VBProjects is null - VBA project object model access may not be enabled"
            }
            
            $projectCount = 0
            foreach ($vbProj in $vbProjects) {
                $projectCount++
                $projectFileName = $vbProj.FileName
                $projectName = $vbProj.Name
                Write-Host -ForegroundColor Cyan "    [$projectCount] Name: $projectName, FileName: $projectFileName"
                
                if ($projectFileName -eq $resolvedPath) {
                    Write-Host -ForegroundColor Yellow "    MATCHED!"
                    $vbProject = $vbProj
                    break
                }
            }
            Write-Host -ForegroundColor Cyan "  total projects found: $projectCount"
        }
        catch {
            Write-Host -ForegroundColor Red "  error accessing VBE.VBProjects: $_"
            Write-Host -ForegroundColor Red "  "
            Write-Host -ForegroundColor Red "  SOLUTION:"
            Write-Host -ForegroundColor Red "  1. Open Excel and go to: File > Options > Trust Center > Trust Center Settings..."
            Write-Host -ForegroundColor Red "  2. Click 'Macro Settings'"
            Write-Host -ForegroundColor Red "  3. Check the checkbox: 'Trust access to the VBA project object model'"
            Write-Host -ForegroundColor Red "  4. Click OK and close Excel completely"
            Write-Host -ForegroundColor Red "  5. Re-open the add-in and try again"
            throw $_
        }
    } else {
        # For workbooks (.xlsm/.xlsx), search through Excel.Workbooks
        Write-Host -ForegroundColor Cyan "  searching Excel.Workbooks (workbook):"
        $workbookCount = $excel.Workbooks.Count
        Write-Host -ForegroundColor Cyan "  total workbooks found: $workbookCount"
        
        foreach ($wb in $excel.Workbooks) {
            $wbFullName = $wb.FullName
            Write-Host -ForegroundColor Cyan "    Workbook: $($wb.Name), FullName: $wbFullName"
            
            if ($wbFullName -eq $resolvedPath) {
                Write-Host -ForegroundColor Yellow "    MATCHED!"
                $macro = $wb
                $vbProject = $wb.VBProject
                break
            }
        }
    }
    
    if ($null -eq $vbProject) {
        throw "NO OPENED WORKBOOK OR ADD-IN FOUND. Please Open Workbook or Add-in."
    }
    
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
