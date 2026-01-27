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
    Write-Host -ForegroundColor Green "- saveSourcePath: $($tmpPath)"

    # check if save source path exists
    Write-Host -ForegroundColor Green "- checking save source folder"
    if (-not (Test-Path $tmpPath)) {
        throw "IMPORT SOURCE FOLDER NOT FOUND: $($tmpPath)"
    }

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
    
    # Search through VBE.VBProjects to find the opened workbook or add-in
    Write-Host -ForegroundColor Cyan "  searching VBE.VBProjects:"
    try {
        $projectCount = 0
        foreach ($vbProj in $excel.VBE.VBProjects) {
            $projectCount++
            $projectFileName = $vbProj.FileName
            $projectName = $vbProj.Name
            Write-Host -ForegroundColor Cyan "    [$projectCount] Name: $projectName, FileName: $projectFileName"
            
            if ($projectFileName -eq $resolvedPath) {
                Write-Host -ForegroundColor Yellow "    MATCHED!"
                # Found the project, save it directly
                $vbProject = $vbProj
                
                # Try to find corresponding workbook object
                foreach ($wb in $excel.workbooks) {
                    if ($wb.Name -eq $projectName) {
                        $macro = $wb
                        Write-Host -ForegroundColor Yellow "    Found in Workbooks"
                        break
                    }
                }
                
                # If not found in Workbooks, just use the VBProject
                if ($null -eq $macro) {
                    Write-Host -ForegroundColor Yellow "    Using VBProject directly (add-in)"
                }
                break
            }
        }
        Write-Host -ForegroundColor Cyan "  total projects found: $projectCount"
    }
    catch {
        Write-Host -ForegroundColor Red "  error accessing VBE.VBProjects: $_"
    }
    
    if ($null -eq $vbProject) {
        throw "NO OPENED WORKBOOK OR ADD-IN FOUND. Please Open Workbook or Add-in."
    }
    
    # Save VBA components from files
    Write-Host -ForegroundColor Green "- saving VBA components"
    
    # Get list of VBA files to save
    $vbaFiles = @()
    if (Test-Path $tmpPath) {
        $vbaFiles = Get-ChildItem -Path $tmpPath -Recurse -Include *.bas, *.cls, *.frm | ForEach-Object { $_.FullName }
    }
    
    Write-Host -ForegroundColor Green "- found VBA files to save: $($vbaFiles.Count)"
    
    # Get list of saveed file names (without extension)
    $saveedFileNames = @()
    foreach ($file in $vbaFiles) {
        $saveedFileNames += [System.IO.Path]::GetFileNameWithoutExtension($file)
    }
    
    # Remove components that are no longer in the save folder
    Write-Host -ForegroundColor Green "- removing deleted components"
    $vbComponents = @($vbProject.VBComponents)
    foreach ($component in $vbComponents) {
        # Skip Document modules (they can't be removed)
        if ($component.Type -eq 100) {
            # 100 = Document module
            continue
        }
        
        if (-not ($saveedFileNames -contains $component.Name)) {
            try {
                Write-Host -ForegroundColor Green "  - removing component: $($component.Name)"
                $vbProject.VBComponents.Remove($component)
            }
            catch {
                Write-Host -ForegroundColor Yellow "  - warning: failed to remove component '$($component.Name)': $_"
            }
        }
    }
    
    # Save VBA files
    Write-Host -ForegroundColor Green "- saving new/updated components"
    $vbComponents = @($vbProject.VBComponents)
    foreach ($file in $vbaFiles) {
        try {
            $fileName = [System.IO.Path]::GetFileName($file)
            $componentName = [System.IO.Path]::GetFileNameWithoutExtension($file)
            $fileExtension = [System.IO.Path]::GetExtension($file).ToLower()
            $filePath = $file

            # For standard modules (.bas), class modules (.cls), and forms (.frm), 
            # remove existing component with same name before saving
            if ($fileExtension -eq ".frm" -or $fileExtension -eq ".bas" -or $fileExtension -eq ".cls") {
                foreach ($component in $vbComponents) {
                    if ($component.Name -eq $componentName) {
                        Write-Host -ForegroundColor Green "  - removing existing component: $componentName"
                        $vbProject.VBComponents.Remove($component)
                        break
                    }
                }
                
                Write-Host -ForegroundColor Green "  - saving: $fileName"
                
                # Remove trailing whitespace and blank lines before import
                $content = [System.IO.File]::ReadAllText($filePath, [System.Text.Encoding]::GetEncoding('shift_jis'))
                
                # Remove blank lines before VBA code starts
                $content = Remove-BlankLinesBeforeVBACode $content
                
                $content = $content -replace '\s+$', ''
                [System.IO.File]::WriteAllText($filePath, $content, [System.Text.Encoding]::GetEncoding('shift_jis'))
                
                $vbProject.VBComponents.Import($filePath) | Out-Null
            }
        }
        catch {
            # Check for .log file and include its content in error message
            $logFile = [System.IO.Path]::ChangeExtension($file, ".log")
            $logContent = ""
            if (Test-Path $logFile) {
                $logContent = Get-Content $logFile -Raw
                Remove-Item $logFile -Force
            }
            
            if ($logContent) {
                throw "FAILED TO IMPORT FILE: $($file) - $logContent"
            }
            else {
                throw "FAILED TO IMPORT FILE: $($file) - $_"
            }
        }
    }
    
    # Save the workbook or add-in
    Write-Host -ForegroundColor Green "- saving workbook/add-in"
    if ($null -ne $macro) {
        $macro.Save()
    }
    else {
        # For add-ins, we need to save through the workbook that contains it
        # Get the parent workbook from VBProject
        $parentWorkbook = $null
        foreach ($wb in $excel.workbooks) {
            try {
                if ($wb.VBProject.Name -eq $vbProject.Name) {
                    $parentWorkbook = $wb
                    break
                }
            }
            catch {
                # Skip if VBProject not accessible
            }
        }
        
        if ($null -ne $parentWorkbook) {
            $parentWorkbook.Save()
        }
        else {
            Write-Host -ForegroundColor Yellow "  WARNING: Could not find workbook to save add-in. Add-in may not be saved."
        }
    }
    
    Write-Host -ForegroundColor Green "- done"
    exit 0
}
catch {
    [Console]::Error.WriteLine("$($_)")
    exit 1
}

