# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $bookPath,
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

try {

    # Display script name
    $scriptName = $MyInvocation.MyCommand.Name
    Write-Host -ForegroundColor Yellow "$($scriptName):"
    Write-Host -ForegroundColor Green "- bookPath: $($bookPath)"
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
        throw "NO EXCEL FOUND. Please start Excel."
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
        throw "NO OPENED WORKBOOK FOUND. Please open workbook."
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
    $vbComponents = @($book.VBProject.VBComponents)
    foreach ($component in $vbComponents) {
        # Skip Document modules (they can't be removed)
        if ($component.Type -eq 100) {
            # 100 = Document module
            continue
        }
        
        if (-not ($saveedFileNames -contains $component.Name)) {
            try {
                Write-Host -ForegroundColor Green "  - removing component: $($component.Name)"
                $book.VBProject.VBComponents.Remove($component)
            }
            catch {
                Write-Host -ForegroundColor Yellow "  - warning: failed to remove component '$($component.Name)': $_"
            }
        }
    }
    
    # Save VBA files
    Write-Host -ForegroundColor Green "- saving new/updated components"
    $vbComponents = @($book.VBProject.VBComponents)
    foreach ($file in $vbaFiles) {
        try {
            $fileName = [System.IO.Path]::GetFileName($file)
            $componentName = [System.IO.Path]::GetFileNameWithoutExtension($file)
            $fileExtension = [System.IO.Path]::GetExtension($file).ToLower()
            
            # For standard modules (.bas), class modules (.cls), and forms (.frm), 
            # remove existing component with same name before saving
            if ($fileExtension -eq ".frm" -or $fileExtension -eq ".bas" -or $fileExtension -eq ".cls") {
                foreach ($component in $vbComponents) {
                    if ($component.Name -eq $componentName) {
                        Write-Host -ForegroundColor Green "  - removing existing component: $componentName"
                        $book.VBProject.VBComponents.Remove($component)
                        break
                    }
                }
                
                Write-Host -ForegroundColor Green "  - saving: $fileName"
                $book.VBProject.VBComponents.Import($file) | Out-Null
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
    
    # Save the workbook
    Write-Host -ForegroundColor Green "- saving workbook"
    $book.Save()
    
    Write-Host -ForegroundColor Green "- done"
    exit 0
}
catch {
    [Console]::Error.WriteLine("$($_)")
    exit 1
}
finally {
}

