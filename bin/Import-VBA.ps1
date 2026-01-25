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
    Write-Host -ForegroundColor Green "- tmpPath: $($tmpPath)"

    # clean temporary path
    Write-Host -ForegroundColor Green "- remove tmpPath"
    if (Test-Path $tmpPath) { 
        Remove-PathToLongDirectory $tmpPath
    }
    
    Write-Host -ForegroundColor Green "- create tmpPath"
    New-Item $tmpPath -itemtype Directory | Out-Null
    Push-Location $tmpPath

    # check if Excel is running
    Write-Host -ForegroundColor Green "- checking Excel running"
    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        throw "FIRST, START EXCEL"
    }
    $book = $null

    try {
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
            throw "Workbook is not open in Excel: $bookPath`nPlease open the workbook in Excel first."
        }

        # done
        Write-Host -ForegroundColor Green "- done"
        exit 0
    }
    catch {
        throw
    }
    finally {
        # Do not close workbook or Excel since they are already open
        # Just cleanup current location
        Pop-Location
    }
}
catch {
    [Console]::Error.WriteLine("$($_)")
    exit 1
}
finally {
}

