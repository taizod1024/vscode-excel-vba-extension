# -*- coding: utf-8 -*-
param(
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

    # output basic information
    $app_name = $myInvocation.MyCommand.name
    Write-Host -ForegroundColor Yellow "$($app_name)"
    Write-Host -ForegroundColor Green "- tmpPath: $($tmpPath)"

    # clear temporary path
    Write-Host -ForegroundColor Green "- remove tmpPath"
    if (Test-Path $tmpPath) { 
        Remove-PathToLongDirectory $tmpPath
    }
    Write-Host -ForegroundColor Green "- create tmpPath"
    New-Item $tmpPath -itemtype Directory | Out-Null

    # change current directory
    Push-Location $tmpPath

    try {
        Write-Host -ForegroundColor Green "- creating file in tmpPath"
        New-Item -ItemType File -Path (Join-Path $tmpPath "temp.txt") | Out-Null

    }
    finally {
    
        # back to directory
        Pop-Location

    }

    # done
    Write-Host -ForegroundColor Green "- done"
    timeout 2
    exit 0
}
catch {
    Write-Host -ForegroundColor Red $_
    exit 1
}
finally {
}
