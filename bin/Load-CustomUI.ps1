# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $macroPath,
    [Parameter(Mandatory = $true)] [string] $tmpPath
)

# Configuration
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Required assemblies for ZIP extraction
Add-Type -AssemblyName System.IO.Compression.FileSystem

try {
    
    # Display script name
    $scriptName = $MyInvocation.MyCommand.Name
    Write-Host -ForegroundColor Yellow "$($scriptName):"
    Write-Host -ForegroundColor Green "- macroPath: $macroPath"
    Write-Host -ForegroundColor Green "- tmpPath: $tmpPath"

    # Check if the macro file exists
    Write-Host -ForegroundColor Green "- checking if macro file exists"
    if (-not (Test-Path $macroPath)) {
        throw "macro FILE NOT FOUND: $macroPath"
    }

    # Clean temporary directory
    Write-Host -ForegroundColor Green "- cleaning tmpPath"
    if (Test-Path $tmpPath) { 
        Remove-Item $tmpPath -Recurse -Force
    }
    Write-Host -ForegroundColor Green "- creating tmpPath"
    New-Item $tmpPath -ItemType Directory | Out-Null

    # Extract customUI files from .xlam (ZIP format)
    Write-Host -ForegroundColor Green "- extracting customUI from Excel Add-in"
    
    # Copy the .xlam file to a temporary location for extraction
    $tempZipPath = Join-Path $env:TEMP "excel_customui_temp_$(Get-Random).zip"
    Copy-Item $macroPath $tempZipPath
    
    try {
        # Open the ZIP archive
        $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($tempZipPath)
        
        # Find customUI XML files in the archive
        $customUIFound = $false
        
        foreach ($entry in $zipArchive.Entries) {
            $entryName = $entry.FullName.ToLower()
            
            # Check for any customUI XML files (more flexible search)
            if ($entryName -match "customui.*\.xml$") {
                $customUIFound = $true
                
                # Extract the file directly to tmpPath (not to a subfolder)
                $fileName = $entry.Name
                $outputPath = Join-Path $tmpPath $fileName
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $outputPath, $true)
                Write-Host -ForegroundColor Cyan "  extracted: $($entry.FullName) to $outputPath"
            }
        }
        
        $zipArchive.Dispose()
        
        if (-not $customUIFound) {
            Write-Host -ForegroundColor Yellow "  no customUI files found in the Add-in"
        }
    }
    finally {
        # Remove temporary ZIP file
        if (Test-Path $tempZipPath) {
            Remove-Item $tempZipPath -Force
        }
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
