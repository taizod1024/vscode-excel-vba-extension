# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $bookPath,
    [Parameter(Mandatory = $true)] [string] $customUIOutputPath
)

# Import common functions
. (Join-Path $PSScriptRoot "Common.ps1")

# Configuration
Add-Type -AssemblyName System.IO.Compression.FileSystem

try {
    
    # Initialize
    Initialize-Script $MyInvocation.MyCommand.Name | Out-Null
    Write-Host -ForegroundColor Green "- bookPath: $bookPath"
    Write-Host -ForegroundColor Green "- customUIOutputPath: $customUIOutputPath"

    # Check if book file exists
    Write-Host -ForegroundColor Green "- checking if book file exists"
    if (-not (Test-Path $bookPath)) {
        throw "BOOK FILE NOT FOUND: $bookPath"
    }

    # Clean temporary directory
    Write-Host -ForegroundColor Green "- cleaning customUIOutputPath"
    if (Test-Path $customUIOutputPath) { 
        Remove-Item $customUIOutputPath -Recurse -Force
    }
    Write-Host -ForegroundColor Green "- creating customUIOutputPath"
    New-Item $customUIOutputPath -ItemType Directory | Out-Null

    # Extract customUI files from .xlam (ZIP format)
    Write-Host -ForegroundColor Green "- extracting customUI from Excel Book"
    
    # Copy the .xlam file to a temporary location for extraction
    $tempZipPath = Join-Path $env:TEMP "excel_xml_temp_$(Get-Random).zip"
    Copy-Item $bookPath $tempZipPath
    
    try {
        # Open the ZIP archive
        $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($tempZipPath)
        
        # Find customUI XML files in the archive
        $customUIFound = $false
        
        foreach ($entry in $zipArchive.Entries) {
            $entryName = $entry.FullName.ToLower()
            
            # Check for any customUI XML files (more flexible search)
            if ($entryName -match "customui.*\.xml$" -and -not $entry.FullName.EndsWith("/")) {
                $customUIFound = $true
                
                # Extract the file directly to tmpPath (not to a subfolder)
                $fileName = $entry.Name
                $outputPath = Join-Path $customUIOutputPath $fileName
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $outputPath, $true)
                Write-Host -ForegroundColor Cyan "  extracted: $($entry.FullName) to $outputPath"
            }
        }
        
        $zipArchive.Dispose()
        
        if (-not $customUIFound) {
            Write-Host -ForegroundColor Yellow "  no customUI files found in Excel Book"
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
