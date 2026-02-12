# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $bookPath,
    [Parameter(Mandatory = $true)] [string] $customUISourcePath
)

# Import common functions
. (Join-Path $PSScriptRoot "Common.ps1")

# Required assemblies for ZIP manipulation
Add-Type -AssemblyName System.IO.Compression.FileSystem

try {

    # Initialize
    Initialize-Script $MyInvocation.MyCommand.Name | Out-Null
    Write-Host "- bookPath: $($bookPath)"
    Write-Host "- customUISourcePath: $($customUISourcePath)"

    # check if source path exists
    Write-Host "- checking custom UI source folder"
    if (-not (Test-Path $customUISourcePath)) {
        throw "Custom UI source folder not found: $customUISourcePath"
    }
    
    # List contents of source folder for debugging
    Write-Host "- contents of $customUISourcePath"
    $sourceContents = Get-ChildItem -Path $customUISourcePath
    if ($sourceContents.Count -gt 0) {
        $sourceContents | ForEach-Object { Write-Host "  - $($_.Name)" }
    }
    else {
        Write-Host "  - (empty folder)"
    }

    # Get list of customUI files directly from source folder (customUI.xml, customUI14.xml, etc.)
    $customUIFiles = Get-ChildItem -Path $customUISourcePath -Filter "customUI*.xml" | ForEach-Object { $_.FullName }
    
    if ($customUIFiles.Count -eq 0) {
        throw "Custom UI XML file not found: Expected customUI.xml or customUI14.xml"
    }

    Write-Host "- found $($customUIFiles.Count) customUI file(s)"

    # Create a temporary directory for backup and work
    $tempDir = Join-Path $env:TEMP "excel_xml_work_$(Get-Random)"
    New-Item $tempDir -ItemType Directory | Out-Null
    
    try {
        # Copy the original macro to temp location
        $tempMacroPath = Join-Path $tempDir "macro.xlam"
        Copy-Item $bookPath $tempMacroPath
        
        Write-Host "- opening Excel Add-in for modification"
        
        # Open the ZIP archive for reading
        # ZipArchiveMode: 0=Read, 1=Create, 2=Update
        $zipArchive = [System.IO.Compression.ZipFile]::Open($tempMacroPath, 2)
        
        try {
            # Remove existing customUI entries from the archive
            Write-Host "- removing existing customUI entries"
            $entriesToRemove = @()
            
            foreach ($entry in $zipArchive.Entries) {
                $entryName = $entry.FullName.ToLower()
                if ($entryName -match "customui.*\.xml$" -and -not $entry.FullName.EndsWith("/")) {
                    $entriesToRemove += $entry
                    Write-Host "  marked for removal: $($entry.FullName)"
                }
            }
            
            # Actually remove the marked entries
            foreach ($entry in $entriesToRemove) {
                $entry.Delete()
            }
            
            # Add new customUI files to the archive
            Write-Host "- adding new customUI files"
            foreach ($file in $customUIFiles) {
                $fileName = Split-Path $file -Leaf
                $entryName = "customUI/$fileName"
                
                # Remove if already exists (shouldn't happen after cleanup above)
                try {
                    $existingEntry = $zipArchive.GetEntry($entryName)
                    if ($null -ne $existingEntry) {
                        $existingEntry.Delete()
                    }
                }
                catch {}
                
                # Add the new file
                $newEntry = $zipArchive.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
                $fileStream = [System.IO.File]::OpenRead($file)
                $entryStream = $newEntry.Open()
                
                try {
                    $fileStream.CopyTo($entryStream)
                    Write-Host "  added: $entryName"
                }
                finally {
                    $entryStream.Close()
                    $fileStream.Close()
                }
            }
        }
        finally {
            $zipArchive.Dispose()
        }
        
        # Replace the original macro with the modified version
        Write-Host "- saving changes to Excel Book"
        Remove-Item $bookPath -Force
        Move-Item $tempMacroPath $bookPath
    }
    finally {
        # Clean up temporary directory
        if (Test-Path $tempDir) {
            Remove-Item $tempDir -Recurse -Force
        }
    }

    Write-Host "- done"
    exit 0
}
catch {
    [Console]::Error.WriteLine("$($_)")
    exit 1
}
