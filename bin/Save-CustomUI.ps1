# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $macroPath,
    [Parameter(Mandatory = $true)] [string] $customUISourcePath
)

# set error action
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Required assemblies for ZIP manipulation
Add-Type -AssemblyName System.IO.Compression.FileSystem

try {

    # Display script name
    $scriptName = $MyInvocation.MyCommand.Name
    Write-Host -ForegroundColor Yellow "$($scriptName):"
    Write-Host -ForegroundColor Green "- macroPath: $($macroPath)"
    Write-Host -ForegroundColor Green "- customUISourcePath: $($customUISourcePath)"

    # check if source path exists
    Write-Host -ForegroundColor Green "- checking custom UI source folder"
    if (-not (Test-Path $customUISourcePath)) {
        throw "CUSTOM UI SOURCE FOLDER NOT FOUND: $($customUISourcePath)"
    }
    
    # List contents of source folder for debugging
    Write-Host -ForegroundColor Green "- contents of $customUISourcePath"
    $sourceContents = Get-ChildItem -Path $customUISourcePath
    if ($sourceContents.Count -gt 0) {
        $sourceContents | ForEach-Object { Write-Host -ForegroundColor Cyan "  - $($_.Name)" }
    }
    else {
        Write-Host -ForegroundColor Yellow "  - (empty folder)"
    }

    # Get list of customUI files directly from source folder (customUI.xml, customUI14.xml, etc.)
    $customUIFiles = Get-ChildItem -Path $customUISourcePath -Filter "customUI*.xml" | ForEach-Object { $_.FullName }
    
    if ($customUIFiles.Count -eq 0) {
        throw "NO CUSTOM UI XML FILES FOUND in $($customUISourcePath). Expected: customUI.xml or customUI14.xml"
    }

    Write-Host -ForegroundColor Green "- found $($customUIFiles.Count) customUI file(s)"

    # Create a temporary directory for backup and work
    $tempDir = Join-Path $env:TEMP "excel_customui_work_$(Get-Random)"
    New-Item $tempDir -ItemType Directory | Out-Null
    
    try {
        # Copy the original macro to temp location
        $tempMacroPath = Join-Path $tempDir "macro.xlam"
        Copy-Item $macroPath $tempMacroPath
        
        Write-Host -ForegroundColor Green "- opening Excel Add-in for modification"
        
        # Open the ZIP archive for reading
        # ZipArchiveMode: 0=Read, 1=Create, 2=Update
        $zipArchive = [System.IO.Compression.ZipFile]::Open($tempMacroPath, 2)
        
        try {
            # Remove existing customUI entries from the archive
            Write-Host -ForegroundColor Green "- removing existing customUI entries"
            $entriesToRemove = @()
            
            foreach ($entry in $zipArchive.Entries) {
                $entryName = $entry.FullName.ToLower()
                if ($entryName -match "customui/customui\.xml$" -or $entryName -match "customui/customui14\.xml$") {
                    $entriesToRemove += $entry
                    Write-Host -ForegroundColor Cyan "  marked for removal: $($entry.FullName)"
                }
            }
            
            # Actually remove the marked entries
            foreach ($entry in $entriesToRemove) {
                $entry.Delete()
            }
            
            # Add new customUI files to the archive
            Write-Host -ForegroundColor Green "- adding new customUI files"
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
                    Write-Host -ForegroundColor Cyan "  added: $entryName"
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
        Write-Host -ForegroundColor Green "- saving changes to Excel Add-in"
        Remove-Item $macroPath -Force
        Move-Item $tempMacroPath $macroPath
    }
    finally {
        # Clean up temporary directory
        if (Test-Path $tempDir) {
            Remove-Item $tempDir -Recurse -Force
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
