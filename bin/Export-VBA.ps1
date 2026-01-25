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

    # output basic information
    $app_name = $myInvocation.MyCommand.name
    Write-Host -ForegroundColor Yellow "$($app_name)"
    Write-Host -ForegroundColor Green "- bookPath: $($bookPath)"
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

    # excel object
    Write-Host -ForegroundColor Green "- open excel"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {

        # book object
        Write-Host -ForegroundColor Green "- open book"
        $book = $excel.Workbooks.Open((Resolve-Path $bookPath).Path)
        
        # vbproject object
        try {
            $vbProject = $book.VBProject
            Write-Host -ForegroundColor Green "- vbProject: $($vbProject.Name)"
        }
        catch {
            throw "VBProject is not accessible. The workbook may be protected or this is an xlsx file."
        }

        # components
        $componentCount = $vbProject.VBComponents.Count
        Write-Host -ForegroundColor Green "- components count: $componentCount"
        if ($componentCount -eq 0) {
            $book.Close($false)
            throw "No VB components found. The VBA Project Object Model needs to be enabled. Please follow these steps: 1. Open Excel, 2. File > Options > Trust Center > Trust Center Settings, 3. Macro Settings > Check 'Trust access to the VBA project object model', 4. Close Excel completely and re-open the workbook"
        }
        
        # VB IDE Object Model uses 1-based indexing
        for ($i = 1; $i -le $componentCount; $i++) {
            $component = $vbProject.VBComponents.Item($i)
            $componentName = $component.Name
            $componentType = $component.Type
            Write-Host -ForegroundColor Green "- [$i] Name=$componentName, Type=$componentType"
            
            # Skip Document Modules as they cannot be exported
            if ($componentType -eq 100) {
                Write-Host -ForegroundColor Yellow "  - skipped (Document Module)"
                continue
            }
            
            $fileExt = switch ($componentType) {
                1 { ".bas" }      # Standard Module
                2 { ".cls" }      # Class Module
                3 { ".frm" }      # Form
                default { ".bas" }
            }
            
            $filePath = Join-Path $tmpPath "$componentName$fileExt"
            Write-Host -ForegroundColor Green "  - filePath: $filePath"
            try {
                # Use explicit method call with [void] to suppress output
                [void]$component.Export($filePath)
                Write-Host -ForegroundColor Cyan "  - exported: $componentName$fileExt"
            }
            catch {
                Write-Host -ForegroundColor Red "  - export failed: $_"
                Write-Host -ForegroundColor Red "  - component type: $($component.GetType().FullName)"
                Write-Host -ForegroundColor Red "  - available methods:"
                $component | Get-Member -MemberType Method | ForEach-Object { Write-Host -ForegroundColor Gray "    - $($_.Name)" }
            }
        }
        $book.Close($false)
        $book = $null
    }
    catch {
        Write-Host -ForegroundColor Red "ERROR: $($_)"
        throw
    }
    finally {
        if ($null -ne $book) {
            Write-Host -ForegroundColor Green "- close book"
            try { $book.Close($false) } catch { }
        }
        if ($null -ne $excel) {
            Write-Host -ForegroundColor Green "- close excel"
            try { $excel.Quit() } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
    # done
    Write-Host -ForegroundColor Green "- done"
    exit 0
}
catch {
    Write-Host -ForegroundColor Red "ERROR: $($_)"
    exit 1
}
finally {
}
