# -*- coding: utf-8 -*-
# Common functions and initialization for Excel VBA scripts

# Initialize error handling and encoding
function Initialize-Script {
    param(
        [string]$scriptName
    )
    
    $ErrorActionPreference = "Stop"
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    
    Write-Host -ForegroundColor Yellow "${scriptName}:"
    
    return @{
        ScriptName = $scriptName
    }
}

# Get the active Excel instance
function Get-ExcelInstance {
    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        throw "EXCEL WORKBOOK NOT FOUND. Please Open Excel first."
    }
    
    return $excel
}

# Resolve macro file path and determine file type
function Get-BookInfo {
    param(
        [string]$bookPath
    )
    
    Write-Host -ForegroundColor Green "- checking if book file exists"
    if (-not (Test-Path $bookPath)) {
        throw "BOOK FILE NOT FOUND: $bookPath"
    }
    
    $resolvedPath = (Resolve-Path $bookPath).Path
    $fileExtension = [System.IO.Path]::GetExtension($resolvedPath).ToLower()
    $isAddIn = ($fileExtension -eq ".xlam")
    
    return @{
        ResolvedPath  = $resolvedPath
        FileExtension = $fileExtension
        IsAddIn       = $isAddIn
    }
}

# Find VB Project from Excel instance
function Find-VBProject {
    param(
        [object]$excel,
        [string]$resolvedPath,
        [bool]$isAddIn
    )
    
    $vbProject = $null
    
    Write-Host -ForegroundColor Green "- checking if workbook/add-in is open in Excel"
    Write-Host -ForegroundColor Cyan "  resolvedPath: $resolvedPath"
    Write-Host -ForegroundColor Cyan "  file extension: $(if ($isAddIn) { '.xlam' } else { 'other' })"
    
    # First try to search through Excel.Workbooks (works for both workbooks and add-ins)
    Write-Host -ForegroundColor Cyan "  searching Excel.Workbooks:"
    $workbookCount = $excel.Workbooks.Count
    Write-Host -ForegroundColor Cyan "  total workbooks found: $workbookCount"
    
    foreach ($wb in $excel.Workbooks) {
        $wbFullName = $wb.FullName
        Write-Host -ForegroundColor Cyan "    Workbook: $($wb.Name), FullName: $wbFullName"
        
        if ($wbFullName -eq $resolvedPath) {
            Write-Host -ForegroundColor Yellow "    MATCHED!"
            $vbProject = $wb.VBProject
            if ($null -ne $vbProject) {
                return @{
                    VBProject = $vbProject
                    Workbook  = $wb
                    Source    = "Workbooks"
                }
            }
        }
    }
    
    # If not found and it's an add-in, try VBE.VBProjects
    if ($null -eq $vbProject -and $isAddIn) {
        Write-Host -ForegroundColor Cyan "  not found in Workbooks, searching VBE.VBProjects (add-in):"
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
                    return @{
                        VBProject = $vbProj
                        Workbook  = $null
                        Source    = "VBE.VBProjects"
                    }
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
    }
    
    if ($null -eq $vbProject) {
        throw "NO OPENED WORKBOOK OR ADD-IN FOUND. Please Open Workbook or Add-in."
    }
    
    return @{
        VBProject = $vbProject
        Workbook  = $null
        Source    = "Unknown"
    }
}

# Get workbook from Excel instance
function Get-Workbook {
    param(
        [object]$excel,
        [string]$excelFilePath
    )
    
    Write-Host -ForegroundColor Green "- checking if Excel file exists"
    if (-not (Test-Path $excelFilePath)) {
        throw "EXCEL FILE NOT FOUND: $($excelFilePath)"
    }
    
    # Check if the workbook is open in Excel
    $fullPath = [System.IO.Path]::GetFullPath($excelFilePath)
    Write-Host -ForegroundColor Green "- checking if workbook is open in Excel"
    
    $workbook = $null
    foreach ($openWorkbook in $excel.Workbooks) {
        if ($openWorkbook.FullName -eq $fullPath) {
            $workbook = $openWorkbook
            break
        }
    }
    
    if ($null -eq $workbook) {
        throw "EXCEL WORKBOOK NOT OPEN: $($fullPath) is not currently open in Excel"
    }
    
    return $workbook
}

# Remove PathToLongDirectory helper
function Remove-PathToLongDirectory {
    param(
        [string]$directory
    )

    # create a temporary (empty) directory
    $parent = [System.IO.Path]::GetTempPath()
    [string]$name = [System.Guid]::NewGuid()
    $tempDirectory = New-Item -ItemType Directory -Path (Join-Path $parent $name)

    robocopy /MIR $tempDirectory.FullName $directory | out-null
    Remove-Item $directory -Force | out-null
    Remove-Item $tempDirectory -Force | out-null
}

# Remove blank lines before VBA code starts
function Remove-BlankLinesBeforeVBACode {
    param([string]$content)
    # Remove blank lines before the first code keyword (excluding Attribute, VERSION, Begin, End)
    # Only replace once using regex match
    $match = [regex]::Match($content, "[\r\n]+(?=\s*(Option|Sub|Function|Const|Private|Public|Dim|Type|Enum|Declare|Global|Static|Param|'))")
    if ($match.Success) {
        return $content.Remove($match.Index, $match.Length).Insert($match.Index, "`r`n")
    }
    return $content
}
