# -*- coding: utf-8 -*-
# Common functions and initialization for Excel VBA scripts

# Initialize error handling and encoding
function Initialize-Script {
    param(
        [string]$scriptName
    )
    
    $ErrorActionPreference = "Stop"
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    
    Write-Host "${scriptName}:"
}

# Get the active Excel instance
function Get-ExcelInstance {
    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        # stop debug
        $indexOfReset = 3
        $excel.VBE.CommandBars("run").Controls($indexOfReset).Execute()
    }
    catch {
        throw "Excel not running."
    }

    return $excel
}

# Get Excel instance or create new one if not running
function Get-OrCreate-ExcelInstance {
    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        Write-Host "- using existing Excel instance"
    }
    catch {
        $excel = New-Object -ComObject Excel.Application
        Write-Host "- created new Excel instance"
    }
    
    return $excel
}

# Resolve macro file path and determine file type
function Get-BookInfo {
    param(
        [string]$bookPath
    )
    
    Write-Host "- checking if workbook file exists"
    
    $fileExtension = [System.IO.Path]::GetExtension($bookPath).ToLower()
    
    # If .url file (cloud-based), skip file existence check and use the path as is
    if ($fileExtension -eq ".url") {
        return @{
            ResolvedPath  = $bookPath
            FileExtension = $fileExtension
            IsAddIn       = $false
        }
    }
    
    # For local files, check existence
    if (-not (Test-Path $bookPath)) {
        throw "Workbook file not found: $bookPath"
    }
    
    $resolvedPath = (Resolve-Path $bookPath).Path
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
    
    Write-Host "- checking if workbook/add-in is open in Excel"
    $searchKeyFileName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath)
    
    # If the file is a .url file (cloud-based), remove the Excel extension (.xlsx/.xlsm/.xlam)
    # Example: "あああ.xlsx.url" -> GetFileNameWithoutExtension -> "あああ.xlsx" -> further remove ".xlsx" -> "あああ"
    $fileExtension = [System.IO.Path]::GetExtension($resolvedPath).ToLower()
    if ($fileExtension -eq ".url") {
        $searchKeyFileName = $searchKeyFileName -ireplace '\.(xlsm|xlsx|xlam)$', ''
    }
    
    Write-Host "  Target: $resolvedPath"
    Write-Host "  Search key (filename without ext): '$searchKeyFileName'"
    Write-Host "  File type: $(if ($isAddIn) { 'Add-in (.xlam)' } else { 'Workbook' })"
    
    # First try to search through Excel.Workbooks (works for both workbooks and add-ins)
    Write-Host "  searching Excel.Workbooks:"
    $workbookCount = $excel.Workbooks.Count
    Write-Host "  total workbooks found: $workbookCount"
    
    foreach ($wb in $excel.Workbooks) {
        $wbFullName = $wb.FullName
        $wbNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($wb.Name)
        Write-Host "    [$($wb.Name)]"
        Write-Host "      FullName: $wbFullName"
        Write-Host "      Name (without ext): '$wbNameWithoutExt'"
        
        if ($wbNameWithoutExt -eq $searchKeyFileName) {
            Write-Host "      Result: MATCHED!"
            $vbProject = $wb.VBProject
            if ($null -ne $vbProject) {
                return @{
                    VBProject = $vbProject
                    Workbook  = $wb
                    Source    = "Workbooks"
                }
            }
        }
        else {
            Write-Host "      Result: No match (compare '$wbNameWithoutExt' vs '$searchKeyFileName')"
        }
    }
    
    # If not found and it's an add-in, try VBE.VBProjects
    if ($isAddIn) {
        Write-Host "  searching VBE.VBProjects (add-in):"
        try {
            $vbe = $excel.VBE
            if ($null -eq $vbe) {
                throw "VBA object model access not enabled."
            }
            
            $vbProjects = $vbe.VBProjects
            if ($null -eq $vbProjects) {
                throw "VBA object model access not enabled."
            }
            
            $projectCount = 0
            foreach ($vbProj in $vbProjects) {
                $projectCount++
                $projectFileName = $vbProj.FileName
                $projectName = $vbProj.Name
                Write-Host "    [$projectCount] Name: '$projectName'"
                Write-Host "      FileName: $projectFileName"
                
                if ($projectFileName -eq $resolvedPath) {
                    Write-Host "      Result: MATCHED!"
                    return @{
                        VBProject = $vbProj
                        Workbook  = $null
                        Source    = "VBE.VBProjects"
                    }
                }
                else {
                    Write-Host "      Result: No match (compare '$projectFileName' vs '$resolvedPath')"
                }
            }
            Write-Host "  total projects found: $projectCount"
        }
        catch {
            Write-Host "  error accessing VBE.VBProjects: $_"
            Write-Host "  "
            Write-Host "  SOLUTION:"
            Write-Host "  1. Open Excel and go to: File > Options > Trust Center > Trust Center Settings..."
            Write-Host "  2. Click 'Macro Settings'"
            Write-Host "  3. Check the checkbox: 'Trust access to the VBA project object model'"
            Write-Host "  4. Click OK and close Excel completely"
            Write-Host "  5. Re-open the add-in and try again"
            throw $_
        }
    }
    
    # No matching workbook found
    Write-Host "  ERROR: No matching workbook/add-in found."
    Write-Host "    Expected file name: '$searchKeyFileName'"
    throw "No workbook open."
}

# Get workbook from Excel instance
function Get-Workbook {
    param(
        [object]$excel,
        [string]$bookPath
    )
    
    Write-Host "- checking if Excel file exists"
    if (-not (Test-Path $bookPath)) {
        throw "Book file not found: $bookPath"
    }
    
    # Check if the workbook is open in Excel
    $fullPath = [System.IO.Path]::GetFullPath($bookPath)
    Write-Host "- checking if workbook is open in Excel"
    
    $workbook = $null
    foreach ($openWorkbook in $excel.Workbooks) {
        if ($openWorkbook.FullName -eq $fullPath) {
            $workbook = $openWorkbook
            break
        }
    }
    
    if ($null -eq $workbook) {
        throw "Workbook not open: $fullPath"
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

# Get VBA code from Document Module by removing metadata
function Get-DocumentModuleCode {
    param([string]$content)
    
    # For Document Modules, strip metadata (VERSION, CLASS, Attribute lines)
    # Keep only the actual VBA code starting from Option Explicit or first actual code
    $lines = $content -split "`r`n"
    $codeLines = @()
    $foundCodeStart = $false
    
    foreach ($line in $lines) {
        # Once we find Option Explicit, collect everything from there on
        if ($line -match '^\s*Option\s+') {
            $foundCodeStart = $true
            $codeLines += $line
            continue
        }
        
        # Collect all lines after code start (including empty lines and everything)
        if ($foundCodeStart) {
            $codeLines += $line
            continue
        }
        
        # Before code starts, skip metadata lines
        # Match: VERSION, Begin, End (at start of line), Attribute, MultiUse, etc.
        if ($line -match '^\s*(VERSION|Begin|End|Attribute|MultiUse)\s*') {
            continue
        }
        
        # Skip empty lines before actual code starts
        if ($line -match '^\s*$') {
            continue
        }
        
        # If we find any other non-empty, non-metadata line, treat it as code start
        $foundCodeStart = $true
        $codeLines += $line
    }
    
    return $codeLines -join "`r`n"
}
