# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $macroPath,
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

# Function to remove blank lines before VBA code starts
function Remove-BlankLinesBeforeVBACode {
    Param([string]$content)
    # Remove blank lines before the first code keyword (excluding Attribute, VERSION, Begin, End)
    # Only replace once using regex match
    $match = [regex]::Match($content, "[\r\n]+(?=\s*(Option|Sub|Function|Const|Private|Public|Dim|Type|Enum|Declare|Global|Static|Param|'))")
    if ($match.Success) {
        return $content.Remove($match.Index, $match.Length).Insert($match.Index, "`r`n")
    }
    return $content
}

try {

    # Display script name
    $scriptName = $MyInvocation.MyCommand.Name
    Write-Host -ForegroundColor Yellow "$($scriptName):"
    Write-Host -ForegroundColor Green "- macroPath: $($macroPath)"
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
        throw "NO EXCEL FOUND. Please Open Excel."
    }
    $macro = $null

    # Check if the workbook or add-in is already open in Excel
    Write-Host -ForegroundColor Green "- checking if workbook/add-in is open in Excel"
    $resolvedPath = (Resolve-Path $macroPath).Path
    Write-Host -ForegroundColor Cyan "  resolvedPath: $resolvedPath"
    $macro = $null
    $vbProject = $null
    
    # Determine if this is an add-in (.xlam) or workbook (.xlsm/.xlsx)
    $fileExtension = [System.IO.Path]::GetExtension($resolvedPath).ToLower()
    Write-Host -ForegroundColor Cyan "  file extension: $fileExtension"
    
    $isAddIn = ($fileExtension -eq ".xlam")
    
    # Search through Excel.Workbooks for both workbooks and add-ins
    Write-Host -ForegroundColor Cyan "  searching Excel.Workbooks:"
    $workbookCount = $excel.Workbooks.Count
    Write-Host -ForegroundColor Cyan "  total workbooks found: $workbookCount"
    
    foreach ($wb in $excel.Workbooks) {
        $wbFullName = $wb.FullName
        Write-Host -ForegroundColor Cyan "    Workbook: $($wb.Name), FullName: $wbFullName"
        
        if ($wbFullName -eq $resolvedPath) {
            Write-Host -ForegroundColor Yellow "    MATCHED!"
            $macro = $wb
            $vbProject = $wb.VBProject
            break
        }
    }
    
    # If not found in Workbooks and it's an add-in, try VBE.VBProjects
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
                    $vbProject = $vbProj
                    break
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
    $vbComponents = @($vbProject.VBComponents)
    foreach ($component in $vbComponents) {
        # Skip Document modules (they can't be removed)
        if ($component.Type -eq 100) {
            # 100 = Document module
            continue
        }
        
        if (-not ($saveedFileNames -contains $component.Name)) {
            try {
                Write-Host -ForegroundColor Green "  - removing component: $($component.Name)"
                $vbProject.VBComponents.Remove($component)
            }
            catch {
                Write-Host -ForegroundColor Yellow "  - warning: failed to remove component '$($component.Name)': $_"
            }
        }
    }
    
    # Save VBA files
    Write-Host -ForegroundColor Green "- saving new/updated components"
    $vbComponents = @($vbProject.VBComponents)
    foreach ($file in $vbaFiles) {
        try {
            $fileName = [System.IO.Path]::GetFileName($file)
            $componentName = [System.IO.Path]::GetFileNameWithoutExtension($file)
            $fileExtension = [System.IO.Path]::GetExtension($file).ToLower()
            $filePath = $file

            # For standard modules (.bas), class modules (.cls), and forms (.frm), 
            # remove existing component with same name before saving
            if ($fileExtension -eq ".frm" -or $fileExtension -eq ".bas" -or $fileExtension -eq ".cls") {
                foreach ($component in $vbComponents) {
                    if ($component.Name -eq $componentName) {
                        Write-Host -ForegroundColor Green "  - removing existing component: $componentName"
                        $vbProject.VBComponents.Remove($component)
                        break
                    }
                }
                
                Write-Host -ForegroundColor Green "  - saving: $fileName"
                
                # Remove trailing whitespace and blank lines before import
                $content = [System.IO.File]::ReadAllText($filePath, [System.Text.Encoding]::GetEncoding('shift_jis'))
                
                # Remove blank lines before VBA code starts
                $content = Remove-BlankLinesBeforeVBACode $content
                
                $content = $content -replace '\s+$', ''
                [System.IO.File]::WriteAllText($filePath, $content, [System.Text.Encoding]::GetEncoding('shift_jis'))
                
                $vbProject.VBComponents.Import($filePath) | Out-Null
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
    
    # Save the workbook or add-in
    Write-Host -ForegroundColor Green "- saving workbook/add-in"
    if ($null -ne $macro) {
        # For workbooks, save through the workbook object
        Write-Host -ForegroundColor Green "  - saving workbook"
        $macro.Save()
    }
    elseif ($isAddIn -and $null -ne $vbProject) {
        # For add-ins (.xlam), VBA components are stored in the Excel runtime
        # The file cannot be saved directly from VBProject
        Write-Host -ForegroundColor Yellow "  - Opening VB Editor for you to save manually..."
        
        try {
            # Show VB Editor and bring it to foreground
            $vbe = $excel.VBE
            $vbe.MainWindow.Visible = $true
            $vbe.MainWindow.SetFocus()
            $vbProject.ActivateVBProject()
        }
        catch {
            Write-Host -ForegroundColor Yellow "  - Could not open VB Editor automatically"
        }
        
        throw "ADD-IN (.XLAM) CANNOT BE SAVED AUTOMATICALLY. Please save manually from Excel using Ctrl+S."
    }
    else {
        Write-Host -ForegroundColor Yellow "  WARNING: Could not find way to save. Please save manually."
    }
    
    # Compile VBA project
    Write-Host -ForegroundColor Green "- compiling VBA project"
    try {
        if ($null -ne $vbProject) {
            # Execute compile command from VBE menu: Debug > Compile
            $vbe = $excel.VBE
            if ($null -ne $vbe) {
                # Make VBE visible temporarily
                $vbe.MainWindow.Visible = $true
                
                # Try to execute "Compile" from Debug menu
                $objVBECommandBar = $vbe.CommandBars
                $compileButton = $objVBECommandBar.FindControl(1, 578)  # 1 = msoControlButton, 578 = Compile ID
                if ($null -ne $compileButton) {
                    $compileButton.Execute()
                    Write-Host -ForegroundColor Green "  - compilation executed"
                }
                else {
                    throw "Could not find compile button"
                }
            }
        }
    }
    catch {
        Write-Host -ForegroundColor Yellow "  - warning: compilation encountered an issue: $_"
    }
    
    Write-Host -ForegroundColor Green "- done"
    exit 0
}
catch {
    [Console]::Error.WriteLine("$($_)")
    exit 1
}

