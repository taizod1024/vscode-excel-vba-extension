# -*- coding: utf-8 -*-
param(
    [Parameter(Mandatory = $true)] [string] $bookPath,
    [Parameter(Mandatory = $true)] [string] $vbaSourcePath
)

# Import common functions
. (Join-Path $PSScriptRoot "Common.ps1")

try {

    # Initialize
    Initialize-Script $MyInvocation.MyCommand.Name | Out-Null
    Write-Host "- bookPath: $($bookPath)"
    Write-Host "- vbaSourcePath: $($vbaSourcePath)"

    # Check if save source path exists
    Write-Host "- checking save source folder"
    if (-not (Test-Path $vbaSourcePath)) {
        throw "VBA source folder not found: $vbaSourcePath"
    }

    # Get Excel instance
    $excel = Get-ExcelInstance
    
    # Bring VBE window to foreground if it exists
    try {
        $vbe = $excel.VBE
        if ($null -ne $vbe) {
            $vbeCaption = $vbe.MainWindow.Caption
            Write-Host "- VBE caption: $vbeCaption"
            
            # Try to activate VBE window using WScript.Shell
            $shell = New-Object -ComObject WScript.Shell
            $shell.AppActivate($vbeCaption) | Out-Null
        }
    }
    catch {
        Write-Host "- Warning: Could not activate VBE window: $_"
    }
    
    $macro = $null

    # Get VB Project
    $macroInfo = Get-BookInfo $bookPath
    $result = Find-VBProject $excel $macroInfo.ResolvedPath $macroInfo.IsAddIn
    $vbProject = $result.VBProject
    $macro = $result.Workbook
    $isAddIn = $macroInfo.IsAddIn
    if (Test-Path $vbaSourcePath) {
        $vbaFiles = Get-ChildItem -Path $vbaSourcePath -Recurse -Include *.bas, *.cls, *.frm | ForEach-Object { $_.FullName }
    }
    
    Write-Host "- found VBA files to save: $($vbaFiles.Count)"
    
    # Get list of saved file names (without extension)
    $savedFileNames = @()
    foreach ($file in $vbaFiles) {
        $savedFileNames += [System.IO.Path]::GetFileNameWithoutExtension($file)
    }
    
    # Remove components that are no longer in the save folder
    Write-Host "- removing deleted components"
    $vbComponents = @($vbProject.VBComponents)  # Snapshot before deletion
    foreach ($component in $vbComponents) {
        # Skip Document modules (they can't be removed)
        if ($component.Type -eq 100) {
            # 100 = Document module
            continue
        }
        
        try {
            Write-Host "  - removing component: $($component.Name)"
            $vbProject.VBComponents.Remove($component)
        }
        catch {
            Write-Host "  - warning: failed to remove component '$($component.Name)': $_"
        }
    }
    
    # Wait for deletion to complete
    Write-Host "- waiting for component removal to complete"
    Start-Sleep -Seconds 1
    
    # Verify no standard modules remain
    Write-Host "- verifying standard modules removal"
    $remainingStandardModules = @()
    foreach ($comp in $vbProject.VBComponents) {
        # Type 1 = Standard Module
        if ($comp.Type -eq 1) {
            $remainingStandardModules += $comp.Name
        }
    }
    
    if ($remainingStandardModules.Count -gt 0) {
        throw "Failed to save VBA. Please retry."
    }
    
    Write-Host "  - confirmed: all old standard modules removed"
    
    # Save VBA files
    Write-Host "- saving new/updated components"
    $vbComponents = @($vbProject.VBComponents)  # Refresh after deletion
    foreach ($file in $vbaFiles) {
        try {
            $componentName = [System.IO.Path]::GetFileNameWithoutExtension($file)
            $fileExtension = [System.IO.Path]::GetExtension($file).ToLower()

            # For standard modules (.bas), class modules (.cls), and forms (.frm), 
            # remove existing component with same name before saving
            if ($fileExtension -eq ".frm" -or $fileExtension -eq ".bas" -or $fileExtension -eq ".cls") {
                $isDocumentModule = $false
                $component = $null
                
                # Find the component by name
                foreach ($comp in $vbComponents) {
                    if ($comp.Name -eq $componentName) {
                        $component = $comp
                        if ($comp.Type -eq 100) {
                            $isDocumentModule = $true
                        }
                        break
                    }
                }
                
                # For Document Modules, clear existing code and import new code
                if ($isDocumentModule -and $null -ne $component) {
                    Write-Host "  - updating Document Module: $componentName"
                    try {
                        # Read file content using Shift-JIS encoding
                        $content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::GetEncoding('shift_jis'))
                        
                        # Get VBA code from Document Module by removing metadata
                        $content = Get-DocumentModuleCode $content
                        
                        # Trim trailing whitespace
                        $content = $content -replace '\s+$', ''
                        
                        # Clear existing code in the Document Module
                        $codeModule = $component.CodeModule
                        if ($codeModule.CountOfLines -gt 0) {
                            $codeModule.DeleteLines(1, $codeModule.CountOfLines)
                        }
                        
                        # Add new code to Document Module
                        $codeModule.AddFromString($content)
                    }
                    catch {
                        throw "Failed to update module: $componentName - $_"
                    }
                }
                else {
                    # For non-Document modules, remove and reimport
                    if ($null -ne $component) {
                        Write-Host "  - removing existing component: $componentName"
                        $vbProject.VBComponents.Remove($component)
                    }
                    
                    Write-Host "  - saving: $componentName"
                    
                    # Remove blank lines before VBA code starts and trim trailing whitespace
                    $content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::GetEncoding('shift_jis'))
                    $content = Remove-BlankLinesBeforeVBACode $content
                    $content = $content -replace '\s+$', ''
                    [System.IO.File]::WriteAllText($file, $content, [System.Text.Encoding]::GetEncoding('shift_jis'))
                    
                    $vbProject.VBComponents.Import($file) | Out-Null
                }
            }
        }
        catch {
            # Check for .log file that may contain detailed error information
            $logFile = [System.IO.Path]::ChangeExtension($file, ".log")
            $logContent = ""
            if (Test-Path $logFile) {
                $logContent = Get-Content $logFile -Raw
                Remove-Item $logFile -Force
            }
            
            if ($logContent) {
                throw "Failed to import file: $file - $logContent"
            }
            else {
                throw "Failed to import file: $file - $_"
            }
        }
    }
    
    # Save the workbook or add-in
    Write-Host "- saving workbook"
    $vbe = $excel.VBE
    $vbe.MainWindow.Visible = $true
    $vbe.MainWindow.SetFocus()
    if ($null -ne $macro) {
        # For workbooks, save through the workbook object
        Write-Host "  - saving workbook"
        $macro.Save()
    }
    elseif ($isAddIn -and $null -ne $vbProject) {
        # For add-ins (.xlam), VBA components are stored in the Excel runtime
        # The file cannot be saved directly from VBProject
        Write-Host "  - Opening VB Editor for you to save manually..."
        $vbProject.Activate
    }
    
    # Compile VBA project
    Write-Host "- compiling VBA project"
    try {
        if ($null -ne $vbProject) {
            $vbe = $excel.VBE
            if ($null -ne $vbe) {
                # Execute compile command from VBE menu: Debug > Compile
                # Parameters: 1 = msoControlButton, 578 = Compile ID
                $objVBECommandBar = $vbe.CommandBars
                $compileButton = $objVBECommandBar.FindControl(1, 578)
                if ($null -ne $compileButton) {
                    $compileButton.Execute()
                    Write-Host "  - compilation executed"
                }
                else {
                    throw "Compile button not found."
                }
            }
        }
    }
    catch {
        Write-Host "  - warning: compilation encountered an issue: $_"
    }

    Write-Host "- done"
    exit 0
}
catch {
    [Console]::Error.WriteLine("$($_)")
    exit 1
}

