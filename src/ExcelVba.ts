import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import child_process from "child_process";

/** Excel VBA extension class */
class ExcelVba {
  /** application id */
  public appId = "excel-vba";

  /** application name */
  public appName = "Excel VBA";

  /** output channel */
  public channel: vscode.OutputChannel;

  /** extension path */
  public extensionPath: string;

  /** constructor */
  constructor() {}

  /** resolve VBA path from selected file */
  public resolveVbaPath(selectedPath: string): string {
    let resolvedPath = selectedPath;

    // Handle temporary Excel files (~$filename.xlsx)
    const fileName = path.basename(selectedPath);
    if (fileName.startsWith("~$")) {
      const dir = path.dirname(selectedPath);
      const actualFileName = fileName.substring(2); // Remove ~$ prefix
      resolvedPath = path.join(dir, actualFileName);
    }

    const ext = path.extname(resolvedPath).toLowerCase();

    // If .url file is selected, treat it as a marker for cloud-based files
    // Use the corresponding local Excel file if it exists
    if (ext === ".url") {
      const dir = path.dirname(resolvedPath);
      const fileNameWithoutExt = path.parse(resolvedPath).name;

      // Try to find .xlsm first, then .xlsx, then .xlam
      const xlsmPath = path.join(dir, `${fileNameWithoutExt}.xlsm`);
      if (fs.existsSync(xlsmPath)) {
        return xlsmPath;
      }

      const xlsxPath = path.join(dir, `${fileNameWithoutExt}.xlsx`);
      if (fs.existsSync(xlsxPath)) {
        return xlsxPath;
      }

      const xlamPath = path.join(dir, `${fileNameWithoutExt}.xlam`);
      if (fs.existsSync(xlamPath)) {
        return xlamPath;
      }

      // If local file doesn't exist, return the .url path itself
      // This will allow CSV/BAS/XML operations to use the corresponding folders
      return resolvedPath;
    }

    // If .xlsm, .xlam or .xlsx is selected, return as is
    if (ext === ".xlsm" || ext === ".xlam" || ext === ".xlsx") {
      return resolvedPath;
    }

    // If .csv is selected, find the parent _CSV folder and the corresponding Excel file
    if (ext === ".csv") {
      const parentDir = path.dirname(resolvedPath);
      let parentName = path.basename(parentDir);

      // Check if parent folder is _csv
      const match = parentName.match(/^(.+)_csv$/i);
      if (match) {
        const macroName = match[1];
        const parentParentDir = path.dirname(parentDir);

        // Try to find .xlsm first, then .xlsx, then .xlam, then .url
        const xlsmPath = path.join(parentParentDir, `${macroName}.xlsm`);
        if (fs.existsSync(xlsmPath)) {
          return xlsmPath;
        }

        const xlsxPath = path.join(parentParentDir, `${macroName}.xlsx`);
        if (fs.existsSync(xlsxPath)) {
          return xlsxPath;
        }

        const xlamPath = path.join(parentParentDir, `${macroName}.xlam`);
        if (fs.existsSync(xlamPath)) {
          return xlamPath;
        }

        const urlPath = path.join(parentParentDir, `${macroName}.url`);
        if (fs.existsSync(urlPath)) {
          return urlPath;
        }
      }
    }

    // If .bas, .cls, .frm is selected, find the parent _bas folder
    if ([".bas", ".cls", ".frm"].includes(ext)) {
      const parentDir = path.dirname(resolvedPath);
      let parentName = path.basename(parentDir);

      // Remove trailing ~ from parent folder name
      if (parentName.endsWith("~")) {
        parentName = parentName.slice(0, -1);
      }

      // Check if parent folder is _bas
      const match = parentName.match(/^(.+)_bas$/i);
      if (match) {
        const macroName = match[1];
        const parentParentDir = path.dirname(parentDir);

        // Try to find .xlsm first, then .xlsx, then .xlam, then .url
        const xlsmPath = path.join(parentParentDir, `${macroName}.xlsm`);
        if (fs.existsSync(xlsmPath)) {
          return xlsmPath;
        }

        const xlsxPath = path.join(parentParentDir, `${macroName}.xlsx`);
        if (fs.existsSync(xlsxPath)) {
          return xlsxPath;
        }

        const xlamPath = path.join(parentParentDir, `${macroName}.xlam`);
        if (fs.existsSync(xlamPath)) {
          return xlamPath;
        }

        const urlPath = path.join(parentParentDir, `${macroName}.url`);
        if (fs.existsSync(urlPath)) {
          return urlPath;
        }
      }
    }

    // If .xml is selected in a _xml folder, find the parent .xlam or .xlsm file
    if (ext === ".xml") {
      const parentDir = path.dirname(resolvedPath);
      let parentName = path.basename(parentDir);

      // Remove trailing ~ from parent folder name
      if (parentName.endsWith("~")) {
        parentName = parentName.slice(0, -1);
      }

      // Check if parent folder is _xml
      const match = parentName.match(/^(.+)_xml$/i);
      if (match) {
        const macroName = match[1];
        const parentParentDir = path.dirname(parentDir);

        // Try to find .xlam first, then .xlsm, then .url
        const xlamPath = path.join(parentParentDir, `${macroName}.xlam`);
        if (fs.existsSync(xlamPath)) {
          return xlamPath;
        }

        const xlsmPath = path.join(parentParentDir, `${macroName}.xlsm`);
        if (fs.existsSync(xlsmPath)) {
          return xlsmPath;
        }

        const urlPath = path.join(parentParentDir, `${macroName}.url`);
        if (fs.existsSync(urlPath)) {
          return urlPath;
        }

        // Default to .xlam if neither exists (will be handled as error later)
        return xlamPath;
      }
    }

    return resolvedPath;
  }

  /** activate extension */
  public activate(context: vscode.ExtensionContext) {
    // init context
    this.channel = vscode.window.createOutputChannel(this.appName, { log: true });
    if (!process.env.WINDIR) {
      this.channel.appendLine(`[ERROR] Failed to activate: Windows directory not found`);
      return;
    }
    this.channel.appendLine(`${this.appName} extension activated`);

    // init vscode
    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.openExcel`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          const selectedPath = uri.fsPath;
          this.channel.appendLine(`[DEBUG] Selected path: ${selectedPath}`);
          const macroPath = this.resolveVbaPath(selectedPath);
          this.channel.appendLine(`[DEBUG] Resolved path: ${macroPath}`);
          await this.openExcelAsync(macroPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.loadVba`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await this.loadVbaAsync(macroPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.saveVba`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await this.saveVbaAsync(macroPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.compareVba`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await this.compareVbaAsync(macroPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.loadCustomUI`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await this.loadCustomUIAsync(macroPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.saveCustomUI`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await this.saveCustomUIAsync(macroPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.runSub`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          const editor = vscode.window.activeTextEditor;
          if (!editor) {
            throw `No active editor found`;
          }
          const macroPath = this.resolveVbaPath(uri.fsPath);
          const subName = this.extractSubNameAtCursor(editor);
          if (!subName) {
            throw `No Sub procedure found at cursor position`;
          }
          await this.saveVbaAsync(macroPath);
          await this.runSubAsync(macroPath, subName);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.loadCsv`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await this.loadCsvAsync(macroPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.saveCsv`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await this.saveCsvAsync(macroPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.newExcel`, async () => {
        this.extensionPath = context.extensionPath;
        try {
          await this.newExcelAsync();
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.createDummyUrlShortcut`, async () => {
        this.extensionPath = context.extensionPath;
        try {
          await this.createDummyUrlShortcutAsync();
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );
  }

  /** load vba */
  public async loadVbaAsync(macroPath: string) {
    const commandName = "Load VBA from Excel Book";
    return vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: commandName,
        cancellable: false,
      },
      async _progress => {
        // setup command
        const macroFileName = path.parse(macroPath).name;
        const macroDir = path.dirname(macroPath);
        const tmpPath = path.join(macroDir, `${macroFileName}_bas~`);
        const scriptPath = `${this.extensionPath}\\bin\\Load-VBA.ps1`;
        this.channel.appendLine("");
        this.channel.appendLine(`${commandName}`);
        this.channel.appendLine(`- File: ${path.basename(macroPath)}`);

        // exec command
        const result = this.execPowerShell(scriptPath, [macroPath, tmpPath]);

        // output result
        if (result.stdout) this.channel.appendLine(`- Output: ${result.stdout}`);
        if (result.exitCode !== 0) {
          throw `${result.stderr}`;
        }

        // Organize loaded files
        const newFolderName = `${macroFileName}_bas`;
        const newPath = path.join(macroDir, newFolderName);

        // Remove existing folder if it exists
        if (fs.existsSync(newPath)) {
          fs.rmSync(newPath, { recursive: true, force: true });
        }

        // Move tmpPath to new location
        fs.renameSync(tmpPath, newPath);

        // Close all diff editors
        await this.closeAllDiffEditors();

        // Open the first file in the explorer view if not already in this folder
        const files = fs.readdirSync(newPath).filter(file => {
          const ext = path.extname(file).toLowerCase();
          return [".bas", ".cls", ".frm", ".xml"].includes(ext);
        });

        if (files.length > 0) {
          // Only open if no active editor or the active editor's file doesn't exist in new folder
          const activeEditor = vscode.window.activeTextEditor;
          let shouldOpen = true;

          if (activeEditor) {
            const activeEditorFileName = path.basename(activeEditor.document.uri.fsPath);
            const activeEditorExists = files.includes(activeEditorFileName);
            shouldOpen = !activeEditorExists;
          }

          if (shouldOpen) {
            const firstFile = path.join(newPath, files[0]);
            const uri = vscode.Uri.file(firstFile);
            await vscode.commands.executeCommand("vscode.open", uri);
          }
        }

        this.channel.appendLine(`[SUCCESS] Loaded files organized`);
      },
    );
  }

  /** Validate that Attribute VB_Name matches file names */
  private async validateVbNames(folderPath: string): Promise<void> {
    const walkDir = (dir: string): string[] => {
      let results: string[] = [];
      const files = fs.readdirSync(dir);

      for (const file of files) {
        const filePath = path.join(dir, file);
        const stat = fs.statSync(filePath);

        if (stat.isDirectory()) {
          results = results.concat(this.validateVbNamesHelper(filePath));
        } else {
          results.push(filePath);
        }
      }
      return results;
    };

    const vbaFiles = walkDir(folderPath).filter(filePath => {
      const ext = path.extname(filePath).toLowerCase();
      return [".bas", ".cls", ".frm"].includes(ext);
    });

    for (const filePath of vbaFiles) {
      const fileName = path.basename(filePath);
      const componentName = path.parse(fileName).name;

      try {
        const content = fs.readFileSync(filePath, { encoding: "utf-8" });
        const attributeMatch = content.match(/Attribute\s+VB_Name\s*=\s*"([^"]+)"/);

        if (attributeMatch) {
          const vbName = attributeMatch[1];
          if (vbName !== componentName) {
            throw new Error(`MISMATCH Attribute VB_Name: "${vbName}" != "${componentName}" in file ${fileName}`);
          }
        }
      } catch (error) {
        if (error instanceof Error) {
          throw error.message;
        }
        throw error;
      }
    }
  }

  private validateVbNamesHelper(dir: string): string[] {
    let results: string[] = [];
    const files = fs.readdirSync(dir);

    for (const file of files) {
      const filePath = path.join(dir, file);
      const stat = fs.statSync(filePath);

      if (stat.isDirectory()) {
        results = results.concat(this.validateVbNamesHelper(filePath));
      } else {
        results.push(filePath);
      }
    }
    return results;
  }

  /** save vba */
  public async saveVbaAsync(macroPath: string) {
    const commandName = "Save VBA to Excel Book";
    return vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: commandName,
        cancellable: false,
      },
      async _progress => {
        // setup command
        const macroFileName = path.parse(macroPath).name;
        const macroDir = path.dirname(macroPath);
        const saveSourcePath = path.join(macroDir, `${macroFileName}_bas`);
        const scriptPath = `${this.extensionPath}\\bin\\Save-VBA.ps1`;
        this.channel.appendLine("");
        this.channel.appendLine(`${commandName}`);
        this.channel.appendLine(`- File: ${path.basename(macroPath)}`);
        this.channel.appendLine(`- Source: ${path.basename(saveSourcePath)}`);

        // Check if save source folder exists
        if (!fs.existsSync(saveSourcePath)) {
          throw `Folder not found: ${path.basename(saveSourcePath)}. Please load VBA first.`;
        }

        // Validate VB_Name attribute matches file names
        await this.validateVbNames(saveSourcePath);

        // exec command
        const result = this.execPowerShell(scriptPath, [macroPath, saveSourcePath]);

        // output result
        if (result.stdout) this.channel.appendLine(`- Output: ${result.stdout}`);
        if (result.exitCode !== 0) {
          throw `${result.stderr}`;
        }

        // Remove temporary folder if it exists
        const tmpPath = path.join(macroDir, `${macroFileName}_bas~`);
        if (fs.existsSync(tmpPath)) {
          fs.rmSync(tmpPath, { recursive: true, force: true });
          this.channel.appendLine(`- Cleaned: Temporary folder removed`);
        }
        this.channel.appendLine(`[SUCCESS] VBA saved`);

        // Show warning for add-in files
        const ext = path.extname(macroPath).toLowerCase();
        if (ext === ".xlam") {
          vscode.window.showWarningMessage(".XLAM CANNOT BE SAVED AUTOMATICALLY. Please save manually from VBE using Ctrl+S.");
        }

        // Close all diff editors
        await this.closeAllDiffEditors();
      },
    );
  }

  /** execute PowerShell script */
  public execPowerShell(scriptPath: string, args: string[], trim = true): { stdout: string; stderr: string; exitCode: number } {
    try {
      const result = child_process.spawnSync("powershell.exe", ["-ExecutionPolicy", "RemoteSigned", "-File", scriptPath, ...args], {
        encoding: "utf8",
        stdio: ["pipe", "pipe", "pipe"],
      });
      let stdout = result.stdout || "";
      let stderr = result.stderr || "";
      if (result.error) {
        return { stdout: "", stderr: result.error.message, exitCode: 1 };
      }
      return {
        stdout: trim ? stdout.trim() : stdout,
        stderr: trim ? stderr.trim() : stderr,
        exitCode: result.status || 0,
      };
    } catch (ex: any) {
      return {
        stdout: "",
        stderr: trim ? (ex.message || "").trim() : ex.message || "",
        exitCode: 1,
      };
    }
  }

  /** open Excel Book */
  public async openExcelAsync(macroPath: string) {
    const commandName = "Open Excel Book";

    // Check if file is .url
    const ext = path.extname(macroPath).toLowerCase();
    if (ext === ".url") {
      vscode.window.showWarningMessage("Cannot open .url files directly. Please open the cloud-hosted Excel file in your web browser.");
      this.channel.appendLine("");
      this.channel.appendLine(`${commandName}`);
      this.channel.appendLine(`- File: ${path.basename(macroPath)}`);
      this.channel.appendLine(`[WARNING] Cannot open .url files. Please open the cloud-hosted file directly.`);
      return;
    }

    this.channel.appendLine("");
    this.channel.appendLine(`${commandName}`);
    this.channel.appendLine(`- File: ${path.basename(macroPath)}`);
    child_process.spawn("cmd.exe", ["/c", "start", "excel.exe", macroPath], { detached: true });
    this.channel.appendLine(`[SUCCESS] Opened in Excel`);
  }

  /** compare VBA with existing folder */
  public async compareVbaAsync(macroPath: string) {
    const commandName = "Compare VBA with Excel Book";
    return vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: commandName,
        cancellable: false,
      },
      async _progress => {
        const macroFileName = path.parse(macroPath).name;
        const macroExtension = path.parse(macroPath).ext.replace(".", "");
        const currentFolderName = `${macroFileName}_${macroExtension}`;
        const macroDir = path.dirname(macroPath);
        const currentPath = path.join(macroDir, currentFolderName);
        const tmpPath = path.join(macroDir, `${macroFileName}_${macroExtension}~`);

        this.channel.appendLine("");
        this.channel.appendLine(`${commandName}`);
        this.channel.appendLine(`- File: ${path.basename(macroPath)}`);
        this.channel.appendLine(`- Current: ${path.basename(currentPath)}`);
        this.channel.appendLine(`- Loading from Excel...`);

        if (!fs.existsSync(currentPath)) {
          throw `Folder not found: ${path.basename(currentPath)}. Please load VBA first.`;
        }

        // Load to temporary folder
        const scriptPath = `${this.extensionPath}\\bin\\Load-VBA.ps1`;
        const result = this.execPowerShell(scriptPath, [macroPath, tmpPath]);

        if (result.exitCode !== 0) {
          throw `${result.stderr}`;
        }

        // Compare files
        const hasDifferences = this.compareDirectories(tmpPath, currentPath);

        // Remove temporary folder only if no differences
        if (!hasDifferences && fs.existsSync(tmpPath)) {
          fs.rmSync(tmpPath, { recursive: true, force: true });
        }

        // Show the channel
        this.channel.show();
      },
    );
  }

  /** compare two directories and output differences */
  private compareDirectories(dir1: string, dir2: string): boolean {
    const files1 = this.getVbaFiles(dir1);
    const files2 = this.getVbaFiles(dir2);

    this.channel.appendLine(`Comparison Results:`);
    this.channel.appendLine(`- Files in Excel: ${files1.length}`);
    this.channel.appendLine(`- Files on disk: ${files2.length}`);

    const added = files1.filter(f => !files2.includes(f));
    const removed = files2.filter(f => !files1.includes(f));
    const common = files1.filter(f => files2.includes(f));

    let modifiedCount = 0;
    let firstModifiedFile: { file1Path: string; file2Path: string; name: string } | null = null;

    if (added.length > 0) {
      this.channel.appendLine(`- [+] Added (${added.length}):`);
      added.forEach(f => {
        const relativePath = f.replace(/\\/g, "/");
        this.channel.appendLine(`    ${relativePath}`);
      });
    }

    if (removed.length > 0) {
      this.channel.appendLine(`- [-] Removed (${removed.length}):`);
      removed.forEach(f => {
        const relativePath = f.replace(/\\/g, "/");
        this.channel.appendLine(`    ${relativePath}`);
      });
    }

    if (common.length > 0) {
      this.channel.appendLine(`- [~] Modified:`);
      common.forEach(f => {
        const file1Path = path.join(dir1, f);
        const file2Path = path.join(dir2, f);
        const content1 = fs.readFileSync(file1Path, "utf8");
        const content2 = fs.readFileSync(file2Path, "utf8");
        if (content1 !== content2) {
          const relativePath = f.replace(/\\/g, "/");
          this.channel.appendLine(`    ${relativePath}`);
          modifiedCount++;
          if (!firstModifiedFile) {
            firstModifiedFile = { file1Path, file2Path, name: relativePath };
          }
        }
      });
    }

    // Summary and return whether differences exist
    const hasDifferences = added.length > 0 || removed.length > 0 || modifiedCount > 0;
    if (hasDifferences) {
      this.channel.appendLine(`[WARN] Differences found: +${added.length} ~${modifiedCount} -${removed.length}`);
    } else {
      this.channel.appendLine(`[SUCCESS] No differences found`);
    }

    // Display first modified file in diff view
    if (firstModifiedFile) {
      this.showDiffAsync(firstModifiedFile.file1Path, firstModifiedFile.file2Path, firstModifiedFile.name);
    }

    return hasDifferences;
  }

  /** show diff between two files */
  private async showDiffAsync(file1Path: string, file2Path: string, title: string) {
    const file1Uri = vscode.Uri.file(file1Path);
    const file2Uri = vscode.Uri.file(file2Path);
    await vscode.commands.executeCommand("vscode.diff", file1Uri, file2Uri, `Compare: ${title}`);
  }

  /** close all diff editors */
  private async closeAllDiffEditors() {
    for (const group of vscode.window.tabGroups.all) {
      const tabsToClose = group.tabs.filter(tab => tab.input instanceof vscode.TabInputTextDiff);
      for (const tab of tabsToClose) {
        try {
          await vscode.window.tabGroups.close(tab);
        } catch (error) {
          // Ignore errors if tab is already closed
          this.channel.appendLine(`- note: tab already closed or not found`);
        }
      }
    }
  }

  /** get all VBA files in directory recursively */
  private getVbaFiles(dir: string, baseDir: string = ""): string[] {
    if (!fs.existsSync(dir)) {
      return [];
    }

    let files: string[] = [];
    const entries = fs.readdirSync(dir, { withFileTypes: true });

    entries.forEach(entry => {
      const fullPath = path.join(dir, entry.name);
      const relativePath = baseDir ? path.join(baseDir, entry.name) : entry.name;

      if (entry.isDirectory()) {
        files = files.concat(this.getVbaFiles(fullPath, relativePath));
      } else if ([".bas", ".cls", ".frm"].includes(path.extname(entry.name))) {
        files.push(relativePath);
      }
    });

    return files;
  }

  /** load customUI */
  public async loadCustomUIAsync(macroPath: string) {
    const ext = path.extname(macroPath).toLowerCase();

    // CustomUI is supported for .xlam (add-ins) and .xlsm (workbooks)
    if (ext !== ".xlam" && ext !== ".xlsm") {
      throw `CustomUI is only supported for .xlam and .xlsm files. Selected file: ${macroPath}`;
    }

    const fileType = ext === ".xlam" ? "Excel Add-in" : "Excel Workbook";
    const commandName = `Load CustomUI from ${fileType}`;
    return vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: commandName,
        cancellable: false,
      },
      async _progress => {
        // setup command
        const macroFileName = path.parse(macroPath).name;
        const macroDir = path.dirname(macroPath);
        const tmpPath = path.join(macroDir, `${macroFileName}_xml~`);
        const scriptPath = `${this.extensionPath}\\bin\\Load-CustomUI.ps1`;
        this.channel.appendLine("");
        this.channel.appendLine(`${commandName}`);
        this.channel.appendLine(`- File: ${path.basename(macroPath)}`);

        // exec command
        const result = this.execPowerShell(scriptPath, [macroPath, tmpPath]);

        // output result
        if (result.stdout) this.channel.appendLine(`- Output: ${result.stdout}`);
        if (result.exitCode !== 0) {
          throw `${result.stderr}`;
        }

        // Organize loaded files
        const newFolderName = `${macroFileName}_xml`;
        const newPath = path.join(macroDir, newFolderName);

        // Remove existing folder if it exists
        if (fs.existsSync(newPath)) {
          fs.rmSync(newPath, { recursive: true, force: true });
        }

        // Move tmpPath to new location
        fs.renameSync(tmpPath, newPath);
        this.channel.appendLine(`- Organized: Files moved`);

        // Verify files exist in new location
        if (fs.existsSync(newPath)) {
          const files = fs.readdirSync(newPath);
          this.channel.appendLine(`[SUCCESS] Loaded ${files.length} file(s)`);
        }

        // Close all diff editors
        await this.closeAllDiffEditors();

        // Open the first file in the explorer view if not already in this folder
        const files = fs.readdirSync(newPath).filter(file => {
          const ext = path.extname(file).toLowerCase();
          return [".xml"].includes(ext);
        });

        if (files.length > 0) {
          // Only open if no active editor or the active editor's file doesn't exist in new folder
          const activeEditor = vscode.window.activeTextEditor;
          let shouldOpen = true;

          if (activeEditor) {
            const activeEditorFileName = path.basename(activeEditor.document.uri.fsPath);
            const activeEditorExists = files.includes(activeEditorFileName);
            shouldOpen = !activeEditorExists;
          }

          if (shouldOpen) {
            const firstFile = path.join(newPath, files[0]);
            const uri = vscode.Uri.file(firstFile);
            await vscode.commands.executeCommand("vscode.open", uri);
          }
        }

        this.channel.appendLine(`[SUCCESS] Loaded ${files.length} file(s)`);
      },
    );
  }

  /** save customUI */
  public async saveCustomUIAsync(macroPath: string) {
    const ext = path.extname(macroPath).toLowerCase();

    // CustomUI is supported for .xlam (add-ins) and .xlsm (workbooks)
    if (ext !== ".xlam" && ext !== ".xlsm") {
      throw `CustomUI is only supported for .xlam and .xlsm files. Selected file: ${macroPath}`;
    }

    const fileType = ext === ".xlam" ? "Excel Add-in" : "Excel Workbook";
    const commandName = `Save CustomUI to ${fileType}`;
    return vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: commandName,
        cancellable: false,
      },
      async _progress => {
        // setup command
        const macroFileName = path.parse(macroPath).name;
        const macroDir = path.dirname(macroPath);
        const saveSourcePath = path.join(macroDir, `${macroFileName}_xml`);
        const scriptPath = `${this.extensionPath}\\bin\\Save-CustomUI.ps1`;
        this.channel.appendLine("");
        this.channel.appendLine(`${commandName}`);
        this.channel.appendLine(`- File: ${path.basename(macroPath)}`);
        this.channel.appendLine(`- Source: ${path.basename(saveSourcePath)}`);

        // Check if save source folder exists
        if (!fs.existsSync(saveSourcePath)) {
          throw `Folder not found: ${path.basename(saveSourcePath)}. Please load CustomUI first.`;
        }

        // exec command
        const result = this.execPowerShell(scriptPath, [macroPath, saveSourcePath]);

        // output result
        if (result.stdout) this.channel.appendLine(`- Output: ${result.stdout}`);
        if (result.exitCode !== 0) {
          throw `${result.stderr}`;
        }

        // Remove temporary folder if it exists
        const tmpPath = path.join(macroDir, `${macroFileName}_xml~`);
        if (fs.existsSync(tmpPath)) {
          fs.rmSync(tmpPath, { recursive: true, force: true });
          this.channel.appendLine(`- Cleaned: Temporary folder removed`);
        }
        this.channel.appendLine(`[SUCCESS] CustomUI saved`);

        // Close all diff editors
        await this.closeAllDiffEditors();
      },
    );
  }

  /** extract sub name at cursor position */
  private extractSubNameAtCursor(editor: vscode.TextEditor): string | null {
    const cursorLine = editor.selection.active.line;
    const document = editor.document;

    // Search backwards and forwards for Sub/Function declaration
    let subName: string | null = null;

    // Search from cursor backwards to find the Sub/Function this cursor is in
    for (let i = cursorLine; i >= 0; i--) {
      const line = document.lineAt(i).text;

      // Match Sub or Function declaration - supports Japanese and other Unicode characters
      const match = line.match(/^\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+([\w\u0080-\uFFFF]+)\s*(?:\(|$)/i);
      if (match) {
        subName = match[1];
        break;
      }

      // Stop if we encounter End Sub/Function before finding declaration
      if (line.match(/^\s*End\s+(?:Sub|Function)\s*$/i) && i !== cursorLine) {
        break;
      }
    }

    return subName;
  }

  /** run sub in Excel */
  public async runSubAsync(macroPath: string, subName: string) {
    const commandName = `Run VBA Sub: ${subName}`;
    return vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: commandName,
        cancellable: false,
      },
      async _progress => {
        const scriptPath = `${this.extensionPath}\\bin\\Run-Sub.ps1`;
        this.channel.appendLine("");
        this.channel.appendLine(`${commandName}`);
        this.channel.appendLine(`- File: ${path.basename(macroPath)}`);
        this.channel.appendLine(`- Sub: ${subName}`);

        // exec command
        const result = this.execPowerShell(scriptPath, [macroPath, subName]);

        // output result
        if (result.stdout) this.channel.appendLine(`- Output: ${result.stdout}`);
        if (result.exitCode !== 0) {
          throw `${result.stderr}`;
        }

        this.channel.appendLine(`[SUCCESS] Sub executed`);
      },
    );
  }

  /** load CSV from sheets */
  public async loadCsvAsync(macroPath: string) {
    const commandName = "Load CSV from Sheets";
    return vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: commandName,
        cancellable: false,
      },
      async _progress => {
        // setup command
        const macroFileName = path.parse(macroPath).name;
        const macroDir = path.dirname(macroPath);
        const csvDir = path.join(macroDir, `${macroFileName}_csv`);
        const scriptPath = `${this.extensionPath}\\bin\\Load-CSV.ps1`;
        this.channel.appendLine("");
        this.channel.appendLine(`${commandName}`);
        this.channel.appendLine(`- File: ${path.basename(macroPath)}`);
        this.channel.appendLine(`- Output: ${path.basename(csvDir)}`);

        // exec command
        const result = this.execPowerShell(scriptPath, [macroPath, csvDir]);

        // output result
        if (result.stdout) this.channel.appendLine(`- Output: ${result.stdout}`);
        if (result.exitCode !== 0) {
          throw `${result.stderr}`;
        }

        this.channel.appendLine(`[SUCCESS] CSV loaded from sheets`);

        // Close all diff editors
        await this.closeAllDiffEditors();

        // Open the first file in the explorer view if not already in this folder
        const files = fs.readdirSync(csvDir).filter(file => {
          const ext = path.extname(file).toLowerCase();
          return [".csv"].includes(ext);
        });

        if (files.length > 0) {
          // Only open if no active editor or the active editor's file doesn't exist in new folder
          const activeEditor = vscode.window.activeTextEditor;
          let shouldOpen = true;

          if (activeEditor) {
            const activeEditorFileName = path.basename(activeEditor.document.uri.fsPath);
            const activeEditorExists = files.includes(activeEditorFileName);
            shouldOpen = !activeEditorExists;
          }

          if (shouldOpen) {
            const firstFile = path.join(csvDir, files[0]);
            const uri = vscode.Uri.file(firstFile);
            await vscode.commands.executeCommand("vscode.open", uri);
          }
        }
      },
    );
  }

  /** save sheets from CSV */
  public async saveCsvAsync(macroPath: string) {
    const commandName = "Save Sheets from CSV";
    return vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: commandName,
        cancellable: false,
      },
      async _progress => {
        // setup command
        const macroFileName = path.parse(macroPath).name;
        const macroDir = path.dirname(macroPath);
        const csvDir = path.join(macroDir, `${macroFileName}_csv`);
        const scriptPath = `${this.extensionPath}\\bin\\Save-CSV.ps1`;
        this.channel.appendLine("");
        this.channel.appendLine(`${commandName}`);
        this.channel.appendLine(`- File: ${path.basename(macroPath)}`);
        this.channel.appendLine(`- Source: ${path.basename(csvDir)}`);

        // Check if CSV directory exists
        if (!fs.existsSync(csvDir)) {
          throw `Folder not found: ${path.basename(csvDir)}. Please export CSV first.`;
        }

        // exec command
        const result = this.execPowerShell(scriptPath, [macroPath, csvDir]);

        // output result
        if (result.stdout) this.channel.appendLine(`- Output: ${result.stdout}`);
        if (result.exitCode !== 0) {
          throw `${result.stderr}`;
        }

        this.channel.appendLine(`[SUCCESS] Sheets saved from CSV`);
      },
    );
  }

  /** create new Excel file */
  public async newExcelAsync() {
    const commandName = "New Excel Book";
    return vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: commandName,
        cancellable: false,
      },
      async _progress => {
        const workspaceFolders = vscode.workspace.workspaceFolders;
        if (!workspaceFolders) {
          throw `No workspace folder found`;
        }

        // Get the first workspace folder
        const workspaceFolder = workspaceFolders[0].uri.fsPath;

        // Ask user for file name
        const fileName = await vscode.window.showInputBox({
          prompt: "Enter new Excel file name",
          placeHolder: "example",
          validateInput: input => {
            if (!input) {
              return "File name cannot be empty";
            }
            if (/[/\\:*?"<>|]/.test(input)) {
              return 'File name cannot contain: / \\ : * ? " < > |';
            }
            return "";
          },
        });

        if (!fileName) {
          throw `File creation cancelled`;
        }

        const filePath = path.join(workspaceFolder, `${fileName}.xlsx`);

        // Check if file already exists
        if (fs.existsSync(filePath)) {
          throw `File already exists: ${fileName}.xlsx`;
        }

        this.channel.appendLine("");
        this.channel.appendLine(`${commandName}`);
        this.channel.appendLine(`- File: ${fileName}.xlsx`);
        this.channel.appendLine(`- Path: ${filePath}`);

        // Use PowerShell to create a new Excel file
        const scriptPath = `${this.extensionPath}\\bin\\New-Excel.ps1`;
        const result = this.execPowerShell(scriptPath, [filePath]);

        // output result
        if (result.stdout) this.channel.appendLine(`- Output: ${result.stdout}`);
        if (result.exitCode !== 0) {
          throw `${result.stderr}`;
        }

        this.channel.appendLine(`[SUCCESS] Created new Excel file`);

        // Open the created file in the file explorer
        const fileUri = vscode.Uri.file(filePath);
        await vscode.commands.executeCommand("vscode.open", fileUri);
      },
    );
  }

  /** create dummy URL shortcut files for cloud-based Excel workbooks */
  public async createDummyUrlShortcutAsync() {
    const commandName = "Create Dummy URL Shortcut for Cloud Files";
    return vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: commandName,
        cancellable: false,
      },
      async _progress => {
        const workspaceFolders = vscode.workspace.workspaceFolders;
        const workspaceFolder = workspaceFolders ? workspaceFolders[0].uri.fsPath : path.join(process.env.USERPROFILE || "", "Desktop");

        this.channel.appendLine("");
        this.channel.appendLine(`${commandName}`);

        // Use PowerShell to create .url files for all open workbooks
        const scriptPath = `${this.extensionPath}\\bin\\Create-DummyUrlShortcuts.ps1`;
        const result = this.execPowerShell(scriptPath, [workspaceFolder]);

        // output result
        if (result.stdout) this.channel.appendLine(`- Output: ${result.stdout}`);
        if (result.exitCode !== 0) {
          throw `${result.stderr}`;
        }

        this.channel.appendLine(`[SUCCESS] Dummy URL shortcuts created for all cloud-based workbooks`);
      },
    );
  }
}
export const excelvba = new ExcelVba();
