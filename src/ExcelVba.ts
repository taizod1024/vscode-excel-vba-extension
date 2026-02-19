import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import child_process from "child_process";

// Import command modules
import { loadVbaAsync } from "./commands/loadVba";
import { saveVbaAsync } from "./commands/saveVba";
import { compareVbaAsync } from "./commands/compareVba";
import { loadCustomUIAsync } from "./commands/loadCustomUI";
import { saveCustomUIAsync } from "./commands/saveCustomUI";
import { loadCsvAsync } from "./commands/loadCsv";
import { saveCsvAsync } from "./commands/saveCsv";
import { compareCsvAsync } from "./commands/compareCsv";
import { openBookAsync } from "./commands/openBook";
import { runSubAsync } from "./commands/runSub";
import { newBookAsync } from "./commands/newBook";
import { newBookWithCustomUIAsMacroAsync } from "./commands/newBookWithCustomUI";
import { newBookWithCustomUIAsAddinAsync } from "./commands/newBookWithCustomUIAsAddin";
import { createUrlShortcutAsync } from "./commands/createUrlShortcut";
import { exportSheetsAsPngAsync } from "./commands/exportSheetAsImage";
import { openSheetFromPngAsync } from "./commands/openSheetFromPng";
import { copyAddinToAppData } from "./utils/fileOperations";

/** Excel VBA extension class */
class ExcelVba {
  /** application id */
  public appId = "excel-vba";

  /** application name */
  public appName = "Excel VBA Extension";

  /** output channel */
  public channel: vscode.OutputChannel;

  /** extension path */
  public extensionPath: string;

  /** constructor */
  constructor() {}

  /** resolve book (Excel file) path from selected file (handles .xlsx, .xlsm, .xlam, .bas, .csv, .xml files) */
  public resolveBookPath(selectedPath: string): string {
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
    // Use the corresponding local Excel file if it exists, otherwise return the .url path itself
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
      // PowerShell will handle it as a cloud-based file
      return resolvedPath;
    }

    // If .xlsm, .xlam or .xlsx is selected, return as is
    if (ext === ".xlsm" || ext === ".xlam" || ext === ".xlsx") {
      return resolvedPath;
    }

    // If .csv is selected, find the parent folder and the corresponding Excel file
    if (ext === ".csv") {
      const parentDir = path.dirname(resolvedPath);
      const parentName = path.basename(parentDir);

      // Check if parent folder is csv (format: aaa_xlsm/csv or aaa.xlsm.url/csv or aaa_xlsx/csv)
      if (parentName === "csv") {
        let prevParentName = path.basename(path.dirname(parentDir));

        // If parent folder ends with .url (cloud-based), remove it to get the actual Excel name
        if (prevParentName.endsWith(".url")) {
          prevParentName = prevParentName.slice(0, -4); // Remove .url
        }

        const match = prevParentName.match(/^(.+?)_(xlsm|xlsx|xlam)$/i);
        if (match) {
          const baseName = match[1];
          const excelExt = match[2];
          const excelFileName = `${baseName}.${excelExt}`;
          const parentParentDir = path.dirname(path.dirname(parentDir));

          // Try to find the exact file first
          const filePath = path.join(parentParentDir, excelFileName);
          if (fs.existsSync(filePath)) {
            return filePath;
          }

          // Also check for .url with the full filename
          const urlPath = path.join(parentParentDir, `${excelFileName}.url`);
          if (fs.existsSync(urlPath)) {
            return urlPath;
          }

          // For cloud-based files, return the Excel file name without .url
          return filePath;
        }

        // Also check for direct Excel file name pattern: aaa.xlsx
        const match2 = prevParentName.match(/^(.+?)\.(xlsm|xlsx|xlam)$/i);
        if (match2) {
          const excelFileName = prevParentName;
          const parentParentDir = path.dirname(path.dirname(parentDir));

          // Try to find the exact file first
          const filePath = path.join(parentParentDir, excelFileName);
          if (fs.existsSync(filePath)) {
            return filePath;
          }

          // Also check for .url with the full filename
          const urlPath = path.join(parentParentDir, `${excelFileName}.url`);
          if (fs.existsSync(urlPath)) {
            return urlPath;
          }

          // For cloud-based files, return the file name as is
          return filePath;
        }
      }
    }

    // If .bas, .cls, .frm is selected, find the parent folder
    if ([".bas", ".cls", ".frm"].includes(ext)) {
      const parentDir = path.dirname(resolvedPath);
      const parentName = path.basename(parentDir);

      // Check if parent folder is bas/csv/png (format: aaa_xlsm/bas or aaa.xlsm.url/bas or aaa_xlsx/bas)
      let prevParentName = path.basename(path.dirname(parentDir));

      // If parent folder ends with .url (cloud-based), remove it to get the actual Excel name
      if (prevParentName.endsWith(".url")) {
        prevParentName = prevParentName.slice(0, -4); // Remove .url
      }

      const match = prevParentName.match(/^(.+?)_(xlsm|xlsx|xlam)$/i);
      if (match) {
        const baseName = match[1];
        const excelExt = match[2];
        const excelFileName = `${baseName}.${excelExt}`;
        const parentParentDir = path.dirname(path.dirname(parentDir));

        // Try to find the exact file first
        const filePath = path.join(parentParentDir, excelFileName);
        if (fs.existsSync(filePath)) {
          return filePath;
        }

        // Also check for .url with the full filename
        const urlPath = path.join(parentParentDir, `${excelFileName}.url`);
        if (fs.existsSync(urlPath)) {
          return urlPath;
        }

        // For cloud-based files, return the Excel file name without .url
        // PowerShell will search for it by name in Excel
        return filePath;
      }

      // Also check for direct Excel file name pattern: aaa.xlsx
      const match2 = prevParentName.match(/^(.+?)\.(xlsm|xlsx|xlam)$/i);
      if (match2) {
        const excelFileName = prevParentName;
        const parentParentDir = path.dirname(path.dirname(parentDir));

        // Try to find the exact file first
        const filePath = path.join(parentParentDir, excelFileName);
        if (fs.existsSync(filePath)) {
          return filePath;
        }

        // Also check for .url with the full filename
        const urlPath = path.join(parentParentDir, `${excelFileName}.url`);
        if (fs.existsSync(urlPath)) {
          return urlPath;
        }

        // For cloud-based files, return the file name as is
        return filePath;
      }
    }

    // If .xml is selected in a xml folder, find the parent Excel file
    if (ext === ".xml") {
      const parentDir = path.dirname(resolvedPath);
      const parentName = path.basename(parentDir);

      // Check if parent folder is xml (format: aaa_xlam/xml or aaa_xlsm/xml or aaa.xlam.url/xml)
      let prevParentName = path.basename(path.dirname(parentDir));

      // If parent folder ends with .url (cloud-based), remove it to get the actual Excel name
      if (prevParentName.endsWith(".url")) {
        prevParentName = prevParentName.slice(0, -4); // Remove .url
      }

      const match = prevParentName.match(/^(.+?)_(xlam|xlsm)$/i);
      if (match) {
        const baseName = match[1];
        const excelExt = match[2];
        const excelFileName = `${baseName}.${excelExt}`;
        const parentParentDir = path.dirname(path.dirname(parentDir));

        // Try to find the exact file first
        const filePath = path.join(parentParentDir, excelFileName);
        if (fs.existsSync(filePath)) {
          return filePath;
        }

        // Also check for .url with the full filename
        const urlPath = path.join(parentParentDir, `${excelFileName}.url`);
        if (fs.existsSync(urlPath)) {
          return urlPath;
        }

        // For cloud-based files, return the Excel file name without .url
        return filePath;
      }

      // Also check for direct Excel file name pattern: aaa.xlam or aaa.xlsm
      const match2 = prevParentName.match(/^(.+?)\.(xlam|xlsm)$/i);
      if (match2) {
        const excelFileName = prevParentName;
        const parentParentDir = path.dirname(path.dirname(parentDir));

        // Try to find the exact file first
        const filePath = path.join(parentParentDir, excelFileName);
        if (fs.existsSync(filePath)) {
          return filePath;
        }

        // Also check for .url with the full filename
        const urlPath = path.join(parentParentDir, `${excelFileName}.url`);
        if (fs.existsSync(urlPath)) {
          return urlPath;
        }

        // For cloud-based files, return the file name as is
        return filePath;
      }
    }

    // If .png is selected, find the parent folder and the corresponding Excel file
    if (ext === ".png") {
      const parentDir = path.dirname(resolvedPath);
      const parentName = path.basename(parentDir);

      // Check if parent folder is png (format: aaa_xlsx/png or aaa.xlsx.url/png)
      let prevParentName = path.basename(path.dirname(parentDir));

      // If parent folder ends with .url (cloud-based), remove it to get the actual Excel name
      if (prevParentName.endsWith(".url")) {
        prevParentName = prevParentName.slice(0, -4); // Remove .url
      }

      const match = prevParentName.match(/^(.+?)_(xlsm|xlsx|xlam)$/i);
      if (match) {
        const baseName = match[1];
        const excelExt = match[2];
        const excelFileName = `${baseName}.${excelExt}`;
        const parentParentDir = path.dirname(path.dirname(parentDir));

        // Try to find the exact file first
        const filePath = path.join(parentParentDir, excelFileName);
        if (fs.existsSync(filePath)) {
          return filePath;
        }

        // Also check for .url with the full filename
        const urlPath = path.join(parentParentDir, `${excelFileName}.url`);
        if (fs.existsSync(urlPath)) {
          return urlPath;
        }

        // For cloud-based files, return the Excel file name without .url
        return filePath;
      }

      // Also check for direct Excel file name pattern: aaa.xlsx
      const match2 = prevParentName.match(/^(.+?)\.(xlsm|xlsx|xlam)$/i);
      if (match2) {
        const excelFileName = prevParentName;
        const parentParentDir = path.dirname(path.dirname(parentDir));

        // Try to find the exact file first
        const filePath = path.join(parentParentDir, excelFileName);
        if (fs.existsSync(filePath)) {
          return filePath;
        }

        // Also check for .url with the full filename
        const urlPath = path.join(parentParentDir, `${excelFileName}.url`);
        if (fs.existsSync(urlPath)) {
          return urlPath;
        }

        // For cloud-based files, return the file name as is
        return filePath;
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
    this.channel.appendLine("");
    this.channel.appendLine(`## ${this.appName} extension activated`);
    this.channel.appendLine(`[DEBUG] Extension Path: ${context.extensionPath}`);

    // Copy Excel addin to AppData
    // Always copy on activation to ensure addin is up-to-date
    const addinPath = path.join(context.extensionPath, "excel", "excel-vba-addin", "excel-vba-addin.xlam");
    this.channel.appendLine(`[DEBUG] Attempting to copy addin from: ${addinPath}`);
    if (copyAddinToAppData(addinPath, this.channel)) {
      this.channel.appendLine(`[DEBUG] Addin copied successfully`);
    } else {
      this.channel.appendLine(`[DEBUG] Addin copy failed`);
    }

    // init vscode
    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.openBook`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          const selectedPath = uri.fsPath;
          this.channel.appendLine(`[DEBUG] Selected path: ${selectedPath}`);
          const bookPath = this.resolveBookPath(selectedPath);
          this.channel.appendLine(`[DEBUG] Resolved path: ${bookPath}`);
          await openBookAsync(bookPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.loadVba`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          const bookPath = this.resolveBookPath(uri.fsPath);
          await loadVbaAsync(bookPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.saveVba`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          const bookPath = this.resolveBookPath(uri.fsPath);
          await saveVbaAsync(bookPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.compareVba`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          const bookPath = this.resolveBookPath(uri.fsPath);
          await compareVbaAsync(bookPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.loadCustomUI`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          const bookPath = this.resolveBookPath(uri.fsPath);
          await loadCustomUIAsync(bookPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.saveCustomUI`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          const bookPath = this.resolveBookPath(uri.fsPath);
          await saveCustomUIAsync(bookPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.runSub`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          const bookPath = this.resolveBookPath(uri.fsPath);
          await saveVbaAsync(bookPath, commandContext);
          await runSubAsync(bookPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.loadCsv`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          const bookPath = this.resolveBookPath(uri.fsPath);
          await loadCsvAsync(bookPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.saveCsv`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          const bookPath = this.resolveBookPath(uri.fsPath);
          await saveCsvAsync(bookPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.compareCsv`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          const bookPath = this.resolveBookPath(uri.fsPath);
          await compareCsvAsync(bookPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.newBook`, async () => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          await newBookAsync(commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.newBookWithCustomUI`, async () => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          await newBookWithCustomUIAsMacroAsync(commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.newBookWithCustomUIAsAddin`, async () => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          await newBookWithCustomUIAsAddinAsync(commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.createUrlShortcut`, async () => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          await createUrlShortcutAsync(commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.exportSheetsAsPng`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          const bookPath = this.resolveBookPath(uri.fsPath);
          await exportSheetsAsPngAsync(bookPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.openSheetFromPng`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        this.channel.show(false);
        try {
          await openSheetFromPngAsync(uri.fsPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );
  }

  /** execute PowerShell */
  public execPowerShell(scriptPath: string, args: string[], trim = true): { stdout: string; stderr: string; exitCode: number } {
    try {
      if (!fs.existsSync(scriptPath)) {
        throw `PowerShell script not found: ${scriptPath}`;
      }

      const result = child_process.spawnSync("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", scriptPath, ...args], {
        encoding: "utf8",
        maxBuffer: 10 * 1024 * 1024,
      });

      let stdout = result.stdout ? result.stdout.toString() : "";
      let stderr = result.stderr ? result.stderr.toString() : "";

      if (trim) {
        stdout = stdout.trim();
        stderr = stderr.trim();
      }

      return {
        stdout,
        stderr,
        exitCode: result.status ?? 1,
      };
    } catch (error) {
      return {
        stdout: "",
        stderr: `${error}`,
        exitCode: 1,
      };
    }
  }

  /** get VBA files from path */
  public getVbaFiles(path: string): Map<string, string> {
    const files = new Map<string, string>();

    if (!fs.existsSync(path)) {
      return files;
    }

    const stat = fs.statSync(path);
    if (stat.isFile()) {
      return files;
    }

    const dirList = fs.readdirSync(path);
    for (const file of dirList) {
      const ext = file.substring(file.lastIndexOf(".")).toLowerCase();
      if ([".bas", ".cls", ".frm"].includes(ext) && !file.endsWith("~")) {
        const filePath = `${path}\\${file}`;
        const content = fs.readFileSync(filePath, "utf-8");
        files.set(file, content);
      }
    }

    return files;
  }

  /** compare directory contents */
  public compareDirectories(workbookFiles: Map<string, string>, diskFiles: Map<string, string>): Array<{ workbook: string; disk: string; name: string }> {
    const differences: Array<{ workbook: string; disk: string; name: string }> = [];

    // Check all files in workbook
    for (const [fileName, content] of workbookFiles) {
      if (!diskFiles.has(fileName)) {
        continue; // File only in workbook, but user can compare different files
      }

      const diskContent = diskFiles.get(fileName)!;
      if (content !== diskContent) {
        differences.push({
          workbook: content,
          disk: diskContent,
          name: fileName,
        });
      }
    }

    return differences;
  }

  /** close all diff editors */
  public async closeAllDiffEditors() {
    const tabs = vscode.window.tabGroups.all.flatMap(group => group.tabs);

    for (const tab of tabs) {
      const input = tab.input;
      if (input instanceof vscode.TabInputTextDiff) {
        await vscode.window.tabGroups.close(tab);
      }
    }
  }

  /** show diff for two strings */
  public async showDiffAsync(originalString: string, modifiedString: string, label: string) {
    // Create temporary files with content
    const tmpDir = path.join(process.env.TEMP || "/tmp", "excel-vba-extension-tmp");

    // Create tmp directory
    if (!fs.existsSync(tmpDir)) {
      fs.mkdirSync(tmpDir, { recursive: true });
    }

    const originalFile = path.join(tmpDir, `original_${label}`);
    const modifiedFile = path.join(tmpDir, `modified_${label}`);

    // Write content to files
    fs.writeFileSync(originalFile, originalString, "utf-8");
    fs.writeFileSync(modifiedFile, modifiedString, "utf-8");
    // Open diff
    const originalUri = vscode.Uri.file(originalFile);
    const modifiedUri = vscode.Uri.file(modifiedFile);

    await vscode.commands.executeCommand("vscode.diff", originalUri, modifiedUri, `${label} (Workbook <-> Disk)`);
  }

  /** deactivate extension */
  public deactivate() {
    // Addin removal is handled by preuninstall script
  }
}

export const excelvba = new ExcelVba();
