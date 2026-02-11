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
import { openBookAsync } from "./commands/openBook";
import { runSubAsync } from "./commands/runSub";
import { newBookAsync } from "./commands/newBook";
import { newBookWithCustomUIAsync } from "./commands/newBookWithCustomUI";
import { createUrlShortcutAsync } from "./commands/createUrlShortcut";

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

        // Try to find .xlsm first, then .xlam, then .xlsx
        const xlsmPath = path.join(parentParentDir, `${macroName}.xlsm`);
        if (fs.existsSync(xlsmPath)) {
          return xlsmPath;
        }

        const xlamPath = path.join(parentParentDir, `${macroName}.xlam`);
        if (fs.existsSync(xlamPath)) {
          return xlamPath;
        }

        const xlsxPath = path.join(parentParentDir, `${macroName}.xlsx`);
        if (fs.existsSync(xlsxPath)) {
          return xlsxPath;
        }

        const urlPath = path.join(parentParentDir, `${macroName}.url`);
        if (fs.existsSync(urlPath)) {
          return urlPath;
        }
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
      vscode.commands.registerCommand(`${this.appId}.openBook`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        try {
          const selectedPath = uri.fsPath;
          this.channel.appendLine(`[DEBUG] Selected path: ${selectedPath}`);
          const macroPath = this.resolveVbaPath(selectedPath);
          this.channel.appendLine(`[DEBUG] Resolved path: ${macroPath}`);
          await openBookAsync(macroPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.loadVba`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await loadVbaAsync(macroPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.saveVba`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await saveVbaAsync(macroPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.compareVba`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await compareVbaAsync(macroPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.loadCustomUI`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await loadCustomUIAsync(macroPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.saveCustomUI`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await saveCustomUIAsync(macroPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.runSub`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await saveVbaAsync(macroPath, commandContext);
          await runSubAsync(macroPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.loadCsv`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await loadCsvAsync(macroPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.saveCsv`, async (uri: vscode.Uri) => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        try {
          const macroPath = this.resolveVbaPath(uri.fsPath);
          await saveCsvAsync(macroPath, commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.newBook`, async () => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
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
        try {
          await newBookWithCustomUIAsync(commandContext);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.createUrlShortcut`, async () => {
        const commandContext = { channel: this.channel, extensionPath: context.extensionPath };
        try {
          await createUrlShortcutAsync(commandContext);
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
  public compareDirectories(
    workbookFiles: Map<string, string>,
    diskFiles: Map<string, string>,
  ): Array<{ workbook: string; disk: string; name: string }> {
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
}

export const excelvba = new ExcelVba();
