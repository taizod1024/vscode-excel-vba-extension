import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import child_process from "child_process";

/** excel-vba-extesnion class */
class ExcelVba {
  /** application id for vscode */
  public appId = "excel-vba";

  /** application name */
  public appName = "Excel VBA";

  /** channel on vscode */
  public channel: vscode.OutputChannel;

  /** project path */
  public projectPath: string;

  /** app path */
  public appPath: string;

  /** extension path */
  public extensionPath: string;

  /** tmp path */
  public tmpPath: string;

  /** constructor */
  constructor() {}

  /** resolve VBA path from selected file */
  public resolveVbaPath(selectedPath: string): string {
    const ext = path.extname(selectedPath).toLowerCase();

    // If .xlsm or .xlam is selected, return as is
    if (ext === ".xlsm" || ext === ".xlam") {
      return selectedPath;
    }

    // If .bas, .cls, .frm is selected, find the parent _xlsm or _xlam folder
    if ([".bas", ".cls", ".frm"].includes(ext)) {
      const parentDir = path.dirname(selectedPath);
      let parentName = path.basename(parentDir);

      // Remove trailing ~ from parent folder name
      if (parentName.endsWith("~")) {
        parentName = parentName.slice(0, -1);
      }

      // Check if parent folder is _xlsm or _xlam
      const match = parentName.match(/^(.+)_(?:xlsm|xlam)$/i);
      if (match) {
        const bookName = match[1];
        const extType = parentName.endsWith("_xlsm") ? "xlsm" : "xlam";
        const bookPath = path.join(path.dirname(parentDir), `${bookName}.${extType}`);
        return bookPath;
      }
    }

    return selectedPath;
  }

  /** activate extension */
  public activate(context: vscode.ExtensionContext) {
    // init context
    this.channel = vscode.window.createOutputChannel(this.appName, { log: true });
    if (!process.env.WINDIR) {
      this.channel.appendLine(`${this.appId} failed, no windir`);
      return;
    }
    this.channel.appendLine(`${this.appId} activated`);

    // init vscode
    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.loadVba`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          const bookPath = this.resolveVbaPath(uri.fsPath);
          await this.loadVbaAsync(bookPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${this.appName}: ${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.saveVba`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          const bookPath = this.resolveVbaPath(uri.fsPath);
          await this.saveVbaAsync(bookPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${this.appName}: ${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.compareVba`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          const bookPath = this.resolveVbaPath(uri.fsPath);
          await this.compareVbaAsync(bookPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${this.appName}: ${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.openExcelBook`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          const bookPath = this.resolveVbaPath(uri.fsPath);
          await this.openExcelBookAsync(bookPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${this.appName}: ${reason}`);
        }
      }),
    );
  }

  /** load vba */
  public async loadVbaAsync(bookPath: string) {
    const commandName = "Load VBA from Excel Book";
    return vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: commandName,
        cancellable: false,
      },
      async _progress => {
        // setup command
        const bookFileName = path.parse(bookPath).name;
        const bookExtension = path.parse(bookPath).ext.replace(".", "");
        const bookDir = path.dirname(bookPath);
        const tmpPath = path.join(bookDir, `${bookFileName}_${bookExtension}~`);
        const scriptPath = `${this.extensionPath}\\bin\\Load-VBA.ps1`;
        this.channel.appendLine(`--------`);
        this.channel.appendLine(`${commandName}:`);
        this.channel.appendLine(`- bookPath: ${bookPath}`);

        // exec command
        const result = this.execPowerShell(scriptPath, [bookPath, tmpPath]);

        // output result
        if (result.stdout) this.channel.appendLine(`- ${result.stdout}`);
        if (result.exitCode !== 0) {
          throw `${result.stderr}`;
        }

        // Organize loaded files
        const newFolderName = `${bookFileName}_${bookExtension}`;
        const newPath = path.join(bookDir, newFolderName);

        // Remove existing folder if it exists
        if (fs.existsSync(newPath)) {
          fs.rmSync(newPath, { recursive: true, force: true });
        }

        // Move tmpPath to new location
        fs.renameSync(tmpPath, newPath);
        this.channel.appendLine(`- organizing loaded files: moved to ${newPath}`);

        // Close all diff editors
        await this.closeAllDiffEditors();
      },
    );
  }

  /** save vba */
  public async saveVbaAsync(bookPath: string) {
    const commandName = "Save VBA to Excel Book";
    return vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: commandName,
        cancellable: false,
      },
      async _progress => {
        // setup command
        const bookFileName = path.parse(bookPath).name;
        const bookExtension = path.parse(bookPath).ext.replace(".", "");
        const bookDir = path.dirname(bookPath);
        const saveSourcePath = path.join(bookDir, `${bookFileName}_${bookExtension}`);
        const scriptPath = `${this.extensionPath}\\bin\\Save-VBA.ps1`;
        this.channel.appendLine(`--------`);
        this.channel.appendLine(`${commandName}:`);
        this.channel.appendLine(`- bookPath: ${bookPath}`);
        this.channel.appendLine(`- saving from: ${saveSourcePath}`);

        // Check if save source folder exists
        if (!fs.existsSync(saveSourcePath)) {
          throw `FOLDER NOT FOUND: ${saveSourcePath}. Plase export VBA first.`;
        }

        // exec command
        const result = this.execPowerShell(scriptPath, [bookPath, saveSourcePath]);

        // output result
        if (result.stdout) this.channel.appendLine(`- ${result.stdout}`);
        if (result.exitCode !== 0) {
          throw `${result.stderr}`;
        }

        // Remove temporary folder if it exists
        const tmpPath = path.join(bookDir, `${bookFileName}_${bookExtension}~`);
        if (fs.existsSync(tmpPath)) {
          fs.rmSync(tmpPath, { recursive: true, force: true });
          this.channel.appendLine(`- removed temporary folder: ${tmpPath}`);
        }

        // Close all diff editors
        await this.closeAllDiffEditors();
      },
    );
  }

  /** execute powershell script */
  public execPowerShell(scriptPath: string, args: string[], trim = true): { stdout: string; stderr: string; exitCode: number } {
    try {
      const result = child_process.spawnSync("powershell.exe", ["-ExecutionPolicy", "RemoteSigned", "-File", scriptPath, ...args], {
        cwd: this.projectPath,
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

  /** open excel book */
  public async openExcelBookAsync(bookPath: string) {
    const commandName = "Open Excel Book";
    try {
      this.channel.appendLine(`--------`);
      this.channel.appendLine(`${commandName}:`);
      this.channel.appendLine(`- bookPath: ${bookPath}`);

      child_process.spawn("cmd.exe", ["/c", "start", bookPath], { detached: true });
      this.channel.appendLine(`- opened in Excel`);
    } catch (reason) {
      throw reason;
    }
  }

  /** compare VBA with existing folder */
  public async compareVbaAsync(bookPath: string) {
    const commandName = "Compare VBA with Excel Book";
    return vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: commandName,
        cancellable: false,
      },
      async _progress => {
        const bookFileName = path.parse(bookPath).name;
        const bookExtension = path.parse(bookPath).ext.replace(".", "");
        const currentFolderName = `${bookFileName}_${bookExtension}`;
        const bookDir = path.dirname(bookPath);
        const currentPath = path.join(bookDir, currentFolderName);
        const tmpPath = path.join(bookDir, `${bookFileName}_${bookExtension}~`);

        this.channel.appendLine(`--------`);
        this.channel.appendLine(`${commandName}:`);
        this.channel.appendLine(`- bookPath: ${bookPath}`);
        this.channel.appendLine(`- comparing folder: ${currentPath}`);
        this.channel.appendLine(`- temporary folder: ${tmpPath}`);

        if (!fs.existsSync(currentPath)) {
          throw `FOLDER NOT FOUND: ${currentPath}. Please load VBA first.`;
        }

        // Load to temporary folder
        const scriptPath = `${this.extensionPath}\\bin\\Load-VBA.ps1`;
        const result = this.execPowerShell(scriptPath, [bookPath, tmpPath]);

        if (result.exitCode !== 0) {
          throw `${result.stderr}`;
        }

        // Compare files
        const hasDifferences = this.compareDirectories(tmpPath, currentPath);

        // Remove temporary folder only if no differences
        if (!hasDifferences && fs.existsSync(tmpPath)) {
          fs.rmSync(tmpPath, { recursive: true, force: true });
          this.channel.appendLine(`- removed temporary folder: ${tmpPath}`);
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

    this.channel.appendLine(`- files in current load: ${files1.length}`);
    this.channel.appendLine(`- files in stored folder: ${files2.length}`);

    const added = files1.filter(f => !files2.includes(f));
    const removed = files2.filter(f => !files1.includes(f));
    const common = files1.filter(f => files2.includes(f));

    let modifiedCount = 0;
    let firstModifiedFile: { file1Path: string; file2Path: string; name: string } | null = null;

    if (added.length > 0) {
      this.channel.appendLine(`- added files:`);
      added.forEach(f => {
        const relativePath = f.replace(/\\/g, "/");
        this.channel.appendLine(`  + ${relativePath}`);
      });
    }

    if (removed.length > 0) {
      this.channel.appendLine(`- removed files:`);
      removed.forEach(f => {
        const relativePath = f.replace(/\\/g, "/");
        this.channel.appendLine(`  - ${relativePath}`);
      });
    }

    if (common.length > 0) {
      this.channel.appendLine(`- comparing common files:`);
      common.forEach(f => {
        const file1Path = path.join(dir1, f);
        const file2Path = path.join(dir2, f);
        const content1 = fs.readFileSync(file1Path, "utf8");
        const content2 = fs.readFileSync(file2Path, "utf8");
        if (content1 !== content2) {
          const relativePath = f.replace(/\\/g, "/");
          this.channel.appendLine(`  ~ ${relativePath} (modified)`);
          modifiedCount++;
          // Store the first modified file
          if (!firstModifiedFile) {
            firstModifiedFile = { file1Path, file2Path, name: relativePath };
          }
        }
      });
    }

    // Summary and return whether differences exist
    const hasDifferences = added.length > 0 || removed.length > 0 || modifiedCount > 0;
    if (hasDifferences) {
      this.channel.appendLine(`- result: ${added.length} added, ${removed.length} removed, ${modifiedCount} modified (differences found)`);
    } else {
      this.channel.appendLine(`- result: no differences`);
    }

    // Show diff for the first modified file
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
      for (const tab of group.tabs) {
        if (tab.input instanceof vscode.TabInputTextDiff) {
          await vscode.window.tabGroups.close(tab);
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
}
export const excelvba = new ExcelVba();
