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
      vscode.commands.registerCommand(`${this.appId}.exportVba`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          await this.exportVbaAsync(uri.fsPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${this.appName}: ${reason}`);
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.importVba`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          await this.importVbaAsync(uri.fsPath);
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
          await this.compareVbaAsync(uri.fsPath);
        } catch (reason) {
          this.channel.appendLine(`ERROR: ${reason}`);
          vscode.window.showErrorMessage(`${this.appName}: ${reason}`);
        }
      }),
    );
  }

  /** export vba */
  public exportVbaAsync(bookPath: string) {
    // setup command
    const commandName = "Export VBA from book";
    const bookFileName = path.parse(bookPath).name;
    const bookExtension = path.parse(bookPath).ext.replace(".", "");
    const bookDir = path.dirname(bookPath);
    const tmpPath = path.join(bookDir, `${bookFileName}_${bookExtension}~`);
    const scriptPath = `${this.extensionPath}\\bin\\Export-VBA.ps1`;
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

    // Organize exported files
    this.organizeExportedFiles(bookPath, tmpPath);
    vscode.window.showInformationMessage(`${commandName}: done`);
  }

  /** import vba */
  public importVbaAsync(bookPath: string) {
    // setup command
    const commandName = "Import VBA to book";
    const bookFileName = path.parse(bookPath).name;
    const bookExtension = path.parse(bookPath).ext.replace(".", "");
    const bookDir = path.dirname(bookPath);
    const tmpPath = path.join(bookDir, `${bookFileName}_${bookExtension}~`);
    const scriptPath = `${this.extensionPath}\\bin\\Import-VBA.ps1`;
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

    vscode.window.showInformationMessage(`${commandName}: done`);
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

  /** organize exported files */
  private organizeExportedFiles(bookPath: string, tmpPath: string) {
    const bookFileName = path.parse(bookPath).name;
    const bookExtension = path.parse(bookPath).ext.replace(".", "");
    const newFolderName = `${bookFileName}_${bookExtension}`;
    const bookDir = path.dirname(bookPath);
    const newPath = path.join(bookDir, newFolderName);

    // Remove existing folder if it exists
    if (fs.existsSync(newPath)) {
      fs.rmSync(newPath, { recursive: true, force: true });
    }

    // Move tmpPath to new location
    fs.renameSync(tmpPath, newPath);
    this.channel.appendLine(`- organizing exported files: moved to ${newPath}`);
  }

  /** compare VBA with existing folder */
  public compareVbaAsync(bookPath: string) {
    const commandName = "Compare VBA with Book";
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

    if (!fs.existsSync(currentPath)) {
      throw `FOLDER NOT FOUND: ${currentPath}. Please export VBA first.`;
    }

    // Export to temporary folder
    const scriptPath = `${this.extensionPath}\\bin\\Export-VBA.ps1`;
    const result = this.execPowerShell(scriptPath, [bookPath, tmpPath]);

    if (result.exitCode !== 0) {
      throw `${result.stderr}`;
    }

    // Compare files
    this.compareDirectories(tmpPath, currentPath);

    this.channel.appendLine(`- temporary folder saved: ${tmpPath}`);

    vscode.window.showInformationMessage(`${commandName}: done`);
  }

  /** compare two directories and output differences */
  private compareDirectories(dir1: string, dir2: string) {
    const files1 = this.getVbaFiles(dir1);
    const files2 = this.getVbaFiles(dir2);

    this.channel.appendLine(`- files in current export: ${files1.length}`);
    this.channel.appendLine(`- files in stored folder: ${files2.length}`);

    const added = files1.filter(f => !files2.includes(f));
    const removed = files2.filter(f => !files1.includes(f));
    const common = files1.filter(f => files2.includes(f));

    let modifiedCount = 0;

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
        }
      });
    }

    // Summary
    if (added.length > 0 || removed.length > 0 || modifiedCount > 0) {
      this.channel.appendLine(`- result: ${added.length} added, ${removed.length} removed, ${modifiedCount} modified (differences found)`);
    } else {
      this.channel.appendLine(`- result: no differences`);
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
