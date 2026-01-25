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
    this.tmpPath = `${process.env.TMP}\\${this.appId}\\`;
    this.channel.appendLine(`tmpPath: ${this.tmpPath}`);

    // init vscode
    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.exportVba`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          await this.exportVbaAsync(uri.fsPath);
        } catch (reason) {
          this.channel.show();
          excelvba.channel.appendLine("**** " + reason + " ****");
        }
      }),
    );

    context.subscriptions.push(
      vscode.commands.registerCommand(`${this.appId}.importVba`, async (uri: vscode.Uri) => {
        this.extensionPath = context.extensionPath;
        try {
          await this.importVbaAsync(uri.fsPath);
        } catch (reason) {
          this.channel.show();
          excelvba.channel.appendLine("**** " + reason + " ****");
        }
      }),
    );
  }

  /** export vba */
  public exportVbaAsync(bookPath: string) {
    // setup command
    const commandName = "Export VBA from book";
    const scriptPath = `${this.extensionPath}\\bin\\Export-VBA.ps1`;
    this.channel.appendLine(`--------`);
    this.channel.appendLine(`${commandName}: bookPath=${bookPath}`);

    // exec command
    const result = this.execPowerShell(scriptPath, [bookPath, this.tmpPath]);

    // output result
    this.channel.appendLine(`exitCode=${result.exitCode}`);
    if (result.text) this.channel.appendLine(`output=${result.text}`);
    if (result.exitCode === 0) {
      vscode.window.showInformationMessage(`${commandName}: done`);
    } else {
      vscode.window.showErrorMessage(`${commandName}: failure`);
      this.channel.show();
    }
  }

  /** import vba */
  public importVbaAsync(bookPath: string) {
    // setup command
    const commandName = "Import VBA to book";
    const scriptPath = `${this.extensionPath}\\bin\\Import-VBA.ps1`;
    this.channel.appendLine(`--------`);
    this.channel.appendLine(`${commandName}: bookPath=${bookPath}`);

    // exec command
    const result = this.execPowerShell(scriptPath, [bookPath, this.tmpPath]);

    // output result
    this.channel.appendLine(`exitCode=${result.exitCode}`);
    if (result.text) this.channel.appendLine(`output=${result.text}`);
    if (result.exitCode === 0) {
      vscode.window.showInformationMessage(`${commandName}: done`);
    } else {
      vscode.window.showErrorMessage(`${commandName}: failure`);
      this.channel.show();
    }
  }

  /** execute powershell script */
  public execPowerShell(scriptPath: string, args: string[], trim = true): { text: string; exitCode: number } {
    try {
      const result = child_process.spawnSync("powershell.exe", ["-ExecutionPolicy", "RemoteSigned", "-File", scriptPath, ...args], {
        cwd: this.projectPath,
        encoding: "utf8",
      });
      let text = result.stdout || "";
      if (result.error) {
        return { text: result.error.message, exitCode: 1 };
      }
      return { text: trim ? text.trim() : text, exitCode: result.status || 0 };
    } catch (ex: any) {
      return {
        text: trim ? (ex.message || "").trim() : ex.message || "",
        exitCode: 1,
      };
    }
  }
}
export const excelvba = new ExcelVba();
