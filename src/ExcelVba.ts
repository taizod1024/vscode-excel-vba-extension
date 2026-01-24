import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import child_process, { ExecFileSyncOptions } from "child_process";

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
          const stats = fs.statSync(uri.fsPath);
          if (stats.isDirectory()) {
            await this.exportVbaAsync(uri.fsPath);
          } else if (stats.isFile()) {
            await this.exportVbaAsync(path.dirname(uri.fsPath));
          }
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
          const stats = fs.statSync(uri.fsPath);
          if (stats.isDirectory()) {
            await this.importVbaAsync(uri.fsPath);
          } else if (stats.isFile()) {
            await this.importVbaAsync(path.dirname(uri.fsPath));
          }
        } catch (reason) {
          this.channel.show();
          excelvba.channel.appendLine("**** " + reason + " ****");
        }
      }),
    );
  }

  /** export vba */
  public async exportVbaAsync(bookPath: string) {
    // show channel
    this.channel.appendLine(`--------`);
    this.channel.appendLine(`exportVbaAsync:`);
    this.channel.appendLine(`bookPath=${bookPath}`);

    // exec command
    let cmd = `powershell -command start-process 'cmd.exe' -ArgumentList '/k "cd /d ${bookPath}"'`;
    this.channel.appendLine(`command=${cmd}`);
    this.execCommand(cmd);
  }

  /** import vba */
  public async importVbaAsync(bookPath: string) {
    // show channel
    this.channel.appendLine(`--------`);
    this.channel.appendLine(`importVbaAsync:`);
    this.channel.appendLine(`bookPath=${bookPath}`);

    // exec command as istrator
    let cmd = `powershell -command start-process 'cmd.exe' -ArgumentList '/c "cd /d ${bookPath} && powershell"'`;
    this.channel.appendLine(`command=${cmd}`);
    this.execCommand(cmd);
  }

  /** execute command */
  public execCommand(cmd: string, trim = true): string {
    let text = null;
    try {
      const options = { cwd: this.projectPath };
      text = child_process.execSync(cmd, options).toString();
      if (trim) text = text.trim();
    } catch (ex) {}
    return text;
  }
}
export const excelvba = new ExcelVba();
