import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";

const commandName = "Export Sheet as Image";

export async function exportSheetAsImageAsync(macroPath: string, context: CommandContext) {
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
      const imageDir = path.join(macroDir, `${macroFileName}_png`);
      const scriptPath = `${context.extensionPath}\\bin\\Export-SheetAsImage.ps1`;
      context.channel.appendLine("");
      context.channel.appendLine(`${commandName}`);
      context.channel.appendLine(`- File: ${path.basename(macroPath)}`);
      context.channel.appendLine(`- Output: ${path.basename(imageDir)}`);

      // exec command
      const result = execPowerShell(scriptPath, [macroPath, imageDir]);

      // output result
      if (result.stdout) context.channel.appendLine(`- Output: ${result.stdout}`);
      if (result.exitCode !== 0) {
        throw `${result.stderr}`;
      }

      vscode.window.showInformationMessage(`Sheets exported as images`);
    }
  );
}