import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";

const commandName = "Save Sheets from CSV";

export async function saveCsvAsync(macroPath: string, context: CommandContext) {
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
      const scriptPath = `${context.extensionPath}\\bin\\Save-CSV.ps1`;
      context.channel.appendLine("");
      context.channel.appendLine(`${commandName}`);
      context.channel.appendLine(`- File: ${path.basename(macroPath)}`);
      context.channel.appendLine(`- Source: ${path.basename(csvDir)}`);

      // Check if CSV directory exists
      if (!fs.existsSync(csvDir)) {
        throw `Folder not found: ${path.basename(csvDir)}. Please export CSV first.`;
      }

      // exec command
      const result = execPowerShell(scriptPath, [macroPath, csvDir]);

      // output result
      if (result.stdout) context.channel.appendLine(`- Output: ${result.stdout}`);
      if (result.exitCode !== 0) {
        throw `${result.stderr}`;
      }

      context.channel.appendLine(`[SUCCESS] Sheets saved from CSV`);
    },
  );
}
