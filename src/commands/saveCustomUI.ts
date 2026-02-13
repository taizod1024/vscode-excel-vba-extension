import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";
import { closeAllDiffEditors } from "../utils/editorOperations";

const commandName = "Save CustomUI to Excel Book";

export async function saveCustomUIAsync(macroPath: string, context: CommandContext) {
  const ext = path.extname(macroPath).toLowerCase();

  // CustomUI is supported for .xlam (add-ins) and .xlsm (workbooks)
  if (ext !== ".xlam" && ext !== ".xlsm") {
    throw `CustomUI is only supported for .xlam and .xlsm files. Selected file: ${macroPath}`;
  }

  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: commandName,
      cancellable: false,
    },
    async _progress => {
      // setup command
      const macroFileName = path.basename(macroPath);
      const macroDir = path.dirname(macroPath);
      const saveSourcePath = path.join(macroDir, `${macroFileName}.xml`);
      const scriptPath = `${context.extensionPath}\\bin\\Save-CustomUI.ps1`;
      context.channel.appendLine("");
      context.channel.appendLine(`${commandName}`);
      context.channel.appendLine(`- File: ${path.basename(macroPath)}`);
      context.channel.appendLine(`- Source: ${path.basename(saveSourcePath)}`);

      // Check if save source folder exists
      if (!fs.existsSync(saveSourcePath)) {
        throw `Folder not found: ${path.basename(saveSourcePath)}. Please load CustomUI first.`;
      }

      // exec command
      const result = execPowerShell(scriptPath, [macroPath, saveSourcePath]);

      // output result
      if (result.stdout) context.channel.appendLine(`- Output: ${result.stdout}`);
      if (result.exitCode !== 0) {
        throw `${result.stderr}`;
      }

      // Remove temporary folder if it exists
      const tmpPath = path.join(macroDir, `${macroFileName}_xml~`);
      if (fs.existsSync(tmpPath)) {
        fs.rmSync(tmpPath, { recursive: true, force: true });
        context.channel.appendLine(`- Cleaned: Temporary folder removed`);
      }
      context.channel.appendLine(`[SUCCESS] CustomUI saved`);

      // Close all diff editors
      await closeAllDiffEditors(context.channel);
    },
  );
}
