import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";
import { validateVbNames } from "../utils/vbValidation";
import { closeAllDiffEditors } from "../utils/editorOperations";

const commandName = "Save VBA to Excel Book";

export async function saveVbaAsync(macroPath: string, context: CommandContext) {
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
      const saveSourcePath = path.join(macroDir, `${macroFileName}.bas`);
      const scriptPath = `${context.extensionPath}\\bin\\Save-VBA.ps1`;
      context.channel.appendLine("");
      context.channel.appendLine(`${commandName}`);
      context.channel.appendLine(`- File: ${path.basename(macroPath)}`);
      context.channel.appendLine(`- Source: ${path.basename(saveSourcePath)}`);

      // Check if save source folder exists
      if (!fs.existsSync(saveSourcePath)) {
        throw `Folder not found: ${path.basename(saveSourcePath)}. Please load VBA first.`;
      }

      // Validate VB_Name attribute matches file names
      await validateVbNames(saveSourcePath, context.channel);

      // exec command
      const result = execPowerShell(scriptPath, [macroPath, saveSourcePath]);

      // output result
      if (result.stdout) context.channel.appendLine(`- Output: ${result.stdout}`);
      if (result.exitCode !== 0) {
        throw `${result.stderr}`;
      }

      // Remove temporary folder if it exists
      const tmpPath = path.join(macroDir, `${macroFileName}_bas~`);
      if (fs.existsSync(tmpPath)) {
        fs.rmSync(tmpPath, { recursive: true, force: true });
        context.channel.appendLine(`- Cleaned: Temporary folder removed`);
      }
      context.channel.appendLine(`[SUCCESS] VBA saved`);

      // Show warning for add-in files
      const ext = path.extname(macroPath).toLowerCase();
      if (ext === ".xlam") {
        vscode.window.showWarningMessage(".XLAM CANNOT BE SAVED AUTOMATICALLY. Please save manually from VBE using Ctrl+S.");
      }

      // Close all diff editors
      await closeAllDiffEditors(context.channel);
    },
  );
}
