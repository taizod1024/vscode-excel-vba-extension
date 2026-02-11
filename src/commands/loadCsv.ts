import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";
import { closeAllDiffEditors } from "../utils/editorOperations";

const commandName = "Load CSV from Sheets";

export async function loadCsvAsync(macroPath: string, context: CommandContext) {
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
      const scriptPath = `${context.extensionPath}\\bin\\Load-CSV.ps1`;
      context.channel.appendLine("");
      context.channel.appendLine(`${commandName}`);
      context.channel.appendLine(`- File: ${path.basename(macroPath)}`);
      context.channel.appendLine(`- Output: ${path.basename(csvDir)}`);

      // exec command
      const result = execPowerShell(scriptPath, [macroPath, csvDir]);

      // output result
      if (result.stdout) context.channel.appendLine(`- Output: ${result.stdout}`);
      if (result.exitCode !== 0) {
        throw `${result.stderr}`;
      }

      context.channel.appendLine(`[SUCCESS] CSV loaded from sheets`);

      // Close all diff editors
      await closeAllDiffEditors(context.channel);

      // Open the first file in the explorer view if not already in this folder
      const files = fs.readdirSync(csvDir).filter(file => {
        const ext = path.extname(file).toLowerCase();
        return [".csv"].includes(ext);
      });

      if (files.length > 0) {
        // Only open if no active editor or the active editor's file doesn't exist in new folder
        const activeEditor = vscode.window.activeTextEditor;
        let shouldOpen = true;

        if (activeEditor) {
          const activeEditorFileName = path.basename(activeEditor.document.uri.fsPath);
          const activeEditorExists = files.includes(activeEditorFileName);
          shouldOpen = !activeEditorExists;
        }

        if (shouldOpen) {
          const firstFile = path.join(csvDir, files[0]);
          const uri = vscode.Uri.file(firstFile);
          await vscode.commands.executeCommand("vscode.open", uri);
        }
      }
    },
  );
}
