import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";
import { closeAllDiffEditors } from "../utils/editorOperations";

const commandName = "Load CustomUI from Excel Book";

export async function loadCustomUIAsync(macroPath: string, context: CommandContext) {
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
      const macroFileName = path.parse(macroPath).name;
      const macroDir = path.dirname(macroPath);
      const tmpPath = path.join(macroDir, `${macroFileName}_xml~`);
      const scriptPath = `${context.extensionPath}\\bin\\Load-CustomUI.ps1`;
      context.channel.appendLine("");
      context.channel.appendLine(`${commandName}`);
      context.channel.appendLine(`- File: ${path.basename(macroPath)}`);

      // exec command
      const result = execPowerShell(scriptPath, [macroPath, tmpPath]);

      // output result
      if (result.stdout) context.channel.appendLine(`- Output: ${result.stdout}`);
      if (result.exitCode !== 0) {
        throw `${result.stderr}`;
      }

      // Organize loaded files
      const newFolderName = `${macroFileName}_xml`;
      const newPath = path.join(macroDir, newFolderName);

      // Remove existing folder if it exists
      if (fs.existsSync(newPath)) {
        fs.rmSync(newPath, { recursive: true, force: true });
      }

      // Move tmpPath to new location
      fs.renameSync(tmpPath, newPath);
      context.channel.appendLine(`- Organized: Files moved`);

      // Verify files exist in new location
      if (fs.existsSync(newPath)) {
        const files = fs.readdirSync(newPath);
      }

      // Close all diff editors
      await closeAllDiffEditors(context.channel);

      // Open the first file in the explorer view if not already in this folder
      const files = fs.readdirSync(newPath).filter(file => {
        const ext = path.extname(file).toLowerCase();
        return [".xml"].includes(ext);
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
          const firstFile = path.join(newPath, files[0]);
          const uri = vscode.Uri.file(firstFile);
          await vscode.commands.executeCommand("vscode.open", uri);
        }
      }

      context.channel.appendLine(`[SUCCESS] Loaded ${files.length} file(s)`);
    },
  );
}
