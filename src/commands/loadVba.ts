import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";
import { closeAllDiffEditors } from "../utils/editorOperations";

const commandName = "Load VBA from Excel Book";

export async function loadVbaAsync(macroPath: string, context: CommandContext) {
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
      const tmpPath = path.join(macroDir, `${macroFileName}.bas~`);
      const scriptPath = `${context.extensionPath}\\bin\\Load-VBA.ps1`;
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
      const newFolderName = `${macroFileName}.bas`;
      const newPath = path.join(macroDir, newFolderName);

      // Remove existing folder if it exists
      if (fs.existsSync(newPath)) {
        fs.rmSync(newPath, { recursive: true, force: true });
      }

      // Move tmpPath to new location
      fs.renameSync(tmpPath, newPath);

      // Close all diff editors
      await closeAllDiffEditors(context.channel);

      // Open the first file in the explorer view if not already in this folder
      const files = fs.readdirSync(newPath).filter(file => {
        const ext = path.extname(file).toLowerCase();
        return [".bas", ".cls", ".frm", ".xml"].includes(ext);
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

      context.channel.appendLine(`[SUCCESS] Loaded files organized`);
    },
  );
}
