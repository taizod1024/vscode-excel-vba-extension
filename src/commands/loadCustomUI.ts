import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { Logger } from "../utils/logger";
import { execPowerShell } from "../utils/execPowerShell";
import { closeAllDiffEditors } from "../utils/editorOperations";

const commandName = "Load CustomUI from Excel Book";

export async function loadCustomUIAsync(bookPath: string, context: CommandContext) {
  const ext = path.extname(bookPath).toLowerCase();

  // CustomUI is supported for .xlam (add-ins) and .xlsm (workbooks)
  if (ext !== ".xlam" && ext !== ".xlsm") {
    throw "CustomUI not supported";
  }

  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: commandName,
      cancellable: false,
    },
    async _progress => {
      const logger = new Logger(context.channel);

      // setup command
      const bookFileName = path.basename(bookPath);
      const bookDir = path.dirname(bookPath);
      const tmpPath = path.join(bookDir, `${bookFileName}.xml~`);
      const scriptPath = `${context.extensionPath}\\bin\\Load-CustomUI.ps1`;

      logger.logCommandStart(commandName, {
        file: bookFileName,
      });

      // exec command
      const result = execPowerShell(scriptPath, [bookPath, tmpPath]);

      // output result
      if (result.stdout) logger.logDetail("Output", result.stdout);
      if (result.exitCode !== 0) {
        // Extract first line of error message for user display
        const errorLine = result.stderr.split("\n")[0].trim() || "Failed to load CustomUI.";
        logger.logError(`${errorLine}:\n${result.stderr}`);
        throw errorLine;
      }

      // Organize loaded files
      const newFolderName = `${bookFileName}.xml`;
      const newPath = path.join(bookDir, newFolderName);

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

      logger.logSuccess(`CustomUI extracted (${files.length} file(s))`);
    },
  );
}
