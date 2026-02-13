import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";
import { Logger } from "../utils/logger";
import { closeAllDiffEditors } from "../utils/editorOperations";
import { getExcelFileName, getFileNameParts } from "../utils/pathResolution";

const commandName = "Load VBA from Excel Book";

export async function loadVbaAsync(bookPath: string, context: CommandContext) {
  // Get display file name (handles .url and VBA component files)
  const excelFileName = getExcelFileName(bookPath);

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
      const { fileNameWithoutExt, excelExt } = getFileNameParts(bookPath);
      const basDir = path.join(bookDir, `${fileNameWithoutExt}_${excelExt}`, "bas");
      const tmpPath = path.join(bookDir, `${fileNameWithoutExt}_${excelExt}`, "bas~");
      const scriptPath = `${context.extensionPath}\\bin\\Load-VBA.ps1`;

      logger.logCommandStart(commandName, {
        file: bookFileName,
      });

      // exec command
      const result = execPowerShell(scriptPath, [bookPath, tmpPath]);

      // output result
      if (result.stdout) logger.logDetail("Output", result.stdout);
      if (result.exitCode !== 0) {
        // Extract first line of error message for user display
        const errorLine = result.stderr.split("\n")[0].trim() || "No book open.";
        throw errorLine;
      }

      // Organize loaded files
      const parentFolderName = `${fileNameWithoutExt}_${excelExt}`;
      const parentPath = path.join(bookDir, parentFolderName);
      const newFolderName = `${fileNameWithoutExt}_${excelExt}`;
      const newPath = path.join(bookDir, newFolderName, "bas");

      // Remove existing folder if it exists
      if (fs.existsSync(newPath)) {
        fs.rmSync(newPath, { recursive: true, force: true });
      }

      // Create parent folder if needed
      if (!fs.existsSync(parentPath)) {
        fs.mkdirSync(parentPath, { recursive: true });
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

      logger.logSuccess(`VBA extracted (${files.length} file(s))`);
    },
  );
}
