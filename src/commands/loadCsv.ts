import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { Logger } from "../utils/logger";
import { execPowerShell } from "../utils/execPowerShell";
import { closeAllDiffEditors } from "../utils/editorOperations";

const commandName = "Load CSV from Excel Book";

export async function loadCsvAsync(bookPath: string, context: CommandContext) {
  // Resolve Excel file name (handle both direct .xlsx and VBA component file selections)
  const fileExtension = path.parse(bookPath).ext.replace(".", "");
  const vbaComponentExtensions = ["bas", "cls", "frm", "frx"];
  let excelFileName = path.basename(bookPath);

  if (vbaComponentExtensions.includes(fileExtension)) {
    // VBA component file selected - extract Excel name from parent folder
    const parentFolderName = path.basename(path.dirname(bookPath));
    const match = parentFolderName.match(/^(.+\.(xlsm|xlsx|xlam))\.bas$/i);
    if (match) {
      excelFileName = match[1];
    }
  }

  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: `[${excelFileName}] ${commandName}`,
      cancellable: false,
    },
    async _progress => {
      const logger = new Logger(context.channel);

      // setup command
      const bookFileName = path.basename(bookPath);
      const bookDir = path.dirname(bookPath);
      const fileNameWithoutExt = path.parse(bookPath).name;
      const excelExt = path.extname(bookPath).slice(1);
      const csvDir = path.join(bookDir, `${fileNameWithoutExt}_${excelExt}`, "csv");
      const scriptPath = `${context.extensionPath}\\bin\\Load-CSV.ps1`;

      logger.logCommandStart(commandName, {
        file: bookFileName,
        output: `${fileNameWithoutExt}_${excelExt}/csv`
      });

      // exec command
      const result = execPowerShell(scriptPath, [bookPath, csvDir]);

      // output result
      if (result.stdout) logger.logDetail("Output", result.stdout);
      if (result.exitCode !== 0) {
        // Extract first line of error message for user display
        const errorLine = result.stderr.split("\n")[0].trim() || "Failed to load CSV.";
        throw errorLine;
      }

      logger.logSuccess(`CSV extracted (${path.basename(path.dirname(csvDir))}/csv folder)`);

      // Create parent folder if it doesn't exist
      const parentDir = path.dirname(csvDir);
      if (!fs.existsSync(parentDir)) {
        fs.mkdirSync(parentDir, { recursive: true });
      }

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
