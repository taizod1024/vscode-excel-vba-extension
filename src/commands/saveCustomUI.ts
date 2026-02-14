import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { Logger } from "../utils/logger";
import { execPowerShell } from "../utils/execPowerShell";
import { closeAllDiffEditors } from "../utils/editorOperations";
import { getActualPath, getExcelFileName, getFileNameParts } from "../utils/pathResolution";

const commandName = "Save CustomUI to Excel Book";

export async function saveCustomUIAsync(bookPath: string, context: CommandContext) {
  // Get display file name (handles .url and VBA component files)
  const excelFileName = getExcelFileName(bookPath);

  // Check file extension (handle .url files)
  const actualPath = getActualPath(bookPath);
  const extForValidation = path.extname(actualPath).toLowerCase();

  // CustomUI is supported for .xlam (add-ins) and .xlsm (workbooks)
  if (extForValidation !== ".xlam" && extForValidation !== ".xlsm") {
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
      const { fileNameWithoutExt, excelExt } = getFileNameParts(bookPath);
      const saveSourcePath = path.join(bookDir, `${fileNameWithoutExt}_${excelExt}`, "xml");
      const scriptPath = `${context.extensionPath}\\bin\\Save-CustomUI.ps1`;

      logger.logCommandStart(commandName, {
        file: bookFileName,
        source: `${fileNameWithoutExt}_${excelExt}/xml`,
      });

      // Check if save source folder exists
      if (!fs.existsSync(saveSourcePath)) {
        throw `CustomUI folder not found`;
      }

      // exec command
      const result = execPowerShell(scriptPath, [bookPath, saveSourcePath]);

      // output result
      if (result.stdout) logger.logDetail("Output", result.stdout.trim());
      if (result.exitCode !== 0) {
        // Extract first line of error message for user display
        const errorLine = result.stderr.split("\n")[0].trim() || "Failed to save CustomUI.";
        throw errorLine;
      }

      // Remove temporary folder if it exists
      const tmpPath = path.join(bookDir, `${fileNameWithoutExt}_${excelExt}`, "xml~");
      if (fs.existsSync(tmpPath)) {
        fs.rmSync(tmpPath, { recursive: true, force: true });
        logger.logInfo("Temporary folder removed");
      }

      logger.logSuccess("CustomUI saved to Excel book");

      // Close all diff editors
      await closeAllDiffEditors(context.channel);
    },
  );
}
