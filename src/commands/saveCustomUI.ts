import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { Logger } from "../utils/logger";
import { execPowerShell } from "../utils/execPowerShell";
import { closeAllDiffEditors } from "../utils/editorOperations";

const commandName = "Save CustomUI to Excel Book";

export async function saveCustomUIAsync(bookPath: string, context: CommandContext) {
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

  const ext = path.extname(bookPath).toLowerCase();

  // CustomUI is supported for .xlam (add-ins) and .xlsm (workbooks)
  if (ext !== ".xlam" && ext !== ".xlsm") {
    throw "CustomUI not supported";
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
      const saveSourcePath = path.join(bookDir, `${fileNameWithoutExt}_${excelExt}`, "xml");
      const scriptPath = `${context.extensionPath}\\bin\\Save-CustomUI.ps1`;

      logger.logCommandStart(commandName, {
        file: bookFileName,
        source: `${fileNameWithoutExt}_${excelExt}/xml`
      });

      // Check if save source folder exists
      if (!fs.existsSync(saveSourcePath)) {
        throw `CustomUI folder not found`;
      }

      // exec command
      const result = execPowerShell(scriptPath, [bookPath, saveSourcePath]);

      // output result
      if (result.stdout) logger.logDetail("Output", result.stdout);
      if (result.exitCode !== 0) {
        // Extract first line of error message for user display
        const errorLine = result.stderr.split("\n")[0].trim() || "Failed to save CustomUI.";
        throw errorLine;
      }

      // Remove temporary folder if it exists
      const tmpPath = path.join(bookDir, `${bookFileName}.xml~`);
      if (fs.existsSync(tmpPath)) {
        fs.rmSync(tmpPath, { recursive: true, force: true });
        logger.logInfo("Temporary folder removed");
      }

      logger.logSuccess("CustomUI saved");

      // Close all diff editors
      await closeAllDiffEditors(context.channel);
    },
  );
}
