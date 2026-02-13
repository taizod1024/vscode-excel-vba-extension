import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";
import { Logger } from "../utils/logger";
import { validateVbNames } from "../utils/vbValidation";
import { closeAllDiffEditors } from "../utils/editorOperations";

const commandName = "Save VBA to Excel Book";

export async function saveVbaAsync(bookPath: string, context: CommandContext) {
  // Resolve Excel file name (handle both direct .xlsx and VBA component file selections)
  // Also handle .url files (OneDrive/cloud files)
  let actualPathForExtension = bookPath;
  const urlExt = path.extname(bookPath).toLowerCase();
  if (urlExt === ".url") {
    actualPathForExtension = bookPath.slice(0, -4); // Remove .url
  }

  const fileExtension = path.parse(actualPathForExtension).ext.replace(".", "");
  const vbaComponentExtensions = ["bas", "cls", "frm", "frx"];
  let excelFileName = path.basename(actualPathForExtension);

  if (vbaComponentExtensions.includes(fileExtension)) {
    // VBA component file selected - extract Excel name from parent folder
    const parentFolderName = path.basename(path.dirname(actualPathForExtension));
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
      let actualBookPath = bookPath;
      const ext = path.extname(bookPath).toLowerCase();
      if (ext === ".url") {
        actualBookPath = bookPath.slice(0, -4); // Remove .url
      }
      const bookFileName = path.basename(bookPath);
      const bookDir = path.dirname(bookPath);
      const fileNameWithoutExt = path.parse(actualBookPath).name;
      const excelExt = path.extname(actualBookPath).slice(1);
      const saveSourcePath = path.join(bookDir, `${fileNameWithoutExt}_${excelExt}`, "bas");
      const scriptPath = `${context.extensionPath}\\bin\\Save-VBA.ps1`;

      logger.logCommandStart(commandName, {
        file: bookFileName,
        source: `${fileNameWithoutExt}_${excelExt}/bas`,
      });

      // Check if save source folder exists
      if (!fs.existsSync(saveSourcePath)) {
        throw `VBA folder not found: ${fileNameWithoutExt}_${excelExt}/bas`;
      }

      // Validate VB_Name attribute matches file names
      await validateVbNames(saveSourcePath, context.channel);

      // exec command
      const result = execPowerShell(scriptPath, [bookPath, saveSourcePath]);

      // output result
      if (result.stdout) logger.logDetail("Output", result.stdout);
      if (result.exitCode !== 0) {
        // Extract first line of error message for user display
        const errorLine = result.stderr.split("\n")[0].trim() || "Failed to save VBA.";
        throw errorLine;
      }

      // Remove temporary folder if it exists
      const tmpPath = path.join(bookDir, `${bookFileName}.bas~`);
      if (fs.existsSync(tmpPath)) {
        fs.rmSync(tmpPath, { recursive: true, force: true });
        logger.logInfo("Temporary folder removed");
      }

      logger.logSuccess("VBA saved");

      // Show warning for add-in files
      if (ext === ".xlam") {
        vscode.window.showWarningMessage(`[${bookFileName}] Save .XLAM in VB Editor (Ctrl+S).`);
      }

      // Close all diff editors
      await closeAllDiffEditors(context.channel);
    },
  );
}
