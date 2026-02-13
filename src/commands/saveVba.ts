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
      const saveSourcePath = path.join(bookDir, `${bookFileName}.bas`);
      const scriptPath = `${context.extensionPath}\\bin\\Save-VBA.ps1`;

      logger.logCommandStart(commandName, {
        file: bookFileName,
        source: `${bookFileName}.bas`,
      });

      // Check if save source folder exists
      if (!fs.existsSync(saveSourcePath)) {
        const errorMsg = `VBA folder not found: ${bookFileName}.bas`;
        logger.logError(errorMsg);
        throw errorMsg;
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
        logger.logError(`${errorLine}:\n${result.stderr}`);
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
      const ext = path.extname(bookPath).toLowerCase();
      if (ext === ".xlam") {
        vscode.window.showWarningMessage("Save .XLAM in VB Editor (Ctrl+S).");
      }

      // Close all diff editors
      await closeAllDiffEditors(context.channel);
    },
  );
}
