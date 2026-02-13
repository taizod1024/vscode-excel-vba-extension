import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { Logger } from "../utils/logger";
import { execPowerShell } from "../utils/execPowerShell";
import { closeAllDiffEditors } from "../utils/editorOperations";

const commandName = "Save CustomUI to Excel Book";

export async function saveCustomUIAsync(bookPath: string, context: CommandContext) {
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
      const saveSourcePath = path.join(bookDir, `${bookFileName}.xml`);
      const scriptPath = `${context.extensionPath}\\bin\\Save-CustomUI.ps1`;
      
      logger.logCommandStart(commandName, {
        File: bookFileName,
        Source: `${bookFileName}.xml`
      });

      // Check if save source folder exists
      if (!fs.existsSync(saveSourcePath)) {
        const errorMsg = `CustomUI folder not found`;
        logger.logError(errorMsg + ` (expected: ${bookFileName}.xml)`);
        throw errorMsg;
      }

      // exec command
      const result = execPowerShell(scriptPath, [bookPath, saveSourcePath]);

      // output result
      if (result.stdout) logger.logDetail("Output", result.stdout);
      if (result.exitCode !== 0) {
        const errorMsg = `PowerShell error`;
        logger.logError(`${errorMsg}: ${result.stderr}`);
        throw errorMsg;
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
}
