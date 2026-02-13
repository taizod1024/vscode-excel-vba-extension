import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { Logger } from "../utils/logger";
import { execPowerShell } from "../utils/execPowerShell";
import { getExcelFileName, getFileNameParts } from "../utils/pathResolution";

const commandName = "Save CSV to Excel Book";

export async function saveCsvAsync(bookPath: string, context: CommandContext) {
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
      const csvDir = path.join(bookDir, `${fileNameWithoutExt}_${excelExt}`, "csv");
      const scriptPath = `${context.extensionPath}\\bin\\Save-CSV.ps1`;

      logger.logCommandStart(commandName, {
        file: bookFileName,
        source: `${fileNameWithoutExt}_${excelExt}/csv`,
      });

      // Check if CSV directory exists
      if (!fs.existsSync(csvDir)) {
        const errorMsg = `CSV folder not found`;
        throw errorMsg;
      }

      // exec command
      const result = execPowerShell(scriptPath, [bookPath, csvDir]);

      // output result
      if (result.stdout) logger.logDetail("Output", result.stdout);
      if (result.exitCode !== 0) {
        // Extract first line of error message for user display
        const errorLine = result.stderr.split("\n")[0].trim() || "Failed to save CSV.";
        logger.logError(errorLine);
        throw errorLine;
      }

      logger.logSuccess("Sheets saved from CSV");
    },
  );
}
