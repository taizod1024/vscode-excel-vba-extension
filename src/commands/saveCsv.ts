import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { Logger } from "../utils/logger";
import { execPowerShell } from "../utils/execPowerShell";

const commandName = "Save Sheets from CSV";

export async function saveCsvAsync(bookPath: string, context: CommandContext) {
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
      const csvDir = path.join(bookDir, `${bookFileName}.csv`);
      const scriptPath = `${context.extensionPath}\\bin\\Save-CSV.ps1`;
      
      logger.logCommandStart(commandName, {
        File: bookFileName,
        Source: `${bookFileName}.csv`
      });

      // Check if CSV directory exists
      if (!fs.existsSync(csvDir)) {
        const errorMsg = `CSV folder not found`;
        logger.logError(errorMsg + ` (expected: ${bookFileName}.csv)`);
        throw errorMsg;
      }

      // exec command
      const result = execPowerShell(scriptPath, [bookPath, csvDir]);

      // output result
      if (result.stdout) logger.logDetail("Output", result.stdout);
      if (result.exitCode !== 0) {
        const errorMsg = `PowerShell error`;
        logger.logError(`${errorMsg}: ${result.stderr}`);
        throw errorMsg;
      }

      logger.logSuccess("Sheets saved from CSV");
    },
  );
}
