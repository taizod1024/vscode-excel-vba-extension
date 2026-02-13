import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { Logger } from "../utils/logger";
import { execPowerShell } from "../utils/execPowerShell";

const commandName = "Export Sheet as PNG";

export async function exportSheetAsPngAsync(bookPath: string, context: CommandContext) {
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
      const imageDir = path.join(bookDir, `${bookFileName}.png`);
      const scriptPath = `${context.extensionPath}\\bin\\Export-SheetAsImage.ps1`;
      
      logger.logCommandStart(commandName, {
        File: bookFileName,
        Output: `${bookFileName}.png`
      });

      // exec command
      const result = execPowerShell(scriptPath, [bookPath, imageDir]);

      // output result
      if (result.stdout) logger.logDetail("Output", result.stdout);
      if (result.exitCode !== 0) {
        // Extract first line of error message for user display
        const errorLine = result.stderr.split('\n')[0].trim() || "Failed to export sheet as PNG.";
        logger.logError(`${errorLine}:\n${result.stderr}`);
        throw errorLine;
      }

      logger.logSuccess("Sheets exported as images");
      vscode.window.showInformationMessage("Sheets exported as PNG.");
    }
  );
}