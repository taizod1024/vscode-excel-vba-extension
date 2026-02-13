import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { Logger } from "../utils/logger";
import { execPowerShell } from "../utils/execPowerShell";

const commandName = "Export Sheet as PNG";

export async function exportSheetAsPngAsync(bookPath: string, context: CommandContext) {
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
      const imageDir = path.join(bookDir, `${fileNameWithoutExt}_${excelExt}`, "png");
      const scriptPath = `${context.extensionPath}\\bin\\Export-SheetAsImage.ps1`;

      logger.logCommandStart(commandName, {
        file: bookFileName,
        output: `${fileNameWithoutExt}_${excelExt}/png`
      });

      // exec command
      const result = execPowerShell(scriptPath, [bookPath, imageDir]);

      // output result
      if (result.stdout) logger.logDetail("Output", result.stdout);
      if (result.exitCode !== 0) {
        // Extract first line of error message for user display
        const errorLine = result.stderr.split("\n")[0].trim() || "Failed to export sheet as PNG.";
        throw errorLine;
      }

      logger.logSuccess("Sheets exported as images");
      vscode.window.showInformationMessage(`[${bookFileName}] Sheets exported as PNG.`);
    }
  );
}
