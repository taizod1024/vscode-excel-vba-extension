import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { Logger } from "../utils/logger";
import { execPowerShell } from "../utils/execPowerShell";
import { getExcelFileName, getFileNameParts } from "../utils/pathResolution";

const commandName = "Export Sheet as PNG";

export async function exportSheetAsPngAsync(bookPath: string, context: CommandContext) {
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
      const imageDir = path.join(bookDir, `${fileNameWithoutExt}_${excelExt}`, "png");
      const scriptPath = `${context.extensionPath}\\bin\\Export-SheetAsPng.ps1`;

      logger.logCommandStart(commandName, {
        file: bookFileName,
        output: `${fileNameWithoutExt}_${excelExt}/png`,
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
      vscode.window.showInformationMessage("Sheets exported as PNG.");

      // Reveal the first PNG file in Explorer
      const files = fs.readdirSync(imageDir).filter(file => {
        const ext = path.extname(file).toLowerCase();
        return ext === ".png";
      });

      if (files.length > 0) {
        const firstFile = path.join(imageDir, files[0]);
        const fileUri = vscode.Uri.file(firstFile);
        await vscode.commands.executeCommand("revealInExplorer", fileUri);
      }
    },
  );
}
