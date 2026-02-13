import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { Logger } from "../utils/logger";
import { execPowerShell } from "../utils/execPowerShell";

const commandName = "Save CSV to Excel Book";

export async function saveCsvAsync(bookPath: string, context: CommandContext) {
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
      const csvDir = path.join(bookDir, `${fileNameWithoutExt}_${excelExt}`, "csv");
      const scriptPath = `${context.extensionPath}\\bin\\Save-CSV.ps1`;

      logger.logCommandStart(commandName, {
        file: bookFileName,
        source: `${fileNameWithoutExt}_${excelExt}/csv`
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
