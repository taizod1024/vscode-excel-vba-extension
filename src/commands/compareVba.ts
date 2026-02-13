import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";
import { Logger } from "../utils/logger";
import { getVbaFiles } from "../utils/fileOperations";
import { showDiffAsync } from "../utils/editorOperations";

const commandName = "Compare VBA with Excel Book";

export async function compareVbaAsync(bookPath: string, context: CommandContext) {
  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: commandName,
      cancellable: false,
    },
    async _progress => {
      const logger = new Logger(context.channel);
      const bookExtension = path.parse(bookPath).ext.replace(".", "");
      const vbaComponentExtensions = ["bas", "cls", "frm", "frx"];
      let bookDir = path.dirname(bookPath);
      let referenceFileName = path.parse(bookPath).name;

      // If VBA component file, find the parent Excel workbook to get the correct folder name
      if (vbaComponentExtensions.includes(bookExtension)) {
        const folderName = path.basename(bookDir);
        const match = folderName.match(/^(.+\.(xlsm|xlsx|xlam))\.bas$/i);
        if (match) {
          const parentDir = path.dirname(bookDir);
          const baseFileName = match[1];

          const excelPath = path.join(parentDir, baseFileName);
          if (fs.existsSync(excelPath)) {
            referenceFileName = path.parse(excelPath).name;
            bookDir = parentDir;
          }
        }
      }

      const currentFolderName = `${referenceFileName}.bas`;
      const currentPath = path.join(bookDir, currentFolderName);
      const tmpPath = path.join(bookDir, `${referenceFileName}.bas~`);

      logger.logCommandStart(commandName, {
        File: path.basename(bookPath),
        Current: currentFolderName
      });
      logger.logInfo("Loading from Excel...");

      if (!fs.existsSync(currentPath)) {
        const errorMsg = `VBA folder not found`;
        logger.logError(errorMsg + ` (expected: ${currentFolderName})`);
        throw errorMsg;
      }

      // Load to temporary folder
      const scriptPath = `${context.extensionPath}\\bin\\Load-VBA.ps1`;
      const result = execPowerShell(scriptPath, [bookPath, tmpPath]);

      if (result.exitCode !== 0) {
        const errorMsg = `PowerShell error`;
        logger.logError(`${errorMsg}: ${result.stderr}`);
        throw errorMsg;
      }

      // Compare files
      const hasDifferences = compareDirectories(tmpPath, currentPath, logger);

      // Remove temporary folder only if no differences
      if (!hasDifferences && fs.existsSync(tmpPath)) {
        fs.rmSync(tmpPath, { recursive: true, force: true });
      }

      // Show the channel
      context.channel.show();
    },
  );
}
logger: Logger): boolean {
  const files1 = getVbaFiles(dir1);
  const files2 = getVbaFiles(dir2);

  logger.logDetail("Files in Excel", `${files1.length}`);
  logger.logDetail("Files on disk", `${files2.length}`);

  const added = files1.filter(f => !files2.includes(f));
  const removed = files2.filter(f => !files1.includes(f));
  const common = files1.filter(f => files2.includes(f));

  let modifiedCount = 0;
  let firstModifiedFile: { file1Path: string; file2Path: string; name: string } | null = null;

  if (added.length > 0) {
    logger.logInfo(`Added (${added.length}):`);
    added.forEach(f => {
      const relativePath = f.replace(/\\/g, "/");
      logger.logRaw(`    + ${relativePath}`);
    });
  }

  if (removed.length > 0) {
    logger.logInfo(`Removed (${removed.length}):`);
    removed.forEach(f => {
      const relativePath = f.replace(/\\/g, "/");
      logger.logRaw(`    - ${relativePath}`);
    });
  }

  if (common.length > 0) {
    common.forEach(f => {
      const file1Path = path.join(dir1, f);
      const file2Path = path.join(dir2, f);
      const content1 = fs.readFileSync(file1Path, "utf8");
      const content2 = fs.readFileSync(file2Path, "utf8");
      if (content1 !== content2) {
        if (modifiedCount === 0) {
          logger.logInfo(`Modified:`);
        }
        const relativePath = f.replace(/\\/g, "/");
        logger.logRaw(`    ~ ${relativePath}`);
        modifiedCount++;
        if (!firstModifiedFile) {
          firstModifiedFile = { file1Path, file2Path, name: relativePath };
        }
      }
    });
  }

  // Summary and return whether differences exist
  const hasDifferences = added.length > 0 || removed.length > 0 || modifiedCount > 0;
  if (hasDifferences) {
    logger.logWarn(`Differences found: +${added.length} ~${modifiedCount} -${removed.length}`);
  } else {
    logger.logSuccess("No differences found"
    context.channel.appendLine(`[SUCCESS] No differences found`);
  }

  // Display first modified file in diff view
  if (firstModifiedFile) {
    showDiffAsync(firstModifiedFile.file1Path, firstModifiedFile.file2Path, firstModifiedFile.name);
  }

  return hasDifferences;
}
