import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";
import { Logger } from "../utils/logger";
import { showDiffAsync } from "../utils/editorOperations";
import { getExcelFileName, getFileNameParts, getActualPath } from "../utils/pathResolution";

const commandName = "Compare CSV with Excel Book";

export async function compareCsvAsync(bookPath: string, context: CommandContext) {
  // Get display file name (handles .url and CSV component files)
  const excelFileName = getExcelFileName(bookPath);

  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: commandName,
      cancellable: false,
    },
    async _progress => {
      const logger = new Logger(context.channel);
      // Get actual path without .url extension
      let actualBookPath = getActualPath(bookPath);
      const bookExtension = path.parse(actualBookPath).ext.replace(".", "");
      const csvComponentExtensions = ["csv"];
      let bookDir = path.dirname(actualBookPath);
      let referenceFileName = path.basename(actualBookPath);

      // If CSV file, find the parent Excel workbook to get the correct folder name
      if (csvComponentExtensions.includes(bookExtension)) {
        const folderName = path.basename(bookDir);
        const match = folderName.match(/^(.+?)_(xlsm|xlsx|xlam)$/i);
        if (match) {
          const parentDir = path.dirname(bookDir);
          const baseName = match[1];
          const excelExtension = match[2];
          const baseFileName = `${baseName}.${excelExtension}`;

          const excelPath = path.join(parentDir, baseFileName);
          if (fs.existsSync(excelPath)) {
            referenceFileName = path.basename(excelPath);
            bookDir = parentDir;
          }
        }
      }

      const { fileNameWithoutExt, excelExt } = getFileNameParts(referenceFileName);
      const currentFolderName = `${fileNameWithoutExt}_${excelExt}`;
      const currentPath = path.join(bookDir, currentFolderName, "csv");
      const tmpPath = path.join(bookDir, currentFolderName, "csv~");

      logger.logCommandStart(commandName, {
        file: path.basename(bookPath),
        current: currentFolderName,
      });
      logger.logInfo("Loading from Excel...");

      if (!fs.existsSync(currentPath)) {
        throw `CSV folder not found`;
      }

      // Load to temporary folder
      const scriptPath = `${context.extensionPath}\\bin\\Load-CSV.ps1`;
      const result = execPowerShell(scriptPath, [bookPath, tmpPath]);

      if (result.exitCode !== 0) {
        // Extract first line of error message for user display
        const errorLine = result.stderr.split("\n")[0].trim() || "Failed to compare CSV.";
        throw errorLine;
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

function getCsvFiles(dir: string): string[] {
  if (!fs.existsSync(dir)) {
    return [];
  }

  const files: string[] = [];

  const items = fs.readdirSync(dir);
  items.forEach(item => {
    const itemPath = path.join(dir, item);
    const stats = fs.statSync(itemPath);

    if (stats.isDirectory()) {
      const subFiles = getCsvFiles(itemPath);
      subFiles.forEach(f => {
        files.push(path.join(item, f));
      });
    } else if (stats.isFile()) {
      const ext = path.extname(item).toLowerCase();
      if (ext === ".csv") {
        files.push(item);
      }
    }
  });

  return files.sort();
}

function compareDirectories(dir1: string, dir2: string, logger: Logger): boolean {
  const files1 = getCsvFiles(dir1);
  const files2 = getCsvFiles(dir2);

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
    logger.logSuccess("No differences found (disk and workbook match)");
  }

  // Display first modified file in diff view
  if (firstModifiedFile) {
    showDiffAsync(firstModifiedFile.file1Path, firstModifiedFile.file2Path, firstModifiedFile.name);
  }

  return hasDifferences;
}
