import * as vscode from "vscode";
import * as fs from "fs";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";
import { getVbaFiles } from "../utils/fileOperations";
import { showDiffAsync } from "../utils/editorOperations";

const commandName = "Compare VBA with Excel Book";

export async function compareVbaAsync(macroPath: string, context: CommandContext) {
  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: commandName,
      cancellable: false,
    },
    async _progress => {
      const macroExtension = path.parse(macroPath).ext.replace(".", "");
      const vbaComponentExtensions = ["bas", "cls", "frm", "frx"];
      let macroDir = path.dirname(macroPath);
      let referenceFileName = path.parse(macroPath).name;

      // If VBA component file, find the parent Excel workbook to get the correct folder name
      if (vbaComponentExtensions.includes(macroExtension)) {
        const folderName = path.basename(macroDir);
        const match = folderName.match(/^(.+)_bas$/i);
        if (match) {
          const parentDir = path.dirname(macroDir);
          const baseFileName = match[1];
          const extensions = [".xlsm", ".xlsx", ".xlam"];

          for (const ext of extensions) {
            const excelPath = path.join(parentDir, baseFileName + ext);
            if (fs.existsSync(excelPath)) {
              referenceFileName = path.parse(excelPath).name;
              macroDir = parentDir;
              break;
            }
          }
        }
      }

      const currentFolderName = `${referenceFileName}_bas`;
      const currentPath = path.join(macroDir, currentFolderName);
      const tmpPath = path.join(macroDir, `${referenceFileName}_bas~`);

      context.channel.appendLine("");
      context.channel.appendLine(`${commandName}`);
      context.channel.appendLine(`- File: ${path.basename(macroPath)}`);
      context.channel.appendLine(`- Current: ${path.basename(currentPath)}`);
      context.channel.appendLine(`- Loading from Excel...`);

      if (!fs.existsSync(currentPath)) {
        throw `Folder not found: ${path.basename(currentPath)}. Please load VBA first.`;
      }

      // Load to temporary folder
      const scriptPath = `${context.extensionPath}\\bin\\Load-VBA.ps1`;
      const result = execPowerShell(scriptPath, [macroPath, tmpPath]);

      if (result.exitCode !== 0) {
        throw `${result.stderr}`;
      }

      // Compare files
      const hasDifferences = compareDirectories(tmpPath, currentPath, context);

      // Remove temporary folder only if no differences
      if (!hasDifferences && fs.existsSync(tmpPath)) {
        fs.rmSync(tmpPath, { recursive: true, force: true });
      }

      // Show the channel
      context.channel.show();
    },
  );
}

function compareDirectories(dir1: string, dir2: string, context: CommandContext): boolean {
  const files1 = getVbaFiles(dir1);
  const files2 = getVbaFiles(dir2);

  context.channel.appendLine(`Comparison Results:`);
  context.channel.appendLine(`- Files in Excel: ${files1.length}`);
  context.channel.appendLine(`- Files on disk: ${files2.length}`);

  const added = files1.filter(f => !files2.includes(f));
  const removed = files2.filter(f => !files1.includes(f));
  const common = files1.filter(f => files2.includes(f));

  let modifiedCount = 0;
  let firstModifiedFile: { file1Path: string; file2Path: string; name: string } | null = null;

  if (added.length > 0) {
    context.channel.appendLine(`- [+] Added (${added.length}):`);
    added.forEach(f => {
      const relativePath = f.replace(/\\/g, "/");
      context.channel.appendLine(`    ${relativePath}`);
    });
  }

  if (removed.length > 0) {
    context.channel.appendLine(`- [-] Removed (${removed.length}):`);
    removed.forEach(f => {
      const relativePath = f.replace(/\\/g, "/");
      context.channel.appendLine(`    ${relativePath}`);
    });
  }

  if (common.length > 0) {
    context.channel.appendLine(`- [~] Modified:`);
    common.forEach(f => {
      const file1Path = path.join(dir1, f);
      const file2Path = path.join(dir2, f);
      const content1 = fs.readFileSync(file1Path, "utf8");
      const content2 = fs.readFileSync(file2Path, "utf8");
      if (content1 !== content2) {
        const relativePath = f.replace(/\\/g, "/");
        context.channel.appendLine(`    ${relativePath}`);
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
    context.channel.appendLine(`[WARN] Differences found: +${added.length} ~${modifiedCount} -${removed.length}`);
  } else {
    context.channel.appendLine(`[SUCCESS] No differences found`);
  }

  // Display first modified file in diff view
  if (firstModifiedFile) {
    showDiffAsync(firstModifiedFile.file1Path, firstModifiedFile.file2Path, firstModifiedFile.name);
  }

  return hasDifferences;
}
