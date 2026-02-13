import * as vscode from "vscode";
const path = require("path");
const fs = require("fs");
import { CommandContext } from "../utils/types";
import { Logger } from "../utils/logger";
import { execPowerShell } from "../utils/execPowerShell";

const commandName = "Open Sheet from PNG";

export async function openSheetFromPngAsync(pngPath: string, context: CommandContext) {
  const logger = new Logger(context.channel);

  logger.logCommandStart(commandName, {
    file: path.basename(pngPath),
    path: pngPath,
  });

  // Extract sheet name from PNG file name (use filename as-is)
  const pngFileName = path.basename(pngPath);
  if (!pngFileName.toLowerCase().endsWith(".png")) {
    throw "File is not a PNG file.";
  }

  const sheetName = pngFileName; // Use full filename including .png extension as sheet name

  // Determine the Excel file path
  // PNG structure: bookPath/aaa_xlsx/png/SheetName.png
  // So we go up 2 directories to get aaa_xlsx, then extract Excel file name
  const pngDir = path.dirname(pngPath); // .../aaa_xlsx/png
  const xlsxDir = path.dirname(pngDir); // .../aaa_xlsx
  const xlsxDirName = path.basename(xlsxDir); // aaa_xlsx

  logger.logDetail("PNG directory", pngDir);
  logger.logDetail("XLSX folder name", xlsxDirName);

  // Extract Excel file name from folder name (aaa_xlsx -> aaa.xlsx)
  const match = xlsxDirName.match(/^(.+?)_(xlsm|xlsx|xlam)$/i);
  if (!match) {
    throw `Invalid folder structure. Expected format: filename_ext (e.g., mybook_xlsx), got: ${xlsxDirName}`;
  }

  const excelFileName = `${match[1]}.${match[2]}`;
  const parentDir = path.dirname(xlsxDir);
  const excelFilePath = path.join(parentDir, excelFileName);

  logger.logDetail("Excel file path", excelFilePath);
  logger.logDetail("Sheet name", sheetName);

  // Check if Excel file exists
  if (!fs.existsSync(excelFilePath)) {
    throw `Excel file not found: ${excelFilePath}`;
  }

  // Execute PowerShell script to open Excel and select the sheet
  const scriptPath = path.join(context.extensionPath, "bin", "Open-SheetFromPng.ps1");
  const result = execPowerShell(scriptPath, [excelFilePath, sheetName]);

  if (result.exitCode !== 0) {
    throw result.stderr || "Failed to open sheet";
  }

  logger.logSuccess(`Opened: ${excelFileName}`);
  logger.logSuccess(`Sheet selected: ${sheetName}`);
}
