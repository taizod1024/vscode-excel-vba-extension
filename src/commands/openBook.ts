import * as vscode from "vscode";
const path = require("path");
import child_process from "child_process";
import { CommandContext } from "../utils/types";
import { readUrlFile } from "../utils/urlFile";
import { getExcelPath } from "../utils/excelPath";
import { Logger } from "../utils/logger";

const commandName = "Open Excel Book";

export async function openBookAsync(bookPath: string, context: CommandContext) {
  const logger = new Logger(context.channel);
  let fileToOpen = bookPath;
  const ext = path.extname(bookPath).toLowerCase();

  // .url ファイルの場合、中身から参照を取得
  if (ext === ".url") {
    const reference = readUrlFile(bookPath);
    if (reference) {
      fileToOpen = reference;
    }
  }

  logger.logCommandStart(commandName, {
    File: path.basename(bookPath),
  });
  // Excel の実行ファイルパスを取得して起動
  const excelPath = getExcelPath();
  child_process.spawn(excelPath, [fileToOpen], { detached: true });
  logger.logSuccess("Opened in Excel");
}
