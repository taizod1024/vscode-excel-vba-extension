import * as vscode from "vscode";
const path = require("path");
import child_process from "child_process";
import { CommandContext } from "../utils/types";
import { readUrlFile } from "../utils/urlFile";
import { getExcelPath } from "../utils/excelPath";

const commandName = "Open Excel Book";

export async function openBookAsync(macroPath: string, context: CommandContext) {
  let fileToOpen = macroPath;
  const ext = path.extname(macroPath).toLowerCase();

  // .url ファイルの場合、中身から参照を取得
  if (ext === ".url") {
    const reference = readUrlFile(macroPath);
    if (reference) {
      fileToOpen = reference;
    }
  }

  context.channel.appendLine("");
  context.channel.appendLine(`${commandName}`);
  context.channel.appendLine(`- File: ${path.basename(macroPath)}`);
  // Excel の実行ファイルパスを取得して起動
  const excelPath = getExcelPath();
  child_process.spawn(excelPath, [fileToOpen], { detached: true });
  context.channel.appendLine(`[SUCCESS] Opened in Excel`);
}
