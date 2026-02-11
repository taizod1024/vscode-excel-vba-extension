import * as vscode from "vscode";
const fs = require("fs");
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";

/** Create new Excel workbook */
export async function newBookAsync(context: CommandContext) {
  const workspaceFolders = vscode.workspace.workspaceFolders;
  if (!workspaceFolders || workspaceFolders.length === 0) {
    throw "No workspace folder is open.";
  }

  const workspaceFolder = workspaceFolders[0].uri.fsPath;

  // Prompt for file name
  const defaultFileName = "NewBook.xlsx";
  const inputPrompt = await vscode.window.showInputBox({
    prompt: "Enter new workbook name",
    value: defaultFileName,
    validateInput: (value: string) => {
      if (value.length === 0) {
        return "File name cannot be empty";
      }
      if (/[/\\:*?"<>|]/.test(value)) {
        return "File name contains invalid characters";
      }
      return "";
    },
  });

  if (inputPrompt === undefined) {
    return; // User cancelled
  }

  const fileName = inputPrompt;
  const filePath = path.join(workspaceFolder, fileName);

  // Check if file already exists
  if (fs.existsSync(filePath)) {
    throw `File already exists: ${filePath}`;
  }

  const commandName = `Create new Excel workbook`;
  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: commandName,
      cancellable: false,
    },
    async _progress => {
      const scriptPath = `${context.extensionPath}\\bin\\New-Excel.ps1`;
      context.channel.appendLine("");
      context.channel.appendLine(`${commandName}`);
      context.channel.appendLine(`- Path: ${filePath}`);

      // exec command
      const result = execPowerShell(scriptPath, [filePath]);

      // output result
      if (result.stdout) context.channel.appendLine(`${result.stdout}`);
      if (result.exitCode !== 0) {
        throw `${result.stderr}`;
      }

      context.channel.appendLine(`[SUCCESS] New workbook created`);

      // Open the newly created file
      const doc = await vscode.workspace.openTextDocument(filePath);
      await vscode.window.showTextDocument(doc);
    },
  );
}
