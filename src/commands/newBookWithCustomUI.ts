import * as vscode from "vscode";
const fs = require("fs");
const path = require("path");
import { CommandContext } from "../utils/types";
import child_process from "child_process";

/** Create new Excel workbook with CustomUI template */
export async function newBookWithCustomUIAsync(context: CommandContext) {
  const workspaceFolders = vscode.workspace.workspaceFolders;
  if (!workspaceFolders || workspaceFolders.length === 0) {
    throw "No workspace folder is open.";
  }

  const workspaceFolder = workspaceFolders[0].uri.fsPath;

  // Check if template file exists
  const templatePath = path.join(context.extensionPath, "excel", "bookWithCustomUI", "bookWithCustomUI.xlsm");

  if (!fs.existsSync(templatePath)) {
    throw `Template file not found: ${templatePath}`;
  }

  // Prompt for file name
  const inputPrompt = await vscode.window.showInputBox({
    prompt: "Enter new book name",
    placeHolder: "Example: MyBook (no extension)",
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

  const fileName = `${inputPrompt}.xlsm`;
  const filePath = path.join(workspaceFolder, fileName);

  // Check if file already exists
  if (fs.existsSync(filePath)) {
    throw `File already exists: ${filePath}`;
  }

  const commandName = `Create new Excel workbook with Custom UI`;
  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: commandName,
      cancellable: false,
    },
    async _progress => {
      context.channel.appendLine("");
      context.channel.appendLine(`${commandName}`);
      context.channel.appendLine(`- Path: ${filePath}`);
      context.channel.appendLine(`- Template: ${templatePath}`);

      // Copy template file
      fs.copyFileSync(templatePath, filePath);

      context.channel.appendLine(`[SUCCESS] New workbook created with Custom UI`);

      // Reveal file in Explorer
      const fileUri = vscode.Uri.file(filePath);
      await vscode.commands.executeCommand("revealInExplorer", fileUri);

      // Open the newly created file with Excel
      try {
        child_process.exec(`start "" "${filePath}"`);
        context.channel.appendLine(`[INFO] Opening file with Excel...`);
      } catch (error) {
        context.channel.appendLine(`[WARNING] Could not open file with Excel: ${error}`);
      }
    },
  );
}
