import * as vscode from "vscode";
const fs = require("fs");
const path = require("path");
import { CommandContext } from "../utils/types";
import child_process from "child_process";
import { Logger } from "../utils/logger";

/** Create new Excel book with CustomUI as Addin template */
export async function newBookWithCustomUIAsAddinAsync(context: CommandContext) {
  const workspaceFolders = vscode.workspace.workspaceFolders;
  if (!workspaceFolders || workspaceFolders.length === 0) {
    throw "No open workspace.";
  }

  const workspaceFolder = workspaceFolders[0].uri.fsPath;

  // Check if template file exists
  const templatePath = path.join(context.extensionPath, "excel", "bookWithCustomUIAsAddin", "bookWithCustomUIAsAddin.xlam");

  if (!fs.existsSync(templatePath)) {
    throw "Template file not found.";
  }

  // Prompt for file name
  const inputPrompt = await vscode.window.showInputBox({
    prompt: "Enter new addin name",
    placeHolder: "Example: MyAddin (no extension)",
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

  const fileName = `${inputPrompt}.xlam`;
  const filePath = path.join(workspaceFolder, fileName);

  // Check if file already exists
  if (fs.existsSync(filePath)) {
    throw "File already exists.";
  }

  const commandName = `New Excel Book with CustomUI as Addin`;
  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: commandName,
      cancellable: false,
    },
    async _progress => {
      const logger = new Logger(context.channel);
      logger.logCommandStart(commandName, {
        path: filePath,
        template: path.basename(templatePath),
      });

      // Copy template file
      fs.copyFileSync(templatePath, filePath);

      logger.logSuccess(`New addin created (${fileName})`);

      // Reveal file in Explorer and open in VS Code
      const fileUri = vscode.Uri.file(filePath);
      await vscode.commands.executeCommand("revealInExplorer", fileUri);
      await vscode.commands.executeCommand("vscode.open", fileUri);

      // Open the newly created file with Excel
      try {
        child_process.exec(`start "" "${filePath}"`);
        logger.logInfo("Opening file with Excel");
      } catch (error) {
        logger.logWarn("Could not open file with Excel");
      }
    },
  );
}
