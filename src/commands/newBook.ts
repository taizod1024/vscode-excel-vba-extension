import * as vscode from "vscode";
const fs = require("fs");
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";
import { Logger } from "../utils/logger";

/** Create new Excel workbook */
export async function newBookAsync(context: CommandContext) {
  const workspaceFolders = vscode.workspace.workspaceFolders;
  if (!workspaceFolders || workspaceFolders.length === 0) {
    throw "No open workspace.";
  }

  const workspaceFolder = workspaceFolders[0].uri.fsPath;

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

  const fileName = `${inputPrompt}.xlsx`;
  const filePath = path.join(workspaceFolder, fileName);

  // Check if file already exists
  if (fs.existsSync(filePath)) {
    throw "File already exists.";
  }

  const commandName = `New Excel Book`;
  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: commandName,
      cancellable: false,
    },
    async _progress => {
      const logger = new Logger(context.channel);
      const scriptPath = `${context.extensionPath}\\bin\\New-Excel.ps1`;
      logger.logCommandStart(commandName, {
        path: filePath,
      });

      // exec command
      const result = execPowerShell(scriptPath, [filePath]);

      // output result
      if (result.stdout) logger.logRaw(result.stdout);
      if (result.exitCode !== 0) {
        logger.logError(result.stderr);
        throw "Failed to create new workbook.";
      }

      logger.logSuccess("New workbook created");

      // Reveal file in Explorer
      const fileUri = vscode.Uri.file(filePath);
      await vscode.commands.executeCommand("revealInExplorer", fileUri);
    },
  );
}
