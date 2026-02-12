import * as vscode from "vscode";
const fs = require("fs");
const path = require("path");
import * as iconv from "iconv-lite";
import { CommandContext } from "../utils/types";

/** Create .url shortcut for cloud-based Excel files */
export async function createUrlShortcutAsync(context: CommandContext) {
  const workspaceFolders = vscode.workspace.workspaceFolders;
  if (!workspaceFolders || workspaceFolders.length === 0) {
    throw "No workspace folder is open.";
  }

  const workspaceFolder = workspaceFolders[0].uri.fsPath;

  // Prompt for Excel cloud URL
  const urlPrompt = await vscode.window.showInputBox({
    prompt: "Enter Excel file URL (OneDrive, SharePoint, etc.)",
    validateInput: (value: string) => {
      if (value.length === 0) {
        return "URL cannot be empty";
      }
      if (!value.startsWith("http")) {
        return "URL must start with http or https";
      }
      return "";
    },
  });

  if (urlPrompt === undefined) {
    return; // User cancelled
  }

  const fileUrl = urlPrompt;

  // Prompt for file name
  const namePrompt = await vscode.window.showInputBox({
    prompt: "Enter shortcut file name",
    placeHolder: "Example: CloudWorkbook (no extension)",
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

  if (namePrompt === undefined) {
    return; // User cancelled
  }

  const fileName = `${namePrompt}.url`;
  const filePath = path.join(workspaceFolder, fileName);

  // Check if file already exists
  if (fs.existsSync(filePath)) {
    throw `File already exists: ${filePath}`;
  }

  const commandName = `Create URL shortcut`;
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
      context.channel.appendLine(`- URL: ${fileUrl}`);

      // Create .url shortcut file content
      const urlContent = `[InternetShortcut]\r\nURL=${fileUrl}\r\n`;

      // Write file in SJIS encoding
      const buffer = iconv.encode(urlContent, "shiftjis");
      fs.writeFileSync(filePath, buffer);

      context.channel.appendLine(`[SUCCESS] URL shortcut created`);

      // Open the newly created file
      const doc = await vscode.workspace.openTextDocument(filePath);
      await vscode.window.showTextDocument(doc);
    },
  );
}
