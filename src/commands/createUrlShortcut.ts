import * as vscode from "vscode";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";

const commandName = "Create URL Shortcut";

/** Create .url shortcut for cloud-based Excel files */
export async function createUrlShortcutAsync(context: CommandContext) {
  const workspaceFolders = vscode.workspace.workspaceFolders;
  if (!workspaceFolders || workspaceFolders.length === 0) {
    throw "No workspace folder is open.";
  }

  const workspaceFolder = workspaceFolders[0].uri.fsPath;
  const scriptPath = `${context.extensionPath}\\bin\\Create-UrlShortcuts.ps1`;

  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: commandName,
      cancellable: false,
    },
    async _progress => {
      context.channel.appendLine("");
      context.channel.appendLine(`${commandName}`);

      // Execute PowerShell script
      const result = execPowerShell(scriptPath, [workspaceFolder], false);

      // Output result
      if (result.stdout) context.channel.appendLine(result.stdout);
      if (result.exitCode !== 0) {
        throw result.stderr || "Failed to create URL shortcut";
      }
    },
  );
}
