import * as vscode from "vscode";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";
import { Logger } from "../utils/logger";

const commandName = "Create URL Shortcut for Cloud Files";

/** Create .url shortcut for cloud-based Excel files */
export async function createUrlShortcutAsync(context: CommandContext) {
  const workspaceFolders = vscode.workspace.workspaceFolders;
  if (!workspaceFolders || workspaceFolders.length === 0) {
    throw "No open workspace.";
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
      const logger = new Logger(context.channel);
      logger.logCommandStart(commandName);

      // Execute PowerShell script
      const result = execPowerShell(scriptPath, [workspaceFolder], false);

      // Output result
      if (result.stdout) logger.logRaw(result.stdout);
      if (result.exitCode !== 0) {
        logger.logError(result.stderr);
        throw "Failed to create URL shortcut.";
      }
    },
  );
}
