import * as vscode from "vscode";
const path = require("path");
import { CommandContext } from "../utils/types";
import { Logger } from "../utils/logger";
import { execPowerShell } from "../utils/execPowerShell";

/** Run VBA Sub command */
export async function runSubAsync(bookPath: string, context: CommandContext) {
  const editor = vscode.window.activeTextEditor;
  if (!editor) {
    throw "No active editor";
  }

  const subName = extractSubNameAtCursor(editor);
  if (!subName) {
    throw "No Sub/Function at cursor";
  }

  const commandName = `Run VBA Sub: ${subName}`;
  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: commandName,
      cancellable: false,
    },
    async _progress => {
      const logger = new Logger(context.channel);
      const scriptPath = `${context.extensionPath}\\bin\\Run-Sub.ps1`;
      logger.logCommandStart(commandName, {
        File: path.basename(bookPath),
        Sub: subName
      });

      // exec command
      const result = execPowerShell(scriptPath, [bookPath, subName]);

      // output result
      if (result.stdout) logger.logDetail("Output", result.stdout);
      if (result.exitCode !== 0) {
        // Extract first line of error message for user display
        const errorLine = result.stderr.split('\n')[0].trim() || "Failed to run Sub.";
        logger.logError(`${errorLine}:\n${result.stderr}`);
        throw errorLine;
      }

      logger.logSuccess("Sub executed");
    },
  );
}

/** Extract Sub/Function name at cursor position */
function extractSubNameAtCursor(editor: vscode.TextEditor): string | null {
  const cursorLine = editor.selection.active.line;
  const document = editor.document;

  // Search backwards and forwards for Sub/Function declaration
  let subName: string | null = null;

  // Search from cursor backwards to find the Sub/Function this cursor is in
  for (let i = cursorLine; i >= 0; i--) {
    const line = document.lineAt(i).text;

    // Match Sub or Function declaration - supports Japanese and other Unicode characters
    const match = line.match(/^\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+([\w\u0080-\uFFFF]+)\s*(?:\(|$)/i);
    if (match) {
      subName = match[1];
      break;
    }

    // Stop if we encounter End Sub/Function before finding declaration
    if (line.match(/^\s*End\s+(?:Sub|Function)\s*$/i) && i !== cursorLine) {
      break;
    }
  }

  return subName;
}
