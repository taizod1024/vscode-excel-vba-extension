import * as vscode from "vscode";
const path = require("path");
import { CommandContext } from "../utils/types";
import { execPowerShell } from "../utils/execPowerShell";

/** Run VBA Sub command */
export async function runSubAsync(macroPath: string, context: CommandContext) {
  const editor = vscode.window.activeTextEditor;
  if (!editor) {
    throw "No editor is active. Please open a VBA file.";
  }

  const subName = extractSubNameAtCursor(editor);
  if (!subName) {
    throw "Could not find Sub or Function declaration at cursor position.";
  }

  const commandName = `Run VBA Sub: ${subName}`;
  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: commandName,
      cancellable: false,
    },
    async _progress => {
      const scriptPath = `${context.extensionPath}\\bin\\Run-Sub.ps1`;
      context.channel.appendLine("");
      context.channel.appendLine(`${commandName}`);
      context.channel.appendLine(`- File: ${path.basename(macroPath)}`);
      context.channel.appendLine(`- Sub: ${subName}`);

      // exec command
      const result = execPowerShell(scriptPath, [macroPath, subName]);

      // output result
      if (result.stdout) context.channel.appendLine(`- Output: ${result.stdout}`);
      if (result.exitCode !== 0) {
        throw `${result.stderr}`;
      }

      context.channel.appendLine(`[SUCCESS] Sub executed`);
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
