import * as vscode from "vscode";

/** Close all diff editors */
export async function closeAllDiffEditors(channel?: vscode.OutputChannel): Promise<void> {
  for (const group of vscode.window.tabGroups.all) {
    const tabsToClose = group.tabs.filter(tab => tab.input instanceof vscode.TabInputTextDiff);
    for (const tab of tabsToClose) {
      try {
        await vscode.window.tabGroups.close(tab);
      } catch (error) {
        // Ignore errors if tab is already closed
        if (channel) {
          channel.appendLine(`- note: tab already closed or not found`);
        }
      }
    }
  }
}

/** Show diff between two files */
export async function showDiffAsync(file1Path: string, file2Path: string, title: string) {
  const file1Uri = vscode.Uri.file(file1Path);
  const file2Uri = vscode.Uri.file(file2Path);
  await vscode.commands.executeCommand("vscode.diff", file1Uri, file2Uri, `Compare: ${title}`);
}
