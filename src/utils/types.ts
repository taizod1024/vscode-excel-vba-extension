import * as vscode from "vscode";

/** PowerShell実行結果 */
export interface PowerShellResult {
  stdout: string;
  stderr: string;
  exitCode: number;
}

/** コマンド処理用の共通コンテキスト */
export interface CommandContext {
  channel: vscode.OutputChannel;
  extensionPath: string;
}
