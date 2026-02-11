import * as vscode from "vscode";
import { excelvba } from "./ExcelVba";
import { execPowerShell } from "./utils/execPowerShell";

export function activate(context: vscode.ExtensionContext) {
  excelvba.activate(context);
  
  // Run Install-Addin.ps1 on startup
  const installAddinPath = vscode.Uri.joinPath(context.extensionUri, "bin", "Install-Addin.ps1").fsPath;
  execPowerShell(installAddinPath, []);
}

export function deactivate() {
  excelvba.deactivate();
}
