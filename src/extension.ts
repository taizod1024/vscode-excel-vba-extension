import * as vscode from "vscode";
import { excelvba } from "./ExcelVba";

// extension entrypoint
export function activate(context: vscode.ExtensionContext) {
  excelvba.activate(context);
}
export function deactivate() {}
