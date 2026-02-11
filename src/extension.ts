import * as vscode from "vscode";
import { excelvba } from "./ExcelVba";

export function activate(context: vscode.ExtensionContext) {
  excelvba.activate(context);
}

export function deactivate() {
  excelvba.deactivate();
}
