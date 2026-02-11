#!/usr/bin/env node

/**
 * Uninstall script for Excel VBA Extension
 * Lifecycle Hook: "vscode:uninstall"
 * Executed: By VS Code after extension uninstall + restart (all installation methods)
 * Purpose: Clean up excel-vba-addin.xlam from %APPDATA%\Microsoft\AddIns
 *
 * Handles both Marketplace and development installations:
 * - Marketplace VSIX: postinstall doesn't run, but activate() copies addin → vscode:uninstall removes it
 * - Local npm install: postinstall copies addin → vscode:uninstall removes it on uninstall
 */

const fs = require("fs");
const path = require("path");
const os = require("os");

function logToFile(message) {
  // Try multiple fallback paths for temp directory
  const tempDirs = [
    process.env.TEMP,
    process.env.TMP,
    process.env.LOCALAPPDATA ? path.join(process.env.LOCALAPPDATA, "Temp") : null,
    os.tmpdir(),
  ].filter(Boolean);

  for (const tempDir of tempDirs) {
    const logFile = path.join(tempDir, "excel-vba-uninstall.log");
    try {
      fs.appendFileSync(logFile, `${new Date().toISOString()} ${message}\n`);
      return; // Success, exit
    } catch (e) {
      // Continue to next fallback
    }
  }
}

function removeAddin() {
  logToFile("[START] Uninstall script initiated");
  
  try {
    const appData = process.env.APPDATA;
    if (!appData) {
      logToFile("[WARNING] APPDATA environment variable not found");
      process.exit(0);
    }

    logToFile(`[DEBUG] APPDATA: ${appData}`);
    const addinFolder = path.join(appData, "Microsoft", "AddIns");
    const addinPath = path.join(addinFolder, "excel-vba-addin.xlam");

    logToFile(`[DEBUG] Checking addin path: ${addinPath}`);
    if (fs.existsSync(addinPath)) {
      fs.unlinkSync(addinPath);
      logToFile(`[SUCCESS] Addin removed from: ${addinPath}`);
      process.exit(0);
    } else {
      logToFile(`[INFO] Addin not found at: ${addinPath}`);
      process.exit(0);
    }
  } catch (error) {
    logToFile(`[ERROR] Failed to uninstall addin: ${error.message}`);
    process.exit(0); // Don't fail the uninstall process
  }
}

removeAddin();
