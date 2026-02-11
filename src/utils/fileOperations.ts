import * as fs from "fs";
const path = require("path");

/** Remove directory recursively */
export function removeDirRecursive(dirPath: string): void {
  if (fs.existsSync(dirPath)) {
    fs.rmSync(dirPath, { recursive: true, force: true });
  }
}

/** Get VBA files from directory recursively */
export function getVbaFiles(dir: string, baseDir: string = ""): string[] {
  if (!fs.existsSync(dir)) {
    return [];
  }

  let files: string[] = [];
  const entries = fs.readdirSync(dir, { withFileTypes: true });

  entries.forEach(entry => {
    const fullPath = path.join(dir, entry.name);
    const relativePath = baseDir ? path.join(baseDir, entry.name) : entry.name;

    if (entry.isDirectory()) {
      files = files.concat(getVbaFiles(fullPath, relativePath));
    } else if ([".bas", ".cls", ".frm"].includes(path.extname(entry.name))) {
      files.push(relativePath);
    }
  });

  return files;
}

/** Copy addin to Excel AppData Add-ins folder */
export function copyAddinToAppData(sourceAddinPath: string, channel: any): boolean {
  try {
    if (!fs.existsSync(sourceAddinPath)) {
      channel.appendLine(`[WARNING] Addin source not found: ${sourceAddinPath}`);
      return false;
    }

    const addinFolder = path.join(process.env.APPDATA || "", "Microsoft", "AddIns");

    if (!fs.existsSync(addinFolder)) {
      fs.mkdirSync(addinFolder, { recursive: true });
    }

    const fileName = path.basename(sourceAddinPath);
    const destAddinPath = path.join(addinFolder, fileName);

    fs.copyFileSync(sourceAddinPath, destAddinPath);
    channel.appendLine(`[INFO] Addin copied to: ${destAddinPath}`);
    return true;
  } catch (error) {
    channel.appendLine(`[ERROR] Failed to copy addin: ${error}`);
    return false;
  }
}
