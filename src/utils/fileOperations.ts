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
