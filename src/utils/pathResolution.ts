import * as fs from "fs";
const path = require("path");

/** Resolve VBA path from selected file */
export function resolveVbaPath(selectedPath: string): string {
  let resolvedPath = selectedPath;

  // Handle temporary Excel files (~$filename.xlsx)
  const fileName = path.basename(selectedPath);
  if (fileName.startsWith("~$")) {
    const dir = path.dirname(selectedPath);
    const actualFileName = fileName.substring(2); // Remove ~$ prefix
    resolvedPath = path.join(dir, actualFileName);
  }

  const ext = path.extname(resolvedPath).toLowerCase();

  // If .url file is selected, treat it as a marker for cloud-based files
  // Use the corresponding local Excel file if it exists
  if (ext === ".url") {
    const dir = path.dirname(resolvedPath);
    const fileNameWithoutExt = path.parse(resolvedPath).name;

    // Try to find .xlsm first, then .xlsx, then .xlam
    const xlsmPath = path.join(dir, `${fileNameWithoutExt}.xlsm`);
    if (fs.existsSync(xlsmPath)) {
      return xlsmPath;
    }

    const xlsxPath = path.join(dir, `${fileNameWithoutExt}.xlsx`);
    if (fs.existsSync(xlsxPath)) {
      return xlsxPath;
    }

    const xlamPath = path.join(dir, `${fileNameWithoutExt}.xlam`);
    if (fs.existsSync(xlamPath)) {
      return xlamPath;
    }

    // If local file doesn't exist, return the .url path itself
    // This will allow CSV/BAS/XML operations to use the corresponding folders
    return resolvedPath;
  }

  // If .xlsm, .xlam or .xlsx is selected, return as is
  if (ext === ".xlsm" || ext === ".xlam" || ext === ".xlsx") {
    return resolvedPath;
  }

  // If .csv is selected, find the parent _CSV folder and the corresponding Excel file
  if (ext === ".csv") {
    const parentDir = path.dirname(resolvedPath);
    let parentName = path.basename(parentDir);

    // Check if parent folder is _csv
    const match = parentName.match(/^(.+)_csv$/i);
    if (match) {
      const macroName = match[1];
      const parentParentDir = path.dirname(parentDir);

      // Try to find .xlsm first, then .xlsx, then .xlam, then .url
      const xlsmPath = path.join(parentParentDir, `${macroName}.xlsm`);
      if (fs.existsSync(xlsmPath)) {
        return xlsmPath;
      }

      const xlsxPath = path.join(parentParentDir, `${macroName}.xlsx`);
      if (fs.existsSync(xlsxPath)) {
        return xlsxPath;
      }

      const xlamPath = path.join(parentParentDir, `${macroName}.xlam`);
      if (fs.existsSync(xlamPath)) {
        return xlamPath;
      }

      const urlPath = path.join(parentParentDir, `${macroName}.url`);
      if (fs.existsSync(urlPath)) {
        return urlPath;
      }
    }
  }

  // If .bas, .cls, .frm is selected, find the parent _bas folder
  if ([".bas", ".cls", ".frm"].includes(ext)) {
    const parentDir = path.dirname(resolvedPath);
    let parentName = path.basename(parentDir);

    // Remove trailing ~ from parent folder name
    if (parentName.endsWith("~")) {
      parentName = parentName.slice(0, -1);
    }

    // Check if parent folder is _bas
    const match = parentName.match(/^(.+)_bas$/i);
    if (match) {
      const macroName = match[1];
      const parentParentDir = path.dirname(parentDir);

      // Try to find .xlsm first, then .xlsx, then .xlam, then .url
      const xlsmPath = path.join(parentParentDir, `${macroName}.xlsm`);
      if (fs.existsSync(xlsmPath)) {
        return xlsmPath;
      }

      const xlsxPath = path.join(parentParentDir, `${macroName}.xlsx`);
      if (fs.existsSync(xlsxPath)) {
        return xlsxPath;
      }

      const xlamPath = path.join(parentParentDir, `${macroName}.xlam`);
      if (fs.existsSync(xlamPath)) {
        return xlamPath;
      }

      const urlPath = path.join(parentParentDir, `${macroName}.url`);
      if (fs.existsSync(urlPath)) {
        return urlPath;
      }
    }
  }

  // If .xml is selected in a _xml folder, find the parent .xlam or .xlsm file
  if (ext === ".xml") {
    const parentDir = path.dirname(resolvedPath);
    let parentName = path.basename(parentDir);

    // Remove trailing ~ from parent folder name
    if (parentName.endsWith("~")) {
      parentName = parentName.slice(0, -1);
    }

    // Check if parent folder is _xml
    const match = parentName.match(/^(.+)_xml$/i);
    if (match) {
      const macroName = match[1];
      const parentParentDir = path.dirname(parentDir);

      // Try to find .xlam first, then .xlsm, then .url
      const xlamPath = path.join(parentParentDir, `${macroName}.xlam`);
      if (fs.existsSync(xlamPath)) {
        return xlamPath;
      }

      const xlsmPath = path.join(parentParentDir, `${macroName}.xlsm`);
      if (fs.existsSync(xlsmPath)) {
        return xlsmPath;
      }

      const urlPath = path.join(parentParentDir, `${macroName}.url`);
      if (fs.existsSync(urlPath)) {
        return urlPath;
      }

      // Default to .xlam if neither exists (will be handled as error later)
      return xlamPath;
    }
  }

  return resolvedPath;
}
