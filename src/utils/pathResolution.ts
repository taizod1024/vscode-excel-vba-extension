import * as fs from "fs";
const path = require("path");

/** Resolve VBA path from selected file */
export function resolveBookPath(selectedPath: string): string {
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
    const fileNameWithoutUrlExt = path.parse(resolvedPath).name; // e.g., "aaa.xlsx"

    // Try to find the file with the extracted name first
    const filePath = path.join(dir, fileNameWithoutUrlExt);
    if (fs.existsSync(filePath)) {
      return filePath;
    }

    // If not found, try old format compatibility (for aaa.url -> aaa.xlsm/xlsx/xlam)
    const xlsmPath = path.join(dir, `${fileNameWithoutUrlExt}.xlsm`);
    if (fs.existsSync(xlsmPath)) {
      return xlsmPath;
    }

    const xlsxPath = path.join(dir, `${fileNameWithoutUrlExt}.xlsx`);
    if (fs.existsSync(xlsxPath)) {
      return xlsxPath;
    }

    const xlamPath = path.join(dir, `${fileNameWithoutUrlExt}.xlam`);
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

  // If .csv is selected, find the parent .CSV folder and the corresponding Excel file
  if (ext === ".csv") {
    const parentDir = path.dirname(resolvedPath);
    let parentName = path.basename(parentDir);

    // Check if parent folder is .csv (format: aaa.xlsx.csv)
    const match = parentName.match(/^(.+\.(xlsm|xlsx|xlam))\.csv$/i);
    if (match) {
      const excelFileName = match[1];
      const parentParentDir = path.dirname(parentDir);

      // Try to find the exact file first
      const filePath = path.join(parentParentDir, excelFileName);
      if (fs.existsSync(filePath)) {
        return filePath;
      }

      // Also check for .url with the full filename
      const urlPath = path.join(parentParentDir, `${excelFileName}.url`);
      if (fs.existsSync(urlPath)) {
        return urlPath;
      }
    }
  }

  // If .bas, .cls, .frm is selected, find the parent .bas folder
  if ([".bas", ".cls", ".frm"].includes(ext)) {
    const parentDir = path.dirname(resolvedPath);
    let parentName = path.basename(parentDir);

    // Remove trailing ~ from parent folder name
    if (parentName.endsWith("~")) {
      parentName = parentName.slice(0, -1);
    }

    // Check if parent folder is .bas (format: aaa.xlsx.bas)
    const match = parentName.match(/^(.+\.(xlsm|xlsx|xlam))\.bas$/i);
    if (match) {
      const excelFileName = match[1];
      const parentParentDir = path.dirname(parentDir);

      // Try to find the exact file first
      const filePath = path.join(parentParentDir, excelFileName);
      if (fs.existsSync(filePath)) {
        return filePath;
      }

      // Also check for .url with the full filename
      const urlPath = path.join(parentParentDir, `${excelFileName}.url`);
      if (fs.existsSync(urlPath)) {
        return urlPath;
      }
    }
  }

  // If .xml is selected in a .xml folder, find the parent .xlam or .xlsm file
  if (ext === ".xml") {
    const parentDir = path.dirname(resolvedPath);
    let parentName = path.basename(parentDir);

    // Remove trailing ~ from parent folder name
    if (parentName.endsWith("~")) {
      parentName = parentName.slice(0, -1);
    }

    // Check if parent folder is .xml (format: aaa.xlam.xml or aaa.xlsm.xml)
    const match = parentName.match(/^(.+\.(xlam|xlsm))\.xml$/i);
    if (match) {
      const excelFileName = match[1];
      const parentParentDir = path.dirname(parentDir);

      // Try to find the exact file first
      const filePath = path.join(parentParentDir, excelFileName);
      if (fs.existsSync(filePath)) {
        return filePath;
      }

      // Also check for .url with the full filename
      const urlPath = path.join(parentParentDir, `${excelFileName}.url`);
      if (fs.existsSync(urlPath)) {
        return urlPath;
      }
    }
  }

  return resolvedPath;
}

/** Remove .url extension from path if present */
export function getActualPath(bookPath: string): string {
  const ext = path.extname(bookPath).toLowerCase();
  if (ext === ".url") {
    return bookPath.slice(0, -4); // Remove .url
  }
  return bookPath;
}

/** Resolve Excel display file name (handles .url and VBA component files) */
export function getExcelFileName(bookPath: string): string {
  const actualPath = getActualPath(bookPath);
  const fileExtension = path.parse(actualPath).ext.replace(".", "");
  const vbaComponentExtensions = ["bas", "cls", "frm", "frx"];
  let excelFileName = path.basename(actualPath);

  if (vbaComponentExtensions.includes(fileExtension)) {
    // VBA component file selected - extract Excel name from parent folder
    const parentFolderName = path.basename(path.dirname(actualPath));
    const match = parentFolderName.match(/^(.+\.(xlsm|xlsx|xlam))\.bas$/i);
    if (match) {
      excelFileName = match[1];
    }
  }

  return excelFileName;
}

/** Get file name parts (without extension and extension only) for path construction */
export function getFileNameParts(bookPath: string): { fileNameWithoutExt: string; excelExt: string } {
  const actualPath = getActualPath(bookPath);
  const fileNameWithoutExt = path.parse(actualPath).name;
  const excelExt = path.extname(actualPath).slice(1);
  return { fileNameWithoutExt, excelExt };
}
