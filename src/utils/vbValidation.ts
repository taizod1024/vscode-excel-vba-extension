import * as fs from "fs";
const path = require("path");

/** Validate that Attribute VB_Name matches file names */
export async function validateVbNames(folderPath: string, channel?: any): Promise<void> {
  const walkDir = (dir: string): string[] => {
    let results: string[] = [];
    const files = fs.readdirSync(dir);

    for (const file of files) {
      const filePath = path.join(dir, file);
      const stat = fs.statSync(filePath);

      if (stat.isDirectory()) {
        results = results.concat(walkDir(filePath));
      } else {
        results.push(filePath);
      }
    }
    return results;
  };

  const vbaFiles = walkDir(folderPath).filter(filePath => {
    const ext = path.extname(filePath).toLowerCase();
    return [".bas", ".cls", ".frm"].includes(ext);
  });

  for (const filePath of vbaFiles) {
    const fileName = path.basename(filePath);
    const componentName = path.parse(fileName).name;

    try {
      const content = fs.readFileSync(filePath, { encoding: "utf-8" });
      const attributeMatch = content.match(/Attribute\s+VB_Name\s*=\s*"([^"]+)"/);

      if (attributeMatch) {
        const vbName = attributeMatch[1];
        if (vbName !== componentName) {
          throw new Error(`MISMATCH Attribute VB_Name: "${vbName}" != "${componentName}" in file ${fileName}`);
        }
      }
    } catch (error) {
      if (error instanceof Error) {
        throw error.message;
      }
      throw error;
    }
  }
}
