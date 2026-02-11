import * as fs from "fs";

/** Read URL reference from .url file */
export function readUrlFile(urlPath: string): string | null {
  try {
    const content = fs.readFileSync(urlPath, "utf8");
    const urlMatch = content.match(/URL=(.+)/);
    if (urlMatch) {
      return urlMatch[1].trim();
    }
  } catch (error) {
    // ファイル読み込みエラー
  }
  return null;
}
