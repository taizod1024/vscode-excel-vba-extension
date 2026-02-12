import * as fs from "fs";
import * as iconv from "iconv-lite";

/** Read URL reference from .url file */
export function readUrlFile(urlPath: string): string | null {
  try {
    const buffer = fs.readFileSync(urlPath);
    const content = iconv.decode(buffer, "shiftjis");
    const urlMatch = content.match(/URL=(.+)/);
    if (urlMatch) {
      return urlMatch[1].trim();
    }
  } catch (error) {
    // ファイル読み込みエラー
  }
  return null;
}
