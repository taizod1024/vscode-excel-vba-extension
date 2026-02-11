import * as fs from "fs";
import child_process from "child_process";

/** Get Excel executable path */
export function getExcelPath(): string {
  try {
    // assoc で .xlsx の関連付けを確認
    const assocResult = child_process.execSync("cmd /c assoc .xlsx", { encoding: "utf8" }).trim();
    const progIdMatch = assocResult.match(/=(.+)$/);

    if (progIdMatch) {
      const progId = progIdMatch[1];
      // ftype で実行コマンドを取得
      const ftypeResult = child_process.execSync(`cmd /c ftype ${progId}`, { encoding: "utf8" }).trim();
      const exePathMatch = ftypeResult.match(/"([^"]+\.exe)"/i);

      if (exePathMatch) {
        const excelPath = exePathMatch[1];
        if (fs.existsSync(excelPath)) {
          return excelPath;
        }
      }
    }
  } catch (error) {
    // assoc/ftype コマンド失敗
  }

  throw new Error("Excel installation not found");
}
