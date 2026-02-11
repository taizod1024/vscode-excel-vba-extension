import child_process from "child_process";
import { PowerShellResult } from "./types";

/** Execute PowerShell script */
export function execPowerShell(scriptPath: string, args: string[], trim = true): PowerShellResult {
  try {
    const result = child_process.spawnSync("powershell.exe", ["-ExecutionPolicy", "RemoteSigned", "-File", scriptPath, ...args], {
      encoding: "utf8",
      stdio: ["pipe", "pipe", "pipe"],
    });
    let stdout = result.stdout || "";
    let stderr = result.stderr || "";
    if (result.error) {
      return { stdout: "", stderr: result.error.message, exitCode: 1 };
    }
    return {
      stdout: trim ? stdout.trim() : stdout,
      stderr: trim ? stderr.trim() : stderr,
      exitCode: result.status || 0,
    };
  } catch (ex: any) {
    return {
      stdout: "",
      stderr: trim ? (ex.message || "").trim() : ex.message || "",
      exitCode: 1,
    };
  }
}
