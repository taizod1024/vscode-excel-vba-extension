import * as vscode from "vscode";

/** Log level */
type LogLevel = "INFO" | "WARN" | "ERROR" | "SUCCESS" | "DEBUG";

/**
 * Structured logger utility for consistent log output
 */
export class Logger {
  constructor(private channel: vscode.OutputChannel) {}

  /**
   * Log command start
   * @param commandName - Name of the command
   * @param details - Optional details as key-value pairs
   */
  public logCommandStart(commandName: string, details?: Record<string, string>): void {
    this.channel.appendLine("");
    this.channel.appendLine(`## ${commandName}`);
    if (details) {
      Object.entries(details).forEach(([key, value]) => {
        this.channel.appendLine(`- ${key}: ${value}`);
      });
    }
  }

  /**
   * Log a key-value pair
   */
  public logDetail(key: string, value: string): void {
    this.channel.appendLine(`- ${key}: ${value}`);
  }

  /**
   * Log a message with level
   */
  public log(level: LogLevel, message: string): void {
    this.channel.appendLine(`[${level}] ${message}`);
  }

  /**
   * Log success
   */
  public logSuccess(message: string): void {
    this.log("SUCCESS", message);
  }

  /**
   * Log info
   */
  public logInfo(message: string): void {
    this.log("INFO", message);
  }

  /**
   * Log warning
   */
  public logWarn(message: string): void {
    this.log("WARN", message);
  }

  /**
   * Log error
   */
  public logError(message: string): void {
    this.log("ERROR", message);
  }

  /**
   * Log raw text
   */
  public logRaw(text: string): void {
    this.channel.appendLine(text);
  }
}
