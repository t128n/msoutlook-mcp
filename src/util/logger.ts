import { appendFileSync, mkdirSync, renameSync, rmSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import type { LogLevel } from "../config/config.ts";

const LEVELS: Record<LogLevel, number> = { debug: 10, info: 20, warn: 30, error: 40 };
const MAX_BYTES = 5 * 1024 * 1024;
const KEEP = 3;

export interface LoggerOptions {
  console?: boolean;
  now?: () => Date;
}

/**
 * Small rotating file log. Callers pass short event names and numeric/boolean
 * fields only. Never pass a mail body, a subject, a token or a recipient list.
 */
export class Logger {
  constructor(
    readonly file: string,
    private level: LogLevel = "info",
    private readonly options: LoggerOptions = {},
  ) {
    try {
      mkdirSync(dirname(file), { recursive: true, mode: 0o700 });
    } catch {
      /* the logger must not take the daemon down */
    }
  }

  setLevel(level: LogLevel): void {
    this.level = level;
  }

  debug(event: string, fields?: Record<string, unknown>): void {
    this.write("debug", event, fields);
  }
  info(event: string, fields?: Record<string, unknown>): void {
    this.write("info", event, fields);
  }
  warn(event: string, fields?: Record<string, unknown>): void {
    this.write("warn", event, fields);
  }
  error(event: string, fields?: Record<string, unknown>): void {
    this.write("error", event, fields);
  }

  private write(level: LogLevel, event: string, fields?: Record<string, unknown>): void {
    if (LEVELS[level] < LEVELS[this.level]) return;
    const at = (this.options.now?.() ?? new Date()).toISOString();
    const tail = fields && Object.keys(fields).length > 0 ? " " + JSON.stringify(fields) : "";
    const line = `${at} ${level.toUpperCase()} ${event}${tail}\n`;
    try {
      this.rotateIfNeeded();
      appendFileSync(this.file, line, { mode: 0o600 });
    } catch {
      /* logging must not break a tool call */
    }
    if (this.options.console) {
      const stream = level === "error" || level === "warn" ? console.error : console.log;
      stream(line.trimEnd());
    }
  }

  private rotateIfNeeded(): void {
    try {
      if (statSync(this.file).size < MAX_BYTES) return;
    } catch {
      return;
    }
    for (let i = KEEP - 1; i >= 1; i--) {
      try {
        renameSync(`${this.file}.${i}`, `${this.file}.${i + 1}`);
      } catch {
        /* no such file */
      }
    }
    try {
      renameSync(this.file, `${this.file}.1`);
    } catch {
      rmSync(this.file, { force: true });
    }
  }
}

export function nullLogger(): Logger {
  return new Logger(join(tmpdir(), "msoutlook-mcp-silent.log"), "error");
}
