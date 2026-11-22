import type { ComBridge } from "../com/bridge.ts";
import type { AppPaths } from "../config/paths.ts";
import { type TaskManager, windowsTaskManager } from "../platform/scheduler.ts";

export interface CommandContext {
  paths: AppPaths;
  platform: NodeJS.Platform;
  env: Record<string, string | undefined>;
  execPath: string;
  compiled: boolean;
  isTTY: boolean;
  out(text?: string): void;
  err(text?: string): void;
  bridgeFactory(): ComBridge;
  tasks: TaskManager;
  now(): Date;
  run(cmd: string[]): { code: number; out: string; err: string };
}
