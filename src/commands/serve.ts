import { errorText } from "../com/errors.ts";
import { loadConfig, readToken } from "../config/config.ts";
import { HealthMonitor } from "../server/health.ts";
import { startMcpServer } from "../server/mcp.ts";
import { Logger } from "../util/logger.ts";
import { VERSION } from "../version.ts";
import type { CommandContext } from "./context.ts";

export interface ServeOptions {
  port?: number;
  console?: boolean;
  signals?: boolean;
}

export interface ServeHandle {
  port: number;
  url: string;
  stop(): void;
  done: Promise<void>;
}

export async function startServe(ctx: CommandContext, options: ServeOptions = {}): Promise<ServeHandle> {
  const config = loadConfig(ctx.paths);
  const token = readToken(ctx.paths);
  const port = options.port ?? config.port;
  const logger = new Logger(ctx.paths.logFile, config.log_level, { console: options.console ?? false });
  const bridge = ctx.bridgeFactory();
  const health = new HealthMonitor({ bridge, version: VERSION });

  let stop!: () => void;
  const done = new Promise<void>((resolve) => {
    stop = () => resolve();
  });

  let server: ReturnType<typeof startMcpServer>;
  try {
    server = startMcpServer({ config: { port }, token, health, version: VERSION, logger });
  } catch (e) {
    bridge.close();
    throw new Error(`could not listen on 127.0.0.1:${port}: ${errorText(e)}`);
  }

  let stopped = false;
  const shutdown = (reason: string) => {
    if (stopped) return;
    stopped = true;
    logger.info("server_stopping", { reason });
    server.stop();
    bridge.close();
    if (options.signals ?? true) {
      process.off("SIGINT", onSigint);
      process.off("SIGTERM", onSigterm);
    }
    stop();
  };
  const onSigint = () => shutdown("SIGINT");
  const onSigterm = () => shutdown("SIGTERM");

  if (options.signals ?? true) {
    process.on("SIGINT", onSigint);
    process.on("SIGTERM", onSigterm);
  }

  logger.info("serve_ready", { port: server.port });
  return {
    port: server.port,
    url: server.url,
    stop: () => shutdown("stop"),
    done,
  };
}

export async function runServe(ctx: CommandContext, options: ServeOptions = {}): Promise<number> {
  try {
    const handle = await startServe(ctx, options);
    ctx.out(`msoutlook-mcp ${VERSION} listening on ${handle.url}`);
    ctx.out("Health: http://127.0.0.1:" + handle.port + "/health (no token)");
    ctx.out("Minimize Outlook; do not close it.");
    await handle.done;
    return 0;
  } catch (e) {
    const error = e as { message: string; hint?: string };
    ctx.err(`error: ${error.message}`);
    if (error.hint) ctx.err(`fix: ${error.hint}`);
    return 1;
  }
}
