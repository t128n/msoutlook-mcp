import { errorText } from "../com/errors.ts";
import { ResilientBridge } from "../com/resilient.ts";
import { loadConfig, readToken } from "../config/config.ts";
import { Guardrails } from "../config/guardrails.ts";
import { createTools } from "../tools/index.ts";
import { EventHub } from "../events/hub.ts";
import { createNewMailEvent } from "../events/new-mail.ts";
import { EventStateStore } from "../events/state.ts";
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
  let health: HealthMonitor | undefined;
  const connection = new ResilientBridge({
    bridge: ctx.bridgeFactory(),
    autoLaunchOutlook: config.auto_launch_outlook,
    onLog: (event, fields) => logger.info(event, fields),
    onStateChange: () => health?.invalidate(),
  });
  health = new HealthMonitor({
    bridge: connection,
    version: VERSION,
    connection: () => connection.snapshot(),
  });

  let stop!: () => void;
  const done = new Promise<void>((resolve) => {
    stop = () => resolve();
  });

  const stateStore = new EventStateStore(ctx.paths.eventsFile, ctx.now);
  const events = new EventHub({
    definitions: [
      createNewMailEvent({
        bridge: connection,
        intervalMs: config.poll_interval_ms,
        now: ctx.now,
      }),
    ],
    store: stateStore,
    logger,
    now: ctx.now,
    onPoll: () => health?.markPoll(),
  });

  const tools = createTools({
    bridge: connection,
    guardrails: Guardrails.fromPaths(ctx.paths),
    health: health!,
    paths: ctx.paths,
    version: VERSION,
    logger,
    now: ctx.now,
  });

  let server: ReturnType<typeof startMcpServer>;
  try {
    server = startMcpServer({
      config: { port },
      token,
      health: health!,
      version: VERSION,
      logger,
      tools,
      events,
    });
  } catch (e) {
    connection.close();
    throw new Error(`could not listen on 127.0.0.1:${port}: ${errorText(e)}`);
  }

  let stopped = false;
  const shutdown = (reason: string) => {
    if (stopped) return;
    stopped = true;
    logger.info("server_stopping", { reason });
    server.stop();
    events.close();
    connection.close();
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
  void connection.start().then((state) => logger.info("outlook_initial_state", { state }));
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
