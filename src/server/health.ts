import type { ComBridge } from "../com/bridge.ts";
import { ComError, errorText } from "../com/errors.ts";

export type OutlookState = "connected" | "waiting_for_outlook" | "error";

export interface HealthSnapshot {
  status: OutlookState;
  version: string;
  lastSuccessfulCall: string | null;
  lastSuccessfulPoll: string | null;
  outlook: {
    state: OutlookState;
    version: string | null;
    profile: string | null;
    default_store: string | null;
    error?: string;
    retryable?: boolean;
  };
}

export interface HealthMonitorOptions {
  bridge: ComBridge;
  version: string;
  /** How long a successful ping is reused for /health. */
  cacheMs?: number;
  /** The ping timeout used by /health. */
  timeoutMs?: number;
  now?: () => Date;
}

/**
 * `/health` asks Outlook directly, but does not hammer it. A retryable COM
 * failure is `waiting_for_outlook`, not `error`: Outlook being closed is an
 * expected state (#6). The background reconnect loop and tool errors land in #6.
 */
export class HealthMonitor {
  private readonly bridge: ComBridge;
  private readonly version: string;
  private readonly cacheMs: number;
  private readonly timeoutMs: number;
  private readonly now: () => Date;
  private cached: { at: number; snapshot: HealthSnapshot } | undefined;
  private lastSuccessfulCall: string | null = null;
  private lastSuccessfulPoll: string | null = null;

  constructor(options: HealthMonitorOptions) {
    this.bridge = options.bridge;
    this.version = options.version;
    this.cacheMs = options.cacheMs ?? 5_000;
    this.timeoutMs = options.timeoutMs ?? 4_000;
    this.now = options.now ?? (() => new Date());
  }

  markCall(): void {
    this.lastSuccessfulCall = this.now().toISOString();
  }

  markPoll(): void {
    this.lastSuccessfulPoll = this.now().toISOString();
  }

  invalidate(): void {
    this.cached = undefined;
  }

  async snapshot(force = false): Promise<HealthSnapshot> {
    if (!force && this.cached && this.now().getTime() - this.cached.at < this.cacheMs) {
      return this.cached.snapshot;
    }
    let snapshot: HealthSnapshot;
    try {
      const ping = await this.bridge.ping({ timeoutMs: this.timeoutMs });
      this.lastSuccessfulCall = this.now().toISOString();
      snapshot = {
        status: "connected",
        version: this.version,
        lastSuccessfulCall: this.lastSuccessfulCall,
        lastSuccessfulPoll: this.lastSuccessfulPoll,
        outlook: {
          state: "connected",
          version: ping.version,
          profile: ping.profile,
          default_store: ping.default_store,
        },
      };
    } catch (e) {
      const retryable = e instanceof ComError ? e.retryable : false;
      const status: OutlookState = retryable ? "waiting_for_outlook" : "error";
      snapshot = {
        status,
        version: this.version,
        lastSuccessfulCall: this.lastSuccessfulCall,
        lastSuccessfulPoll: this.lastSuccessfulPoll,
        outlook: {
          state: status,
          version: null,
          profile: null,
          default_store: null,
          error: errorText(e),
          retryable,
        },
      };
    }
    this.cached = { at: this.now().getTime(), snapshot };
    return snapshot;
  }
}
