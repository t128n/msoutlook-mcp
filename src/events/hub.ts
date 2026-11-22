import {
  DEFAULT_NEXT_POLL_MS,
  MAX_EVENTS_PER_POLL,
  MAX_PAYLOAD_BYTES,
  MAX_SUBSCRIPTIONS,
  MIN_NEXT_POLL_MS,
  type EventDefinition,
  type EventOccurrence,
  type PollResult,
  validateEventArguments,
} from "../../packages/events/src/index.ts";
import { errorText } from "../com/errors.ts";
import type { Logger } from "../util/logger.ts";
import { EventStateStore, subscriptionKey } from "./state.ts";

export class EventHubError extends Error {
  constructor(
    readonly kind: "not_found" | "invalid_params" | "resource_exhausted",
    message: string,
  ) {
    super(message);
    this.name = "EventHubError";
  }
}

export type HubMessage =
  | { type: "active"; cursor: string | null; truncated: boolean; heartbeatMs: number }
  | { type: "event"; event: EventOccurrence; cursor: string | null }
  | { type: "heartbeat"; cursor: string | null }
  | { type: "error"; error: { code: number; message: string; data: Record<string, unknown> } };

export interface EventStream {
  messages: ReadableStream<HubMessage>;
  cancel(): void;
}

export interface EventHubOptions {
  definitions: EventDefinition<never, unknown>[] | EventDefinition<any, any>[];
  store: EventStateStore;
  logger?: Logger;
  onPoll?: () => void;
  now?: () => Date;
  heartbeatMs?: number;
  backoffInitialMs?: number;
  backoffMaxMs?: number;
  maxSubscriptions?: number;
  /** Test hook; production keeps the 1 s floor from the spec. */
  minNextPollMs?: number;
}

const SUB_ID = "io.modelcontextprotocol/subscriptionId";

export class EventHub {
  private readonly definitions = new Map<string, EventDefinition<any, any>>();
  private readonly store: EventStateStore;
  private readonly logger: Logger | undefined;
  private readonly onPoll: (() => void) | undefined;
  private readonly now: () => Date;
  private readonly heartbeatMs: number;
  private readonly backoffInitialMs: number;
  private readonly backoffMaxMs: number;
  private readonly maxSubscriptions: number;
  private readonly minNextPollMs: number;
  private readonly subscriptions = new Map<string, SharedSubscription>();

  constructor(options: EventHubOptions) {
    for (const definition of options.definitions) this.definitions.set(definition.name, definition);
    this.store = options.store;
    this.logger = options.logger;
    this.onPoll = options.onPoll;
    this.now = options.now ?? (() => new Date());
    this.heartbeatMs = clamp(options.heartbeatMs ?? 15_000, 1_000, 30_000);
    this.backoffInitialMs = options.backoffInitialMs ?? 2_000;
    this.backoffMaxMs = options.backoffMaxMs ?? 60_000;
    this.maxSubscriptions = options.maxSubscriptions ?? MAX_SUBSCRIPTIONS;
    this.minNextPollMs = options.minNextPollMs ?? MIN_NEXT_POLL_MS;
  }

  list(): { name: string; description: string; delivery: string[]; inputSchema: unknown; payloadSchema: unknown }[] {
    return [...this.definitions.values()].map((definition) => ({
      name: definition.name,
      description: definition.description,
      delivery: definition.delivery,
      inputSchema: definition.inputSchema,
      payloadSchema: definition.payloadSchema,
    }));
  }

  has(name: string): boolean {
    return this.definitions.has(name);
  }

  private definition(name: string): EventDefinition<any, any> {
    const definition = this.definitions.get(name);
    if (!definition) throw new EventHubError("not_found", `unknown event "${name}"`);
    return definition;
  }

  async poll(
    name: string,
    args: Record<string, unknown> | undefined,
    cursor: string | null,
    maxEvents = MAX_EVENTS_PER_POLL,
  ): Promise<PollResult> {
    const definition = this.definition(name);
    const normalizedArgs = args ?? {};
    try {
      validateEventArguments(definition, normalizedArgs);
    } catch (e) {
      throw new EventHubError("invalid_params", errorText(e));
    }
    const effective = cursor ?? this.store.get(name, normalizedArgs);
    const result = await definition.poll(effective, normalizedArgs, {
      limit: clamp(maxEvents, 1, MAX_EVENTS_PER_POLL),
    });
    this.store.set(name, normalizedArgs, result.cursor);
    this.store.markPoll();
    this.onPoll?.();
    return this.capPayloads(result);
  }

  async openStream(
    name: string,
    args: Record<string, unknown> | undefined,
    cursor: string | null,
    heartbeatMs?: number,
  ): Promise<{ stream: EventStream; heartbeatMs: number }> {
    const definition = this.definition(name);
    const normalizedArgs = args ?? {};
    try {
      validateEventArguments(definition, normalizedArgs);
    } catch (e) {
      throw new EventHubError("invalid_params", errorText(e));
    }
    const key = subscriptionKey(name, normalizedArgs);
    let subscription = this.subscriptions.get(key);
    if (!subscription) {
      if (this.subscriptions.size >= this.maxSubscriptions) {
        throw new EventHubError("resource_exhausted", `at most ${this.maxSubscriptions} event subscriptions`);
      }
      subscription = new SharedSubscription({
        key,
        definition,
        args: normalizedArgs,
        store: this.store,
        now: this.now,
        logger: this.logger,
        initialCursor: cursor ?? this.store.get(name, normalizedArgs),
        heartbeatMs: clamp(heartbeatMs ?? this.heartbeatMs, 1_000, 30_000),
        backoffInitialMs: this.backoffInitialMs,
        backoffMaxMs: this.backoffMaxMs,
        minNextPollMs: this.minNextPollMs,
        onPoll: this.onPoll,
        onEmpty: () => this.subscriptions.delete(key),
      });
      this.subscriptions.set(key, subscription);
      await subscription.initialize();
    }

    let controller: ReadableStreamDefaultController<HubMessage> | undefined;
    let unsubscribe: (() => void) | undefined;
    let closed = false;

    const messages = new ReadableStream<HubMessage>({
      start(ctrl) {
        controller = ctrl;
        unsubscribe = subscription!.attach((message) => {
          if (closed) return;
          try {
            ctrl.enqueue(message);
          } catch {
            closed = true;
          }
        });
      },
      cancel() {
        closed = true;
        unsubscribe?.();
      },
    });

    const cancel = () => {
      if (closed) return;
      closed = true;
      unsubscribe?.();
      try {
        controller?.close();
      } catch {
        /* already closed */
      }
    };

    return { stream: { messages, cancel }, heartbeatMs: subscription.heartbeatMs };
  }

  close(): void {
    for (const subscription of this.subscriptions.values()) subscription.stop();
    this.subscriptions.clear();
  }

  private capPayloads(result: PollResult): PollResult {
    const events: EventOccurrence[] = [];
    let dropped = false;
    for (const event of result.events) {
      const bytes = Buffer.byteLength(JSON.stringify(event.data), "utf8");
      if (bytes > MAX_PAYLOAD_BYTES) {
        dropped = true;
        continue;
      }
      events.push(event);
    }
    if (dropped) this.logger?.warn("event_payload_dropped", { name: "new_mail" });
    return dropped ? { ...result, events, truncated: true } : result;
  }
}

interface SharedOptions {
  key: string;
  definition: EventDefinition<any, any>;
  args: Record<string, unknown>;
  store: EventStateStore;
  now: () => Date;
  logger?: Logger;
  initialCursor: string | null;
  heartbeatMs: number;
  backoffInitialMs: number;
  backoffMaxMs: number;
  minNextPollMs: number;
  onPoll?: () => void;
  onEmpty?: () => void;
}

/**
 * One poll loop per distinct (name, arguments), shared by every stream that
 * asks for it. Listeners are added and removed without restarting the loop
 * while at least one stream remains.
 */
class SharedSubscription {
  readonly heartbeatMs: number;
  private readonly key: string;
  private readonly definition: EventDefinition<any, any>;
  private readonly args: Record<string, unknown>;
  private readonly store: EventStateStore;
  private readonly now: () => Date;
  private readonly logger: Logger | undefined;
  private readonly backoffInitialMs: number;
  private readonly backoffMaxMs: number;
  private readonly minNextPollMs: number;
  private readonly onPoll: (() => void) | undefined;
  private readonly onEmpty: (() => void) | undefined;
  private readonly listeners = new Set<(message: HubMessage) => void>();
  private cursor: string | null;
  private timer: ReturnType<typeof setTimeout> | undefined;
  private heartbeatTimer: ReturnType<typeof setInterval> | undefined;
  private backoffMs: number;
  private lastSent = 0;
  private stopped = false;
  private initialError: string | null = null;

  constructor(options: SharedOptions) {
    this.key = options.key;
    this.definition = options.definition;
    this.args = options.args;
    this.store = options.store;
    this.now = options.now;
    this.logger = options.logger;
    this.onPoll = options.onPoll;
    this.cursor = options.initialCursor;
    this.heartbeatMs = options.heartbeatMs;
    this.backoffInitialMs = options.backoffInitialMs;
    this.backoffMaxMs = options.backoffMaxMs;
    this.minNextPollMs = options.minNextPollMs;
    this.onEmpty = options.onEmpty;
    this.backoffMs = options.backoffInitialMs;
  }

  async initialize(): Promise<void> {
    if (this.cursor !== null) return;
    try {
      const result = await this.definition.poll(null, this.args, { limit: 1 });
      this.cursor = result.cursor;
      this.store.set(this.definition.name, this.args, result.cursor);
    } catch (e) {
      // Outlook may be away. Attach reports this error without closing the
      // stream; the loop retries with a null cursor and starts from the first
      // poll that succeeds.
      this.initialError = errorText(e);
      this.logger?.warn("event_initial_poll_failed", { event: this.definition.name, error: this.initialError });
    }
  }

  attach(listener: (message: HubMessage) => void): () => void {
    this.listeners.add(listener);
    listener({
      type: "active",
      cursor: this.cursor,
      truncated: false,
      heartbeatMs: this.heartbeatMs,
    });
    if (this.initialError) {
      listener({
        type: "error",
        error: { code: -32603, message: "UpstreamError", data: { reason: this.initialError } },
      });
      this.initialError = null;
    }
    if (this.listeners.size === 1) this.start();
    return () => {
      this.listeners.delete(listener);
      if (this.listeners.size === 0) this.stop();
    };
  }

  private start(): void {
    if (this.stopped) {
      this.stopped = false;
    }
    this.lastSent = this.now().getTime();
    this.heartbeatTimer = setInterval(() => {
      if (this.listeners.size === 0) return;
      if (this.now().getTime() - this.lastSent < this.heartbeatMs) return;
      this.broadcast({ type: "heartbeat", cursor: this.cursor });
      this.lastSent = this.now().getTime();
    }, Math.min(this.heartbeatMs, 5_000));
    (this.heartbeatTimer as { unref?: () => void }).unref?.();
    this.schedule(0);
  }

  stop(): void {
    this.stopped = true;
    if (this.timer) clearTimeout(this.timer);
    this.timer = undefined;
    if (this.heartbeatTimer) clearInterval(this.heartbeatTimer);
    this.heartbeatTimer = undefined;
    if (this.listeners.size === 0) this.onEmpty?.();
  }

  private schedule(delay: number): void {
    if (this.stopped) return;
    if (this.timer) clearTimeout(this.timer);
    this.timer = setTimeout(() => {
      this.timer = undefined;
      void this.tick();
    }, Math.max(0, delay));
    (this.timer as { unref?: () => void }).unref?.();
  }

  private async tick(): Promise<void> {
    if (this.stopped) return;
    try {
      const result = await this.definition.poll(this.cursor, this.args, { limit: MAX_EVENTS_PER_POLL });
      this.cursor = result.cursor;
      this.store.set(this.definition.name, this.args, result.cursor);
      this.store.markPoll();
      this.onPoll?.();
      this.backoffMs = this.backoffInitialMs;
      if (result.truncated) {
        this.broadcast({ type: "active", cursor: this.cursor, truncated: true, heartbeatMs: this.heartbeatMs });
      }
      for (let i = 0; i < result.events.length; i++) {
        const event = result.events[i]!;
        if (Buffer.byteLength(JSON.stringify(event.data), "utf8") > MAX_PAYLOAD_BYTES) {
          this.logger?.warn("event_payload_dropped", { event: this.definition.name });
          continue;
        }
        // Only the last event carries the end-of-batch cursor. Earlier events
        // have none, so a disconnect mid-batch resumes and re-delivers instead
        // of skipping the rest.
        const isLast = i === result.events.length - 1;
        this.broadcast({ type: "event", event, cursor: isLast ? result.cursor : null });
        this.lastSent = this.now().getTime();
      }
      this.schedule(
        result.hasMore ? 0 : Math.max(this.minNextPollMs, result.nextPollMs || DEFAULT_NEXT_POLL_MS),
      );
    } catch (e) {
      this.logger?.warn("event_poll_failed", { event: this.definition.name, error: errorText(e) });
      this.broadcast({
        type: "error",
        error: {
          code: -32603,
          message: "UpstreamError",
          data: { reason: errorText(e) },
        },
      });
      const delay = this.backoffMs;
      this.backoffMs = Math.min(this.backoffMaxMs, this.backoffMs * 2);
      this.schedule(delay);
    }
  }

  private broadcast(message: HubMessage): void {
    for (const listener of [...this.listeners]) {
      try {
        listener(message);
      } catch {
        this.listeners.delete(listener);
      }
    }
    if (message.type !== "heartbeat") this.lastSent = this.now().getTime();
    if (this.listeners.size === 0) this.stop();
  }
}

function clamp(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return Math.max(min, Math.min(max, value));
}

export { SUB_ID };
