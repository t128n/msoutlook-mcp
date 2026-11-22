/**
 * A tiny local implementation of the MCP Events authoring interface.
 *
 * It is vendored here on purpose: t128n/github-events-mcp#3 (the shared
 * package) does not exist yet. The interface is the one agreed for that
 * package: `defineEvent` plus a `poll` that returns
 * `{ events, cursor, hasMore, nextPollMs, truncated? }`.
 *
 * Protocol: `io.modelcontextprotocol/events`, revision 2026-07-28.
 */

export type Delivery = "poll" | "push";

export interface JsonSchema {
  [key: string]: unknown;
}

export interface EventDescriptor {
  name: string;
  description: string;
  delivery: Delivery[];
  inputSchema: JsonSchema;
  payloadSchema: JsonSchema;
}

export interface EventOccurrence<T = unknown> {
  /** Stable upstream id, for client-side deduplication within a subscription. */
  eventId: string;
  name: string;
  timestamp: string;
  data: T;
}

export interface PollResult<T = unknown> {
  events: EventOccurrence<T>[];
  /** Opaque position after this batch. */
  cursor: string;
  hasMore: boolean;
  nextPollMs: number;
  /** True only when the server started later than the supplied cursor. */
  truncated?: boolean;
}

export interface PollOptions {
  /** Batch cap for this poll. */
  limit: number;
  signal?: AbortSignal;
}

export interface EventDefinition<A = Record<string, unknown>, T = unknown> extends EventDescriptor {
  poll(cursor: string | null, args: A, options: PollOptions): Promise<PollResult<T>>;
}

export function defineEvent<A = Record<string, unknown>, T = unknown>(
  definition: EventDefinition<A, T>,
): EventDefinition<A, T> {
  return definition;
}

/** Hard limits the server applies to every event type. */
export const MAX_EVENTS_PER_POLL = 50;
export const MAX_PAYLOAD_BYTES = 64 * 1024;
export const MAX_SUBSCRIPTIONS = 32;
export const DEFAULT_NEXT_POLL_MS = 30_000;
export const MIN_NEXT_POLL_MS = 1_000;

/** A small JSON Schema check for subscription arguments. */
export function validateEventArguments(definition: EventDescriptor, args: unknown): asserts args is Record<string, unknown> {
  if (args === undefined || args === null) return;
  if (typeof args !== "object" || Array.isArray(args)) {
    throw new Error("arguments must be an object");
  }
  const value = args as Record<string, unknown>;
  const schema = definition.inputSchema as {
    properties?: Record<string, { type?: string | string[]; enum?: unknown[]; items?: { type?: string } }>;
    required?: string[];
  };
  for (const key of schema.required ?? []) {
    if (value[key] === undefined) throw new Error(`missing required argument "${key}"`);
  }
  for (const [key, spec] of Object.entries(schema.properties ?? {})) {
    const item = value[key];
    if (item === undefined) continue;
    if (spec.enum && !spec.enum.includes(item)) {
      throw new Error(`argument "${key}" must be one of: ${spec.enum.join(", ")}`);
    }
    if (spec.type) {
      const types = Array.isArray(spec.type) ? spec.type : [spec.type];
      const actual = Array.isArray(item) ? "array" : item === null ? "null" : typeof item;
      const normalized = actual === "number" && types.includes("integer") && Number.isInteger(item) ? "integer" : actual;
      if (!types.includes(normalized) && !types.includes(actual)) {
        throw new Error(`argument "${key}" must be ${types.join(" or ")}`);
      }
      if (actual === "array" && spec.items?.type) {
        if (!(item as unknown[]).every((v) => typeof v === spec.items!.type)) {
          throw new Error(`argument "${key}" must be an array of ${spec.items.type}`);
        }
      }
    }
  }
}
