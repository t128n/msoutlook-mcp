import { readFileSync } from "node:fs";
import { writePrivateFile } from "../config/paths.ts";

export interface CursorRecord {
  cursor: string;
  updated_at: string;
}

interface StateFile {
  version: 1;
  cursors: Record<string, CursorRecord>;
  last_poll_at: string | null;
}

const EMPTY: StateFile = { version: 1, cursors: {}, last_poll_at: null };

/** Canonical JSON so the same arguments always map to the same subscription key. */
export function canonicalJson(value: unknown): string {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(",")}]`;
  const object = value as Record<string, unknown>;
  return `{${Object.keys(object)
    .sort()
    .map((key) => `${JSON.stringify(key)}:${canonicalJson(object[key])}`)
    .join(",")}}`;
}

export function subscriptionKey(name: string, args: unknown): string {
  return `${name}:${canonicalJson(args ?? {})}`;
}

/**
 * The last cursor per subscription lives on disk. A daemon restart resumes
 * there, so old mail is not replayed and new mail is not skipped.
 */
export class EventStateStore {
  private state: StateFile = { ...EMPTY, cursors: {} };
  private readonly file: string;
  private readonly now: () => Date;

  constructor(file: string, now: () => Date = () => new Date()) {
    this.file = file;
    this.now = now;
    this.load();
  }

  private load(): void {
    try {
      const parsed = JSON.parse(readFileSync(this.file, "utf8")) as Partial<StateFile>;
      if (parsed && parsed.version === 1 && parsed.cursors && typeof parsed.cursors === "object") {
        this.state = {
          version: 1,
          cursors: parsed.cursors as Record<string, CursorRecord>,
          last_poll_at: parsed.last_poll_at ?? null,
        };
      }
    } catch {
      /* no state yet, or unreadable: start empty */
    }
  }

  get(name: string, args: unknown): string | null {
    return this.state.cursors[subscriptionKey(name, args)]?.cursor ?? null;
  }

  set(name: string, args: unknown, cursor: string): void {
    const key = subscriptionKey(name, args);
    this.state.cursors[key] = { cursor, updated_at: this.now().toISOString() };
    this.save();
  }

  markPoll(): void {
    this.state.last_poll_at = this.now().toISOString();
    this.save();
  }

  get lastPollAt(): string | null {
    return this.state.last_poll_at;
  }

  snapshot(): StateFile {
    return JSON.parse(JSON.stringify(this.state)) as StateFile;
  }

  private save(): void {
    writePrivateFile(this.file, JSON.stringify(this.state, null, 2) + "\n");
  }
}
