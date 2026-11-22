/**
 * Opaque cursors are base64url-encoded JSON. Shape validation belongs to the
 * caller; this only owns the envelope and the "invalid cursor" fallback.
 */
export function encodeJsonCursor(value: unknown): string {
  return Buffer.from(JSON.stringify(value), "utf8").toString("base64url");
}

export function decodeJsonCursor(value: string): unknown {
  try {
    return JSON.parse(Buffer.from(value, "base64url").toString("utf8")) as unknown;
  } catch {
    /* invalid cursor */
  }
  return null;
}
