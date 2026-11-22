/** UTF-8 byte truncation that never splits a multi-byte character. */
export function truncateUtf8(value: string, maxBytes: number): { text: string; truncated: boolean } {
  if (maxBytes <= 0) return { text: value, truncated: false };
  const bytes = new TextEncoder().encode(value);
  if (bytes.length <= maxBytes) return { text: value, truncated: false };
  let cut = maxBytes;
  while (cut > 0 && (bytes[cut]! & 0xc0) === 0x80) cut -= 1;
  return { text: new TextDecoder().decode(bytes.slice(0, cut)), truncated: true };
}

export function coverage(items: { received_at: string | null }[]): {
  oldest_received: string | null;
  newest_received: string | null;
} {
  const dates = items.map((item) => item.received_at).filter((value): value is string => !!value).sort();
  return {
    oldest_received: dates[0] ?? null,
    newest_received: dates[dates.length - 1] ?? null,
  };
}

export function splitRecipients(value: string | undefined | null): string[] {
  if (!value) return [];
  return value
    .replace(/\n/g, ",")
    .replace(/;/g, ",")
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);
}

export function nowIso(now: () => Date): string {
  return now().toISOString();
}
