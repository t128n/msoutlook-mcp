export function parseVersion(value: string): number[] {
  const clean = value.trim().replace(/^v/i, "").split(/[+-]/)[0] ?? "";
  return clean
    .split(".")
    .map((part) => Number.parseInt(part.replace(/\D.*$/, ""), 10))
    .map((n) => (Number.isFinite(n) ? n : 0));
}

/** -1 when a < b, 0 when equal, 1 when a > b. */
export function compareVersions(a: string, b: string): number {
  const left = parseVersion(a);
  const right = parseVersion(b);
  const length = Math.max(left.length, right.length);
  for (let i = 0; i < length; i++) {
    const x = left[i] ?? 0;
    const y = right[i] ?? 0;
    if (x < y) return -1;
    if (x > y) return 1;
  }
  return 0;
}
