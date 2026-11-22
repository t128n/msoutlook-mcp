/**
 * Recipient allowlist.
 *
 * Entries:
 *   user@example.com   exact address, matched literally (case-insensitive)
 *   example.com        any address at that exact domain
 *   *@example.com      any local part at that exact domain
 *   *.example.com      any subdomain (one or more labels), not the apex
 *   /pattern/          a regular expression, case-insensitive
 *
 * A plain entry never becomes a regex, so a typo stays a typo. An invalid
 * regex is an error: the config must not load.
 */

export type AllowlistRule =
  | { kind: "exact"; value: string }
  | { kind: "domain"; value: string }
  | { kind: "address-wildcard"; domain: string }
  | { kind: "subdomain"; domain: string }
  | { kind: "regex"; source: string; regex: RegExp };

export interface AllowlistResult {
  ok: boolean;
  unrestricted: boolean;
  rejected: string[];
  accepted: string[];
}

export class AllowlistError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "AllowlistError";
  }
}

/** `Name <a@b>` -> `a@b`; other text is trimmed unchanged. */
export function extractAddress(value: string): string {
  const trimmed = value.trim();
  const match = /<([^<>]+)>\s*$/.exec(trimmed);
  return (match?.[1] ?? trimmed).trim();
}

function domainOnly(value: string, entry: string): string {
  const domain = value.toLowerCase();
  if (!domain || domain.includes("@") || domain.includes("*") || domain.includes("/") || /\s/.test(domain)) {
    throw new AllowlistError(`invalid allowlist entry "${entry}"`);
  }
  return domain;
}

export function compileAllowlist(entries: readonly string[]): AllowlistRule[] {
  const rules: AllowlistRule[] = [];
  for (const raw of entries) {
    const entry = raw.trim();
    if (!entry) continue;

    if (entry.startsWith("/")) {
      if (entry.length < 3 || !entry.endsWith("/")) {
        throw new AllowlistError(`invalid regular expression entry "${entry}": expected /pattern/`);
      }
      const source = entry.slice(1, -1);
      if (!source) throw new AllowlistError(`invalid regular expression entry "${entry}": empty pattern`);
      let regex: RegExp;
      try {
        regex = new RegExp(source, "i");
      } catch (e) {
        throw new AllowlistError(
          `invalid regular expression entry "${entry}": ${e instanceof Error ? e.message : String(e)}`,
        );
      }
      rules.push({ kind: "regex", source, regex });
      continue;
    }

    if (entry.startsWith("*@")) {
      rules.push({ kind: "address-wildcard", domain: domainOnly(entry.slice(2), entry) });
      continue;
    }

    if (entry.startsWith("*.")) {
      rules.push({ kind: "subdomain", domain: domainOnly(entry.slice(2), entry) });
      continue;
    }

    if (entry.startsWith("@")) {
      rules.push({ kind: "domain", value: domainOnly(entry.slice(1), entry) });
      continue;
    }

    if (entry.includes("@")) {
      const address = entry.toLowerCase();
      if (/\s/.test(address) || address.endsWith("@") || address.startsWith("@")) {
        throw new AllowlistError(`invalid allowlist entry "${entry}"`);
      }
      rules.push({ kind: "exact", value: address });
      continue;
    }

    rules.push({ kind: "domain", value: domainOnly(entry, entry) });
  }
  return rules;
}

export function matchesAllowlist(rules: readonly AllowlistRule[], value: string): boolean {
  const address = extractAddress(value).toLowerCase();
  if (!address.includes("@")) return false;
  const domain = address.slice(address.lastIndexOf("@") + 1);
  for (const rule of rules) {
    switch (rule.kind) {
      case "exact":
        if (address === rule.value) return true;
        break;
      case "domain":
        if (domain === rule.value) return true;
        break;
      case "address-wildcard":
        if (domain === rule.domain) return true;
        break;
      case "subdomain":
        if (domain !== rule.domain && domain.endsWith("." + rule.domain)) return true;
        break;
      case "regex":
        if (rule.regex.test(address)) return true;
        break;
    }
  }
  return false;
}

export function checkRecipients(entries: readonly string[], recipients: readonly string[]): AllowlistResult {
  const rules = compileAllowlist(entries);
  if (rules.length === 0) {
    return { ok: true, unrestricted: true, rejected: [], accepted: [...recipients] };
  }
  const rejected: string[] = [];
  const accepted: string[] = [];
  for (const recipient of recipients) {
    if (matchesAllowlist(rules, recipient)) accepted.push(recipient);
    else rejected.push(recipient);
  }
  return { ok: rejected.length === 0, unrestricted: false, rejected, accepted };
}
