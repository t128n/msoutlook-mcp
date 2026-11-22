/**
 * COM failures translated into something a caller can act on.
 *
 * The PowerShell helper reports the HRESULT of the COM exception (and of any
 * inner exception). Everything that means "Outlook is not there" is marked
 * retryable so #6 can pause tools and resume after a reconnect.
 */

export type ComErrorKind =
  | "not_installed"
  | "not_running"
  | "connection_lost"
  | "timeout"
  | "access_denied"
  | "com";

export interface ComErrorOptions {
  code?: number | undefined;
  detail?: string | undefined;
  kind?: ComErrorKind | undefined;
}

const MESSAGES: Record<ComErrorKind, string> = {
  not_installed:
    "Classic Outlook is not installed or has no COM surface. The new Outlook is not supported; install Outlook Classic.",
  not_running:
    "Outlook isn't running. Minimize Outlook instead of closing it, then retry. The server reconnects on its own.",
  connection_lost:
    "The connection to Outlook was lost, usually because Outlook was closed. Minimize Outlook instead of closing it. The server reconnects with backoff; retry this call later.",
  timeout:
    "Outlook did not answer in time. A dialog in Outlook may be waiting for an answer (profile, sign-in or a security prompt). Close the dialog and retry.",
  access_denied:
    "Windows or your administrator blocked COM automation of Outlook. Ask IT to allow this program to automate Outlook, and make sure Outlook and this server run as the same, non-elevated user.",
  com: "Outlook returned an error.",
};

export const RETRYABLE_KINDS: ReadonlySet<ComErrorKind> = new Set(["not_running", "connection_lost", "timeout"]);

export class ComError extends Error {
  readonly code: number | undefined;
  readonly kind: ComErrorKind;
  readonly retryable: boolean;
  readonly detail: string | undefined;

  constructor(messageOrKind: string | ComErrorOptions, kind: ComErrorKind = "com") {
    const opts: ComErrorOptions = typeof messageOrKind === "string" ? { detail: messageOrKind, kind } : messageOrKind;
    const resolvedKind = opts.kind ?? kind;
    super(opts.detail ? `${MESSAGES[resolvedKind]} (${opts.detail})` : MESSAGES[resolvedKind]);
    this.name = "ComError";
    this.code = opts.code;
    this.kind = resolvedKind;
    this.retryable = RETRYABLE_KINDS.has(resolvedKind);
    this.detail = opts.detail;
  }

  toJSON(): { code: string; message: string; retryable: boolean; hresult?: string } {
    const out: { code: string; message: string; retryable: boolean; hresult?: string } = {
      code: this.retryable ? "OUTLOOK_UNAVAILABLE" : "OUTLOOK_ERROR",
      message: this.message,
      retryable: this.retryable,
    };
    if (this.code !== undefined) out.hresult = hresult(this.code);
    return out;
  }
}

/**
 * JSON that is pure ASCII. Non-ASCII UTF-16 code units become \uXXXX, so the
 * PowerShell helper's console code page can never change the bytes.
 */
export function asciiJson(value: unknown): string {
  return JSON.stringify(value).replace(/[^\x00-\x7e]/g, (c) => "\\u" + c.charCodeAt(0).toString(16).padStart(4, "0"));
}

export function hresult(code: number): string {
  return "0x" + (code >>> 0).toString(16).toUpperCase().padStart(8, "0");
}

/** Maps a Windows HRESULT to a readable kind. Unknown codes stay generic COM errors. */
export function classifyHresult(code: number | undefined): ComErrorKind {
  if (code === undefined) return "com";
  switch (code >>> 0) {
    case 0x80040154: // REGDB_E_CLASSNOTREG
    case 0x800401f3: // CO_E_CLASSSTRING
      return "not_installed";
    case 0x800401e3: // MK_E_UNAVAILABLE: GetActiveObject, nothing is running
    case 0x800401e4: // MK_E_SYNTAX
      return "not_running";
    case 0x800706ba: // RPC_S_SERVER_UNAVAILABLE: the Outlook process went away
    case 0x800706be: // RPC_S_CALL_FAILED
    case 0x80010108: // RPC_E_DISCONNECTED
    case 0x800a01a8: // "Object required": a stale object from the old process
    case 0x8002801d: // library not registered, observed after an Office repair
      return "connection_lost";
    case 0x80070005: // E_ACCESSDENIED, Object Model Guard
    case 0x80080005: // CO_E_SERVER_EXEC_FAILURE
      return "access_denied";
    default:
      return "com";
  }
}

export function comErrorFromAnswer(code: number | undefined, detail: string | undefined): ComError {
  return new ComError({ code, detail, kind: classifyHresult(code) });
}

export function connectionLost(detail?: string): ComError {
  return new ComError({ kind: "connection_lost", detail });
}

export function timeoutError(detail?: string): ComError {
  return new ComError({ kind: "timeout", detail });
}

export function isComError(e: unknown): e is ComError {
  return e instanceof ComError;
}

/** A plain error whose message the tool layer can show without leaking internals. */
export function errorText(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}
