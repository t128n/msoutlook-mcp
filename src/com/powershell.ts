import { join } from "node:path";
import type { ComBridge } from "./bridge.ts";
import { ComError, asciiJson, comErrorFromAnswer, connectionLost, timeoutError } from "./errors.ts";
import { OUTLOOK_HELPER_SCRIPT } from "./script.ts";
import type {
  BusyInterval,
  ContactDetail,
  ContactSummary,
  CreateDraftArgs,
  CreateEventArgs,
  CreateEventResult,
  CreateRuleArgs,
  DraftResult,
  EventDetail,
  EventSummary,
  FolderNode,
  ListMailArgs,
  ListResult,
  MailDetail,
  MailRecipient,
  MailSummary,
  PingResult,
  PrepareReplyArgs,
  RuleSummary,
  SearchMailArgs,
} from "./types.ts";

export { asciiJson };

interface HelperProcess {
  stdin: { write(chunk: string): number; flush(): number | Promise<number> };
  stdout: ReadableStream<Uint8Array>;
  stderr: ReadableStream<Uint8Array>;
  exited: Promise<number>;
  kill(): void;
}

interface Pending {
  resolve: (value: unknown) => void;
  reject: (error: Error) => void;
  timer: ReturnType<typeof setTimeout>;
}

interface Answer {
  id?: number;
  ok?: boolean;
  result?: unknown;
  error?: { code?: number; message?: string };
}

export interface PowerShellBridgeOptions {
  /** Exact command. Tests pass a Bun stub that speaks the same JSON-lines protocol. */
  command?: string[];
  /** Append the encoded helper script to `command` (default false when command is set). */
  appendScript?: boolean;
  /** Default per-call timeout. */
  timeoutMs?: number;
  /** Heavy calls (calendar scans, contact search, attachment saves) get this timeout. */
  heavyTimeoutMs?: number;
}

const DEFAULT_TIMEOUT = 30_000;
const HEAVY_TIMEOUT = 120_000;

function powershellPath(): string {
  const root = process.env.SystemRoot ?? "C:\\Windows";
  return join(root, "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
}

function encodedArgs(): string[] {
  return [
    "-NoLogo",
    "-NoProfile",
    "-NonInteractive",
    "-Sta",
    "-ExecutionPolicy",
    "Bypass",
    "-EncodedCommand",
    Buffer.from(OUTLOOK_HELPER_SCRIPT, "utf16le").toString("base64"),
  ];
}

function array<T>(value: unknown): T[] {
  if (value === null || value === undefined) return [];
  return Array.isArray(value) ? (value as T[]) : [value as T];
}

/**
 * Host side of the persistent PowerShell helper.
 *
 * A call that times out (a modal Outlook prompt looks like a hang) kills the
 * helper. The next call starts a fresh process, which also drops every cached
 * COM object. The helper itself resets its `Outlook.Application` after any
 * error, so a closed Outlook reconnects on the next call.
 */
export class PowerShellComBridge implements ComBridge {
  private readonly options: PowerShellBridgeOptions;
  private proc: HelperProcess | undefined;
  private nextId = 1;
  private pending = new Map<number, Pending>();
  private stderr = "";
  private closed = false;

  constructor(options: PowerShellBridgeOptions = {}) {
    this.options = options;
  }

  private command(): string[] {
    if (!this.options.command) return [powershellPath(), ...encodedArgs()];
    return this.options.appendScript ? [...this.options.command, ...encodedArgs()] : [...this.options.command];
  }

  private start(): HelperProcess {
    if (this.proc) return this.proc;
    if (this.closed) throw connectionLost("the bridge was closed");
    let proc: HelperProcess;
    try {
      proc = Bun.spawn(this.command(), {
        stdin: "pipe",
        stdout: "pipe",
        stderr: "pipe",
        windowsHide: true,
      }) as unknown as HelperProcess;
    } catch (e) {
      throw new ComError({
        kind: "not_running",
        detail: "could not start Windows PowerShell: " + (e instanceof Error ? e.message : String(e)),
      });
    }
    this.proc = proc;
    this.stderr = "";
    void this.pump(proc);
    void this.collectStderr(proc);
    void proc.exited.then((code) => {
      if (this.proc !== proc) return;
      this.proc = undefined;
      const detail = `the Outlook helper exited (code ${code})` + (this.stderr ? `: ${this.stderr.trim()}` : "");
      this.failAll(connectionLost(detail));
    });
    return proc;
  }

  private async collectStderr(proc: HelperProcess): Promise<void> {
    const decoder = new TextDecoder();
    try {
      for await (const chunk of proc.stderr) {
        this.stderr = (this.stderr + decoder.decode(chunk)).slice(-2000);
      }
    } catch {
      /* the process is gone */
    }
  }

  private async pump(proc: HelperProcess): Promise<void> {
    const decoder = new TextDecoder();
    let buffer = "";
    try {
      for await (const chunk of proc.stdout) {
        buffer += decoder.decode(chunk, { stream: true });
        let newline = buffer.indexOf("\n");
        while (newline >= 0) {
          const line = buffer.slice(0, newline).trim();
          buffer = buffer.slice(newline + 1);
          if (line.startsWith("{")) this.answer(line);
          newline = buffer.indexOf("\n");
        }
      }
    } catch {
      /* handled by the exit hook */
    }
  }

  private answer(line: string): void {
    let message: Answer;
    try {
      message = JSON.parse(line) as Answer;
    } catch {
      return;
    }
    if (typeof message.id !== "number") return;
    const pending = this.pending.get(message.id);
    if (!pending) return;
    this.pending.delete(message.id);
    clearTimeout(pending.timer);
    if (message.ok) {
      pending.resolve(message.result);
    } else {
      pending.reject(comErrorFromAnswer(message.error?.code, message.error?.message));
    }
  }

  private failAll(error: Error): void {
    for (const [id, pending] of this.pending) {
      clearTimeout(pending.timer);
      pending.reject(error);
      this.pending.delete(id);
    }
  }

  private kill(reason: Error): void {
    const proc = this.proc;
    this.proc = undefined;
    this.failAll(reason);
    try {
      proc?.kill();
    } catch {
      /* already gone */
    }
  }

  /** @internal exposed so the protocol test can drive arbitrary stub ops. */
  request<T>(op: string, params: Record<string, unknown> = {}, timeoutMs?: number): Promise<T> {
    const proc = this.start();
    const id = this.nextId++;
    const limit = timeoutMs ?? this.options.timeoutMs ?? DEFAULT_TIMEOUT;
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        const error = timeoutError(`no answer to "${op}" within ${limit} ms`);
        this.kill(error);
        reject(error);
      }, limit);
      this.pending.set(id, { resolve: resolve as (value: unknown) => void, reject, timer });
      try {
        proc.stdin.write(asciiJson({ id, op, ...params }) + "\n");
        void proc.stdin.flush();
      } catch (e) {
        this.pending.delete(id);
        clearTimeout(timer);
        reject(connectionLost(e instanceof Error ? e.message : String(e)));
      }
    });
  }

  private heavy<T>(op: string, params: Record<string, unknown> = {}): Promise<T> {
    return this.request<T>(op, params, this.options.heavyTimeoutMs ?? HEAVY_TIMEOUT);
  }

  async ping(options: { launch?: boolean; timeoutMs?: number } = {}): Promise<PingResult> {
    return this.request<PingResult>("ping", { launch: options.launch === true }, options.timeoutMs ?? 60_000);
  }

  async listFolders(args: { maxDepth: number; mailbox?: string | null }): Promise<FolderNode[]> {
    const result = await this.request<unknown>("listFolders", {
      maxDepth: args.maxDepth,
      mailbox: args.mailbox ?? null,
    });
    const normalize = (node: FolderNode): FolderNode => ({
      ...node,
      children: array<FolderNode>(node.children).map(normalize),
    });
    return array<FolderNode>(result).map(normalize);
  }

  async listMail(args: ListMailArgs): Promise<ListResult<MailSummary>> {
    const result = await this.request<ListResult<MailSummary>>("listMail", {
      folder: args.folder ?? "inbox",
      mailbox: args.mailbox ?? null,
      limit: args.limit,
      filter: args.filter ?? null,
    });
    return { ...result, items: array<MailSummary>(result.items) };
  }

  async searchMail(args: SearchMailArgs): Promise<ListResult<MailSummary>> {
    const result = await this.request<ListResult<MailSummary>>("searchMail", {
      folder: args.folder ?? "inbox",
      mailbox: args.mailbox ?? null,
      limit: args.limit,
      query: args.query,
      scope: args.scope,
      filter: args.filter,
    });
    return { ...result, items: array<MailSummary>(result.items) };
  }

  async readMail(args: { entryId: string; maxBodyBytes: number; includeHtml: boolean }): Promise<MailDetail> {
    const result = await this.request<MailDetail>("readMail", {
      entryId: args.entryId,
      maxBodyBytes: args.maxBodyBytes,
      includeHtml: args.includeHtml,
    });
    return {
      ...result,
      recipients: array<MailRecipient>(result.recipients),
      attachments: array(result.attachments),
    };
  }

  downloadAttachment(args: {
    entryId: string;
    index: number;
    destDir?: string | null;
  }): Promise<{ saved_to: string; size: number | null }> {
    return this.heavy("downloadAttachment", {
      entryId: args.entryId,
      index: args.index,
      destDir: args.destDir ?? null,
    });
  }

  async createDraft(args: CreateDraftArgs): Promise<DraftResult> {
    const result = await this.request<DraftResult>("createDraft", { ...args });
    return { ...result, attachments: array<string>(result.attachments) };
  }

  async getRecipients(args: { entryId: string }): Promise<MailRecipient[]> {
    const result = await this.request<unknown>("getRecipients", { entryId: args.entryId });
    return array<MailRecipient>(result);
  }

  sendDraft(args: { entryId: string }): Promise<{ ok: true }> {
    return this.request<{ ok: true }>("sendDraft", { entryId: args.entryId });
  }

  async prepareReply(args: PrepareReplyArgs): Promise<DraftResult & { recipients: MailRecipient[] }> {
    const result = await this.request<DraftResult & { recipients: MailRecipient[] }>("prepareReply", {
      ...args,
    });
    return {
      ...result,
      attachments: array<string>(result.attachments),
      recipients: array<MailRecipient>(result.recipients),
    };
  }

  async moveMail(args: {
    entryId: string;
    targetFolder: string;
    mailbox?: string | null;
  }): Promise<{ new_entry_id: string | null; target_folder: string | null }> {
    return this.request("moveMail", {
      entryId: args.entryId,
      targetFolder: args.targetFolder,
      mailbox: args.mailbox ?? null,
    });
  }

  markRead(args: { entryId: string; read: boolean }): Promise<{ ok: true }> {
    return this.request("markRead", { entryId: args.entryId, read: args.read });
  }

  flagMail(args: { entryId: string; flag: boolean }): Promise<{ ok: true }> {
    return this.request("flagMail", { entryId: args.entryId, flag: args.flag });
  }

  async listEvents(args: {
    daysAhead: number;
    daysBack: number;
    limit: number;
    calendar?: string | null;
  }): Promise<EventSummary[]> {
    const result = await this.heavy<unknown>("listEvents", { ...args });
    return array<EventSummary>(result);
  }

  async readEvent(args: { entryId: string; maxBodyBytes: number }): Promise<EventDetail> {
    const result = await this.request<EventDetail>("readEvent", { ...args });
    return { ...result, recipients: array<MailRecipient>(result.recipients) };
  }

  async selfBusy(args: { startIso: string; daysAhead: number }): Promise<BusyInterval[]> {
    const result = await this.heavy<unknown>("selfBusy", { ...args });
    return array<BusyInterval>(result);
  }

  attendeeBusy(args: {
    smtp: string;
    startIso: string;
    slotMinutes: number;
  }): Promise<{ busy_string: string | null; error: string | null }> {
    return this.heavy("attendeeBusy", { ...args });
  }

  createEventDraft(args: CreateEventArgs): Promise<CreateEventResult> {
    return this.request("createEventDraft", { ...args });
  }

  async getEventRecipients(args: { entryId: string }): Promise<MailRecipient[]> {
    const result = await this.request<unknown>("getEventRecipients", { entryId: args.entryId });
    return array<MailRecipient>(result);
  }

  sendMeetingInvites(args: { entryId: string }): Promise<{ ok: true }> {
    return this.request("sendMeetingInvites", { entryId: args.entryId });
  }

  respondMeeting(args: {
    entryId: string;
    response: "accepted" | "tentative" | "declined";
    sendResponse: boolean;
    comment?: string | null;
  }): Promise<{ ok: true; action: "sent" | "saved" }> {
    return this.request("respondMeeting", {
      entryId: args.entryId,
      response: args.response,
      sendResponse: args.sendResponse,
      comment: args.comment ?? null,
    });
  }

  async listRules(): Promise<RuleSummary[]> {
    const result = await this.request<unknown>("listRules");
    return array<RuleSummary>(result);
  }

  toggleRule(args: { name: string; enabled: boolean }): Promise<{ ok: true }> {
    return this.request("toggleRule", args);
  }

  createRule(args: CreateRuleArgs): Promise<{ rule: RuleSummary }> {
    return this.request("createRule", { ...args });
  }

  deleteRule(args: { name: string }): Promise<{ ok: true }> {
    return this.request("deleteRule", args);
  }

  async searchContacts(args: { query: string; limit: number }): Promise<{
    contacts: ContactSummary[];
    hint?: string | null;
  }> {
    const result = await this.heavy<{ contacts?: unknown; hint?: string | null }>("searchContacts", args);
    return { contacts: array<ContactSummary>(result.contacts), hint: result.hint ?? null };
  }

  getContactDetails(args: { smtpAddress: string }): Promise<ContactDetail> {
    return this.heavy("getContactDetails", args);
  }

  close(): void {
    this.closed = true;
    const proc = this.proc;
    this.proc = undefined;
    this.failAll(connectionLost("the bridge was closed"));
    try {
      proc?.kill();
    } catch {
      /* already gone */
    }
  }
}
