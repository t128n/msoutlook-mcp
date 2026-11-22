import type { ComBridge } from "./bridge.ts";
import { ComError, connectionLost, errorText } from "./errors.ts";
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

export type ConnectionState = "connected" | "waiting_for_outlook" | "error";

export interface ConnectionSnapshot {
  state: ConnectionState;
  last_success_at: string | null;
  last_error: { kind: string; message: string; retryable: boolean } | null;
  attempts: number;
  next_retry_ms: number | null;
  next_retry_at: string | null;
}

export interface ResilientBridgeOptions {
  bridge: ComBridge;
  /** Decided in #6: off by default. The logon task already starts Outlook. */
  autoLaunchOutlook?: boolean;
  initialDelayMs?: number;
  maxDelayMs?: number;
  factor?: number;
  pingTimeoutMs?: number;
  onStateChange?: (state: ConnectionState, error: ComError | null) => void;
  onLog?: (event: string, fields?: Record<string, unknown>) => void;
  now?: () => Date;
}

/**
 * Wraps a ComBridge so a closed Outlook is a state, not a crash.
 *
 * - A retryable COM error moves to `waiting_for_outlook` and starts a probe
 *   with exponential backoff (2 s to 60 s).
 * - A successful probe moves to `connected`. Cached COM objects were dropped
 *   when the connection failed, so the next call reconnects from scratch.
 * - Calls while waiting fail immediately with a retryable error; they do not
 *   queue behind the backoff.
 * - A non-retryable error (not installed, access denied) moves to `error` and
 *   does not spin.
 * - Outline never starts Outlook unless `autoLaunchOutlook` is true.
 */
export class ResilientBridge implements ComBridge {
  private readonly bridge: ComBridge;
  private readonly autoLaunch: boolean;
  private readonly initialDelay: number;
  private readonly maxDelay: number;
  private readonly factor: number;
  private readonly pingTimeout: number;
  private readonly onChange: ((state: ConnectionState, error: ComError | null) => void) | undefined;
  private readonly log: ((event: string, fields?: Record<string, unknown>) => void) | undefined;
  private readonly now: () => Date;

  private stateValue: ConnectionState = "waiting_for_outlook";
  private lastSuccess: string | null = null;
  private lastError: ComError | null = null;
  private attempts = 0;
  private timer: ReturnType<typeof setTimeout> | undefined;
  private nextRetryMs: number | null = null;
  private nextRetryAt: string | null = null;
  private started = false;
  private stopped = false;
  private startPromise: Promise<ConnectionState> | undefined;

  constructor(options: ResilientBridgeOptions) {
    this.bridge = options.bridge;
    this.autoLaunch = options.autoLaunchOutlook ?? false;
    this.initialDelay = options.initialDelayMs ?? 2_000;
    this.maxDelay = options.maxDelayMs ?? 60_000;
    this.factor = options.factor ?? 2;
    this.pingTimeout = options.pingTimeoutMs ?? 5_000;
    this.onChange = options.onStateChange;
    this.log = options.onLog;
    this.now = options.now ?? (() => new Date());
  }

  get state(): ConnectionState {
    return this.stateValue;
  }

  snapshot(): ConnectionSnapshot {
    return {
      state: this.stateValue,
      last_success_at: this.lastSuccess,
      last_error: this.lastError
        ? { kind: this.lastError.kind, message: this.lastError.message, retryable: this.lastError.retryable }
        : null,
      attempts: this.attempts,
      next_retry_ms: this.nextRetryMs,
      next_retry_at: this.nextRetryAt,
    };
  }

  /** Runs one probe now. Used by start, backoff and tests. */
  async checkNow(): Promise<ConnectionState> {
    if (this.stopped) return this.stateValue;
    this.attempts = 0;
    await this.probe();
    return this.stateValue;
  }

  start(): Promise<ConnectionState> {
    if (this.stopped) return Promise.resolve(this.stateValue);
    if (!this.startPromise) this.startPromise = this.begin();
    return this.startPromise;
  }

  private async begin(): Promise<ConnectionState> {
    this.started = true;
    await this.probe();
    return this.stateValue;
  }

  private async probe(): Promise<void> {
    if (this.stopped) return;
    try {
      const ping = await this.bridge.ping({ launch: this.autoLaunch, timeoutMs: this.pingTimeout });
      void ping;
      this.succeed();
    } catch (e) {
      this.fail(e);
    }
  }

  private succeed(): void {
    const changed = this.stateValue !== "connected";
    this.clearTimer();
    this.attempts = 0;
    this.lastSuccess = this.now().toISOString();
    this.lastError = null;
    this.stateValue = "connected";
    if (changed) {
      this.log?.("outlook_connected", {});
      this.onChange?.("connected", null);
    }
  }

  private fail(error: unknown): ComError {
    const com = error instanceof ComError ? error : connectionLost(errorText(error));
    // Stale COM objects from the old Outlook process are useless. Drop them all.
    try {
      this.bridge.reset();
    } catch {
      /* the next call will start something new anyway */
    }
    const hadError = this.lastError !== null;
    this.lastError = com;
    if (com.retryable) {
      if (this.stateValue !== "waiting_for_outlook" || !hadError) {
        this.log?.("outlook_disconnected", { kind: com.kind });
        this.onChange?.("waiting_for_outlook", com);
      }
      this.stateValue = "waiting_for_outlook";
      this.schedule();
    } else {
      this.clearTimer();
      this.log?.("outlook_error", { kind: com.kind });
      this.stateValue = "error";
      this.onChange?.("error", com);
    }
    return com;
  }

  private schedule(): void {
    if (this.stopped || this.timer) return;
    this.attempts += 1;
    const delay = Math.min(this.maxDelay, Math.round(this.initialDelay * this.factor ** (this.attempts - 1)));
    this.nextRetryMs = delay;
    this.nextRetryAt = new Date(this.now().getTime() + delay).toISOString();
    this.log?.("outlook_reconnect_scheduled", { delay_ms: delay, attempt: this.attempts });
    this.timer = setTimeout(() => {
      this.timer = undefined;
      this.nextRetryMs = null;
      this.nextRetryAt = null;
      void this.probe();
    }, delay);
    (this.timer as { unref?: () => void }).unref?.();
  }

  private clearTimer(): void {
    if (this.timer) clearTimeout(this.timer);
    this.timer = undefined;
    this.nextRetryMs = null;
    this.nextRetryAt = null;
  }

  private unavailable(): ComError {
    if (this.lastError) {
      return new ComError({ kind: this.lastError.kind, detail: this.lastError.detail ?? this.lastError.message });
    }
    return new ComError({ kind: "connection_lost", detail: "Outlook is not reachable right now" });
  }

  private async run<T>(action: (bridge: ComBridge) => Promise<T>): Promise<T> {
    if (!this.started) await this.start();
    if (this.stateValue === "waiting_for_outlook") throw this.unavailable();
    if (this.stateValue === "error" && this.lastError) throw this.lastError;
    try {
      const value = await action(this.bridge);
      this.succeed();
      return value;
    } catch (e) {
      throw this.fail(e);
    }
  }

  reset(): void {
    this.bridge.reset();
    this.clearTimer();
    this.stateValue = "waiting_for_outlook";
    this.started = true;
    this.schedule();
  }

  close(): void {
    this.stopped = true;
    this.clearTimer();
    this.bridge.close();
  }

  ping(options: { launch?: boolean; timeoutMs?: number } = {}): Promise<PingResult> {
    return this.run((b) => b.ping(options));
  }

  listFolders(args: { maxDepth: number; mailbox?: string | null }): Promise<FolderNode[]> {
    return this.run((b) => b.listFolders(args));
  }

  listMail(args: ListMailArgs): Promise<ListResult<MailSummary>> {
    return this.run((b) => b.listMail(args));
  }

  searchMail(args: SearchMailArgs): Promise<ListResult<MailSummary>> {
    return this.run((b) => b.searchMail(args));
  }

  readMail(args: { entryId: string; maxBodyBytes: number; includeHtml: boolean }): Promise<MailDetail> {
    return this.run((b) => b.readMail(args));
  }

  downloadAttachment(args: {
    entryId: string;
    index: number;
    destDir?: string | null;
  }): Promise<{ saved_to: string; size: number | null }> {
    return this.run((b) => b.downloadAttachment(args));
  }

  createDraft(args: CreateDraftArgs): Promise<DraftResult> {
    return this.run((b) => b.createDraft(args));
  }

  getRecipients(args: { entryId: string }): Promise<MailRecipient[]> {
    return this.run((b) => b.getRecipients(args));
  }

  sendDraft(args: { entryId: string }): Promise<{ ok: true }> {
    return this.run((b) => b.sendDraft(args));
  }

  prepareReply(args: PrepareReplyArgs): Promise<DraftResult & { recipients: MailRecipient[] }> {
    return this.run((b) => b.prepareReply(args));
  }

  moveMail(args: {
    entryId: string;
    targetFolder: string;
    mailbox?: string | null;
  }): Promise<{ new_entry_id: string | null; target_folder: string | null }> {
    return this.run((b) => b.moveMail(args));
  }

  markRead(args: { entryId: string; read: boolean }): Promise<{ ok: true }> {
    return this.run((b) => b.markRead(args));
  }

  flagMail(args: { entryId: string; flag: boolean }): Promise<{ ok: true }> {
    return this.run((b) => b.flagMail(args));
  }

  listEvents(args: {
    daysAhead: number;
    daysBack: number;
    limit: number;
    calendar?: string | null;
  }): Promise<EventSummary[]> {
    return this.run((b) => b.listEvents(args));
  }

  readEvent(args: { entryId: string; maxBodyBytes: number }): Promise<EventDetail> {
    return this.run((b) => b.readEvent(args));
  }

  selfBusy(args: { startIso: string; daysAhead: number }): Promise<BusyInterval[]> {
    return this.run((b) => b.selfBusy(args));
  }

  attendeeBusy(args: {
    smtp: string;
    startIso: string;
    slotMinutes: number;
  }): Promise<{ busy_string: string | null; error: string | null }> {
    return this.run((b) => b.attendeeBusy(args));
  }

  createEventDraft(args: CreateEventArgs): Promise<CreateEventResult> {
    return this.run((b) => b.createEventDraft(args));
  }

  getEventRecipients(args: { entryId: string }): Promise<MailRecipient[]> {
    return this.run((b) => b.getEventRecipients(args));
  }

  sendMeetingInvites(args: { entryId: string }): Promise<{ ok: true }> {
    return this.run((b) => b.sendMeetingInvites(args));
  }

  respondMeeting(args: {
    entryId: string;
    response: "accepted" | "tentative" | "declined";
    sendResponse: boolean;
    comment?: string | null;
  }): Promise<{ ok: true; action: "sent" | "saved" }> {
    return this.run((b) => b.respondMeeting(args));
  }

  listRules(): Promise<RuleSummary[]> {
    return this.run((b) => b.listRules());
  }

  toggleRule(args: { name: string; enabled: boolean }): Promise<{ ok: true }> {
    return this.run((b) => b.toggleRule(args));
  }

  createRule(args: CreateRuleArgs): Promise<{ rule: RuleSummary }> {
    return this.run((b) => b.createRule(args));
  }

  deleteRule(args: { name: string }): Promise<{ ok: true }> {
    return this.run((b) => b.deleteRule(args));
  }

  searchContacts(args: { query: string; limit: number }): Promise<{
    contacts: ContactSummary[];
    hint?: string | null;
  }> {
    return this.run((b) => b.searchContacts(args));
  }

  getContactDetails(args: { smtpAddress: string }): Promise<ContactDetail> {
    return this.run((b) => b.getContactDetails(args));
  }
}
