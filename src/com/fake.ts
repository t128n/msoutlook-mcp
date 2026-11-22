/**
 * In-memory Outlook used by every test that does not need COM.
 *
 * It implements the whole `ComBridge` surface, including Outlook's habit of
 * failing when the process is gone. Tests run on macOS, Linux and CI.
 */
import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import type { ComBridge } from "./bridge.ts";
import { ComError, connectionLost, timeoutError } from "./errors.ts";
import {
  DEFAULT_MAX_ATTACHMENT_BYTES,
  DEFAULT_MAX_ATTACHMENT_TOTAL_BYTES,
  MAX_TRANSPORT_HEADER_CHARS,
} from "./limits.ts";
import type {
  AttachmentInfo,
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
  ReadMailArgs,
  RuleSummary,
  SearchMailArgs,
} from "./types.ts";

export interface FakeAttachment {
  filename: string;
  size: number;
  content?: Uint8Array;
  content_id?: string;
  mime_type?: string;
}

export interface FakeMail {
  entry_id: string;
  folder: string;
  received_at: string;
  sent_at: string | null;
  sender: string;
  sender_address: string;
  sender_smtp: string;
  to: string;
  cc: string;
  bcc?: string;
  subject: string;
  body: string;
  html?: string;
  is_read: boolean;
  importance: number;
  internet_message_id: string | null;
  transport_headers?: string | null;
  attachments: FakeAttachment[];
  flagged?: boolean;
}

export interface FakeEvent {
  entry_id: string;
  subject: string;
  start: string;
  end: string;
  location: string;
  organizer: string;
  body: string;
  required: string;
  optional: string;
  is_recurring?: boolean;
  is_all_day?: boolean;
  meeting_status?: number;
  response_status?: number;
  invite_sent?: boolean;
}

export interface FakeRule extends RuleSummary {}

export interface FakeContact extends ContactDetail {}

export interface FakeState {
  running: boolean;
  version: string;
  profile: string;
  default_store: string;
  user_smtp: string;
  mail: Map<string, FakeMail>;
  events: Map<string, FakeEvent>;
  rules: FakeRule[];
  contacts: FakeContact[];
  attendee_busy: Map<string, string>;
  next_id: number;
  latency_ms: number;
  last_ping_launch: boolean;
  resets: number;
}

export interface FakeBridgeOptions {
  running?: boolean;
  version?: string;
  profile?: string;
  store?: string;
  user?: string;
}

let counter = 0;

function entryId(prefix: string, state: FakeState): string {
  state.next_id += 1;
  return `${prefix}-${state.next_id}-${(counter += 1)}`;
}

function splitAddresses(value: string | undefined | null): string[] {
  if (!value) return [];
  return value
    .replace(/\n/g, ",")
    .replace(/;/g, ",")
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      const match = /<([^>]+)>/.exec(part);
      return (match?.[1] ?? part).trim();
    });
}

function domainOf(address: string | null | undefined): string {
  if (!address || !address.includes("@")) return "";
  return address.split("@").pop()!.toLowerCase();
}

export class FakeComBridge implements ComBridge {
  state: FakeState;
  failNext: ComError | undefined;

  constructor(options: FakeBridgeOptions = {}) {
    this.state = {
      running: options.running ?? true,
      version: options.version ?? "16.0.17928.20156",
      profile: options.profile ?? "Outlook",
      default_store: options.store ?? "Personal Mailbox",
      user_smtp: options.user ?? "user@example.com",
      mail: new Map(),
      events: new Map(),
      rules: [],
      contacts: [],
      attendee_busy: new Map(),
      next_id: 0,
      latency_ms: 0,
      last_ping_launch: false,
      resets: 0,
    };
  }

  /** The process went away: the next call fails the way GetActiveObject does. */
  closeOutlook(): void {
    this.state.running = false;
  }

  openOutlook(): void {
    this.state.running = true;
  }

  seedMail(input: Partial<FakeMail> & { subject?: string }): FakeMail {
    const id = input.entry_id ?? entryId("mail", this.state);
    const received = input.received_at ?? new Date().toISOString();
    const mail: FakeMail = {
      entry_id: id,
      folder: input.folder ?? "inbox",
      received_at: received,
      sent_at: input.sent_at ?? received,
      sender: input.sender ?? "Sender Example",
      sender_address: input.sender_address ?? "/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP/CN=RECIPIENTS/CN=someone",
      sender_smtp: input.sender_smtp ?? "sender@example.com",
      to: input.to ?? "user@example.com",
      cc: input.cc ?? "",
      bcc: input.bcc ?? "",
      subject: input.subject ?? "Example subject",
      body: input.body ?? "Example body.",
      html: input.html,
      is_read: input.is_read ?? false,
      importance: input.importance ?? 1,
      internet_message_id: input.internet_message_id ?? `<${id}@example.com>`,
      transport_headers: input.transport_headers ?? null,
      attachments: input.attachments ?? [],
      flagged: input.flagged ?? false,
    };
    this.state.mail.set(id, mail);
    return mail;
  }

  seedEvent(input: Partial<FakeEvent> & { subject?: string; start: string }): FakeEvent {
    const id = input.entry_id ?? entryId("event", this.state);
    const start = new Date(input.start);
    const duration = (input as { duration?: number }).duration ?? 30;
    const end = input.end ?? new Date(start.getTime() + duration * 60_000).toISOString();
    const event: FakeEvent = {
      entry_id: id,
      subject: input.subject ?? "Example event",
      start: start.toISOString(),
      end,
      location: input.location ?? "Room 1",
      organizer: input.organizer ?? "user@example.com",
      body: input.body ?? "Example event body.",
      required: input.required ?? "",
      optional: input.optional ?? "",
      is_recurring: input.is_recurring ?? false,
      is_all_day: input.is_all_day ?? false,
      meeting_status: input.meeting_status ?? 0,
      response_status: input.response_status ?? 0,
      invite_sent: input.invite_sent ?? false,
    };
    this.state.events.set(id, event);
    return event;
  }

  seedRule(rule: Partial<FakeRule> & { name: string }): FakeRule {
    const value: FakeRule = {
      name: rule.name,
      enabled: rule.enabled ?? true,
      execution_order: rule.execution_order ?? this.state.rules.length + 1,
      is_local_rule: rule.is_local_rule ?? false,
      rule_type: rule.rule_type ?? 0,
      conditions: rule.conditions ?? [],
      actions: rule.actions ?? [],
    };
    this.state.rules.push(value);
    return value;
  }

  seedContact(contact: Partial<FakeContact> & { name: string }): FakeContact {
    const value: FakeContact = {
      name: contact.name,
      smtp_address: contact.smtp_address ?? `${contact.name.toLowerCase().replace(/\s+/g, ".")}@example.com`,
      alias: contact.alias ?? null,
      job_title: contact.job_title ?? null,
      department: contact.department ?? null,
      office_location: contact.office_location ?? null,
      company_name: contact.company_name ?? "Example Corp",
      business_phone: contact.business_phone ?? null,
      mobile_phone: contact.mobile_phone ?? null,
      street_address: contact.street_address ?? null,
      city: contact.city ?? null,
      state: contact.state ?? null,
      postal_code: contact.postal_code ?? null,
      manager: contact.manager ?? null,
    };
    this.state.contacts.push(value);
    return value;
  }

  private async ready(): Promise<void> {
    const failure = this.failNext;
    this.failNext = undefined;
    if (failure) throw failure;
    if (this.state.latency_ms > 0) await Bun.sleep(this.state.latency_ms);
    if (!this.state.running) throw new ComError({ kind: "not_running" });
  }

  private locate(entryId: string): FakeMail {
    const mail = this.state.mail.get(entryId);
    if (!mail) throw new ComError({ kind: "com", detail: `ENTRY_ID_INVALID: ${entryId}` });
    return mail;
  }

  private locateEvent(entryId: string): FakeEvent {
    const event = this.state.events.get(entryId);
    if (!event) throw new ComError({ kind: "com", detail: `ENTRY_ID_INVALID: ${entryId}` });
    return event;
  }

  private folderMatches(folder: string, requested: string | undefined): boolean {
    const want = (requested ?? "inbox").toLowerCase();
    const have = folder.toLowerCase();
    if (want === "inbox") return have === "inbox" || have.endsWith("\\inbox");
    if (want === "sent") return have === "sent" || have.endsWith("\\sent items");
    if (want === "drafts") return have === "drafts" || have.endsWith("\\drafts");
    return have === want;
  }

  private afterCursor(mail: FakeMail, cursor: ListMailArgs["cursor"], ascending: boolean): boolean {
    if (!cursor) return true;
    const seconds = Math.floor(new Date(mail.received_at).getTime() / 1000);
    const at = Math.floor(new Date(cursor.at).getTime() / 1000);
    const cmp = seconds - at;
    if (ascending) {
      if (cmp < 0) return false;
      return cmp > 0 || !cursor.ids.includes(mail.entry_id);
    }
    if (cmp > 0) return false;
    return cmp < 0 || !cursor.ids.includes(mail.entry_id);
  }

  private summary(mail: FakeMail): MailSummary {
    return {
      entry_id: mail.entry_id,
      received_at: mail.received_at,
      sent_at: mail.sent_at,
      sender: mail.sender,
      sender_address: mail.sender_address,
      sender_smtp: mail.sender_smtp,
      to: mail.to,
      cc: mail.cc,
      subject: mail.subject,
      is_read: mail.is_read,
      has_attachments: mail.attachments.length > 0,
      size_bytes: mail.body.length + mail.subject.length,
      importance: mail.importance,
      folder: mail.folder,
      internet_message_id: mail.internet_message_id,
    };
  }

  private recipientsOf(mail: FakeMail): MailRecipient[] {
    const out: MailRecipient[] = [];
    for (const address of splitAddresses(mail.to)) out.push({ name: address, address, smtp: address, type: 1 });
    for (const address of splitAddresses(mail.cc)) out.push({ name: address, address, smtp: address, type: 2 });
    for (const address of splitAddresses(mail.bcc)) out.push({ name: address, address, smtp: address, type: 3 });
    return out;
  }

  async ping(options: { launch?: boolean; timeoutMs?: number } = {}): Promise<PingResult> {
    this.state.last_ping_launch = options.launch === true;
    await this.ready();
    const inbox = [...this.state.mail.values()].filter((m) => this.folderMatches(m.folder, "inbox"));
    return {
      version: this.state.version,
      profile: this.state.profile,
      default_store: this.state.default_store,
      user_smtp: this.state.user_smtp,
      inbox_count: inbox.length,
      inbox_unread: inbox.filter((m) => !m.is_read).length,
      calendar_count: this.state.events.size,
      accounts: [{ display: this.state.profile, smtp: this.state.user_smtp }],
    };
  }

  async listFolders(args: { maxDepth: number; mailbox?: string | null }): Promise<FolderNode[]> {
    await this.ready();
    if (args.mailbox && args.mailbox.toLowerCase() !== this.state.default_store.toLowerCase()) {
      throw new ComError({ kind: "com", detail: `MAILBOX_NOT_FOUND: ${args.mailbox}` });
    }
    const names = ["Inbox", "Sent Items", "Drafts", "Calendar", "Archive"];
    const children = names.slice(0, args.maxDepth > 0 ? names.length : 0).map((name) => ({
      name,
      path: `\\\\${this.state.default_store}\\${name}`,
      item_count: [...this.state.mail.values()].filter((m) => this.folderMatches(m.folder, name)).length,
      unread_count: 0,
      children: [],
    }));
    return [
      {
        name: this.state.default_store,
        path: `\\\\${this.state.default_store}`,
        item_count: this.state.mail.size,
        unread_count: 0,
        children,
        store: this.state.default_store,
      },
    ];
  }

  async listMail(args: ListMailArgs): Promise<ListResult<MailSummary>> {
    await this.ready();
    const since = args.since_iso ? new Date(args.since_iso).getTime() : undefined;
    const matching = [...this.state.mail.values()]
      .filter((m) => this.folderMatches(m.folder, args.folder))
      .filter((m) => (args.unread_only ? !m.is_read : true))
      .filter((m) => (since === undefined ? true : new Date(m.received_at).getTime() >= since))
      .filter((m) => this.afterCursor(m, args.cursor, args.sort_ascending === true));
    const ordered = matching.sort((a, b) =>
      args.sort_ascending
        ? a.received_at.localeCompare(b.received_at)
        : b.received_at.localeCompare(a.received_at),
    );
    const page = ordered.slice(0, args.limit);
    return {
      items: page.map((m) => this.summary(m)),
      truncated: ordered.length > page.length,
      folder: `\\\\${this.state.default_store}\\Inbox`,
      mailbox_used: this.state.default_store,
    };
  }

  async searchMail(args: SearchMailArgs): Promise<ListResult<MailSummary>> {
    await this.ready();
    const q = args.query.toLowerCase();
    const matching = [...this.state.mail.values()]
      .filter((m) => this.folderMatches(m.folder, args.folder))
      .filter((m) => {
        const subject = m.subject.toLowerCase().includes(q);
        const body = m.body.toLowerCase().includes(q);
        const sender = m.sender.toLowerCase().includes(q) || m.sender_smtp.toLowerCase().includes(q);
        if (args.scope === "subject") return subject;
        if (args.scope === "sender") return sender;
        if (args.scope === "all") return subject || body || sender;
        return subject || body;
      })
      .filter((m) => this.afterCursor(m, args.cursor, false))
      .sort((a, b) => b.received_at.localeCompare(a.received_at));
    const page = matching.slice(0, args.limit);
    return {
      items: page.map((m) => this.summary(m)),
      truncated: matching.length > page.length,
      folder: `\\\\${this.state.default_store}\\Inbox`,
      mailbox_used: this.state.default_store,
    };
  }

  async readMail(args: ReadMailArgs): Promise<MailDetail> {
    await this.ready();
    const mail = this.locate(args.entryId);
    const encoder = new TextEncoder();
    let body = mail.body;
    let bodyTruncated = false;
    if (args.maxBodyBytes > 0 && encoder.encode(body).length > args.maxBodyBytes) {
      body = new TextDecoder().decode(encoder.encode(body).slice(0, args.maxBodyBytes));
      bodyTruncated = true;
    }
    const maxAttachmentBytes = args.maxAttachmentBytes ?? DEFAULT_MAX_ATTACHMENT_BYTES;
    const maxAttachmentTotalBytes = args.maxAttachmentTotalBytes ?? DEFAULT_MAX_ATTACHMENT_TOTAL_BYTES;
    let totalBytes = 0;
    const attachments = mail.attachments.map((a, index) => {
      const contentId = a.content_id ?? "";
      const info: AttachmentInfo = {
        index: index + 1,
        filename: a.filename,
        size: a.size,
        content_id: contentId,
        mime_type: a.mime_type ?? "",
        is_inline: contentId !== "",
      };
      if (args.attachmentMode !== "base64") return info;
      const bytes = a.content;
      if (!bytes) {
        info.data_base64 = "";
        info.skipped_reason = "unavailable";
      } else if (bytes.length > maxAttachmentBytes) {
        info.data_base64 = "";
        info.skipped_reason = "too_large";
      } else if (totalBytes + bytes.length > maxAttachmentTotalBytes) {
        info.data_base64 = "";
        info.skipped_reason = "total_cap";
      } else {
        info.data_base64 = Buffer.from(bytes).toString("base64");
        totalBytes += bytes.length;
      }
      return info;
    });
    const out: MailDetail = {
      ...this.summary(mail),
      body,
      body_truncated: bodyTruncated,
      recipients: this.recipientsOf(mail),
      attachments,
    };
    if (args.includeHtml) {
      out.html = mail.html ?? `<p>${mail.body}</p>`;
      out.html_truncated = false;
    }
    if (args.includeTransportHeaders) {
      const full = mail.transport_headers ?? "";
      let cut = full.length > MAX_TRANSPORT_HEADER_CHARS ? MAX_TRANSPORT_HEADER_CHARS : full.length;
      // Do not split a surrogate pair at the cap.
      const last = full.charCodeAt(cut - 1);
      if (cut > 0 && last >= 0xd800 && last <= 0xdbff) cut -= 1;
      out.transport_headers = full.slice(0, cut);
      out.transport_headers_truncated = cut < full.length;
    }
    return out;
  }

  async downloadAttachment(args: {
    entryId: string;
    index: number;
    destDir?: string | null;
  }): Promise<{ saved_to: string; size: number | null }> {
    await this.ready();
    const mail = this.locate(args.entryId);
    const attachment = mail.attachments[args.index - 1];
    if (!attachment) throw new ComError({ kind: "com", detail: `ATTACHMENT_INDEX_OUT_OF_RANGE: ${args.index}` });
    const dir = args.destDir ?? join(Bun.env.TMPDIR ?? "/tmp", "msoutlook-mcp-downloads");
    mkdirSync(dir, { recursive: true });
    let target = join(dir, attachment.filename);
    const dot = attachment.filename.lastIndexOf(".");
    const stem = dot > 0 ? attachment.filename.slice(0, dot) : attachment.filename;
    const ext = dot > 0 ? attachment.filename.slice(dot) : "";
    let n = 1;
    while (existsSync(target)) target = join(dir, `${stem}_${n++}${ext}`);
    writeFileSync(target, attachment.content ?? new TextEncoder().encode("fixture attachment"));
    return { saved_to: target, size: attachment.size };
  }

  async createDraft(args: CreateDraftArgs): Promise<DraftResult> {
    await this.ready();
    const id = entryId("draft", this.state);
    const mail: FakeMail = {
      entry_id: id,
      folder: args.from_mailbox ? `\\\\${args.from_mailbox}\\Drafts` : "drafts",
      received_at: new Date().toISOString(),
      sent_at: null,
      sender: this.state.user_smtp,
      sender_address: this.state.user_smtp,
      sender_smtp: this.state.user_smtp,
      to: args.to.join("; "),
      cc: args.cc.join("; "),
      bcc: args.bcc.join("; "),
      subject: args.subject,
      body: args.body,
      html: args.html ? args.body : undefined,
      is_read: true,
      importance: 1,
      internet_message_id: null,
      attachments: args.attachments.map((path) => ({ filename: path.split(/[\\/]/).pop() ?? path, size: 0 })),
    };
    this.state.mail.set(id, mail);
    return {
      entry_id: id,
      drafts_folder: args.from_mailbox ? `\\\\${args.from_mailbox}\\Drafts` : "drafts",
      attachments: args.attachments,
    };
  }

  async getRecipients(args: { entryId: string }): Promise<MailRecipient[]> {
    await this.ready();
    return this.recipientsOf(this.locate(args.entryId));
  }

  async sendDraft(args: { entryId: string }): Promise<{ ok: true }> {
    await this.ready();
    const mail = this.locate(args.entryId);
    mail.folder = "sent";
    mail.sent_at = new Date().toISOString();
    mail.received_at = mail.sent_at;
    return { ok: true };
  }

  async prepareReply(args: PrepareReplyArgs): Promise<DraftResult & { recipients: MailRecipient[] }> {
    await this.ready();
    const original = this.locate(args.entry_id);
    const id = entryId("reply", this.state);
    const to = [original.sender_smtp];
    const cc = args.reply_all ? splitAddresses(original.cc) : [];
    const reply: FakeMail = {
      entry_id: id,
      folder: args.from_mailbox ? `\\\\${args.from_mailbox}\\Drafts` : "drafts",
      received_at: new Date().toISOString(),
      sent_at: null,
      sender: this.state.user_smtp,
      sender_address: this.state.user_smtp,
      sender_smtp: this.state.user_smtp,
      to: to.join("; "),
      cc: cc.join("; "),
      subject: original.subject.startsWith("RE:") ? original.subject : `RE: ${original.subject}`,
      body: args.body,
      html: args.html ? args.body : undefined,
      is_read: true,
      importance: 1,
      internet_message_id: null,
      attachments: [],
    };
    this.state.mail.set(id, reply);
    return {
      entry_id: id,
      drafts_folder: reply.folder,
      attachments: [],
      recipients: this.recipientsOf(reply),
    };
  }

  async moveMail(args: {
    entryId: string;
    targetFolder: string;
    mailbox?: string | null;
  }): Promise<{ new_entry_id: string | null; target_folder: string | null }> {
    await this.ready();
    const mail = this.locate(args.entryId);
    const oldId = mail.entry_id;
    this.state.mail.delete(oldId);
    mail.entry_id = entryId("moved", this.state);
    mail.folder = args.targetFolder;
    this.state.mail.set(mail.entry_id, mail);
    return { new_entry_id: mail.entry_id, target_folder: args.targetFolder };
  }

  async markRead(args: { entryId: string; read: boolean }): Promise<{ ok: true }> {
    await this.ready();
    this.locate(args.entryId).is_read = args.read;
    return { ok: true };
  }

  async flagMail(args: { entryId: string; flag: boolean }): Promise<{ ok: true }> {
    await this.ready();
    this.locate(args.entryId).flagged = args.flag;
    return { ok: true };
  }

  async listEvents(args: {
    daysAhead: number;
    daysBack: number;
    limit: number;
    calendar?: string | null;
  }): Promise<EventSummary[]> {
    await this.ready();
    const start = Date.now() - args.daysBack * 86_400_000;
    const end = Date.now() + args.daysAhead * 86_400_000;
    return [...this.state.events.values()]
      .filter((event) => {
        const at = new Date(event.start).getTime();
        return at >= start && at <= end;
      })
      .sort((a, b) => a.start.localeCompare(b.start))
      .slice(0, args.limit)
      .map((event) => this.eventSummary(event));
  }

  private eventSummary(event: FakeEvent): EventSummary {
    return {
      entry_id: event.entry_id,
      subject: event.subject,
      start: event.start,
      end: event.end,
      duration_minutes: Math.round((new Date(event.end).getTime() - new Date(event.start).getTime()) / 60_000),
      location: event.location,
      organizer: event.organizer,
      is_recurring: event.is_recurring ?? false,
      is_all_day: event.is_all_day ?? false,
      meeting_status: event.meeting_status ?? 0,
      response_status: event.response_status ?? 0,
      required: event.required,
      optional: event.optional,
    };
  }

  async readEvent(args: { entryId: string; maxBodyBytes: number }): Promise<EventDetail> {
    await this.ready();
    const event = this.locateEvent(args.entryId);
    return {
      ...this.eventSummary(event),
      body: event.body,
      body_truncated: false,
      recipients: splitAddresses(event.required)
        .map((address) => ({ name: address, address, smtp: address, type: 1 }))
        .concat(
          splitAddresses(event.optional).map((address) => ({ name: address, address, smtp: address, type: 2 })),
        ),
    };
  }

  async selfBusy(args: { startIso: string; daysAhead: number }): Promise<BusyInterval[]> {
    await this.ready();
    const start = new Date(args.startIso).getTime();
    const end = start + args.daysAhead * 86_400_000;
    return [...this.state.events.values()]
      .filter((event) => !event.is_all_day && (event.meeting_status ?? 0) !== 5)
      .filter((event) => new Date(event.end).getTime() >= start && new Date(event.start).getTime() <= end)
      .map((event) => ({ start: event.start, end: event.end }));
  }

  async attendeeBusy(args: {
    smtp: string;
    startIso: string;
    slotMinutes: number;
  }): Promise<{ busy_string: string | null; error: string | null }> {
    await this.ready();
    const busy = this.state.attendee_busy.get(args.smtp);
    if (busy === undefined) return { busy_string: null, error: "unresolved (not found in the GAL)" };
    return { busy_string: busy, error: null };
  }

  async createEventDraft(args: CreateEventArgs): Promise<CreateEventResult> {
    await this.ready();
    const id = entryId("event-draft", this.state);
    const start = new Date(args.start_iso);
    const event: FakeEvent = {
      entry_id: id,
      subject: args.subject,
      start: start.toISOString(),
      end: new Date(start.getTime() + args.duration_minutes * 60_000).toISOString(),
      location: args.location ?? "",
      organizer: this.state.user_smtp,
      body: args.body ?? "",
      required: args.required_attendees.join("; "),
      optional: args.optional_attendees.join("; "),
      meeting_status: args.required_attendees.length + args.optional_attendees.length > 0 ? 1 : 0,
      invite_sent: false,
    };
    this.state.events.set(id, event);
    return {
      entry_id: id,
      subject: args.subject,
      start: start.toISOString(),
      duration_minutes: args.duration_minutes,
      required: args.required_attendees,
      optional: args.optional_attendees,
      is_meeting: event.meeting_status === 1,
    };
  }

  async getEventRecipients(args: { entryId: string }): Promise<MailRecipient[]> {
    await this.ready();
    const event = this.locateEvent(args.entryId);
    return splitAddresses(event.required)
      .map((address) => ({ name: address, address, smtp: address, type: 1 }))
      .concat(
        splitAddresses(event.optional).map((address) => ({ name: address, address, smtp: address, type: 2 })),
      );
  }

  async sendMeetingInvites(args: { entryId: string }): Promise<{ ok: true }> {
    await this.ready();
    const event = this.locateEvent(args.entryId);
    if ((event.meeting_status ?? 0) === 0) throw new ComError({ kind: "com", detail: "NOT_A_MEETING" });
    event.invite_sent = true;
    return { ok: true };
  }

  async respondMeeting(args: {
    entryId: string;
    response: "accepted" | "tentative" | "declined";
    sendResponse: boolean;
    comment?: string | null;
  }): Promise<{ ok: true; action: "sent" | "saved" }> {
    await this.ready();
    const event = this.locateEvent(args.entryId);
    event.response_status = args.response === "accepted" ? 3 : args.response === "tentative" ? 2 : 4;
    return { ok: true, action: args.sendResponse ? "sent" : "saved" };
  }

  async listRules(): Promise<RuleSummary[]> {
    await this.ready();
    return this.state.rules.map((rule) => ({ ...rule, conditions: [...rule.conditions], actions: [...rule.actions] }));
  }

  async toggleRule(args: { name: string; enabled: boolean }): Promise<{ ok: true }> {
    await this.ready();
    const rule = this.state.rules.find((r) => r.name === args.name);
    if (!rule) throw new ComError({ kind: "com", detail: `RULE_NOT_FOUND: ${args.name}` });
    rule.enabled = args.enabled;
    return { ok: true };
  }

  async createRule(args: CreateRuleArgs): Promise<{ rule: RuleSummary }> {
    await this.ready();
    if (this.state.rules.some((r) => r.name === args.name)) {
      throw new ComError({ kind: "com", detail: `RULE_NAME_EXISTS: ${args.name}` });
    }
    const conditions: string[] = [];
    if (args.from_addresses.length) conditions.push(`From IN [${args.from_addresses.join(", ")}]`);
    if (args.sent_to.length) conditions.push(`SentTo IN [${args.sent_to.join(", ")}]`);
    if (args.subject_contains.length) conditions.push(`Subject CONTAINS ${args.subject_contains.join(", ")}`);
    if (args.body_or_subject_contains.length)
      conditions.push(`BodyOrSubject CONTAINS ${args.body_or_subject_contains.join(", ")}`);
    if (args.has_attachment) conditions.push("HasAttachment");
    if (args.importance_high) conditions.push("Importance = 2");
    const actions: string[] = [];
    if (args.move_to_folder) actions.push(`Move -> ${args.move_to_folder}`);
    if (args.copy_to_folder) actions.push(`Copy -> ${args.copy_to_folder}`);
    if (args.assign_categories.length) actions.push(`AssignToCategory ${args.assign_categories.join(", ")}`);
    if (args.forward_to.length) actions.push(`Forward -> [${args.forward_to.join(", ")}]`);
    if (args.stop_processing) actions.push("StopProcessingMoreRules");
    const rule: FakeRule = {
      name: args.name,
      enabled: args.enabled,
      execution_order: this.state.rules.length + 1,
      is_local_rule: false,
      rule_type: 0,
      conditions,
      actions,
    };
    this.state.rules.push(rule);
    return { rule };
  }

  async deleteRule(args: { name: string }): Promise<{ ok: true }> {
    await this.ready();
    const index = this.state.rules.findIndex((r) => r.name === args.name);
    if (index < 0) throw new ComError({ kind: "com", detail: `RULE_NOT_FOUND: ${args.name}` });
    this.state.rules.splice(index, 1);
    return { ok: true };
  }

  async searchContacts(args: { query: string; limit: number }): Promise<{
    contacts: ContactSummary[];
    hint?: string | null;
  }> {
    await this.ready();
    const q = args.query.toLowerCase();
    const contacts = this.state.contacts
      .filter(
        (c) =>
          (c.name ?? "").toLowerCase().includes(q) ||
          (c.smtp_address ?? "").toLowerCase().includes(q) ||
          (c.alias ?? "").toLowerCase().includes(q),
      )
      .slice(0, args.limit)
      .map((c) => ({ ...c }));
    return { contacts, hint: null };
  }

  async getContactDetails(args: { smtpAddress: string }): Promise<ContactDetail> {
    await this.ready();
    const contact = this.state.contacts.find(
      (c) => (c.smtp_address ?? "").toLowerCase() === args.smtpAddress.toLowerCase(),
    );
    if (!contact) throw new ComError({ kind: "com", detail: `CONTACT_NOT_FOUND: ${args.smtpAddress}` });
    return { ...contact };
  }

  reset(): void {
    this.state.resets += 1;
  }

  close(): void {
    this.state.running = false;
  }
}

/** Convenience for tests that need a quick failure. */
export function failWithTimeout(message = "test timeout"): ComError {
  return timeoutError(message);
}

export function failWithDisconnect(message = "test disconnect"): ComError {
  return connectionLost(message);
}

export function senderDomain(mail: FakeMail): string {
  return domainOf(mail.sender_smtp);
}
