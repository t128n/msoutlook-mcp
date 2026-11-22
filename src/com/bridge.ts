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
  ReadMailArgs,
  RuleSummary,
  SearchMailArgs,
} from "./types.ts";

/**
 * The only surface the rest of the server knows.
 *
 * The real implementation is a persistent Windows PowerShell process; the fake
 * implementation is used by every test that does not need COM. Keep COM object
 * lifetime inside the implementation: nothing here returns a live COM handle.
 */
export interface ComBridge {
  /** Connects through GetActiveObject. With `launch` it may start Outlook (off by default). */
  ping(options?: { launch?: boolean; timeoutMs?: number }): Promise<PingResult>;

  /** Drops every cached COM object. The next call reconnects from scratch. */
  reset(): void;

  listFolders(args: { maxDepth: number; mailbox?: string | null }): Promise<FolderNode[]>;

  listMail(args: ListMailArgs): Promise<ListResult<MailSummary>>;
  searchMail(args: SearchMailArgs): Promise<ListResult<MailSummary>>;
  readMail(args: ReadMailArgs): Promise<MailDetail>;
  downloadAttachment(args: {
    entryId: string;
    index: number;
    destDir?: string | null;
  }): Promise<{ saved_to: string; size: number | null }>;

  createDraft(args: CreateDraftArgs): Promise<DraftResult>;
  getRecipients(args: { entryId: string }): Promise<MailRecipient[]>;
  sendDraft(args: { entryId: string }): Promise<{ ok: true }>;
  prepareReply(args: PrepareReplyArgs): Promise<DraftResult & { recipients: MailRecipient[] }>;
  moveMail(args: {
    entryId: string;
    targetFolder: string;
    mailbox?: string | null;
  }): Promise<{ new_entry_id: string | null; target_folder: string | null }>;
  markRead(args: { entryId: string; read: boolean }): Promise<{ ok: true }>;
  flagMail(args: { entryId: string; flag: boolean }): Promise<{ ok: true }>;

  listEvents(args: {
    daysAhead: number;
    daysBack: number;
    limit: number;
    calendar?: string | null;
  }): Promise<EventSummary[]>;
  readEvent(args: { entryId: string; maxBodyBytes: number }): Promise<EventDetail>;
  selfBusy(args: { startIso: string; daysAhead: number }): Promise<BusyInterval[]>;
  attendeeBusy(args: {
    smtp: string;
    startIso: string;
    slotMinutes: number;
  }): Promise<{ busy_string: string | null; error: string | null }>;
  createEventDraft(args: CreateEventArgs): Promise<CreateEventResult>;
  getEventRecipients(args: { entryId: string }): Promise<MailRecipient[]>;
  sendMeetingInvites(args: { entryId: string }): Promise<{ ok: true }>;
  respondMeeting(args: {
    entryId: string;
    response: "accepted" | "tentative" | "declined";
    sendResponse: boolean;
    comment?: string | null;
  }): Promise<{ ok: true; action: "sent" | "saved" }>;

  listRules(): Promise<RuleSummary[]>;
  toggleRule(args: { name: string; enabled: boolean }): Promise<{ ok: true }>;
  createRule(args: CreateRuleArgs): Promise<{ rule: RuleSummary }>;
  deleteRule(args: { name: string }): Promise<{ ok: true }>;

  searchContacts(args: { query: string; limit: number }): Promise<{
    contacts: ContactSummary[];
    hint?: string | null;
  }>;
  getContactDetails(args: { smtpAddress: string }): Promise<ContactDetail>;

  close(): void;
}
