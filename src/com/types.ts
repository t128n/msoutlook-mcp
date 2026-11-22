/**
 * Public types of the COM bridge.
 *
 * Field names deliberately match the JSON produced by the original Python
 * project (Astral0/outlook-com-mcp) so the ported tools keep the same output
 * shape. The original is MIT licensed; see NOTICE.
 */

export interface PingResult {
  version: string | null;
  profile: string | null;
  default_store: string | null;
  user_smtp: string | null;
  inbox_count: number | null;
  inbox_unread: number | null;
  calendar_count: number | null;
  accounts: { display: string | null; smtp: string | null }[];
}

export interface FolderNode {
  name: string | null;
  path: string | null;
  item_count: number;
  unread_count: number;
  children: FolderNode[];
  store?: string | null;
}

export interface MailSummary {
  entry_id: string;
  received_at: string | null;
  sent_at: string | null;
  sender: string | null;
  sender_address: string | null;
  sender_smtp: string | null;
  to: string | null;
  cc: string | null;
  subject: string;
  is_read: boolean;
  has_attachments: boolean;
  size_bytes: number | null;
  importance: number | null;
  folder: string | null;
  internet_message_id?: string | null;
}

export interface MailRecipient {
  name: string | null;
  address: string | null;
  smtp: string | null;
  type: number | null;
}

export interface AttachmentInfo {
  index: number;
  filename: string | null;
  size: number | null;
}

export interface MailDetail extends MailSummary {
  body: string;
  body_truncated: boolean;
  html?: string;
  html_truncated?: boolean;
  recipients: MailRecipient[];
  attachments: AttachmentInfo[];
}

export interface ListMailArgs {
  folder?: string;
  mailbox?: string | null;
  limit: number;
  /** DASL clause without the leading `@SQL=`. Built in TypeScript. */
  filter?: string;
  /** Structured copies for the fake bridge and for callers that do not filter. */
  unread_only?: boolean;
  since_iso?: string | null;
}

export interface SearchMailArgs {
  folder?: string;
  mailbox?: string | null;
  limit: number;
  query: string;
  scope: "subject_body" | "subject" | "sender" | "all";
  filter: string;
}

export interface CreateDraftArgs {
  to: string[];
  cc: string[];
  bcc: string[];
  subject: string;
  body: string;
  html: boolean;
  attachments: string[];
  from_mailbox?: string | null;
}

export interface PrepareReplyArgs {
  entry_id: string;
  body: string;
  reply_all: boolean;
  html: boolean;
  from_mailbox?: string | null;
}

export interface DraftResult {
  entry_id: string;
  drafts_folder: string | null;
  attachments: string[];
}

export interface EventSummary {
  entry_id: string;
  subject: string;
  start: string | null;
  end: string | null;
  duration_minutes: number | null;
  location: string;
  organizer: string | null;
  is_recurring: boolean;
  is_all_day: boolean;
  meeting_status: number | null;
  response_status: number | null;
  required: string;
  optional: string;
}

export interface EventDetail extends EventSummary {
  body: string;
  body_truncated: boolean;
  recipients: MailRecipient[];
}

export interface BusyInterval {
  start: string;
  end: string;
}

export interface CreateEventArgs {
  subject: string;
  start_iso: string;
  duration_minutes: number;
  body?: string | null;
  location?: string | null;
  required_attendees: string[];
  optional_attendees: string[];
}

export interface CreateEventResult {
  entry_id: string;
  subject: string;
  start: string;
  duration_minutes: number;
  required: string[];
  optional: string[];
  is_meeting: boolean;
}

export interface RuleSummary {
  name: string | null;
  enabled: boolean;
  execution_order: number | null;
  is_local_rule: boolean;
  rule_type: number | null;
  conditions: string[];
  actions: string[];
}

export interface CreateRuleArgs {
  name: string;
  from_addresses: string[];
  subject_contains: string[];
  body_or_subject_contains: string[];
  sent_to: string[];
  has_attachment: boolean;
  importance_high: boolean;
  move_to_folder?: string | null;
  copy_to_folder?: string | null;
  assign_categories: string[];
  forward_to: string[];
  stop_processing: boolean;
  enabled: boolean;
  apply_now: boolean;
}

export interface ContactSummary {
  name: string | null;
  smtp_address: string | null;
  alias?: string | null;
  job_title?: string | null;
  department?: string | null;
  office_location?: string | null;
  company_name?: string | null;
  business_phone?: string | null;
  mobile_phone?: string | null;
}

export interface ContactDetail extends ContactSummary {
  street_address?: string | null;
  city?: string | null;
  state?: string | null;
  postal_code?: string | null;
  manager?: { name: string | null; smtp_address: string | null } | null;
}

export interface ListResult<T> {
  items: T[];
  truncated: boolean;
  folder?: string | null;
  mailbox_used?: string | null;
}
