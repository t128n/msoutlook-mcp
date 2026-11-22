/** Default and maximum attachment payload sizes, shared by the tool and the fake. */
export const DEFAULT_MAX_ATTACHMENT_BYTES = 2 * 1024 * 1024;
export const DEFAULT_MAX_ATTACHMENT_TOTAL_BYTES = 8 * 1024 * 1024;
export const MAX_ATTACHMENT_CAP = 64 * 1024 * 1024;

/** Transport header block cap. The host enforces UTF-8 bytes; the helper caps characters. */
export const MAX_TRANSPORT_HEADER_BYTES = 64 * 1024;
