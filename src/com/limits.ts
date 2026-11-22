/** Default and maximum attachment payload sizes, shared by the tool, fake and helper. */
export const DEFAULT_MAX_ATTACHMENT_BYTES = 2 * 1024 * 1024;
export const DEFAULT_MAX_ATTACHMENT_TOTAL_BYTES = 8 * 1024 * 1024;
export const MAX_ATTACHMENT_CAP = 64 * 1024 * 1024;

/**
 * Transport header block cap. The host enforces UTF-8 bytes; the helper and
 * fake cap UTF-16 characters before the block crosses the JSON-lines pipe.
 */
export const MAX_TRANSPORT_HEADER_BYTES = 64 * 1024;
export const MAX_TRANSPORT_HEADER_CHARS = 64 * 1024;
