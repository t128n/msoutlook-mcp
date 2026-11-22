/**
 * Tool handlers ported from Astral0/outlook-com-mcp (MIT). The original
 * server.py is the source of the tool inventory and behaviour; see NOTICE.
 */
import { calendarReadTools } from "./calendar-read.ts";
import { calendarWriteTools } from "./calendar-write.ts";
import { contactTools } from "./contacts.ts";
import { healthTools } from "./health.ts";
import { mailReadTools } from "./mail-read.ts";
import { mailWriteTools } from "./mail-write.ts";
import { ruleTools } from "./rules.ts";
import { ToolRegistry, type ToolContext } from "./registry.ts";

export * from "./registry.ts";
export { createTools };

/** All 27 tools: reads, guarded writes, rules and contacts. */
function createTools(ctx: ToolContext): ToolRegistry {
  return new ToolRegistry([
    ...mailReadTools(ctx),
    ...mailWriteTools(ctx),
    ...healthTools(ctx),
    ...calendarReadTools(ctx),
    ...calendarWriteTools(ctx),
    ...ruleTools(ctx),
    ...contactTools(ctx),
  ]);
}
