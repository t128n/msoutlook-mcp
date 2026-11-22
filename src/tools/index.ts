import { calendarReadTools } from "./calendar-read.ts";
import { calendarWriteTools } from "./calendar-write.ts";
import { healthTools } from "./health.ts";
import { mailReadTools } from "./mail-read.ts";
import { mailWriteTools } from "./mail-write.ts";
import { ToolRegistry, type ToolContext } from "./registry.ts";

export * from "./registry.ts";
export { createTools };

/** Phases 1 and 2: read tools plus guarded drafts and sends. */
function createTools(ctx: ToolContext): ToolRegistry {
  return new ToolRegistry([
    ...mailReadTools(ctx),
    ...mailWriteTools(ctx),
    ...healthTools(ctx),
    ...calendarReadTools(ctx),
    ...calendarWriteTools(ctx),
  ]);
}
