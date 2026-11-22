import { calendarReadTools } from "./calendar-read.ts";
import { healthTools } from "./health.ts";
import { mailReadTools } from "./mail-read.ts";
import { ToolRegistry, type ToolContext } from "./registry.ts";

export * from "./registry.ts";
export { createTools };

/** Phase 1: read mail and calendar, health_check, whoami, guardrails_status. */
function createTools(ctx: ToolContext): ToolRegistry {
  return new ToolRegistry([...mailReadTools(ctx), ...healthTools(ctx), ...calendarReadTools(ctx)]);
}
