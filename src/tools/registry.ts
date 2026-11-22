import type { ComBridge } from "../com/bridge.ts";
import { ComError, errorText, hresult } from "../com/errors.ts";
import type { AppPaths } from "../config/paths.ts";
import type { Guardrails } from "../config/guardrails.ts";
import type { HealthMonitor } from "../server/health.ts";
import type { Logger } from "../util/logger.ts";

export type JsonSchema = Record<string, unknown>;

export interface ToolResponse {
  data: unknown;
  isError?: boolean;
}

export interface ToolDefinition {
  name: string;
  description: string;
  inputSchema: JsonSchema;
  handler: (args: Record<string, unknown>) => Promise<ToolResponse>;
}

export interface ToolContext {
  bridge: ComBridge;
  guardrails: Guardrails;
  health: HealthMonitor;
  paths: AppPaths;
  version: string;
  logger: Logger;
  now: () => Date;
}

export class ToolArgumentError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ToolArgumentError";
  }
}

export function ok(data: unknown): ToolResponse {
  return { data };
}

export function fail(code: string, message?: string, detail?: Record<string, unknown>): ToolResponse {
  return {
    data: { error: code, ...(message ? { message } : {}), ...(detail ?? {}) },
    isError: true,
  };
}

export function responseForError(e: unknown): ToolResponse {
  if (e instanceof ToolArgumentError) return fail("INVALID_ARGUMENT", e.message);
  if (e instanceof ComError) {
    const detail = e.detail ?? "";
    const known = /^([A-Z][A-Z0-9_]+):/.exec(detail)?.[1];
    return fail(known ?? "OUTLOOK_UNAVAILABLE", e.message, {
      retryable: e.retryable,
      ...(e.code !== undefined ? { hresult: hresult(e.code) } : {}),
      ...(known ? { detail } : {}),
    });
  }
  return fail("INTERNAL_ERROR", errorText(e));
}

export class ToolRegistry {
  private readonly tools = new Map<string, ToolDefinition>();

  constructor(definitions: ToolDefinition[]) {
    for (const tool of definitions) {
      if (this.tools.has(tool.name)) throw new Error(`duplicate tool ${tool.name}`);
      this.tools.set(tool.name, tool);
    }
  }

  list(): { name: string; description: string; inputSchema: JsonSchema }[] {
    return [...this.tools.values()].map((tool) => ({
      name: tool.name,
      description: tool.description,
      inputSchema: tool.inputSchema,
    }));
  }

  has(name: string): boolean {
    return this.tools.has(name);
  }

  async call(name: string, args: Record<string, unknown> = {}): Promise<ToolResponse> {
    const tool = this.tools.get(name);
    if (!tool) return responseForError(new ToolArgumentError(`unknown tool "${name}"`));
    try {
      return await tool.handler(args);
    } catch (e) {
      return responseForError(e);
    }
  }
}

export function tool(
  name: string,
  description: string,
  inputSchema: JsonSchema,
  handler: (args: Record<string, unknown>) => Promise<ToolResponse>,
): ToolDefinition {
  return { name, description, inputSchema, handler };
}
