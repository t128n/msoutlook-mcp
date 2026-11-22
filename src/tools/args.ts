import { ToolArgumentError } from "./registry.ts";

export function requireString(args: Record<string, unknown>, key: string): string {
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") {
    throw new ToolArgumentError(`${key} must be a non-empty string`);
  }
  return value.trim();
}

export function optionalString(args: Record<string, unknown>, key: string): string | undefined {
  const value = args[key];
  if (value === undefined || value === null) return undefined;
  if (typeof value !== "string") throw new ToolArgumentError(`${key} must be a string`);
  return value.trim() || undefined;
}

export function integer(
  args: Record<string, unknown>,
  key: string,
  fallback: number,
  min: number,
  max: number,
): number {
  const value = args[key];
  if (value === undefined || value === null || value === "") return fallback;
  const parsed = typeof value === "number" ? value : Number(value);
  if (!Number.isInteger(parsed)) throw new ToolArgumentError(`${key} must be an integer`);
  if (parsed < min || parsed > max) throw new ToolArgumentError(`${key} must be between ${min} and ${max}`);
  return parsed;
}

export function boolean(args: Record<string, unknown>, key: string, fallback: boolean): boolean {
  const value = args[key];
  if (value === undefined || value === null) return fallback;
  if (typeof value === "boolean") return value;
  if (value === "true" || value === "1" || value === "on") return true;
  if (value === "false" || value === "0" || value === "off") return false;
  throw new ToolArgumentError(`${key} must be true or false`);
}

export function stringList(args: Record<string, unknown>, key: string): string[] {
  const value = args[key];
  if (value === undefined || value === null || value === "") return [];
  if (Array.isArray(value)) {
    if (value.some((v) => typeof v !== "string")) throw new ToolArgumentError(`${key} must be a list of strings`);
    return (value as string[]).map((v) => v.trim()).filter(Boolean);
  }
  if (typeof value !== "string") throw new ToolArgumentError(`${key} must be a list of strings`);
  return value
    .split(/[,;]/)
    .map((v) => v.trim())
    .filter(Boolean);
}

export function oneOf<T extends string>(
  args: Record<string, unknown>,
  key: string,
  allowed: readonly T[],
  fallback: T,
): T {
  const value = args[key];
  if (value === undefined || value === null || value === "") return fallback;
  if (typeof value !== "string" || !allowed.includes(value as T)) {
    throw new ToolArgumentError(`${key} must be one of: ${allowed.join(", ")}`);
  }
  return value as T;
}

export function requireConfirm(args: Record<string, unknown>): void {
  if (args.confirm !== true) {
    throw new ToolArgumentError("confirm must be true for this state-changing operation");
  }
}

export function schemaObject(
  properties: Record<string, unknown>,
  required: string[] = [],
): Record<string, unknown> {
  return { type: "object", properties, required };
}

export const stringProp = (description: string, extra: Record<string, unknown> = {}): Record<string, unknown> => ({
  type: "string",
  description,
  ...extra,
});

export const numberProp = (description: string, extra: Record<string, unknown> = {}): Record<string, unknown> => ({
  type: "integer",
  description,
  ...extra,
});

export const booleanProp = (description: string, extra: Record<string, unknown> = {}): Record<string, unknown> => ({
  type: "boolean",
  description,
  ...extra,
});

export const arrayProp = (description: string, extra: Record<string, unknown> = {}): Record<string, unknown> => ({
  type: "array",
  items: { type: "string" },
  description,
  ...extra,
});
