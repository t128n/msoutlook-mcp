import { describe, expect, test } from "bun:test";
import { ComError, classifyHresult, hresult } from "../src/com/errors.ts";

describe("HRESULT mapping", () => {
  test("classic Outlook not registered", () => {
    expect(classifyHresult(0x80040154)).toBe("not_installed");
  });

  test("Outlook not running", () => {
    expect(classifyHresult(0x800401e3)).toBe("not_running");
  });

  test("Outlook process went away", () => {
    expect(classifyHresult(0x800706ba)).toBe("connection_lost");
    expect(classifyHresult(0x80010108)).toBe("connection_lost");
    expect(classifyHresult(0x800a01a8)).toBe("connection_lost");
  });

  test("access denied and Object Model Guard", () => {
    expect(classifyHresult(0x80070005)).toBe("access_denied");
    expect(classifyHresult(0x80080005)).toBe("access_denied");
  });

  test("unknown codes stay generic", () => {
    expect(classifyHresult(0x80004005)).toBe("com");
    expect(classifyHresult(undefined)).toBe("com");
  });

  test("signed JSON numbers are normalised", () => {
    expect(classifyHresult(-2147023174)).toBe("connection_lost");
  });

  test("only connection-shaped errors are retryable", () => {
    expect(new ComError({ kind: "connection_lost" }).retryable).toBe(true);
    expect(new ComError({ kind: "not_running" }).retryable).toBe(true);
    expect(new ComError({ kind: "timeout" }).retryable).toBe(true);
    expect(new ComError({ kind: "access_denied" }).retryable).toBe(false);
    expect(new ComError({ kind: "not_installed" }).retryable).toBe(false);
  });

  test("messages say what to do without leaking internals", () => {
    const error = new ComError({ kind: "connection_lost", detail: "RPC server unavailable" });
    expect(error.message).toContain("Minimize Outlook");
    const json = error.toJSON();
    expect(json.code).toBe("OUTLOOK_UNAVAILABLE");
    expect(json.retryable).toBe(true);
    expect(json.message).toBe(error.message);
    expect("hresult" in json).toBe(false);
  });

  test("HRESULT formatting is stable", () => {
    expect(hresult(0x800706ba)).toBe("0x800706BA");
  });
});
