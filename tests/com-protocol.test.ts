import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { PowerShellComBridge, asciiJson } from "../src/com/index.ts";
import { ComError } from "../src/com/errors.ts";

const stub = join(import.meta.dir, "fixtures", "com-stub.ts");

function bridge(options: { timeoutMs?: number } = {}): PowerShellComBridge {
  return new PowerShellComBridge({
    command: [process.execPath, stub],
    timeoutMs: options.timeoutMs ?? 2_000,
  });
}

describe("JSON-lines host", () => {
  test("sends requests and matches answers by id", async () => {
    const com = bridge();
    try {
      const ping = await com.ping();
      expect(ping.version).toBe("16.0.0");
      expect(ping.profile).toBe("Stub");
    } finally {
      com.close();
    }
  });

  test("keeps concurrent calls apart", async () => {
    const com = bridge();
    try {
      const [echo, slow] = await Promise.all([
        com["request"]<{ value: string }>("echo", { value: "hello" }, 2_000),
        com["request"]<{ slow: boolean }>("slow", {}, 2_000),
      ]);
      expect(echo.value).toBe("hello");
      expect(slow.slow).toBe(true);
    } finally {
      com.close();
    }
  });

  test("ignores non-JSON noise on stdout", async () => {
    const com = bridge();
    try {
      const result = await com["request"]<{ clean: boolean }>("noise", {}, 2_000);
      expect(result.clean).toBe(true);
    } finally {
      com.close();
    }
  });

  test("maps a COM HRESULT to a retryable connection error", async () => {
    const com = bridge();
    try {
      const failure = com["request"]("fail", {}, 2_000);
      await expect(failure).rejects.toBeInstanceOf(ComError);
      await failure.catch((error: ComError) => {
        expect(error.kind).toBe("connection_lost");
        expect(error.retryable).toBe(true);
        expect(error.code).toBe(0x800706ba);
      });
    } finally {
      com.close();
    }
  });

  test("a timeout kills the helper and the next call starts a fresh one", async () => {
    const com = bridge({ timeoutMs: 120 });
    try {
      const failure = com["request"]("hang", {}, 120);
      await expect(failure).rejects.toBeInstanceOf(ComError);
      await failure.catch((error: ComError) => {
        expect(error.kind).toBe("timeout");
        expect(error.retryable).toBe(true);
      });
      const ping = await com.ping();
      expect(ping.version).toBe("16.0.0");
    } finally {
      com.close();
    }
  });

  test("an exited helper fails pending calls", async () => {
    const com = bridge();
    try {
      const failure = com["request"]("crash", {}, 2_000);
      await expect(failure).rejects.toBeInstanceOf(ComError);
      await failure.catch((error: ComError) => {
        expect(["connection_lost", "not_running"]).toContain(error.kind);
      });
    } finally {
      com.close();
    }
  });

  test("non-ASCII values survive the ASCII-only protocol", async () => {
    const com = bridge();
    try {
      const value = "Grüße – 日本語 – 🚀";
      const result = await com["request"]<{ value: string }>("echo", { value }, 2_000);
      expect(result.value).toBe(value);
    } finally {
      com.close();
    }
  });

  test("requests on the wire are ASCII", async () => {
    const wire = asciiJson({ value: "Grüße 🚀" });
    expect(wire).toMatch(/^[\x00-\x7e]*$/);
    expect(JSON.parse(wire).value).toBe("Grüße 🚀");
  });
});
