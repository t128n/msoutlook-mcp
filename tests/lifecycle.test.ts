import { describe, expect, test } from "bun:test";
import { ComError, connectionLost } from "../src/com/errors.ts";
import { FakeComBridge } from "../src/com/fake.ts";
import { ResilientBridge } from "../src/com/resilient.ts";
import { HealthMonitor } from "../src/server/health.ts";

function bridge(options: { autoLaunch?: boolean; initial?: number; max?: number } = {}) {
  const inner = new FakeComBridge();
  const events: { event: string; fields?: Record<string, unknown> }[] = [];
  const value = new ResilientBridge({
    bridge: inner,
    autoLaunchOutlook: options.autoLaunch ?? false,
    initialDelayMs: options.initial ?? 10,
    maxDelayMs: options.max ?? 40,
    pingTimeoutMs: 100,
    onLog: (event, fields) => events.push({ event, fields }),
  });
  return { inner, value, events };
}

describe("Outlook lifecycle", () => {
  test("starts connected and reports the state", async () => {
    const { value } = bridge();
    expect(await value.start()).toBe("connected");
    expect(value.state).toBe("connected");
    expect(value.snapshot().last_success_at).not.toBeNull();
  });

  test("a closed Outlook becomes waiting_for_outlook, not a crash", async () => {
    const { inner, value } = bridge();
    await value.start();
    inner.closeOutlook();
    await expect(value.listMail({ folder: "inbox", limit: 5 })).rejects.toBeInstanceOf(ComError);
    expect(value.state).toBe("waiting_for_outlook");
    expect(value.snapshot().last_error?.retryable).toBe(true);
  });

  test("tool calls while waiting fail fast with a retryable error", async () => {
    const { inner, value } = bridge();
    inner.closeOutlook();
    await value.start();
    expect(value.state).toBe("waiting_for_outlook");
    const started = Date.now();
    await expect(value.listMail({ folder: "inbox", limit: 5 })).rejects.toMatchObject({
      retryable: true,
    });
    expect(Date.now() - started).toBeLessThan(50);
    expect(value.snapshot().attempts).toBeGreaterThan(0);
  });

  test("reconnects with backoff and drops cached COM objects", async () => {
    const { inner, value, events } = bridge();
    inner.closeOutlook();
    await value.start();
    expect(value.state).toBe("waiting_for_outlook");
    expect(inner.state.resets).toBeGreaterThanOrEqual(1);
    const attemptsAfterStart = value.snapshot().attempts;
    expect(attemptsAfterStart).toBeGreaterThanOrEqual(1);

    inner.openOutlook();
    await Bun.sleep(80);
    expect(value.state).toBe("connected");
    expect(value.snapshot().attempts).toBe(0);
    const names = events.map((e) => e.event);
    expect(names).toContain("outlook_disconnected");
    expect(names).toContain("outlook_reconnect_scheduled");
    expect(names).toContain("outlook_connected");
  });

  test("the backoff delay grows but is capped", async () => {
    const { inner, value } = bridge({ initial: 10, max: 20 });
    inner.closeOutlook();
    await value.start();
    expect(value.snapshot().next_retry_ms).toBe(10);
    await Bun.sleep(15);
    expect(value.snapshot().next_retry_ms).toBe(20);
    // Still closed: the next attempt stays at the cap.
    await Bun.sleep(25);
    expect(["waiting_for_outlook", "connected"]).toContain(value.state);
    expect(value.snapshot().next_retry_ms === null || value.snapshot().next_retry_ms === 20).toBe(true);
  });

  test("reconnects without restarting the daemon", async () => {
    const { inner, value } = bridge({ initial: 5, max: 10 });
    inner.closeOutlook();
    await value.start();
    inner.openOutlook();
    await Bun.sleep(30);
    const ping = await value.ping();
    expect(ping.version).toBe("16.0.17928.20156");
    expect(value.state).toBe("connected");
  });

  test("does not launch Outlook unless explicitly enabled", async () => {
    const off = bridge();
    off.inner.closeOutlook();
    await off.value.start();
    expect(off.inner.state.last_ping_launch).toBe(false);

    const on = bridge({ autoLaunch: true });
    on.inner.closeOutlook();
    await on.value.start();
    expect(on.inner.state.last_ping_launch).toBe(true);
  });

  test("a non-retryable error stops the retry loop with state error", async () => {
    const { inner, value } = bridge({ initial: 5, max: 10 });
    inner.failNext = new ComError({ kind: "not_installed" });
    await value.start();
    expect(value.state).toBe("error");
    const attempts = value.snapshot().attempts;
    await Bun.sleep(20);
    expect(value.snapshot().attempts).toBe(attempts);
    expect(value.state).toBe("error");
  });

  test("a timeout is retryable and reported as waiting", async () => {
    const { inner, value } = bridge();
    await value.start();
    inner.failNext = new ComError({ kind: "timeout" });
    await expect(value.ping()).rejects.toMatchObject({ kind: "timeout" });
    expect(value.state).toBe("waiting_for_outlook");
  });

  test("in-flight calls fail cleanly when Outlook closes", async () => {
    const { inner, value } = bridge();
    await value.start();
    inner.state.latency_ms = 20;
    const call = value.listMail({ folder: "inbox", limit: 5 });
    inner.closeOutlook();
    await expect(call).rejects.toBeInstanceOf(ComError);
    expect(value.state).toBe("waiting_for_outlook");
  });

  test("health follows the wrapper and exposes the retry state", async () => {
    const { inner, value } = bridge();
    const health = new HealthMonitor({
      bridge: value,
      version: "0.1.0",
      cacheMs: 0,
      timeoutMs: 100,
      connection: () => value.snapshot(),
    });
    expect((await health.snapshot()).status).toBe("connected");
    inner.closeOutlook();
    const snapshot = await health.snapshot(true);
    expect(snapshot.status).toBe("waiting_for_outlook");
    expect(snapshot.outlook.retryable).toBe(true);
    expect(snapshot.outlook.reconnect_attempts).toBeGreaterThan(0);
    expect(snapshot.outlook.next_retry_at).not.toBeNull();
  });

  test("a successful call refreshes last_success_at", async () => {
    const { value } = bridge();
    await value.start();
    const before = value.snapshot().last_success_at;
    await Bun.sleep(5);
    await value.listMail({ folder: "inbox", limit: 1 });
    expect(value.snapshot().last_success_at).not.toBe(before);
  });

  test("the wrapper forwards every ComBridge operation", async () => {
    const { inner, value } = bridge();
    const mail = inner.seedMail({ subject: "one", folder: "inbox" });
    inner.seedEvent({ subject: "event", start: new Date(Date.now() + 3_600_000).toISOString() });
    inner.seedContact({ name: "Ada Lovelace", smtp_address: "ada@example.com" });
    await value.start();
    expect((await value.listMail({ folder: "inbox", limit: 5 })).items[0]!.subject).toBe("one");
    expect((await value.readMail({ entryId: mail.entry_id, maxBodyBytes: 0, includeHtml: false })).subject).toBe("one");
    expect((await value.listEvents({ daysAhead: 2, daysBack: 0, limit: 5 })).length).toBe(1);
    expect((await value.searchContacts({ query: "ada", limit: 5 })).contacts.length).toBe(1);
    expect(await value.listRules()).toEqual([]);
  });
});

describe("backoff edge cases", () => {
  test("connectionLost is retryable", () => {
    expect(connectionLost("x").retryable).toBe(true);
  });
});
