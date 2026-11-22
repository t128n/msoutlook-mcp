/**
 * A tiny process that speaks the same JSON-lines protocol as the real
 * PowerShell helper. The host-side protocol test spawns it, so the host logic
 * is exercised on any OS without COM.
 */
let buffer = "";

function reply(value: unknown): void {
  process.stdout.write(JSON.stringify(value) + "\n");
}

async function handle(request: Record<string, unknown>): Promise<void> {
  const id = request.id;
  const op = request.op;
  if (op === "ping") {
    reply({ id, ok: true, result: { version: "16.0.0", profile: "Stub", inbox_count: 0 } });
    return;
  }
  if (op === "echo") {
    reply({ id, ok: true, result: { value: request.value } });
    return;
  }
  if (op === "fail") {
    reply({ id, ok: false, error: { code: 0x800706ba, message: "stub connection lost" } });
    return;
  }
  if (op === "noise") {
    process.stdout.write("not json\n");
    reply({ id, ok: true, result: { clean: true } });
    return;
  }
  if (op === "slow") {
    await Bun.sleep(60);
    reply({ id, ok: true, result: { slow: true } });
    return;
  }
  if (op === "hang") {
    await new Promise(() => {});
    return;
  }
  if (op === "crash") {
    process.stderr.write("stub crashed on purpose");
    process.exit(1);
  }
  if (op === "quit") {
    reply({ id, ok: true, result: { bye: true } });
    setTimeout(() => process.exit(0), 10);
    return;
  }
  reply({ id, ok: false, error: { code: 0x80004005, message: `unknown op ${String(op)}` } });
}

const decoder = new TextDecoder();
for await (const chunk of Bun.stdin.stream()) {
  buffer += decoder.decode(chunk, { stream: true });
  let newline = buffer.indexOf("\n");
  while (newline >= 0) {
    const line = buffer.slice(0, newline).trim();
    buffer = buffer.slice(newline + 1);
    if (line) {
      try {
        await handle(JSON.parse(line) as Record<string, unknown>);
      } catch {
        process.stdout.write("stub could not parse a line\n");
      }
    }
    newline = buffer.indexOf("\n");
  }
}
