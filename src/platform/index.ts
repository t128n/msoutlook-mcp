import { basename } from "node:path";

export const isWindows = process.platform === "win32";

/**
 * `bun run src/main.ts` has the Bun binary as execPath; a compiled exe is
 * itself. `install` and `upgrade` must not copy Bun into the install folder.
 */
export function isCompiled(execPath = process.execPath): boolean {
  const name = basename(execPath).toLowerCase();
  return name !== "bun" && name !== "bun.exe";
}
