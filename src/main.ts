import { VERSION } from "./version.ts";

/** Placeholder entry point for the COM-bridge PR. The CLI lands in #3. */
function main(argv: string[]): number {
  const command = argv[0];
  if (command === "version" || command === "--version") {
    console.log(VERSION);
    return 0;
  }
  console.log(`msoutlook-mcp ${VERSION}`);
  console.log("Commands: version, --help");
  return 0;
}

if (import.meta.main) process.exit(main(process.argv.slice(2)));
