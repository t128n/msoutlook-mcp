import { main } from "./cli.ts";

if (import.meta.main) {
  const code = await main(process.argv.slice(2));
  process.exit(code);
}
