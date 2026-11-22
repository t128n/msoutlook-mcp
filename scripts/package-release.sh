#!/usr/bin/env bash
# Build the release archives. Run by .github/workflows/release.yml on a tag,
# and locally to inspect the assets before tagging.
set -euo pipefail
cd "$(dirname "$0")/.."
ROOT="$(pwd)"

OUT="dist/release-assets"
STAGE="dist/release-stage"
rm -rf "$OUT" "$STAGE"
mkdir -p "$OUT" "$STAGE/windows" "$STAGE/darwin-arm64" "$STAGE/darwin-x64" "$STAGE/linux-x64"

echo "Building windows-x64"
bun build --compile --target=bun-windows-x64 --outfile "$STAGE/windows/msoutlook-mcp.exe" src/main.ts

echo "Building darwin-arm64"
bun build --compile --target=bun-darwin-arm64 --outfile "$STAGE/darwin-arm64/msoutlook-mcp" src/main.ts

echo "Building darwin-x64"
bun build --compile --target=bun-darwin-x64 --outfile "$STAGE/darwin-x64/msoutlook-mcp" src/main.ts

echo "Building linux-x64"
bun build --compile --target=bun-linux-x64 --outfile "$STAGE/linux-x64/msoutlook-mcp" src/main.ts

for dir in "$STAGE"/*; do
  cp README.md LICENSE NOTICE "$dir/"
done

echo "Packaging"
(cd "$STAGE/windows" && zip -q -r "$ROOT/$OUT/msoutlook-mcp-x86_64-pc-windows-msvc.zip" .)
tar -czf "$ROOT/$OUT/msoutlook-mcp-aarch64-apple-darwin.tar.gz" -C "$STAGE/darwin-arm64" .
tar -czf "$ROOT/$OUT/msoutlook-mcp-x86_64-apple-darwin.tar.gz" -C "$STAGE/darwin-x64" .
tar -czf "$ROOT/$OUT/msoutlook-mcp-x86_64-unknown-linux-gnu.tar.gz" -C "$STAGE/linux-x64" .

echo "Checksums"
(
  cd "$OUT"
  if command -v sha256sum >/dev/null 2>&1; then
    sha256sum ./*.zip ./*.tar.gz | sed "s#  \./#  #" > SHA256SUMS
  else
    shasum -a 256 ./*.zip ./*.tar.gz | sed "s#  \./#  #" > SHA256SUMS
  fi
)

ls -lh "$OUT"
