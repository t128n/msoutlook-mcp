# Releases

A release is built only when a human pushes a version tag. Pushing a tag is the
switch; no schedule, no automatic release from `main`.

## Cut a release

```sh
git checkout main
git pull
git tag -a v0.1.0 -m "msoutlook-mcp 0.1.0"
git push origin v0.1.0
```

`.github/workflows/release.yml` then:

1. installs Bun 1.4.4,
2. runs `bun test` and `bunx tsc --noEmit`,
3. runs `scripts/package-release.sh`,
4. creates or updates the GitHub release named after the tag and uploads every asset.

If `docs/release-notes/<version>.md` exists it is used as the body; otherwise
GitHub generates notes from the merged PRs. Add notes before tagging.

## Assets

| Asset | Platform |
|---|---|
| `msoutlook-mcp-x86_64-pc-windows-msvc.zip` | Windows 10/11 x64 — the supported target |
| `msoutlook-mcp-aarch64-apple-darwin.tar.gz` | macOS arm64 — development build |
| `msoutlook-mcp-x86_64-apple-darwin.tar.gz` | macOS x64 — development build |
| `msoutlook-mcp-x86_64-unknown-linux-gnu.tar.gz` | Linux x64 — development build |
| `SHA256SUMS` | SHA-256 for the archives |

Each archive contains the executable plus `README.md`, `LICENSE` and `NOTICE`.
The asset names carry the OS and architecture so mise's GitHub backend selects
the right one.

## mise

```sh
mise use -g github:t128n/msoutlook-mcp            # latest
mise use -g github:t128n/msoutlook-mcp@0.1.0      # pinned
mise x -- msoutlook-mcp version
```

A version is the tag without its leading `v`.

## Verify locally before tagging

```sh
bash scripts/package-release.sh
./dist/release-stage/darwin-arm64/msoutlook-mcp version   # adjust to the host
```

The script builds all four targets, writes the archives and `SHA256SUMS` to
`dist/release-assets/`.
