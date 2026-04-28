# cowork-gateway

Tiny **read-only diagnostic** that prints the right Excel / Claude
Desktop config for talking to your local **9router** through whatever
HTTPS proxy you have running (Caddy, mitmproxy, anything).

Replaces the manual "what port is the proxy on this time, what's my
9router API key again?" lookup. Scans, prints, done.

## Usage

```bash
npx github:dmdfami/cowork-gateway
```

Output:

```
=== Local gateway scan ===
  ✓ 9router      :20128  (http, reachable)
  ✓ HTTPS proxy  :20443  (caddy pid=89752)

=== Excel paste-ready config ===

  Cấu hình Anthropic Claude extension trong Excel
  ================================================

  Gateway URL:   https://127.0.0.1:20443/v1
  Token:         sk-...
  Auth header:   x-api-key
  API format:    anthropic
  Model:         cc/claude-sonnet-4-6
```

## Save to file

```bash
npx github:dmdfami/cowork-gateway --save
```

Writes the config block to `~/Desktop/excel-claude-config.txt` so you
can copy-paste it into Excel later without remembering port numbers.

## JSON mode (for scripts)

```bash
npx github:dmdfami/cowork-gateway --json
```

Returns `{nineRouter, httpsProxy, apiKey, models}` for programmatic use.

## What this tool does NOT do

- Does **not** install services, certs, or launchd plists.
- Does **not** modify Excel preferences or sideload add-ins.
- Does **not** start 9router or Caddy.
- Does **not** touch `~/.9router/db.json` (read-only).

It only reads and reports. Use your existing setup scripts (or set up
Caddy manually) to get an HTTPS proxy running. This tool will then
discover and print the config.

## Why?

Port numbers, API keys, and which proxy is in front of 9router this
session are easy to forget. Excel re-prompts for the gateway URL more
often than you'd like. This tool removes the "look up the values" step
to a one-liner you can run from any terminal.

## Compatibility

- macOS (uses `lsof` to scan listeners)
- Node 18+ (uses native `https.request` only — zero npm deps)
- Assumes 9router is installed at `~/.9router/`

## License

MIT
