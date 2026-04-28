# cowork-gateway

Run **Claude for Office (Excel)** and **Claude Desktop 3p** against any
Anthropic-compatible gateway — 9router, litellm, openrouter — without
fighting the three quirks that block this setup out of the box.

## What it solves

| Problem | Fix |
|---|---|
| Office add-ins refuse plain HTTP (mixed content) | Single-process Node HTTPS server with locally-trusted self-signed cert |
| Excel only accepts model names matching `claude*` | Wizard step where you pick exactly which `claude-*` aliases to expose |
| 9router's Windsurf-emulating `cc/` provider mangles tool names with `_ide` suffix → `UnknownToolError` | Streaming response rewrite strips the suffix transparently |

## Install

```bash
npx github:dmdfami/cowork-gateway init
```

You'll be asked for three things:

1. **Gateway base URL** — e.g. `http://127.0.0.1:20128/v1` (your 9router)
2. **API key** — the `sk-...` token from your gateway
3. **Model IDs** — comma-separated, e.g. `claude-sonnet-4-6,claude-opus-4-7`

That's it. The wizard will:

- Generate a self-signed cert for `127.0.0.1`
- Trust it in your login keychain
- Write a launchd plist so the proxy auto-starts at login
- Print the exact settings to paste into Excel and Claude Desktop

## Use in Excel

After `init`, open **Claude for Office** → Settings → Configure Gateway:

```
URL:        https://127.0.0.1:20443/v1
Token:      sk-...
AuthHeader: x-api-key
APIFormat:  anthropic
Model:      claude-sonnet-4-6
```

Apply, restart Excel, ask Claude to build a spreadsheet — tools work.

## Use in Claude Desktop 3p

Settings → Configure third-party inference → Gateway:

```
Base URL:     https://127.0.0.1:20443/v1
API key:      sk-...
Auth scheme:  bearer
```

## Manage

```bash
cowork-gateway start
cowork-gateway stop
cowork-gateway status
cowork-gateway uninstall
```

Logs at `/tmp/cowork-gateway.log`.

## Architecture

```
Excel ──HTTPS──▶  cowork-gateway (Node, :20443)  ──HTTP──▶  upstream gateway
                  ↳ self-signed TLS termination
                  ↳ strips `_ide` suffix from tool_use names in response stream
```

Single Node process, zero npm dependencies, ~300 LOC.

## Why does `_ide` get added in the first place?

9router's `cc/` provider authenticates against Anthropic by impersonating
the Cascade/Windsurf IDE. Cascade has a fixed whitelist of 20 tool names
(`view_file`, `run_command`, etc.); any tool not on the list gets `_ide`
appended outbound so upstream auth accepts it. 9router never reverses the
mangle on the response, so clients see `foo_ide` and reject it. This
gateway un-mangles the response so Excel/Claude Desktop see the original
`foo` they sent.

## Compatibility

- macOS 12+ (Apple Silicon or Intel)
- Node 18+
- System `openssl` (preinstalled on macOS)

Linux/Windows support — PRs welcome.

## License

MIT
