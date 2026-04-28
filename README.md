# excel-claude

One-shot setup for the **Anthropic Claude add-in in Excel** when running
against a local **9router** gateway. Replaces the manual "look up port,
look up API key, paste into Excel, trust the cert, hope it works" dance
with a single command.

## Usage

```bash
npx dmdfami/excel-claude
```

That's it. The tool will:

1. **Scan** — find 9router (`~/.9router/db.json`) and your local HTTPS
   proxy port automatically (Caddy / mitmproxy / anything on `:2044x`).
2. **Print** — emit the Excel-paste-ready config block.
3. **Trust cert** — if the proxy's root CA isn't in your login keychain
   yet, install it now (one-time macOS Touch ID / password prompt).
4. **Inject** — write the gateway URL + API key directly into Excel's
   `pivot.claude.ai` LocalStorage. On next Excel launch the Anthropic
   Claude add-in finds the gateway pre-filled.

If Excel is open when you run it, step 4 is skipped (LocalStorage is
locked); you'll see a friendly nudge to ⌘Q Excel and re-run, or just
paste the printed config manually.

## Flags

```
--no-inject / --print-only   Skip the LocalStorage write, just print config
--save                       Also write config to ~/Desktop/excel-claude-config.txt
--json                       Machine-readable JSON only (implies --no-inject)
```

## What this does NOT do

- Doesn't install services, generate certs, or run launchd plists.
- Doesn't start 9router or the HTTPS proxy — your existing setup must
  already serve the gateway. This tool just discovers and consumes it.
- Doesn't touch any settings other than the two Anthropic-controlled
  LocalStorage keys (`claude.inference.profile` and the matching
  `_OfficeRuntime_Storage_*` mirror).

## Why a tool exists

Three things forget themselves between sessions:

| | What you have to remember |
|---|---|
| HTTPS proxy port | `20443`? `20444`? changes per machine |
| 9router API key | `sk-b694a92...` 32 chars |
| Cert trust state | did you `security add-trusted-cert` on this Mac yet? |

The tool reads all three from local sources, then makes Excel itself
remember them too — by writing into the same LocalStorage slot the
Anthropic UI uses. Run it once on every new Mac you set up.

## Compatibility

- macOS (uses `lsof`, `security`, `/usr/bin/curl`, `/usr/bin/python3`)
- Node 18+ for the npx entry point — zero npm dependencies
- Assumes 9router installed at `~/.9router/`

## License

MIT
