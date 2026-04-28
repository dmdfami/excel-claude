#!/usr/bin/env node
/*
 * excel-claude — read-only diagnostic / config printer.
 *
 * Scans the local machine for:
 *   1. 9router (~/.9router/db.json + listener on :20128)
 *   2. An HTTPS proxy in front of 9router (typically :20443) that
 *      successfully forwards /v1/models with the 9router API key.
 *
 * Emits an Excel-paste-ready config block. Does NOT install services,
 * does NOT generate certs, does NOT touch Excel preferences. Pure read.
 *
 * Run: npx github:dmdfami/excel-claude
 *
 * Flags:
 *   --save     Also write config to ~/Desktop/excel-claude-config.txt
 *   --json     Emit machine-readable JSON instead of human output
 *   --inject   Write config directly into Excel's pivot.claude.ai LocalStorage
 *              (requires Excel to be quit). On next launch the Anthropic
 *              Claude add-in finds the gateway URL + API key pre-filled —
 *              no manual paste.
 */

const fs        = require('fs');
const os        = require('os');
const path      = require('path');
const http      = require('http');
const https     = require('https');
const { spawnSync } = require('child_process');

const HOME = os.homedir();
const NINER_DB = path.join(HOME, '.9router', 'db.json');
const SAVE_PATH = path.join(HOME, 'Desktop', 'excel-claude-config.txt');

const c = {
    cyan:   (s) => `\x1b[36m${s}\x1b[0m`,
    green:  (s) => `\x1b[32m${s}\x1b[0m`,
    yellow: (s) => `\x1b[33m${s}\x1b[0m`,
    red:    (s) => `\x1b[31m${s}\x1b[0m`,
    dim:    (s) => `\x1b[2m${s}\x1b[0m`,
    bold:   (s) => `\x1b[1m${s}\x1b[0m`,
};

const log = (...a) => console.log(c.cyan('[excel-claude]'), ...a);
const die = (m) => { console.error(c.red('[excel-claude] ERROR:'), m); process.exit(1); };

function probe({ host, port, scheme, path: p, apiKey, timeout = 3000, strictTls = false }) {
    return new Promise((resolve) => {
        const lib = scheme === 'https' ? https : http;
        const req = lib.request({
            host, port, path: p, method: 'GET',
            headers: apiKey ? { Authorization: `Bearer ${apiKey}` } : {},
            // strictTls=false: just confirm something answers (used for discovery)
            // strictTls=true: also require cert to chain to a trusted root
            //   (Excel/Claude WebView enforces this, our scan must too).
            rejectUnauthorized: strictTls,
            timeout,
        }, (res) => {
            let body = '';
            res.on('data', (d) => { body += d.toString(); if (body.length > 4096) body = body.slice(0, 4096); });
            res.on('end', () => resolve({ ok: res.statusCode === 200, status: res.statusCode, body }));
        });
        req.on('error', (e) => resolve({ ok: false, error: e.message, code: e.code }));
        req.on('timeout', () => { req.destroy(); resolve({ ok: false, error: 'timeout' }); });
        req.end();
    });
}

function listListeners(portRange) {
    const r = spawnSync('lsof', ['-nP', '-iTCP', '-sTCP:LISTEN'], { encoding: 'utf8' });
    if (r.status !== 0) return [];
    const lines = r.stdout.split('\n').slice(1).filter(Boolean);
    const out = [];
    for (const line of lines) {
        const cols = line.split(/\s+/);
        const cmd  = cols[0];
        const pid  = cols[1];
        const name = cols[cols.length - 2] || cols[cols.length - 1];
        const m = name.match(/:(\d+)$/);
        if (!m) continue;
        const port = Number(m[1]);
        if (port >= portRange[0] && port <= portRange[1]) {
            out.push({ cmd, pid: Number(pid), port });
        }
    }
    return out;
}

// Verify a cert is trusted by macOS keychain (matches what Excel/WebKit do).
// Use system curl on macOS — it uses Apple Secure Transport, which reads the
// keychain. Node's built-in TLS only knows its bundled CA list, so it would
// reject self-signed-but-keychain-trusted certs and produce false negatives.
function verifyCertViaSystemCurl(url, apiKey) {
    const r = spawnSync('/usr/bin/curl',
        ['-s', '-o', '/dev/null', '-w', '%{http_code}|%{ssl_verify_result}',
         '--max-time', '5', url,
         '-H', `Authorization: Bearer ${apiKey}`],
        { encoding: 'utf8' });
    if (r.status !== 0) return { trusted: false, error: r.stderr || 'curl error', http: null };
    const [http, sslVerify] = (r.stdout || '').split('|');
    // ssl_verify_result === '0' means the cert chain validated against the
    // system trust store. Anything else (or empty) means Excel will reject.
    return { trusted: http === '200' && sslVerify === '0', http, sslVerify };
}

async function findHttpsProxy(apiKey) {
    const candidates = [20443, 20444, 20445, 20446, 20447, 20448, 20449, 20450];
    const listeners = listListeners([20440, 20460]);
    for (const l of listeners) {
        if (!candidates.includes(l.port)) candidates.push(l.port);
    }
    for (const port of [...new Set(candidates)]) {
        // Discovery probe — does the port serve /v1/models with our key?
        const loose = await probe({ host: '127.0.0.1', port, scheme: 'https',
            path: '/v1/models', apiKey, strictTls: false });
        if (!loose.ok) continue;
        // Cert trust probe — system curl == Excel's view of trust.
        const verdict = verifyCertViaSystemCurl(`https://127.0.0.1:${port}/v1/models`, apiKey);
        return {
            port,
            listener: listeners.find((l) => l.port === port),
            certTrusted: verdict.trusted,
            certError:   verdict.trusted ? null : `http=${verdict.http} ssl_verify=${verdict.sslVerify}`,
        };
    }
    return null;
}

async function scan() {
    if (!fs.existsSync(NINER_DB)) {
        return { error: '9router not installed (~/.9router/db.json missing). Install with: npm install -g 9router' };
    }
    let db;
    try { db = JSON.parse(fs.readFileSync(NINER_DB, 'utf8')); }
    catch (e) { return { error: '9router db.json unreadable: ' + e.message }; }
    const apiKey = db?.apiKeys?.[0]?.key;
    if (!apiKey) return { error: '9router has no apiKey. Open dashboard http://127.0.0.1:20128/dashboard' };

    const niner   = await probe({ host: '127.0.0.1', port: 20128, scheme: 'http',  path: '/v1/models', apiKey });
    const httpsP  = await findHttpsProxy(apiKey);

    return {
        nineRouter: { port: 20128, alive: niner.ok, status: niner.status || niner.error },
        httpsProxy: httpsP
            ? {
                port: httpsP.port,
                listener: httpsP.listener,
                alive: true,
                certTrusted: httpsP.certTrusted,
                certError: httpsP.certError,
              }
            : { alive: false, hint: 'No HTTPS proxy reaching 9router. Excel/Office add-ins need HTTPS — start your Caddy/cowork proxy first.' },
        apiKey,
        models: ['cc/claude-sonnet-4-6', 'cc/claude-opus-4-7', 'cc/claude-haiku-4-5-20251001'],
    };
}

function buildConfigText(s) {
    if (!s.httpsProxy?.alive) return null;
    const url = `https://127.0.0.1:${s.httpsProxy.port}/v1`;
    return `Cấu hình Anthropic Claude extension trong Excel
================================================

Gateway URL:   ${url}
Token:         ${s.apiKey}
Auth header:   x-api-key
API format:    anthropic
Model:         ${s.models[0]}

Other models you can pick:
  ${s.models.slice(1).join('\n  ')}

(Port ${s.httpsProxy.port} = HTTPS proxy in front of 9router :${s.nineRouter.port};
Excel add-ins reject plain HTTP, that's why we don't use :${s.nineRouter.port} directly.)
`;
}

function printHuman(s) {
    if (s.error) {
        console.error(c.red('✗ ' + s.error));
        process.exit(1);
    }
    console.log(c.bold('=== Local gateway scan ==='));
    const ok9   = s.nineRouter.alive;
    const okHt  = s.httpsProxy.alive;
    console.log(`  ${ok9  ? c.green('✓') : c.red('✗')} 9router      :${s.nineRouter.port}  ${c.dim(`(http, ${ok9  ? 'reachable' : s.nineRouter.status})`)}`);
    if (okHt) {
        const l = s.httpsProxy.listener;
        const desc = l ? `${l.cmd} pid=${l.pid}` : 'unknown';
        console.log(`  ${c.green('✓')} HTTPS proxy  :${s.httpsProxy.port}  ${c.dim(`(${desc})`)}`);
        if (s.httpsProxy.certTrusted === false) {
            console.log(`  ${c.red('✗')} TLS cert      ${c.yellow('NOT trusted by system keychain — Excel will reject')}`);
            console.log(`     ${c.dim('error: ' + s.httpsProxy.certError)}`);
            console.log();
            console.log(c.bold('  Fix: install the proxy\'s root CA into login keychain.'));
            console.log(c.dim('  If you used 9router\'s rootCA (~/.9router/mitm/rootCA.crt):'));
            console.log(c.cyan('    security add-trusted-cert -r trustRoot \\'));
            console.log(c.cyan('      -k ~/Library/Keychains/login.keychain-db \\'));
            console.log(c.cyan('      ~/.9router/mitm/rootCA.crt'));
            console.log(c.dim('  Then re-run this command and try Excel again.'));
        } else if (s.httpsProxy.certTrusted === true) {
            console.log(`  ${c.green('✓')} TLS cert      ${c.dim('trusted by system keychain')}`);
        }
    } else {
        console.log(`  ${c.red('✗')} HTTPS proxy   ${c.dim(s.httpsProxy.hint)}`);
    }
    console.log();

    const cfg = buildConfigText(s);
    if (cfg && s.httpsProxy.certTrusted === false) {
        console.log(c.yellow('⚠ Cert untrusted — Excel will reject the connection. Fix above before pasting.'));
        console.log();
    }
    if (cfg) {
        console.log(c.bold('=== Excel paste-ready config ==='));
        console.log();
        console.log(cfg.split('\n').map((l) => '  ' + l).join('\n'));
        console.log();
        console.log(c.dim('Save to file: excel-claude --save'));
        console.log(c.dim('Re-run anytime to re-emit if you forget the values.'));
    } else {
        console.log(c.yellow('Cannot emit config — HTTPS proxy not reachable.'));
        console.log(c.dim('Set up Caddy or another HTTPS proxy in front of 9router :' + s.nineRouter.port));
        process.exit(2);
    }
}

// Find pivot.claude.ai LocalStorage SQLite path by scanning WebKit origin folders.
function findPivotClaudeLocalStorage() {
    const ROOT = path.join(HOME, 'Library/Containers/com.microsoft.Excel/Data/Library/WebKit/WebsiteData/Default');
    if (!fs.existsSync(ROOT)) return null;
    for (const d of fs.readdirSync(ROOT)) {
        const inner = path.join(ROOT, d, d);
        const originFile = path.join(inner, 'origin');
        if (!fs.existsSync(originFile)) continue;
        const origin = fs.readFileSync(originFile, 'utf8');
        if (origin.includes('pivot.claude.ai')) {
            const ls = path.join(inner, 'LocalStorage', 'localstorage.sqlite3');
            if (fs.existsSync(ls)) return ls;
        }
    }
    return null;
}

// Write Claude inference profile directly into Excel's LocalStorage so
// Anthropic's add-in finds it pre-filled on next launch. Reverse-engineered
// from the live storage shape:
//
//   key: claude.inference.profile
//   key: _OfficeRuntime_Storage_claude.inference.profile (Office runtime mirror)
//   value (UTF-16-LE):
//     {"kind":"gateway","url":"https://127.0.0.1:PORT/v1",
//      "token":"sk-...","authHeader":"x-api-key","apiFormat":"anthropic"}
function injectIntoExcel(s) {
    if (!s.httpsProxy?.alive) die('HTTPS proxy not detected — nothing to inject');
    if (!s.httpsProxy.certTrusted) die('Cert not trusted by keychain — Excel would reject anyway');

    const lsPath = findPivotClaudeLocalStorage();
    if (!lsPath) die('pivot.claude.ai LocalStorage not found. Open Anthropic Claude add-in in Excel once first to create it.');

    // Excel must be quit — SQLite WAL is locked otherwise.
    const pg = spawnSync('pgrep', ['-f', 'Microsoft Excel'], { encoding: 'utf8' });
    if ((pg.stdout || '').trim()) {
        die('Excel is currently running. Quit it completely (⌘Q) before --inject (LocalStorage DB is locked).');
    }

    const profile = JSON.stringify({
        kind: 'gateway',
        url: `https://127.0.0.1:${s.httpsProxy.port}/v1`,
        token: s.apiKey,
        authHeader: 'x-api-key',
        apiFormat: 'anthropic',
    });
    const profileB64 = Buffer.from(profile, 'utf8').toString('base64');

    // Use python3 (preinstalled on macOS) — Node has no built-in sqlite3.
    const pyScript = `
import sqlite3, base64, sys
v = base64.b64decode("${profileB64}").decode('utf-8').encode('utf-16-le')
con = sqlite3.connect(${JSON.stringify(lsPath)})
for k in ("claude.inference.profile", "_OfficeRuntime_Storage_claude.inference.profile"):
    con.execute("INSERT OR REPLACE INTO ItemTable (key, value) VALUES (?, ?)", (k, sqlite3.Binary(v)))
con.commit()
print("ok", flush=True)
`;
    const pyResult = spawnSync('/usr/bin/python3', ['-c', pyScript], { encoding: 'utf8' });
    if (pyResult.status !== 0 || !(pyResult.stdout || '').includes('ok')) {
        die(`Inject failed: ${pyResult.stderr || pyResult.stdout}`);
    }

    log(c.green(`✓ Config injected into Excel LocalStorage`));
    log(c.dim(`  ${lsPath}`));
    log(c.dim('  Open Excel → Anthropic Claude — gateway URL + token will be pre-filled.'));
}

(async () => {
    const args = process.argv.slice(2);
    const wantJson = args.includes('--json');
    const wantSave = args.includes('--save');
    const wantInject = args.includes('--inject');

    const s = await scan();

    if (wantJson) {
        console.log(JSON.stringify(s, null, 2));
        return;
    }

    printHuman(s);

    if (wantSave) {
        const cfg = buildConfigText(s);
        if (cfg) {
            fs.writeFileSync(SAVE_PATH, cfg);
            console.log(c.green(`\n✓ Saved to ${SAVE_PATH}`));
        }
    }
    if (wantInject) {
        console.log();
        injectIntoExcel(s);
    }
})().catch((e) => {
    console.error(c.red('Unexpected error: ' + e.message));
    process.exit(1);
});
