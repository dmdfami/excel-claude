#!/usr/bin/env node
/*
 * cowork-gateway — read-only diagnostic / config printer.
 *
 * Scans the local machine for:
 *   1. 9router (~/.9router/db.json + listener on :20128)
 *   2. An HTTPS proxy in front of 9router (typically :20443) that
 *      successfully forwards /v1/models with the 9router API key.
 *
 * Emits an Excel-paste-ready config block. Does NOT install services,
 * does NOT generate certs, does NOT touch Excel preferences. Pure read.
 *
 * Run: npx github:dmdfami/cowork-gateway
 *
 * Flags:
 *   --save   Also write config to ~/Desktop/excel-claude-config.txt
 *   --json   Emit machine-readable JSON instead of human output
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

function probe({ host, port, scheme, path: p, apiKey, timeout = 3000 }) {
    return new Promise((resolve) => {
        const lib = scheme === 'https' ? https : http;
        const req = lib.request({
            host, port, path: p, method: 'GET',
            headers: apiKey ? { Authorization: `Bearer ${apiKey}` } : {},
            rejectUnauthorized: false,
            timeout,
        }, (res) => {
            let body = '';
            res.on('data', (d) => { body += d.toString(); if (body.length > 4096) body = body.slice(0, 4096); });
            res.on('end', () => resolve({ ok: res.statusCode === 200, status: res.statusCode, body }));
        });
        req.on('error', (e) => resolve({ ok: false, error: e.message }));
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

async function findHttpsProxy(apiKey) {
    // Common Caddy / cowork-gateway / mitmproxy ports first, then scan 20440-20460.
    const candidates = [20443, 20444, 20445, 20446, 20447, 20448, 20449, 20450];
    const listeners = listListeners([20440, 20460]);
    for (const l of listeners) {
        if (!candidates.includes(l.port)) candidates.push(l.port);
    }
    for (const port of [...new Set(candidates)]) {
        const r = await probe({ host: '127.0.0.1', port, scheme: 'https', path: '/v1/models', apiKey });
        if (r.ok) return { port, listener: listeners.find((l) => l.port === port) };
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
            ? { port: httpsP.port, listener: httpsP.listener, alive: true }
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
    } else {
        console.log(`  ${c.red('✗')} HTTPS proxy   ${c.dim(s.httpsProxy.hint)}`);
    }
    console.log();

    const cfg = buildConfigText(s);
    if (cfg) {
        console.log(c.bold('=== Excel paste-ready config ==='));
        console.log();
        console.log(cfg.split('\n').map((l) => '  ' + l).join('\n'));
        console.log();
        console.log(c.dim('Save to file: cowork-gateway --save'));
        console.log(c.dim('Re-run anytime to re-emit if you forget the values.'));
    } else {
        console.log(c.yellow('Cannot emit config — HTTPS proxy not reachable.'));
        console.log(c.dim('Set up Caddy or another HTTPS proxy in front of 9router :' + s.nineRouter.port));
        process.exit(2);
    }
}

(async () => {
    const args = process.argv.slice(2);
    const wantJson = args.includes('--json');
    const wantSave = args.includes('--save');

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
})().catch((e) => {
    console.error(c.red('Unexpected error: ' + e.message));
    process.exit(1);
});
