#!/usr/bin/env node
/*
 * cowork-gateway — HTTPS bridge that lets Claude for Office (Excel) and
 * Claude Desktop 3p talk to any Anthropic-compatible gateway (9router,
 * litellm, etc.). Solves three concrete pains:
 *
 *   1. Office add-ins block plain HTTP (mixed content) → we serve HTTPS
 *      with a locally-trusted self-signed cert.
 *   2. Excel rejects model names that don't match `claude*` → use the
 *      "model" wizard step to expose only safe aliases.
 *   3. 9router's Windsurf-emulating provider (`cc/`) appends `_ide` to
 *      tool names; we strip the suffix on the way back so clients see
 *      the original names.
 *
 * Zero npm dependencies — relies on Node 18+ stdlib, system openssl,
 * and macOS `security`/`launchctl`.
 */

const fs        = require('fs');
const os        = require('os');
const path      = require('path');
const http      = require('http');
const https     = require('https');
const readline  = require('readline');
const { spawnSync } = require('child_process');

const HOME       = os.homedir();
const CFG_DIR    = path.join(HOME, '.config', 'cowork-gateway');
const CFG_FILE   = path.join(CFG_DIR, 'config.json');
const CERT_FILE  = path.join(CFG_DIR, 'cert.crt');
const KEY_FILE   = path.join(CFG_DIR, 'cert.key');
const PLIST_FILE = path.join(HOME, 'Library', 'LaunchAgents', 'com.cowork-gateway.plist');
const LOG_FILE   = '/tmp/cowork-gateway.log';
const LABEL      = 'com.cowork-gateway';

const c = {
    cyan:   (s) => `\x1b[36m${s}\x1b[0m`,
    green:  (s) => `\x1b[32m${s}\x1b[0m`,
    yellow: (s) => `\x1b[33m${s}\x1b[0m`,
    red:    (s) => `\x1b[31m${s}\x1b[0m`,
    dim:    (s) => `\x1b[2m${s}\x1b[0m`,
};

const log  = (...a) => console.log(c.cyan('[cowork]'), ...a);
const die  = (m) => { console.error(c.red('[cowork] ERROR:'), m); process.exit(1); };

const sh = (cmd, args, opts = {}) => {
    const r = spawnSync(cmd, args, { encoding: 'utf8', ...opts });
    if (r.status !== 0 && !opts.allowFail) {
        throw new Error(`${cmd} ${args.join(' ')} failed: ${r.stderr || r.stdout}`);
    }
    return r;
};

// ---------- prompt helpers ----------
const ask = (question, defaultValue = '') => new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    const prompt = defaultValue
        ? `${question} ${c.dim(`[${defaultValue}]`)} `
        : `${question} `;
    rl.question(prompt, (answer) => {
        rl.close();
        resolve((answer.trim() || defaultValue).trim());
    });
});

// ---------- detect ----------
const DEFAULT_MODELS = ['claude-sonnet-4-6', 'claude-opus-4-7', 'claude-haiku-4-5'];

// Look for an existing 9router install on the local machine and pull its
// baseUrl + first apiKey straight from db.json. Verifies the gateway is
// actually reachable before returning. Returns null if nothing usable.
function detectLocal9router() {
    const dbPath = path.join(HOME, '.9router', 'db.json');
    if (!fs.existsSync(dbPath)) return null;
    let db;
    try { db = JSON.parse(fs.readFileSync(dbPath, 'utf8')); }
    catch (_) { return null; }
    const apiKey = db?.apiKeys?.[0]?.key;
    if (!apiKey) return null;
    const baseUrl = 'http://127.0.0.1:20128/v1';
    const r = sh('curl', ['-s', '-o', '/dev/null', '-w', '%{http_code}',
        '--max-time', '3', `${baseUrl}/models`,
        '-H', `Authorization: Bearer ${apiKey}`], { allowFail: true });
    if (r.stdout.trim() !== '200') return null;
    return { baseUrl, apiKey, source: '~/.9router/db.json' };
}

function checkPortFree(port) {
    const r = sh('lsof', ['-nP', `-iTCP:${port}`, '-sTCP:LISTEN'],
                 { allowFail: true });
    if (r.stdout.trim()) {
        die(`Port ${port} already in use:\n${r.stdout}\nStop the existing service or re-run init with a different port.`);
    }
}

// ---------- init ----------
async function cmdInit() {
    const args = process.argv.slice(3);
    const auto = args.includes('--auto') || args.includes('-y') || args.includes('--yes');
    const portArg = args.find((a) => a.startsWith('--port='));
    const portFromArg = portArg ? Number(portArg.split('=')[1]) : null;

    let baseUrl, apiKey, models, httpsPort;

    const detected = detectLocal9router();
    if (detected) {
        log(`${c.green('✓')} Detected 9router → ${detected.baseUrl}`);
        log(`  API key from ${detected.source}`);
        if (auto) {
            log(`Running ${c.cyan('--auto')} mode — using detected config, no prompts.`);
            ({ baseUrl, apiKey } = detected);
            models = DEFAULT_MODELS;
            httpsPort = portFromArg || 20443;
        } else {
            const confirm = await ask(`Use detected config? ${c.dim('[Y/n]')}`, 'Y');
            if (confirm.toLowerCase().startsWith('y')) {
                ({ baseUrl, apiKey } = detected);
                const modelStr = await ask('Model IDs (comma-sep):',
                                            DEFAULT_MODELS.join(','));
                models = modelStr.split(',').map((s) => s.trim()).filter(Boolean);
                httpsPort = portFromArg || Number(await ask('HTTPS port:', '20443'));
            }
        }
    }

    // Wizard fallback (no detection or user declined)
    if (!apiKey) {
        if (auto) die('No 9router detected and --auto specified. Run without --auto for wizard.');
        log('No local 9router detected, or you declined. Manual setup:\n');
        baseUrl  = await ask('1. Gateway base URL:', 'http://127.0.0.1:20128/v1');
        apiKey   = await ask('2. API key (sk-...):');
        const modelStr = await ask('3. Model IDs (comma-sep):',
                                    DEFAULT_MODELS.join(','));
        models = modelStr.split(',').map((s) => s.trim()).filter(Boolean);
        httpsPort = portFromArg || Number(await ask('4. HTTPS port:', '20443'));
    }

    if (!baseUrl || !apiKey) die('baseUrl and apiKey are required');
    checkPortFree(httpsPort);

    const config = { baseUrl, apiKey, models, httpsPort, toolNameUnmangle: true };
    fs.mkdirSync(CFG_DIR, { recursive: true });
    fs.writeFileSync(CFG_FILE, JSON.stringify(config, null, 2));
    fs.chmodSync(CFG_FILE, 0o600);
    log(`Config saved: ${CFG_FILE}`);

    generateCert();
    trustCert();
    writePlist();
    reloadLaunchd();

    await sleep(2000);
    smokeTest(config);
    printExcelInstructions(config);
}

// ---------- cert ----------
function generateCert() {
    if (fs.existsSync(CERT_FILE) && fs.existsSync(KEY_FILE)) {
        // Check expiry > 30 days
        const r = sh('openssl', ['x509', '-in', CERT_FILE, '-noout', '-checkend', '2592000'],
                     { allowFail: true });
        if (r.status === 0) {
            log('Cert OK (>30 days valid)');
            return;
        }
    }
    log('Generating self-signed cert for 127.0.0.1...');
    const cnf = `[req]\ndistinguished_name=req\n[v3]\nsubjectAltName=IP:127.0.0.1,DNS:localhost\nextendedKeyUsage=serverAuth\nbasicConstraints=critical,CA:false`;
    const cnfFile = path.join(CFG_DIR, 'openssl.cnf');
    fs.writeFileSync(cnfFile, cnf);
    sh('openssl', ['req', '-x509', '-nodes', '-newkey', 'rsa:2048',
        '-keyout', KEY_FILE, '-out', CERT_FILE,
        '-days', '825', '-subj', '/CN=127.0.0.1',
        '-extensions', 'v3', '-config', cnfFile]);
    fs.chmodSync(KEY_FILE, 0o600);
    fs.unlinkSync(cnfFile);
    log(`Cert: ${CERT_FILE}`);
}

function trustCert() {
    // Check if already trusted in login keychain
    const r = sh('security', ['find-certificate', '-c', '127.0.0.1',
        path.join(HOME, 'Library/Keychains/login.keychain-db')], { allowFail: true });
    if (r.status === 0) {
        log('Cert already trusted in login keychain');
        return;
    }
    log('Trusting cert in login keychain...');
    sh('security', ['add-trusted-cert', '-r', 'trustRoot',
        '-k', path.join(HOME, 'Library/Keychains/login.keychain-db'),
        CERT_FILE]);
}

// ---------- launchd ----------
function writePlist() {
    const nodeBin = process.execPath;
    const plist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>${LABEL}</string>
  <key>ProgramArguments</key>
  <array>
    <string>${nodeBin}</string>
    <string>${path.resolve(__filename)}</string>
    <string>serve</string>
  </array>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><true/>
  <key>StandardOutPath</key><string>${LOG_FILE}</string>
  <key>StandardErrorPath</key><string>${LOG_FILE}</string>
</dict>
</plist>`;
    fs.mkdirSync(path.dirname(PLIST_FILE), { recursive: true });
    fs.writeFileSync(PLIST_FILE, plist);
    log(`launchd plist: ${PLIST_FILE}`);
}

function reloadLaunchd() {
    sh('launchctl', ['unload', PLIST_FILE], { allowFail: true });
    sh('launchctl', ['load', PLIST_FILE]);
    log('Service loaded — auto-starts on login');
}

// ---------- serve (run by launchd) ----------
function cmdServe() {
    if (!fs.existsSync(CFG_FILE)) die('Not initialized — run `cowork-gateway init` first');
    const cfg = JSON.parse(fs.readFileSync(CFG_FILE, 'utf8'));
    const upstream = new URL(cfg.baseUrl);

    const NAME_RE = /"name"(\s*:\s*)"([^"]*?)_ide"/g;
    const stripIde = (s) => s.replace(NAME_RE, '"name"$1"$2"');

    const server = https.createServer({
        cert: fs.readFileSync(CERT_FILE),
        key:  fs.readFileSync(KEY_FILE),
    }, (clientReq, clientRes) => {
        // Forward to upstream
        const upstreamPath = (upstream.pathname.replace(/\/$/, '') + clientReq.url).replace(/\/+/g, '/');
        const opts = {
            host: upstream.hostname,
            port: upstream.port || (upstream.protocol === 'https:' ? 443 : 80),
            method: clientReq.method,
            path: upstreamPath,
            headers: { ...clientReq.headers, host: upstream.host },
        };
        const lib = upstream.protocol === 'https:' ? https : http;
        const upstreamReq = lib.request(opts, (upstreamRes) => {
            const headers = { ...upstreamRes.headers };
            delete headers['content-length']; // body may shrink after rewrite
            clientRes.writeHead(upstreamRes.statusCode, headers);

            if (!cfg.toolNameUnmangle) {
                upstreamRes.pipe(clientRes);
                return;
            }

            let buf = '';
            upstreamRes.setEncoding('utf8');
            upstreamRes.on('data', (chunk) => {
                buf += chunk;
                const split = Math.max(buf.lastIndexOf('\n'), buf.lastIndexOf('}'));
                if (split === -1) return;
                clientRes.write(stripIde(buf.slice(0, split + 1)));
                buf = buf.slice(split + 1);
            });
            upstreamRes.on('end', () => {
                if (buf) clientRes.write(stripIde(buf));
                clientRes.end();
            });
            upstreamRes.on('error', (err) => {
                console.error('upstream error:', err.message);
                clientRes.destroy();
            });
        });
        upstreamReq.on('error', (err) => {
            console.error('upstream connect error:', err.message);
            try {
                clientRes.writeHead(502, { 'content-type': 'text/plain' });
                clientRes.end(`upstream unreachable: ${err.message}`);
            } catch (_) {}
        });
        clientReq.pipe(upstreamReq);
        clientReq.on('error', () => upstreamReq.destroy());
    });

    server.listen(cfg.httpsPort, '127.0.0.1', () => {
        console.log(`[cowork-gateway] HTTPS https://127.0.0.1:${cfg.httpsPort} → ${cfg.baseUrl} (unmangle=${cfg.toolNameUnmangle})`);
    });
}

// ---------- start/stop/status/uninstall ----------
function cmdStart()  { sh('launchctl', ['load', PLIST_FILE], { allowFail: true }); log('started'); }
function cmdStop()   { sh('launchctl', ['unload', PLIST_FILE], { allowFail: true }); log('stopped'); }
function cmdStatus() {
    const r = sh('launchctl', ['list', LABEL], { allowFail: true });
    console.log(r.stdout || c.yellow('not loaded'));
    if (fs.existsSync(LOG_FILE)) {
        console.log(c.dim('--- last log lines ---'));
        const lines = fs.readFileSync(LOG_FILE, 'utf8').split('\n').slice(-5);
        console.log(lines.join('\n'));
    }
}
function cmdUninstall() {
    sh('launchctl', ['unload', PLIST_FILE], { allowFail: true });
    [PLIST_FILE, CERT_FILE, KEY_FILE, CFG_FILE].forEach((f) => {
        if (fs.existsSync(f)) fs.unlinkSync(f);
    });
    sh('security', ['delete-certificate', '-c', '127.0.0.1',
        path.join(HOME, 'Library/Keychains/login.keychain-db')], { allowFail: true });
    log('uninstalled. Config dir kept at ' + CFG_DIR);
}

// ---------- helpers ----------
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function smokeTest(cfg) {
    const url = `https://127.0.0.1:${cfg.httpsPort}/models`;
    const r = sh('curl', ['-sk', '-o', '/dev/null', '-w', '%{http_code}', '--max-time', '5',
        url, '-H', `Authorization: Bearer ${cfg.apiKey}`], { allowFail: true });
    if (r.stdout.trim() === '200') {
        log(c.green('✓ smoke test passed — gateway reachable'));
    } else {
        console.log(c.yellow(`! smoke test inconclusive (HTTP ${r.stdout.trim() || '???'}). Check ${LOG_FILE}`));
    }
}

function printExcelInstructions(cfg) {
    const endpoint = `https://127.0.0.1:${cfg.httpsPort}/v1`;
    console.log(`
${c.green('✓ Setup complete.')}

${c.cyan('Claude for Office (Excel) — Settings → Configure Gateway:')}
  ${c.dim('URL:')}        ${endpoint}
  ${c.dim('Token:')}      ${cfg.apiKey}
  ${c.dim('AuthHeader:')} x-api-key
  ${c.dim('APIFormat:')}  anthropic
  ${c.dim('Model:')}      ${cfg.models[0]}  ${c.dim(`(or any of: ${cfg.models.join(', ')})`)}

${c.cyan('Claude Desktop 3p — Configure third-party inference → Gateway:')}
  ${c.dim('Base URL:')}     ${endpoint}
  ${c.dim('API key:')}      ${cfg.apiKey}
  ${c.dim('Auth scheme:')}  bearer

${c.dim(`Logs: ${LOG_FILE}`)}
${c.dim(`Manage: cowork-gateway [start|stop|status|uninstall]`)}
`);
}

// ---------- entry ----------
const cmd = process.argv[2] || 'help';
const handlers = {
    init: cmdInit, serve: cmdServe,
    start: cmdStart, stop: cmdStop, status: cmdStatus, uninstall: cmdUninstall,
};
if (handlers[cmd]) {
    Promise.resolve(handlers[cmd]()).catch((e) => die(e.message));
} else {
    console.log(`cowork-gateway — Anthropic-compatible HTTPS bridge for Excel & Claude Desktop

Commands:
  init [--auto] [--port=N]
              Setup. Auto-detects 9router from ~/.9router/db.json.
              --auto: zero-prompt mode using detected config + defaults.
  start       Load launchd service
  stop        Unload launchd service
  status      Show service + recent log
  uninstall   Remove plist, cert, trust entry
  serve       (internal) Run the HTTPS proxy in foreground

Quick start (with 9router already running locally):
  npx github:dmdfami/cowork-gateway init --auto

Manual mode (will prompt for baseUrl + apiKey):
  npx github:dmdfami/cowork-gateway init
`);
}
