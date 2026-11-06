'use strict';

require('./deny-production-network');
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { EventEmitter, once } = require('node:events');
const { spawn, spawnSync } = require('node:child_process');
const { startSupportDevelopment, isSupportDevelopmentFile } = require('../scripts/support-dev');
const { startSupportSandbox } = require('../support-sandbox-server');

async function fixture(t, { failInitially = false } = {}) {
    const root = await fs.mkdtemp(path.join(os.tmpdir(), 'support-development-test-'));
    const childFile = path.join(root, 'child.cjs');
    const eventsFile = path.join(root, 'events.txt');
    const failureFile = path.join(root, 'fail.txt');
    const storeModule = require.resolve('../integrations/support-card-store');
    await fs.writeFile(childFile, `
        const fs = require('node:fs/promises');
        const { acquireSupportSandboxLock } = require(${JSON.stringify(storeModule)});
        async function main() {
            if (await fs.stat(${JSON.stringify(failureFile)}).then(() => true, () => false)) {
                process.exitCode = 1;
                process.disconnect();
                return;
            }
            const release = await acquireSupportSandboxLock(${JSON.stringify(root)});
            await fs.appendFile(${JSON.stringify(eventsFile)}, 'start ' + process.pid + '\\n');
            let stopping;
            const stop = () => {
                stopping ??= new Promise(resolve => setTimeout(resolve, 40)).then(release).then(async () => {
                    await fs.appendFile(${JSON.stringify(eventsFile)}, 'stop ' + process.pid + '\\n');
                    if (process.connected) process.disconnect();
                });
            };
            process.once('disconnect', stop);
            process.on('message', message => { if (message.type === 'support:stop') stop(); });
            process.send({type:'ready'});
        }
        main().catch(() => { process.exitCode = 1; if (process.connected) process.disconnect(); });
    `);
    if (failInitially) await fs.writeFile(failureFile, 'fail');
    const children = [];
    const watchers = [];
    const messages = [];
    const development = startSupportDevelopment({
        root,
        spawnChild() {
            const child = spawn(process.execPath, [childFile], { cwd: root, stdio: ['ignore', 'ignore', 'pipe', 'ipc'], windowsHide: true });
            child.ready = once(child, 'message');
            child.completed = once(child, 'exit');
            children.push(child);
            return child;
        },
        watchDirectory(directory, listener) {
            const watcher = new EventEmitter();
            watcher.directory = directory;
            watcher.listener = listener;
            watcher.closed = false;
            watcher.close = () => { watcher.closed = true; };
            watchers.push(watcher);
            return watcher;
        },
        debounceMs: 5,
        log: (message) => messages.push(message),
        report: (message) => messages.push(message),
    });
    t.after(async () => {
        await development.stop();
        await fs.rm(root, { recursive: true, force: true });
    });
    return { root, failureFile, eventsFile, development, children, watchers, messages };
}

test('support development watches only its backend and ignored local configuration', () => {
    for (const filename of ['support-app.js', 'support-sandbox-server.js', '.env', '.env.support-sandbox', 'domains/support-management.js', 'integrations/support-users.js', 'integrations/pagarme.js', 'integrations/support-card-store.js', 'integrations\\pagarme.js']) {
        assert.equal(isSupportDevelopmentFile(filename), true, filename);
    }
    for (const filename of ['server.js', 'app.js', 'domains/learning-platform.js', 'integrations/microsoft-graph.js', 'node_modules/express/index.js', '.support-sandbox-data/sandbox.lock', 'test/support-dev.test.js', 'apps/support-management/style.css', 'scripts/support-dev.js']) {
        assert.equal(isSupportDevelopmentFile(filename), false, filename);
    }
});

test('importing the development runner and sandbox server does not watch, spawn, or listen', () => {
    const result = spawnSync(process.execPath, ['-e', `
        require('node:fs').watch = () => { throw new Error('Unexpected watcher'); };
        require('node:child_process').spawn = () => { throw new Error('Unexpected child'); };
        require('node:net').Server.prototype.listen = () => { throw new Error('Unexpected listener'); };
        require('./scripts/support-dev');
        require('./support-sandbox-server');
    `], { cwd: path.resolve(__dirname, '..'), encoding: 'utf8', windowsHide: true });
    assert.equal(result.status, 0, result.stderr);
    assert.equal(result.stdout, '');
});

test('development restart drains the child and releases its real lock before spawning the replacement', async (t) => {
    const f = await fixture(t);
    await f.children[0].ready;
    const firstPid = f.children[0].pid;
    await Promise.all([f.development.restart(), f.development.restart()]);
    assert.equal(f.children.length, 2);
    await f.children[1].ready;
    const secondPid = f.children[1].pid;
    assert.notEqual(firstPid, secondPid);
    assert.equal(await fs.readFile(path.join(f.root, 'sandbox.lock'), 'utf8'), String(secondPid));
    assert.deepEqual((await fs.readFile(f.eventsFile, 'utf8')).trim().split('\n'), [`start ${firstPid}`, `stop ${firstPid}`, `start ${secondPid}`]);
    await Promise.all([f.development.stop(), f.development.stop()]);
    assert.equal(f.children[1].exitCode, 0);
    assert.equal(f.watchers.every((watcher) => watcher.closed), true);
    await assert.rejects(fs.stat(path.join(f.root, 'sandbox.lock')), { code: 'ENOENT' });
    await f.development.restart();
    assert.equal(f.children.length, 2);
});

test('a failed startup stays watchable and a configuration save starts a fresh child', async (t) => {
    const f = await fixture(t, { failInitially: true });
    assert.equal((await f.children[0].completed)[0], 1);
    assert.equal(f.watchers.every((watcher) => !watcher.closed), true);
    await fs.rm(f.failureFile);
    const watcher = f.watchers.find((item) => item.directory === f.root);
    watcher.listener('change', '.env.support-sandbox');
    const deadline = Date.now() + 3000;
    while (f.children.length < 2) {
        if (Date.now() >= deadline) throw new Error('Configuration save did not restart child');
        await new Promise((resolve) => setTimeout(resolve, 5));
    }
    await f.children[1].ready;
    assert.equal(f.messages.some((message) => message.includes('Save a watched')), true);
});

test('unresponsive children are preserved instead of force-killed or replaced', async () => {
    const child = new EventEmitter();
    child.connected = true;
    child.send = () => {};
    child.kill = () => { throw new Error('Force termination is forbidden'); };
    const watchers = [];
    let spawned = 0;
    const development = startSupportDevelopment({
        root: '.', spawnChild: () => { spawned++; return child; },
        watchDirectory: () => { const watcher = new EventEmitter(); watcher.close = () => {}; watchers.push(watcher); return watcher; },
        log: () => {}, report: () => {}, stopTimeoutMs: 5,
    });
    await assert.rejects(development.restart(), /left running/);
    assert.equal(spawned, 1);
    await assert.rejects(development.stop(), /left running/);
});

test('shutdown waits for a child that has already disconnected cleanly', async () => {
    const child = new EventEmitter();
    child.connected = false;
    child.send = () => { throw new Error('Disconnected child must not receive IPC'); };
    const development = startSupportDevelopment({
        spawnChild: () => child,
        watchDirectory: () => { const watcher = new EventEmitter(); watcher.close = () => {}; return watcher; },
        log: () => {}, report: () => {}, stopTimeoutMs: 1000,
    });
    const stopping = development.stop();
    setTimeout(() => child.emit('exit', 0, null), 5);
    await stopping;
});

test('sandbox preflight names missing settings without exposing supplied values or constructing clients', async (t) => {
    t.mock.method(require('dotenv'), 'config', () => ({}));
    await assert.rejects(startSupportSandbox({ environment: { CLIENT_SECRET: 'private-example-value' } }), (error) => {
        assert.equal(error.code, 'SUPPORT_SETUP');
        assert.match(error.message, /SUPPORT_PAGARME_PUBLIC_KEY/);
        assert.match(error.message, /SUPPORT_PAGARME_SECRET_KEY/);
        assert.equal(error.message.includes('private-example-value'), false);
        return true;
    });
    const environment = { CLIENT_ID: 'synthetic', TENANT_ID: 'synthetic', CLIENT_SECRET: 'private-example-value', SUPPORT_PAGARME_PUBLIC_KEY: 'pk_live_secret-example', SUPPORT_PAGARME_SECRET_KEY: 'sk_live_secret-example' };
    await assert.rejects(startSupportSandbox({ environment }), (error) => {
        assert.equal(error.code, 'SUPPORT_SETUP');
        assert.match(error.message, /must be Pagar.me test keys/);
        assert.equal(error.message.includes('secret-example'), false);
        return true;
    });
    const existingName = { CLIENT_ID: 'synthetic', TENANT_ID: 'synthetic', CLIENT_SECRET: 'private-example-value', PAGARME_API_KEY_SANDBOX: 'sk_test_example' };
    await assert.rejects(startSupportSandbox({ environment: existingName }), (error) => {
        assert.equal(error.code, 'SUPPORT_SETUP');
        assert.equal(error.message, 'Missing local settings: SUPPORT_PAGARME_PUBLIC_KEY. Add them to the ignored local environment files.');
        return true;
    });
    await assert.rejects(startSupportSandbox({ environment: { ...existingName, SUPPORT_PAGARME_PUBLIC_KEY: 'pk_test_example', PAGARME_API_KEY_SANDBOX: 'sk_live_private' } }), (error) => {
        assert.equal(error.code, 'SUPPORT_SETUP');
        assert.match(error.message, /must be Pagar.me test keys/);
        assert.equal(error.message.includes('sk_live_private'), false);
        return true;
    });
});
