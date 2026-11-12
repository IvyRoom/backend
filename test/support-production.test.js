'use strict';

require('./deny-production-network');
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { createSupportProductionRuntime, productionSettings, acquireProductionLock, createSupportBackup } = require('../support-production');
const runtimes = new Map();

async function directory(t) {
    const parent = await fs.mkdtemp(path.join(os.tmpdir(), 'support-production-'));
    t.after(async () => {
        for (const runtime of runtimes.get(parent) || []) await runtime.stop();
        runtimes.delete(parent);
        const resolved = path.resolve(parent);
        assert.ok(resolved.startsWith(`${path.resolve(os.tmpdir())}${path.sep}`));
        await fs.rm(resolved, { recursive: true, force: true });
    });
    return parent;
}

function environment(directory, extra = {}) {
    return { SUPPORT_MANAGEMENT_ENABLED: 'true', SUPPORT_DATA_DIRECTORY: directory,
        PAGARME_PUBLIC_KEY_PRODUCTION: 'pk_live_synthetic', PAGARME_API_KEY_PRODUCTION: 'sk_live_synthetic', ...extra };
}

function fixture(env, overrides = {}) {
    const calls = [];
    let options;
    const fakeStore = () => ({ read: async () => ({}), write: async value => value });
    const runtime = createSupportProductionRuntime({ environment: env, graphClient: { synthetic: true },
        factories: {
            createUsers: args => { calls.push(['users', args]); return {}; },
            createPagarme: args => { calls.push(['pagarme', args]); return {}; },
            createAppointments: args => { calls.push(['appointments', args]); return {}; },
            createNotifications: args => { calls.push(['notifications', args]); return {}; },
            createCardStore: fakeStore, createBookingStore: fakeStore,
        },
        createApplication(value) { options = value; calls.push(['application']); return (req, res) => { res.statusCode = 204; }; },
        ...overrides,
    });
    if (runtime && env.SUPPORT_DATA_DIRECTORY) runtimes.set(env.SUPPORT_DATA_DIRECTORY, [...(runtimes.get(env.SUPPORT_DATA_DIRECTORY) || []), runtime]);
    return { runtime, calls, options: () => options };
}

test('production runtime is disabled by default without loading credentials, filesystem or factories', () => {
    const { runtime, calls } = fixture({});
    assert.equal(runtime, undefined);
    assert.deepEqual(calls, []);
    assert.equal(productionSettings({ SUPPORT_MANAGEMENT_ENABLED: 'false' }), null);
});

test('enabled production requires matching live credentials and a private absolute storage location', async t => {
    const root = await directory(t);
    assert.equal(productionSettings(environment(root, { PAGARME_PUBLIC_KEY_PRODUCTION: 'pk_productionSynthetic',
        PAGARME_API_KEY_PRODUCTION: 'sk_productionSynthetic' })).publicKey, 'pk_productionSynthetic');
    for (const change of [
        { SUPPORT_MANAGEMENT_ENABLED: 'TRUE' }, { PAGARME_PUBLIC_KEY_PRODUCTION: 'pk_test_synthetic' },
        { PAGARME_API_KEY_PRODUCTION: 'sk_test_synthetic' }, { PAGARME_API_KEY_PRODUCTION: '' },
        { SUPPORT_DATA_DIRECTORY: '.' }, { SUPPORT_DATA_DIRECTORY: path.resolve(__dirname, '..') },
        { SUPPORT_DATA_DIRECTORY: path.parse(root).root },
    ]) assert.throws(() => productionSettings(environment(root, change)), /Support production configuration or storage unavailable/u);
    assert.equal(productionSettings(environment(root)).directory, root);
});

test('Azure HOME default uses persistent private data and rejects unsafe local-cache or disabled-storage defaults', async t => {
    const root = await directory(t);
    const settings = environment('', { HOME: root, WEBSITE_SITE_NAME: 'synthetic-site', WEBSITE_INSTANCE_ID: 'synthetic-instance' });
    assert.equal(productionSettings(settings).directory, path.join(root, 'data', 'support-management-production'));
    for (const change of [{ WEBSITE_LOCAL_CACHE_OPTION: 'Always' }, { WEBSITES_ENABLE_APP_SERVICE_STORAGE: 'false' },
        { WEBSITE_INSTANCE_ID: '' }, { HOME: '' }]) assert.throws(() => productionSettings({ ...settings, ...change }));
    assert.throws(() => productionSettings(environment('', { HOME: root })));
});

test('runtime construction is lazy and first use opens durable stores once with production-only dependencies', async t => {
    const root = await directory(t);
    const { runtime, calls, options } = fixture(environment(root));
    t.after(() => runtime.stop());
    assert.deepEqual(calls, []);
    assert.deepEqual(await fs.readdir(root), []);
    const [first, second] = await Promise.all([runtime.ready(), runtime.ready()]);
    assert.equal(first, second);
    assert.equal(calls.filter(([name]) => name === 'application').length, 1);
    assert.equal(options().environment, 'production');
    assert.equal(options().origin, 'https://machadogestao.com');
    assert.equal(options().apiOnly, true);
    assert.equal(options().preload, true);
    assert.equal(calls.find(([name]) => name === 'pagarme')[1].environment, 'production');
    assert.equal(calls.find(([name]) => name === 'users')[1].environment, 'production');
    assert.equal(options().publicKey, 'pk_live_synthetic');
    const marker = JSON.parse(await fs.readFile(path.join(root, 'environment.json'), 'utf8'));
    assert.equal(marker.environment, 'production');
    assert.match(marker.keyFingerprint, /^[a-f0-9]{64}$/u);
    assert.equal(JSON.stringify(marker).includes('pk_live_'), false);
    assert.ok((await fs.readdir(root)).includes('sessions.json'));
    await runtime.stop();
    assert.equal((await fs.readdir(root)).includes('production.lock'), false);
});

test('separate production workers cannot initialize the same metadata directory concurrently', async t => {
    const root = await directory(t);
    const first = fixture(environment(root)).runtime;
    const second = fixture(environment(root)).runtime;
    t.after(() => first.stop());
    t.after(() => second.stop());
    await first.ready();
    await assert.rejects(second.ready(), { code: 'support_production.storage_locked' });
    await first.stop();
    await second.ready();
});

test('a dead same-instance owner is recovered while foreign owners and uncertain recovery are preserved', async t => {
    const root = await directory(t);
    const filename = path.join(root, 'production.lock');
    await fs.writeFile(filename, JSON.stringify({ instance: 'same', pid: 2147483647, token: 'old' }));
    const release = await acquireProductionLock(root, 'same');
    assert.equal(JSON.parse(await fs.readFile(filename, 'utf8')).pid, process.pid);
    await release();
    const foreign = JSON.stringify({ instance: 'other-host', pid: 2147483647, token: 'foreign' });
    await fs.writeFile(filename, foreign);
    await assert.rejects(acquireProductionLock(root, 'same'), { code: 'support_production.storage_locked' });
    assert.equal(await fs.readFile(filename, 'utf8'), foreign);
    await fs.writeFile(path.join(root, 'production-recovery.lock'), 'review required');
    await assert.rejects(acquireProductionLock(root, 'same'), { code: 'support_production.storage_locked' });
    assert.equal(await fs.readFile(path.join(root, 'production-recovery.lock'), 'utf8'), 'review required');
});

test('production refuses unlabeled sandbox state and mismatched account keys without overwriting files', async t => {
    const root = await directory(t);
    await fs.writeFile(path.join(root, 'sessions.json'), '{}');
    const sandbox = fixture(environment(root)).runtime;
    t.after(() => sandbox.stop());
    await assert.rejects(sandbox.ready(), { code: 'support_production.environment_mismatch' });
    assert.equal(await fs.readFile(path.join(root, 'sessions.json'), 'utf8'), '{}');
    await fs.rm(path.join(root, 'sessions.json'));
    const first = fixture(environment(root)).runtime;
    await first.ready();
    await first.stop();
    const changed = fixture(environment(root, { PAGARME_PUBLIC_KEY_PRODUCTION: 'pk_live_anotherAccount' })).runtime;
    t.after(() => changed.stop());
    await assert.rejects(changed.ready(), { code: 'support_production.environment_mismatch' });
});

test('failed initialization releases its owner and returns a fixed unavailable response without internal details', async t => {
    const root = await directory(t);
    const { runtime } = fixture(environment(root), { createApplication() { throw new Error('private configuration'); } });
    t.after(() => runtime.stop());
    const response = { status(code) { this.code = code; return this; }, json(body) { this.body = body; return this; } };
    await runtime.middleware({}, response, () => { throw new Error('must not fall through'); });
    assert.equal(response.code, 503);
    assert.deepEqual(response.body, { error: 'support_management.service_unavailable' });
    assert.equal((await fs.readdir(root)).includes('production.lock'), false);
});

test('initialization can retry after a metadata-only setup failure without replaying integrations', async t => {
    const root = await directory(t);
    let attempts = 0;
    const { runtime } = fixture(environment(root), { createApplication() {
        if (!attempts++) throw new Error('synthetic setup failure');
        return () => {};
    } });
    await assert.rejects(runtime.ready(), /synthetic setup failure/u);
    const application = await runtime.ready();
    assert.equal(typeof application, 'function');
    assert.equal(attempts, 2);
});

test('stop waits for pending initialization, releases its owner and prevents later reinitialization', async t => {
    const root = await directory(t);
    let entered;
    let release;
    const pending = new Promise(resolve => { release = resolve; });
    const started = new Promise(resolve => { entered = resolve; });
    const { runtime } = fixture(environment(root), { factories: {
        openSessionStore: async () => { entered(); await pending; return {}; },
        createUsers: () => ({}), createPagarme: () => ({}), createAppointments: () => ({}), createNotifications: () => ({}),
    }, createApplication: () => (() => {}) });
    const ready = runtime.ready();
    await started;
    let stopped = false;
    const stop = runtime.stop().then(() => { stopped = true; });
    await Promise.resolve();
    assert.equal(stopped, false);
    release();
    await ready;
    await stop;
    await assert.rejects(runtime.ready(), { code: 'support_production.stopped' });
    assert.equal((await fs.readdir(root)).includes('production.lock'), false);
});

test('daily snapshots retain seven dates, omit locks and unknown files, and preserve the beginning-of-day state', async t => {
    const root = await directory(t);
    let now = Date.parse('2026-11-01T15:00:00Z');
    const backup = createSupportBackup(root, () => now);
    await fs.writeFile(path.join(root, 'sessions.json'), '{"v":1}');
    await fs.writeFile(path.join(root, 'production.lock'), 'owner');
    await fs.writeFile(path.join(root, 'unrelated.txt'), 'not app metadata');
    for (let day = 1; day <= 9; day++) {
        now = Date.parse(`2026-11-${String(day).padStart(2, '0')}T15:00:00Z`);
        await Promise.all([backup.ensure(), backup.ensure()]);
    }
    assert.deepEqual(await fs.readdir(path.join(root, 'backups')), ['2026-11-03', '2026-11-04', '2026-11-05', '2026-11-06', '2026-11-07', '2026-11-08', '2026-11-09']);
    assert.deepEqual(await fs.readdir(path.join(root, 'backups', '2026-11-09')), ['sessions.json']);
    await fs.writeFile(path.join(root, 'sessions.json'), '{"v":2}');
    await backup.ensure();
    assert.equal(await fs.readFile(path.join(root, 'backups', '2026-11-09', 'sessions.json'), 'utf8'), '{"v":1}');
});

test('daily snapshots and retention accept a private directory reached through an ancestor alias', async t => {
    const root = await directory(t);
    const physical = path.join(root, 'physical');
    const alias = path.join(root, 'alias');
    await fs.mkdir(physical);
    try { await fs.symlink(physical, alias, process.platform === 'win32' ? 'junction' : 'dir'); }
    catch (error) {
        if (['EPERM', 'EACCES'].includes(error.code)) { t.skip('Directory links require additional filesystem permission'); return; }
        throw error;
    }
    const data = path.join(alias, 'data');
    await fs.mkdir(data);
    assert.notEqual(path.relative(await fs.realpath(data), data), '', 'The supplied ancestor alias must differ from its canonical path');
    await fs.writeFile(path.join(data, 'sessions.json'), '{"v":1}');
    let now;
    const backup = createSupportBackup(data, () => now);
    for (let day = 1; day <= 9; day++) {
        now = Date.parse(`2026-11-${String(day).padStart(2, '0')}T15:00:00Z`);
        await backup.ensure();
    }
    const canonicalBackups = path.join(physical, 'data', 'backups');
    assert.deepEqual(await fs.readdir(canonicalBackups), ['2026-11-03', '2026-11-04', '2026-11-05', '2026-11-06', '2026-11-07', '2026-11-08', '2026-11-09']);
    assert.equal(await fs.readFile(path.join(canonicalBackups, '2026-11-09', 'sessions.json'), 'utf8'), '{"v":1}');
});

test('a new Brasília day snapshots metadata before its first mutation', async t => {
    const root = await directory(t);
    let now = Date.parse('2026-11-01T15:00:00Z');
    const { runtime, options } = fixture(environment(root), { now: () => now,
        factories: {
            createUsers: () => ({}), createPagarme: () => ({}), createAppointments: () => ({}), createNotifications: () => ({}),
        } });
    t.after(() => runtime.stop());
    await runtime.ready();
    const code = `support_${'a'.repeat(44)}`;
    await options().cardStore.write(code, { customerId: 'cus_before' });
    now = Date.parse('2026-11-02T03:00:01Z');
    await options().cardStore.write(code, { customerId: 'cus_after' });
    assert.deepEqual(JSON.parse(await fs.readFile(path.join(root, 'backups', '2026-11-02', `${code}.json`), 'utf8')), { customerId: 'cus_before' });
});

test('linked backup roots or dated snapshots cannot redirect writes outside private storage', async t => {
    const root = await directory(t);
    const outside = path.join(root, 'outside');
    const data = path.join(root, 'data');
    await fs.mkdir(outside);
    await fs.mkdir(data);
    await fs.writeFile(path.join(outside, 'preserve.txt'), 'unchanged');
    try { await fs.symlink(outside, path.join(data, 'backups'), process.platform === 'win32' ? 'junction' : 'dir'); }
    catch (error) {
        if (['EPERM', 'EACCES'].includes(error.code)) { t.skip('Directory links require additional filesystem permission'); return; }
        throw error;
    }
    await assert.rejects(createSupportBackup(data).ensure(), { code: 'support_production.invalid_backup_directory' });
    assert.deepEqual(await fs.readdir(outside), ['preserve.txt']);
    const dated = path.join(root, 'dated');
    await fs.mkdir(path.join(dated, 'backups'), { recursive: true });
    await fs.symlink(outside, path.join(dated, 'backups', '2026-11-01'), process.platform === 'win32' ? 'junction' : 'dir');
    await assert.rejects(createSupportBackup(dated, () => Date.parse('2026-11-01T15:00:00Z')).ensure(), { code: 'support_production.invalid_backup_directory' });
    assert.deepEqual(await fs.readdir(outside), ['preserve.txt']);
});
