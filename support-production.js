'use strict';

const crypto = require('node:crypto');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { createSupportApp } = require('./support-app');
const { createSupportUsersAdapter } = require('./integrations/support-users');
const { createPagarmeAdapter } = require('./integrations/pagarme');
const { createSupportCardStore } = require('./integrations/support-card-store');
const { openSupportSessionStore } = require('./integrations/support-session-store');
const { createSupportAppointmentsAdapter } = require('./integrations/support-appointments');
const { createSupportBookingStore } = require('./integrations/support-booking-store');
const { createSupportBookingNotifications } = require('./integrations/support-booking-notifications');

const ROOT = path.resolve(__dirname);
const SNAPSHOT_FILE = /^(?:environment|sessions|availability-sync|support_[a-f0-9]{44}|reservation_[a-f0-9-]{36})\.json$/u;

function failure(code) {
    const error = new Error('Support production configuration or storage unavailable');
    error.code = `support_production.${code}`;
    return error;
}

function within(parent, target) {
    const relative = path.relative(parent, target);
    return relative === '' || (!relative.startsWith(`..${path.sep}`) && relative !== '..' && !path.isAbsolute(relative));
}

function productionSettings(environment) {
    const enabled = environment.SUPPORT_MANAGEMENT_ENABLED;
    if (enabled === undefined || enabled === '' || enabled === 'false') return null;
    if (enabled !== 'true') throw failure('invalid_enabled');
    const publicKey = environment.PAGARME_PUBLIC_KEY_PRODUCTION?.trim();
    const secretKey = environment.PAGARME_API_KEY_PRODUCTION?.trim();
    if (!/^pk_(?!test_)[A-Za-z0-9_-]+$/u.test(publicKey || '') || !/^sk_(?!test_)[A-Za-z0-9_-]+$/u.test(secretKey || '')) {
        throw failure('invalid_keys');
    }
    let directory = environment.SUPPORT_DATA_DIRECTORY?.trim();
    if (!directory) {
        if (!environment.WEBSITE_SITE_NAME || !environment.WEBSITE_INSTANCE_ID || !environment.HOME
            || String(environment.WEBSITE_LOCAL_CACHE_OPTION || '').toLowerCase() === 'always'
            || String(environment.WEBSITES_ENABLE_APP_SERVICE_STORAGE || '').toLowerCase() === 'false') {
            throw failure('data_directory_required');
        }
        directory = path.join(environment.HOME, 'data', 'support-management-production');
    }
    if (!path.isAbsolute(directory)) throw failure('invalid_data_directory');
    directory = path.resolve(directory);
    if (directory === path.parse(directory).root || within(ROOT, directory) || within(directory, ROOT)
        || (environment.HOME && directory === path.resolve(environment.HOME))) throw failure('invalid_data_directory');
    return { publicKey, secretKey, directory,
        instance: environment.WEBSITE_INSTANCE_ID || os.hostname() };
}

async function privateDirectory(directory) {
    await fs.mkdir(directory, { recursive: true, mode: 0o700 });
    const actual = await fs.realpath(directory);
    const deployment = await fs.realpath(ROOT);
    if ((await fs.lstat(directory)).isSymbolicLink() || within(deployment, actual) || within(actual, deployment)) throw failure('invalid_data_directory');
    return actual;
}

function processExists(pid) {
    try { process.kill(pid, 0); return true; }
    catch (error) { if (error.code === 'ESRCH') return false; return true; }
}

async function acquireProductionLock(directory, instance) {
    const filename = path.join(directory, 'production.lock');
    const recoveryPath = path.join(directory, 'production-recovery.lock');
    const owner = { instance, pid: process.pid, token: crypto.randomUUID() };
    async function create() {
        const handle = await fs.open(filename, 'wx', 0o600);
        try { await handle.writeFile(JSON.stringify(owner)); }
        catch (error) { await handle.close(); await fs.rm(filename, { force: true }); throw error; }
        return handle;
    }
    let handle;
    try { handle = await create(); }
    catch (error) {
        if (error.code !== 'EEXIST') throw error;
        let recovery;
        try { recovery = await fs.open(recoveryPath, 'wx', 0o600); }
        catch { throw failure('storage_locked'); }
        try {
            let previous;
            try { previous = JSON.parse(await fs.readFile(filename, 'utf8')); }
            catch { throw failure('storage_locked'); }
            // A PID has meaning only on its own instance; never remove another host's owner.
            if (previous.instance !== instance || !Number.isInteger(previous.pid) || previous.pid < 1
                || processExists(previous.pid)) throw failure('storage_locked');
            await fs.rm(filename);
            handle = await create();
        } finally {
            await recovery.close();
            await fs.rm(recoveryPath, { force: true });
        }
    }
    let released;
    return () => released ??= (async () => {
        await handle.close();
        const current = JSON.parse(await fs.readFile(filename, 'utf8'));
        if (current.token !== owner.token) throw failure('storage_owner_changed');
        await fs.rm(filename);
    })();
}

async function bindEnvironment(directory, publicKey) {
    const filename = path.join(directory, 'environment.json');
    const identity = { v: 1, environment: 'production', keyFingerprint: crypto.createHash('sha256').update(publicKey).digest('hex') };
    try {
        const previous = JSON.parse(await fs.readFile(filename, 'utf8'));
        if (JSON.stringify(previous) !== JSON.stringify(identity)) throw failure('environment_mismatch');
    } catch (error) {
        if (error.code !== 'ENOENT') throw failure('environment_mismatch');
        if ((await fs.readdir(directory)).some(name => name.endsWith('.json'))) throw failure('environment_mismatch');
        await fs.writeFile(filename, JSON.stringify(identity), { flag: 'wx', mode: 0o600 });
    }
}

function createSupportBackup(directory, now = Date.now) {
    const root = path.join(directory, 'backups');
    let pending;
    let savedDate;
    async function snapshot(date) {
        await fs.mkdir(root, { recursive: true, mode: 0o700 });
        if ((await fs.lstat(root)).isSymbolicLink() || path.relative(await fs.realpath(root), root) !== '') throw failure('invalid_backup_directory');
        const target = path.join(root, date);
        if (!await fs.lstat(target).then(stat => {
            if (!stat.isDirectory() || stat.isSymbolicLink()) throw failure('invalid_backup_directory');
            return true;
        }, error => { if (error.code === 'ENOENT') return false; throw error; })) {
            const temporary = path.join(root, `.pending-${crypto.randomUUID()}`);
            await fs.mkdir(temporary, { mode: 0o700 });
            try {
                for (const name of (await fs.readdir(directory)).filter(name => SNAPSHOT_FILE.test(name)).sort()) {
                    const source = path.join(directory, name);
                    const stat = await fs.lstat(source);
                    if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 4 * 1024 * 1024) throw failure('invalid_backup_source');
                    await fs.writeFile(path.join(temporary, name), await fs.readFile(source), { flag: 'wx', mode: 0o600 });
                }
                await fs.rename(temporary, target);
            } catch (error) {
                if (within(root, temporary) && temporary !== root) await fs.rm(temporary, { recursive: true, force: true });
                throw error;
            }
        }
        const names = (await fs.readdir(root, { withFileTypes: true }))
            .filter(entry => entry.isDirectory() && !entry.isSymbolicLink() && /^\d{4}-\d{2}-\d{2}$/u.test(entry.name))
            .map(entry => entry.name).sort().reverse();
        for (const name of names.slice(7)) {
            const expired = path.resolve(root, name);
            if (!within(root, expired) || expired === root) throw failure('invalid_backup_directory');
            await fs.rm(expired, { recursive: true });
        }
    }
    return {
        async ensure() {
            const date = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date(now()));
            if (savedDate === date) return;
            if (pending) { await pending; return this.ensure(); }
            pending = snapshot(date);
            try { await pending; savedDate = date; }
            finally { pending = undefined; }
        },
    };
}

function backedStore(store, mutations, backup) {
    return new Proxy(store, { get(target, name) {
        const value = Reflect.get(target, name);
        if (typeof value !== 'function') return value;
        if (!mutations.includes(name)) return value.bind(target);
        return async (...args) => { await backup.ensure(); return value.apply(target, args); };
    } });
}

function createSupportProductionRuntime({ environment = process.env, graphClient, createApplication = createSupportApp,
    factories = {}, now = Date.now } = {}) {
    const settings = productionSettings(environment);
    if (!settings) return undefined;
    let initialization;
    let release;
    let stopped = false;
    async function initialize() {
        const directory = await privateDirectory(settings.directory);
        release = await acquireProductionLock(directory, settings.instance);
        try {
            await bindEnvironment(directory, settings.publicKey);
            const backup = createSupportBackup(directory, now);
            await backup.ensure();
            const sessionStore = backedStore(await (factories.openSessionStore || openSupportSessionStore)(directory), ['set', 'delete', 'prune'], backup);
            const bookingStore = backedStore((factories.createBookingStore || createSupportBookingStore)(directory), ['write', 'remove', 'writeAvailabilityIntent', 'clearAvailabilityIntent'], backup);
            const cardStore = backedStore((factories.createCardStore || createSupportCardStore)(directory), ['write'], backup);
            const appointments = (factories.createAppointments || createSupportAppointmentsAdapter)({ graphClient, availabilityJournal: {
                read: () => bookingStore.readAvailabilityIntent(), write: value => bookingStore.writeAvailabilityIntent(value),
                clear: () => bookingStore.clearAvailabilityIntent(),
            } });
            return createApplication({ environment: 'production', origin: 'https://machadogestao.com', apiOnly: true, preload: true,
                users: (factories.createUsers || createSupportUsersAdapter)({ graphClient, environment: 'production' }),
                pagarme: (factories.createPagarme || createPagarmeAdapter)({ publicKey: settings.publicKey, secretKey: settings.secretKey, environment: 'production' }),
                sessionStore, bookingStore, cardStore, appointments,
                bookingNotifications: (factories.createNotifications || createSupportBookingNotifications)({ graphClient }),
                publicKey: settings.publicKey, now });
        } catch (error) {
            await release();
            release = undefined;
            throw error;
        }
    }
    function ready() {
        if (stopped) return Promise.reject(failure('stopped'));
        if (!initialization) {
            const pending = initialize();
            initialization = pending;
            // Metadata initialization can retry after an overlapping worker releases its owner.
            pending.catch(() => { if (initialization === pending) initialization = undefined; });
        }
        return initialization;
    }
    return {
        ready,
        async middleware(req, res, next) {
            try { return (await ready())(req, res, next); }
            catch { return res.status(503).json({ error: 'support_management.service_unavailable' }); }
        },
        async stop() {
            stopped = true;
            const pending = initialization;
            if (pending) await pending.catch(() => {});
            if (release) { await release(); release = undefined; }
        },
    };
}

module.exports = { createSupportProductionRuntime, productionSettings, acquireProductionLock, createSupportBackup };
