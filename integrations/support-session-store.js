'use strict';

const crypto = require('node:crypto');
const fs = require('node:fs/promises');
const path = require('node:path');

const MAX_SESSIONS = 500;
const MAX_FILE_BYTES = 2 * 1024 * 1024;
const HEX = /^[a-f0-9]{64}$/u;
const normalize = (value) => value.trim().replace(/\s+/gu, ' ').toLowerCase();

function exactKeys(value, keys) {
    return value && typeof value === 'object' && !Array.isArray(value)
        && Object.keys(value).length === keys.length && keys.every((key) => Object.hasOwn(value, key));
}

function validateSession(record) {
    if (!exactKeys(record, ['email', 'company', 'fingerprint', 'expiresAt'])
        || typeof record.email !== 'string' || !record.email || record.email.length > 254 || normalize(record.email) !== record.email
        || typeof record.company !== 'string' || !record.company || record.company.length > 4096 || normalize(record.company) !== record.company
        || typeof record.fingerprint !== 'string' || !HEX.test(record.fingerprint)
        || !Number.isSafeInteger(record.expiresAt) || record.expiresAt <= 0) {
        throw new Error('Invalid support session record');
    }
    return { email: record.email, company: record.company, fingerprint: record.fingerprint, expiresAt: record.expiresAt };
}

function validateHash(hash) {
    if (typeof hash !== 'string' || !HEX.test(hash)) throw new Error('Invalid support session hash');
}

function readSnapshot(snapshot) {
    if (!exactKeys(snapshot, ['v', 'salt', 'sessions']) || snapshot.v !== 1
        || typeof snapshot.salt !== 'string' || !HEX.test(snapshot.salt)
        || !Array.isArray(snapshot.sessions) || snapshot.sessions.length > MAX_SESSIONS) {
        throw new Error('Invalid support session snapshot');
    }
    const sessions = new Map();
    for (const entry of snapshot.sessions) {
        if (!exactKeys(entry, ['tokenHash', 'email', 'company', 'fingerprint', 'expiresAt'])) throw new Error('Invalid support session snapshot');
        validateHash(entry.tokenHash);
        if (sessions.has(entry.tokenHash)) throw new Error('Invalid support session snapshot');
        const { tokenHash, ...record } = entry;
        sessions.set(tokenHash, validateSession(record));
    }
    return { salt: snapshot.salt, sessions };
}

function createStore({ salt, sessions = new Map(), persist = async () => {} }) {
    let current = sessions;
    let pending = Promise.resolve();
    function update(change) {
        const operation = pending.then(async () => {
            const next = new Map(current);
            if (!change(next)) return;
            await persist(next);
            current = next;
        });
        pending = operation.catch(() => {});
        return operation;
    }
    return {
        get salt() { return salt; },
        get size() { return current.size; },
        async get(hash) {
            validateHash(hash);
            await pending;
            const record = current.get(hash);
            return record ? { ...record } : undefined;
        },
        async set(hash, record) {
            validateHash(hash);
            const safeRecord = validateSession(record);
            return update((next) => {
                if (!next.has(hash) && next.size >= MAX_SESSIONS) throw new Error('Support session limit reached');
                next.set(hash, safeRecord);
                return true;
            });
        },
        async delete(hash) {
            validateHash(hash);
            return update((next) => next.delete(hash));
        },
        async prune(nowMs) {
            if (!Number.isSafeInteger(nowMs) || nowMs < 0) throw new Error('Invalid support session time');
            return update((next) => {
                let changed = false;
                for (const [hash, record] of next) {
                    if (record.expiresAt <= nowMs) { next.delete(hash); changed = true; }
                }
                return changed;
            });
        },
    };
}

function createMemorySupportSessionStore() {
    return createStore({ salt: crypto.randomBytes(32).toString('hex') });
}

async function openSupportSessionStore(directory) {
    const filename = path.join(directory, 'sessions.json');
    let snapshot;
    try {
        await fs.mkdir(directory, { recursive: true });
        const file = await fs.open(filename, 'r');
        try {
            const stat = await file.stat();
            if (!stat.isFile() || stat.size > MAX_FILE_BYTES) throw new Error('Invalid support session storage');
            snapshot = readSnapshot(JSON.parse(await file.readFile('utf8')));
        } finally {
            await file.close();
        }
    } catch (error) {
        if (error.code !== 'ENOENT') throw new Error('Support session storage unavailable');
        snapshot = { salt: crypto.randomBytes(32).toString('hex'), sessions: new Map() };
    }
    async function persist(sessions) {
        const temporary = `${filename}.${crypto.randomUUID()}.tmp`;
        try {
            const contents = JSON.stringify({ v: 1, salt: snapshot.salt,
                sessions: Array.from(sessions, ([tokenHash, record]) => ({ tokenHash, ...record })),
            });
            if (Buffer.byteLength(contents) > MAX_FILE_BYTES) throw new Error('Invalid support session storage');
            await fs.writeFile(temporary, contents, { flag: 'wx', mode: 0o600 });
            await fs.rename(temporary, filename);
        } catch {
            await fs.rm(temporary, { force: true }).catch(() => {});
            throw new Error('Support session storage unavailable');
        }
    }
    // The salt must survive a restart even when no account has signed in yet.
    if (!await fs.stat(filename).then(() => true, (error) => { if (error.code === 'ENOENT') return false; throw error; })) await persist(snapshot.sessions);
    return createStore({ ...snapshot, persist });
}

module.exports = { openSupportSessionStore, createMemorySupportSessionStore };
