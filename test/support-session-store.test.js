'use strict';

require('./deny-production-network');
const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { openSupportSessionStore, createMemorySupportSessionStore } = require('../integrations/support-session-store');

const hash = value => crypto.createHash('sha256').update(value).digest('hex');
const tokenHash = hash('a'.repeat(43));
const record = { email: 'admin@example.test', company: 'empresa a', fingerprint: 'f'.repeat(64), expiresAt: 1_900_000_000_000 };

async function directory(t) {
    const root = await fs.mkdtemp(path.join(os.tmpdir(), 'support-session-state-test-'));
    t.after(() => fs.rm(root, { recursive: true, force: true }));
    return root;
}

test('session store preserves its salt, token hashes and original expiry across reopen', async (t) => {
    const root = await directory(t);
    const store = await openSupportSessionStore(root);
    assert.match(store.salt, /^[a-f0-9]{64}$/u);
    assert.equal(store.size, 0);
    const empty = await openSupportSessionStore(root);
    assert.equal(empty.salt, store.salt, 'The fingerprint salt must survive even before the first login');
    await store.set(tokenHash, record);
    const reopened = await openSupportSessionStore(root);
    assert.equal(reopened.salt, store.salt);
    assert.equal(reopened.size, 1);
    assert.deepEqual(await reopened.get(tokenHash), record);
    const persisted = await fs.readFile(path.join(root, 'sessions.json'), 'utf8');
    assert.equal(persisted.includes('a'.repeat(43)), false, 'A bearer token must never be persisted');
    assert.equal(persisted.includes(tokenHash), true);
    assert.equal(persisted.includes('password'), false);
    assert.equal(persisted.includes('cvv'), false);
});

test('memory and disk stores isolate caller-owned records and returned records', async (t) => {
    const stores = [createMemorySupportSessionStore(), await openSupportSessionStore(await directory(t))];
    for (const store of stores) {
        const original = { ...record };
        await store.set(tokenHash, original);
        original.email = 'changed@example.test';
        const returned = await store.get(tokenHash);
        assert.deepEqual(returned, record);
        returned.company = 'another company';
        assert.deepEqual(await store.get(tokenHash), record);
    }
});

test('concurrent session mutations are serialized without losing logins or reviving a logout', async (t) => {
    const root = await directory(t);
    const store = await openSupportSessionStore(root);
    const hashes = Array.from({ length: 12 }, (_, index) => hash(`session-${index}`));
    await Promise.all(hashes.map(key => store.set(key, record)));
    assert.equal(store.size, hashes.length);
    const addition = hash('new-session');
    await Promise.all([
        ...hashes.slice(0, 5).map(key => store.delete(key)),
        store.set(addition, record),
    ]);
    const reopened = await openSupportSessionStore(root);
    assert.equal(reopened.size, 8);
    for (const key of hashes.slice(0, 5)) assert.equal(await reopened.get(key), undefined);
    for (const key of [...hashes.slice(5), addition]) assert.deepEqual(await reopened.get(key), record);
});

test('logout deletion and expiry pruning survive restart without extending expiry', async (t) => {
    const root = await directory(t);
    const store = await openSupportSessionStore(root);
    const expiresNow = hash('expires-now');
    const expiresLater = hash('expires-later');
    await store.set(tokenHash, record);
    await store.set(expiresNow, { ...record, expiresAt: record.expiresAt - 1 });
    await store.set(expiresLater, { ...record, expiresAt: record.expiresAt + 1 });
    await store.delete(tokenHash);
    await store.prune(record.expiresAt - 1);
    const reopened = await openSupportSessionStore(root);
    assert.equal(await reopened.get(tokenHash), undefined);
    assert.equal(await reopened.get(expiresNow), undefined);
    assert.equal((await reopened.get(expiresLater)).expiresAt, record.expiresAt + 1);
    await reopened.prune(record.expiresAt + 1);
    assert.equal((await openSupportSessionStore(root)).size, 0);
});

test('failed persistence does not commit a login or logout in memory and permits a later retry', async (t) => {
    const root = await directory(t);
    const store = await openSupportSessionStore(root);
    await store.set(tokenHash, record);
    const target = path.join(root, 'sessions.json');
    const backup = path.join(root, 'previous-session-state.json');
    await fs.rename(target, backup);
    await fs.mkdir(target);
    const addition = hash('failed-login');
    await assert.rejects(store.set(addition, record));
    assert.equal(await store.get(addition), undefined);
    assert.equal(store.size, 1);
    await assert.rejects(store.delete(tokenHash));
    assert.deepEqual(await store.get(tokenHash), record);
    await fs.rmdir(target);
    await fs.rename(backup, target);
    await store.set(addition, record);
    await store.delete(tokenHash);
    const reopened = await openSupportSessionStore(root);
    assert.equal(await reopened.get(tokenHash), undefined);
    assert.deepEqual(await reopened.get(addition), record);
});

test('session stores refuse raw tokens, unsafe fields and malformed records before mutation', async (t) => {
    const root = await directory(t);
    const stores = [createMemorySupportSessionStore(), await openSupportSessionStore(root)];
    for (const store of stores) {
        for (const key of ['a'.repeat(43), '../escape', tokenHash.toUpperCase(), '', null]) {
            await assert.rejects(store.set(key, record));
            await assert.rejects(store.get(key));
            await assert.rejects(store.delete(key));
        }
        for (const invalid of [
            null, [], {}, { ...record, password: 'private-password' }, { ...record, token: 'a'.repeat(43) },
            { ...record, cvv: '123' }, { ...record, fingerprint: 'not-a-hash' },
            { ...record, expiresAt: 0 }, { ...record, expiresAt: -1 }, { ...record, expiresAt: 1.5 },
            { ...record, expiresAt: Number.MAX_SAFE_INTEGER + 1 }, { ...record, email: '' },
            { ...record, email: 'a'.repeat(255) }, { ...record, company: '' },
            { ...record, company: 'a'.repeat(4097) },
        ]) await assert.rejects(store.set(tokenHash, invalid));
        for (const now of [-1, 1.5, NaN]) await assert.rejects(store.prune(now));
        assert.equal(store.size, 0);
    }
});

test('corrupt or unknown disk state fails closed without overwriting it', async (t) => {
    const root = await directory(t);
    const store = await openSupportSessionStore(root);
    const target = path.join(root, 'sessions.json');
    const valid = { v: 1, salt: store.salt, sessions: [{ tokenHash, ...record }] };
    const invalid = [
        '{broken', JSON.stringify(null), JSON.stringify([]), JSON.stringify({ ...valid, v: 2 }),
        JSON.stringify({ ...valid, salt: 'bad' }), JSON.stringify({ ...valid, password: 'private-password' }),
        JSON.stringify({ ...valid, sessions: {} }),
        JSON.stringify({ ...valid, sessions: [{ ...record, tokenHash: 'a'.repeat(43) }] }),
        JSON.stringify({ ...valid, sessions: [{ tokenHash, ...record, token: 'a'.repeat(43) }] }),
        JSON.stringify({ ...valid, sessions: [{ tokenHash, ...record }, { tokenHash, ...record }] }),
        JSON.stringify({ ...valid, sessions: Array.from({ length: 501 }, (_, index) => ({ tokenHash: hash(String(index)), ...record })) }),
        ' '.repeat(2 * 1024 * 1024 + 1),
    ];
    for (const text of invalid) {
        await fs.writeFile(target, text);
        await assert.rejects(openSupportSessionStore(root));
        assert.equal(await fs.readFile(target, 'utf8'), text);
    }
});

test('session count stays bounded while a full store can revoke an existing session', async (t) => {
    const root = await directory(t);
    const store = await openSupportSessionStore(root);
    const sessions = Array.from({ length: 500 }, (_, index) => ({ tokenHash: hash(String(index)), ...record }));
    await fs.writeFile(path.join(root, 'sessions.json'), JSON.stringify({ v: 1, salt: store.salt, sessions }));
    const full = await openSupportSessionStore(root);
    assert.equal(full.size, 500);
    await assert.rejects(full.set(hash('overflow'), record));
    assert.equal(full.size, 500);
    await full.delete(sessions[0].tokenHash);
    await full.set(hash('replacement'), record);
    assert.equal((await openSupportSessionStore(root)).size, 500);
});
