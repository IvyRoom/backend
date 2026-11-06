'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { createSupportCardStore, acquireSupportSandboxLock } = require('../integrations/support-card-store');

const company = `support_${'a'.repeat(44)}`;

async function directory(t) {
    const result = await fs.mkdtemp(path.join(os.tmpdir(), 'support-card-state-test-'));
    t.after(() => fs.rm(result, { recursive: true, force: true }));
    return result;
}

test('sandbox lock refuses concurrent instances and releases idempotently before a restart', async (t) => {
    const root = await directory(t);
    const release = await acquireSupportSandboxLock(root);
    const lockPath = path.join(root, 'sandbox.lock');
    assert.equal(await fs.readFile(lockPath, 'utf8'), String(process.pid));
    await assert.rejects(acquireSupportSandboxLock(root), { code: 'EEXIST' });
    assert.equal(await fs.readFile(lockPath, 'utf8'), String(process.pid));
    await Promise.all([release(), release()]);
    assert.deepEqual(await fs.readdir(root), []);
    const releaseNext = await acquireSupportSandboxLock(root);
    await release();
    assert.equal(await fs.readFile(lockPath, 'utf8'), String(process.pid));
    await releaseNext();
    assert.deepEqual(await fs.readdir(root), []);
});

test('sandbox lock preserves an existing lock requiring manual review', async (t) => {
    const root = await directory(t);
    const lockPath = path.join(root, 'sandbox.lock');
    await fs.writeFile(lockPath, 'existing-instance');
    await assert.rejects(acquireSupportSandboxLock(root), { code: 'EEXIST' });
    assert.equal(await fs.readFile(lockPath, 'utf8'), 'existing-instance');
});

test('sandbox lock closes and removes its file when writing the lock fails', async (t) => {
    const root = await directory(t);
    const open = fs.open;
    const mocked = t.mock.method(fs, 'open', async (...args) => {
        const handle = await open(...args);
        handle.writeFile = async () => { throw new Error('simulated lock write failure'); };
        return handle;
    });
    await assert.rejects(acquireSupportSandboxLock(root), /simulated lock write failure/);
    mocked.mock.restore();
    assert.deepEqual(await fs.readdir(root), []);
    const release = await acquireSupportSandboxLock(root);
    await release();
});

test('card reference store accepts only valid reference and pending states across reloads', async (t) => {
    const root = await directory(t);
    const store = createSupportCardStore(root);
    for (const record of [
        {}, { pending: 'customer' }, { customerId: 'cus_Test1' },
        { customerId: 'cus_Test1', pending: 'card' },
        { customerId: 'cus_Test1', cardId: 'card_Test1' },
        { customerId: 'cus_Test1', cardId: 'card_Test1', pending: 'card' },
    ]) {
        await store.write(company, record);
        assert.deepEqual(await createSupportCardStore(root).read(company), record);
    }
});

test('card reference store refuses extra, sensitive, malformed or inconsistent fields on write and read', async (t) => {
    const root = await directory(t);
    const store = createSupportCardStore(root);
    const target = path.join(root, `${company}.json`);
    const previous = { customerId: 'cus_Test1', cardId: 'card_Test1' };
    const invalid = [
        null, [], 'text', 1,
        { ...previous, number: '4000000000000010' },
        { ...previous, cvv: '123' },
        { ...previous, token: 'token_Test1' },
        { ...previous, billingAddress: {} },
        { customerId: 'cus_Path/../Other' },
        { customerId: 'cus_Test_with_underscores' },
        { customerId: 123 },
        { customerId: 'cus_' },
        { customerId: 'cus_Test1', cardId: 'card_' },
        { customerId: 'cus_Test1', cardId: 123 },
        { cardId: 'card_Test1' },
        { pending: 'card' },
        { pending: 'unknown' },
        { pending: true },
        { customerId: 'cus_Test1', pending: 'customer' },
    ];
    for (const record of invalid) {
        await store.write(company, previous);
        await assert.rejects(store.write(company, record), { message: 'Invalid card reference record' });
        assert.deepEqual(await store.read(company), previous);
        assert.deepEqual(await fs.readdir(root), [`${company}.json`]);
        await fs.writeFile(target, JSON.stringify(record));
        await assert.rejects(store.read(company), { message: 'Card reference storage unavailable' });
    }
});
