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
        { customerId: 'cus_Test1', cardId: 'card_Test1', cardRevision: 'a'.repeat(64) },
        { customerId: 'cus_Test1', cardId: 'card_Test1', pending: 'card' },
        { customerId: 'cus_Test1', cardId: 'card_Test1', pending: 'workbook' },
        { customerId: 'cus_Test1', cardId: 'card_Test1', pending: 'delete' },
        { customerId: 'cus_Test1', cardId: 'card_Test1', pending: 'delete_workbook' },
        { customerId: 'cus_Test1', cardId: 'card_Test1', workbookSynced: true },
        { customerId: 'cus_Test1', cardId: 'card_Test1', workbookSynced: false },
        { customerId: 'cus_Test1', cardId: 'card_Test1', verification: 'pagarme_verify_card_v1' },
        { customerId: 'cus_Test1', cardId: 'card_Test1', verification: 'pagarme_verify_card_v1', pending: 'workbook' },
        { customerId: 'cus_Test1', cardId: 'card_Test1', verification: 'pagarme_verify_card_v1', workbookSynced: true },
        { customerId: 'cus_Test1', cardId: 'card_Test1', verification: 'pagarme_verify_card_v1', workbookSynced: true, workbookReferencesSynced: true },
        { customerId: 'cus_Test1', cardId: 'card_Test1', verification: 'pagarme_verify_card_v1', workbookSynced: true, workbookReferencesSynced: true, pending: 'card' },
        { customerId: 'cus_Test1', cardId: 'card_Test1', verification: 'pagarme_verify_card_v1', workbookSynced: true, workbookReferencesSynced: true, pending: 'delete' },
        { customerId: 'cus_Test1', cardId: 'card_Test1', cardRevision: 'a'.repeat(64), verification: 'pagarme_verify_card_v1', workbookSynced: true, workbookReferencesSynced: true, pending: 'delete' },
        { customerId: 'cus_Test1', cardId: 'card_Test1', cardRevision: 'a'.repeat(64), pending: 'delete_workbook' },
        { customerId: 'cus_Test1', cardId: 'card_Test1', verification: 'pagarme_verify_card_v1', workbookSynced: true, pending: 'card' },
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
        { ...previous, cardRevision: '' },
        { ...previous, cardRevision: 'a'.repeat(63) },
        { ...previous, cardRevision: 'a'.repeat(65) },
        { ...previous, cardRevision: 'A'.repeat(64) },
        { ...previous, cardRevision: `${'a'.repeat(63)}\n` },
        { ...previous, cardRevision: 123 },
        { cardRevision: 'a'.repeat(64) },
        { customerId: 'cus_Test1', cardRevision: 'a'.repeat(64) },
        { cardId: 'card_Test1', cardRevision: 'a'.repeat(64) },
        { ...previous, billingAddress: {} },
        { customerId: 'cus_Path/../Other' },
        { customerId: 'cus_Test_with_underscores' },
        { customerId: 123 },
        { customerId: 'cus_' },
        { customerId: 'cus_Test1', cardId: 'card_' },
        { customerId: 'cus_Test1', cardId: 123 },
        { cardId: 'card_Test1' },
        { pending: 'card' },
        { pending: 'workbook' },
        { customerId: 'cus_Test1', pending: 'workbook' },
        { pending: 'delete' },
        { pending: 'delete_workbook' },
        { customerId: 'cus_Test1', pending: 'delete' },
        { customerId: 'cus_Test1', pending: 'delete_workbook' },
        { cardId: 'card_Test1', pending: 'delete' },
        { cardId: 'card_Test1', pending: 'delete_workbook' },
        { workbookSynced: true },
        { customerId: 'cus_Test1', workbookSynced: true },
        { ...previous, workbookSynced: 'true' },
        { ...previous, workbookSynced: true, pending: 'workbook' },
        { workbookReferencesSynced: true },
        { ...previous, workbookReferencesSynced: true },
        { ...previous, workbookSynced: false, workbookReferencesSynced: true },
        { ...previous, workbookSynced: true, workbookReferencesSynced: false },
        { ...previous, workbookSynced: true, workbookReferencesSynced: 'true' },
        { ...previous, workbookSynced: true, workbookReferencesSynced: true, pending: 'workbook' },
        { verification: 'pagarme_verify_card_v1' },
        { customerId: 'cus_Test1', verification: 'pagarme_verify_card_v1' },
        { ...previous, verification: true },
        { ...previous, verification: 'active' },
        { ...previous, verification: 'pagarme_verify_card_v2' },
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

test('reference recovery matches both identifiers and preserves verification and pending state across reloads', async (t) => {
    const root = await directory(t);
    const store = createSupportCardStore(root);
    const references = { customerId: 'cus_Test1', cardId: 'card_Test1' };
    const record = { ...references, cardRevision: 'a'.repeat(64), verification: 'pagarme_verify_card_v1', pending: 'card' };
    await store.write(company, record);
    const reloaded = createSupportCardStore(root);
    assert.deepEqual(await reloaded.findByReferences(references), record);
    assert.equal(await reloaded.findByReferences({ ...references, customerId: 'cus_Other' }), null);
    assert.equal(await reloaded.findByReferences({ ...references, cardId: 'card_Other' }), null);
    assert.equal(await reloaded.findByReferences({ ...references, customerId: 'cus_test1' }), null);
    assert.deepEqual(await reloaded.read(company), record);
});

test('reference recovery returns unverified records without inventing verification evidence', async (t) => {
    const root = await directory(t);
    const store = createSupportCardStore(root);
    const record = { customerId: 'cus_Test1', cardId: 'card_Test1' };
    await store.write(company, record);
    assert.deepEqual(await store.findByReferences(record), record);
    assert.equal(Object.hasOwn(await store.findByReferences(record), 'verification'), false);
});

test('reference recovery tolerates absent storage and ignores unrelated filenames', async (t) => {
    const root = await directory(t);
    const references = { customerId: 'cus_Test1', cardId: 'card_Test1' };
    assert.equal(await createSupportCardStore(path.join(root, 'not-created')).findByReferences(references), null);
    await fs.writeFile(path.join(root, 'sessions.json'), '{private malformed JSON');
    await fs.writeFile(path.join(root, `${company}.json.partial.tmp`), '{private malformed JSON');
    await fs.writeFile(path.join(root, 'support_not-a-valid-company.json'), '{private malformed JSON');
    assert.equal(await createSupportCardStore(root).findByReferences(references), null);
});

test('reference recovery rejects multiple exact matches instead of selecting a record', async (t) => {
    const root = await directory(t);
    const store = createSupportCardStore(root);
    const references = { customerId: 'cus_Test1', cardId: 'card_Test1' };
    const secondCompany = `support_${'b'.repeat(44)}`;
    await store.write(company, { ...references, verification: 'pagarme_verify_card_v1' });
    await store.write(secondCompany, references);
    await assert.rejects(store.findByReferences(references), { message: 'Ambiguous local card references' });
    assert.deepEqual((await fs.readdir(root)).sort(), [`${company}.json`, `${secondCompany}.json`].sort());
});

test('reference recovery fails closed on malformed stored records or lookup identifiers', async (t) => {
    const root = await directory(t);
    const store = createSupportCardStore(root);
    const references = { customerId: 'cus_Test1', cardId: 'card_Test1' };
    for (const invalid of [{}, { customerId: 'cus_Test1' }, { cardId: 'card_Test1' }, { ...references, cardId: 'card_bad/id' }]) {
        await assert.rejects(store.findByReferences(invalid));
    }
    for (const content of ['{malformed JSON', JSON.stringify({ ...references, verification: 'active' }), JSON.stringify({ ...references, cvv: '123' })]) {
        await fs.writeFile(path.join(root, `${company}.json`), content);
        await assert.rejects(store.findByReferences(references));
    }
});
