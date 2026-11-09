'use strict';

require('./deny-production-network');

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { createSupportBookingStore } = require('../integrations/support-booking-store');

const ID = 'd0b07b46-1311-41fb-984f-b338fd6b8a16';
const OTHER_ID = 'e0b07b46-1311-41fb-984f-b338fd6b8a16';

function record(overrides = {}) {
    return { id: ID, date: '2026-11-10', companyId: '12345678000195', name: 'João da Silva',
        company: 'Empresa de Consultoria', email: 'user@example.invalid', priceCents: 100000,
        createdAt: Date.parse('2026-11-09T15:00:00Z'), stage: 'workbook', ...overrides };
}

async function directory(t) {
    const root = await fs.mkdtemp(path.join(os.tmpdir(), 'support-booking-store-test-'));
    t.after(() => fs.rm(root, { recursive: true, force: true }));
    return root;
}

function filename(root, id = ID) { return path.join(root, `reservation_${id}.json`); }

test('booking store construction and missing reads do not create storage', async (t) => {
    const root = path.join(await directory(t), 'not-created');
    const store = createSupportBookingStore(root);
    assert.equal(await store.read(ID), null);
    assert.deepEqual(await store.pendingDates(), new Set());
    await assert.rejects(fs.stat(root), { code: 'ENOENT' });
});

test('valid booking stages survive reloads and pending dates exclude completed records', async (t) => {
    const root = await directory(t);
    const store = createSupportBookingStore(root);
    for (const stage of ['workbook', 'event', 'confirmation', 'complete']) {
        const value = record({ stage, ...(['confirmation', 'complete'].includes(stage)
            ? { eventId: 'provider-event-id', iCalUid: 'provider-calendar-uid' } : {}) });
        await store.write(value);
        assert.deepEqual(await createSupportBookingStore(root).read(ID), value);
        assert.deepEqual(await store.pendingDates(), new Set(stage === 'complete' ? [] : ['2026-11-10']));
        assert.deepEqual(await fs.readdir(root), [`reservation_${ID}.json`]);
    }
    await store.write(record({ stage: 'event' }));
    await store.write(record({ id: OTHER_ID, stage: 'event' }));
    assert.deepEqual(await store.pendingDates(), new Set(['2026-11-10']));
    await store.remove(ID);
    assert.equal(await store.read(ID), null);
    assert.deepEqual(await store.pendingDates(), new Set(['2026-11-10']));
    await store.remove(OTHER_ID);
    await store.remove(OTHER_ID);
    assert.deepEqual(await fs.readdir(root), []);
});

test('journal rejects unsafe or malformed records before replacing an existing booking', async (t) => {
    const root = await directory(t);
    const store = createSupportBookingStore(root);
    const previous = record();
    await store.write(previous);
    const invalid = [null, [], 'text', 42,
        record({ number: 'synthetic-sensitive-field' }), record({ id: '../escape' }), record({ id: [ID] }),
        record({ id: null }), record({ date: '2026-02-31' }), record({ date: '2026-13-10' }),
        record({ companyId: '123' }), record({ companyId: 12345678000195 }), record({ email: '' }),
        record({ email: 'user@example.invalid\r\nInjected: bad' }), record({ name: '' }),
        record({ name: 'x'.repeat(256) }), record({ name: 'Name\rInjected' }), record({ company: '' }),
        record({ company: 'Company\nInjected' }), record({ company: 'x'.repeat(4097) }),
        record({ priceCents: -1 }), record({ priceCents: 1.5 }), record({ createdAt: 0 }),
        record({ createdAt: NaN }), record({ stage: 'unknown' }), record({ stage: 'confirmation' }),
        record({ stage: 'complete', eventId: 'event-only' }), record({ eventId: 'event\r\nInjected' }),
        record({ iCalUid: 'uid\nInjected' }), record({ eventId: 'x'.repeat(2049) }), record({ iCalUid: 'x'.repeat(1025) })];
    for (const value of invalid) {
        await assert.rejects(store.write(value), /Invalid support booking record/u);
        assert.deepEqual(await store.read(ID), previous);
        assert.deepEqual(await fs.readdir(root), [`reservation_${ID}.json`]);
    }
    await store.write(record({ date: '2028-02-29' }));
    assert.equal((await store.read(ID)).date, '2028-02-29');
});

test('read and remove cannot escape the journal directory through an invalid ID', async (t) => {
    const root = await directory(t);
    const store = createSupportBookingStore(root);
    const sentinel = path.join(root, 'sentinel.txt');
    await fs.writeFile(sentinel, 'keep');
    for (const id of ['../sentinel', '..\\sentinel', `/absolute/${ID}`, '', null, [ID], `${ID}.json`, ID.toUpperCase()]) {
        await assert.rejects(store.read(id), /Support booking storage unavailable/u);
        await assert.rejects(store.remove(id), /Invalid support booking ID/u);
    }
    assert.equal(await fs.readFile(sentinel, 'utf8'), 'keep');
    assert.deepEqual(await fs.readdir(root), ['sentinel.txt']);
});

test('corrupt, malformed or mismatched records fail closed instead of freeing their pending date', async (t) => {
    const root = await directory(t);
    const store = createSupportBookingStore(root);
    for (const value of ['not-json', '{}', JSON.stringify(record({ id: OTHER_ID })),
        JSON.stringify(record({ stage: 'complete' }))]) {
        await fs.writeFile(filename(root), value);
        await assert.rejects(store.read(ID), /Support booking storage unavailable/u);
        await assert.rejects(store.pendingDates(), /Support booking storage unavailable/u);
    }
});

test('pending-date scan ignores temporary and unrelated files but checks every final journal entry', async (t) => {
    const root = await directory(t);
    const store = createSupportBookingStore(root);
    await store.write(record({ stage: 'event' }));
    await store.write(record({ id: OTHER_ID, date: '2026-11-11', stage: 'complete', eventId: 'event', iCalUid: 'uid' }));
    await fs.writeFile(path.join(root, 'reservation_incomplete.json.random.tmp'), 'not-json');
    await fs.writeFile(path.join(root, 'unrelated.json'), 'not-json');
    assert.deepEqual(await store.pendingDates(), new Set(['2026-11-10']));
    await fs.writeFile(path.join(root, 'reservation_invalid-id.json'), '{}');
    await assert.rejects(store.pendingDates(), /Support booking storage unavailable/u);
});

test('atomic journal replacement writes an exclusive temporary file before rename', async (t) => {
    const root = await directory(t);
    const store = createSupportBookingStore(path.join(root, 'new-directory'));
    const writes = [];
    const renames = [];
    const writeFile = fs.writeFile;
    const rename = fs.rename;
    t.mock.method(fs, 'writeFile', async (target, body, options) => {
        writes.push({ target, body, options });
        return writeFile(target, body, options);
    });
    t.mock.method(fs, 'rename', async (source, target) => {
        renames.push({ source, target });
        return rename(source, target);
    });
    await store.write(record());
    assert.equal(writes.length, 1);
    assert.equal(renames.length, 1);
    const target = filename(path.join(root, 'new-directory'));
    assert.equal(renames[0].target, target);
    assert.equal(renames[0].source, writes[0].target);
    assert.ok(writes[0].target.startsWith(`${target}.`));
    assert.ok(writes[0].target.endsWith('.tmp'));
    assert.deepEqual(writes[0].options, { flag: 'wx', mode: 0o600 });
    assert.deepEqual(JSON.parse(writes[0].body), record());
    assert.deepEqual(await fs.readdir(path.join(root, 'new-directory')), [`reservation_${ID}.json`]);
});

test('failed replacement preserves the previous final record and removes the temporary file', async (t) => {
    const root = await directory(t);
    const store = createSupportBookingStore(root);
    await store.write(record());
    const rename = t.mock.method(fs, 'rename', async () => { throw new Error('synthetic rename failure'); });
    await assert.rejects(store.write(record({ stage: 'event' })), /synthetic rename failure/u);
    rename.mock.restore();
    assert.deepEqual(await store.read(ID), record());
    assert.deepEqual(await fs.readdir(root), [`reservation_${ID}.json`]);
    await store.write(record({ stage: 'event' }));
    assert.equal((await store.read(ID)).stage, 'event');
});

test('failed temporary-file write cleans partial content without creating a final record', async (t) => {
    const root = await directory(t);
    const store = createSupportBookingStore(root);
    const writeFile = fs.writeFile;
    const write = t.mock.method(fs, 'writeFile', async (target, body, options) => {
        await writeFile(target, body.slice(0, 12), options);
        throw new Error('synthetic disk-full failure');
    });
    await assert.rejects(store.write(record()), /synthetic disk-full failure/u);
    write.mock.restore();
    assert.equal(await store.read(ID), null);
    assert.deepEqual(await fs.readdir(root), []);
});
