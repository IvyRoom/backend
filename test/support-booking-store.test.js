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
    assert.equal(await store.readAvailabilityIntent(), null);
    await store.clearAvailabilityIntent();
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
        record({ acknowledgedAt: 0 }), record({ acknowledgedAt: -1 }), record({ acknowledgedAt: 1.5 }),
        record({ acknowledgedAt: Number.MAX_SAFE_INTEGER + 1 }), record({ acknowledgedAt: '1800000000000' }),
        record({ meetingUrl: 'http://teams.microsoft.com/meeting' }), record({ meetingUrl: 'https://teams.microsoft.com.attacker.invalid/meeting' }),
        record({ meetingUrl: 'https://attacker.invalid/meeting' }), record({ meetingUrl: 'https://user:password@teams.microsoft.com/meeting' }),
        record({ meetingUrl: null }), record({ meetingUrl: 42 }), record({ meetingUrl: `https://teams.microsoft.com/${'x'.repeat(4096)}` }),
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

test('acknowledgment time and Teams meeting links survive journal reload without becoming required for older bookings', async (t) => {
    const root = await directory(t);
    const store = createSupportBookingStore(root);
    await store.write(record());
    assert.deepEqual(await createSupportBookingStore(root).read(ID), record());
    for (const hostname of ['teams.microsoft.com', 'teams.live.com', 'teams.cloud.microsoft']) {
        const value = record({ acknowledgedAt: 1_800_000_000_000, meetingUrl: `https://${hostname}/meet/synthetic` });
        await store.write(value);
        assert.deepEqual(await createSupportBookingStore(root).read(ID), value);
    }
});

test('availability intent persists an ordered batch separately from reservations and clears idempotently', async (t) => {
    const root = await directory(t);
    const store = createSupportBookingStore(root);
    const intent = { dates: ['2026-11-10', '2026-11-11', '2026-11-12'] };
    await store.write(record());
    await store.writeAvailabilityIntent(intent);
    assert.deepEqual(await createSupportBookingStore(root).readAvailabilityIntent(), intent);
    assert.deepEqual(await store.pendingDates(), new Set(['2026-11-10']));
    assert.deepEqual(await store.read(ID), record());
    const calendarWindow = { dates: Array.from({ length: 62 }, (_, index) => new Date(Date.UTC(2028, 0, index + 1)).toISOString().slice(0, 10)) };
    await store.writeAvailabilityIntent(calendarWindow);
    assert.deepEqual(await store.readAvailabilityIntent(), calendarWindow);
    await store.clearAvailabilityIntent();
    await store.clearAvailabilityIntent();
    assert.equal(await store.readAvailabilityIntent(), null);
    assert.deepEqual(await fs.readdir(root), [`reservation_${ID}.json`]);
});

test('invalid availability intents cannot replace the last durable batch', async (t) => {
    const root = await directory(t);
    const store = createSupportBookingStore(root);
    const previous = { dates: ['2028-02-29', '2028-03-01'] };
    await store.writeAvailabilityIntent(previous);
    for (const value of [null, [], 'text', 42, {}, { dates: [] }, { dates: '2026-11-10' },
        { dates: [null] }, { dates: [20261110] }, { dates: ['2026-02-29'] }, { dates: ['2026-02-31'] },
        { dates: ['2026-13-01'] }, { dates: ['2026-11-10T00:00:00Z'] }, { dates: ['../escape'] },
        { dates: ['2026-11-10', '2026-11-10'] }, { dates: ['2026-11-11', '2026-11-10'] },
        { dates: ['2026-11-10'], extra: true },
        { dates: Array.from({ length: 63 }, (_, index) => new Date(Date.UTC(2026, 10, index + 1)).toISOString().slice(0, 10)) }]) {
        await assert.rejects(store.writeAvailabilityIntent(value), /Invalid availability intent/u);
        assert.deepEqual(await store.readAvailabilityIntent(), previous);
        assert.deepEqual(await fs.readdir(root), ['availability-sync.json']);
    }
});

test('corrupt availability intent fails closed and a failed replacement preserves the previous batch', async (t) => {
    const root = await directory(t);
    const store = createSupportBookingStore(root);
    const target = path.join(root, 'availability-sync.json');
    for (const value of ['not-json', '{}', JSON.stringify({ dates: ['2026-02-31'] }),
        JSON.stringify({ dates: ['2026-11-11', '2026-11-10'] })]) {
        await fs.writeFile(target, value);
        await assert.rejects(store.readAvailabilityIntent(), /Availability storage unavailable/u);
        assert.equal(await fs.readFile(target, 'utf8'), value);
    }
    const previous = { dates: ['2026-11-10'] };
    await store.writeAvailabilityIntent(previous);
    const rename = t.mock.method(fs, 'rename', async () => { throw new Error('synthetic intent replacement failure'); });
    await assert.rejects(store.writeAvailabilityIntent({ dates: ['2026-11-11'] }), /synthetic intent replacement failure/u);
    rename.mock.restore();
    assert.deepEqual(await store.readAvailabilityIntent(), previous);
    assert.deepEqual(await fs.readdir(root), ['availability-sync.json']);
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
