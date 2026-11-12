'use strict';

require('./deny-production-network');

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const { createSupportBookingHandlers, calendarDays, eventPayload, confirmation } = require('../domains/support-bookings');
const { createSupportManagementHandlers, companyCode } = require('../domains/support-management');

const NOW = Date.parse('2026-11-09T15:00:00Z');
const ID = 'd0b07b46-1311-41fb-984f-b338fd6b8a16';
const OTHER_ID = 'e0b07b46-1311-41fb-984f-b338fd6b8a16';
const COMPANY_ID = '12345678000195';
const CONTACT = 'contato@machadogestao.com';
const WEEKDAYS = ['2026-11-10', '2026-11-11', '2026-11-12', '2026-11-13', '2026-11-16',
    '2026-11-17', '2026-11-18', '2026-11-19', '2026-11-23', '2026-11-24',
    '2026-11-25', '2026-11-26', '2026-11-27'];
const MASKED_CARD = { brand: 'Visa', lastFour: '1111' };
const MEETING_URL = 'https://teams.microsoft.com/l/meetup-join/synthetic-meeting?context=example&tenant=example';
const USER = { companyId: COMPANY_ID, company: 'Empresa de Consultoria', name: 'João da Silva',
    email: 'USER@example.invalid', referenceRate: 1000, bookingAllowed: true };

function savedRecord(overrides = {}) {
    return { id: ID, date: WEEKDAYS[0], companyId: COMPANY_ID, company: USER.company,
        name: USER.name, email: USER.email.toLowerCase(), priceCents: 100000, createdAt: NOW,
        stage: 'workbook', ...overrides };
}

function response(user = USER, directory = [user]) {
    return { locals: { supportUser: user, supportUsers: directory }, statusCode: 200, body: null,
        status(value) { this.statusCode = value; return this; },
        json(value) { this.body = value; return this; } };
}

function fixture({ rows = WEEKDAYS.map((date) => ({ date, status: '' })), records = [],
    faults = {}, reserveDate, writeRecord, user = USER, directory, users, card = MASKED_CARD,
    syncAvailability = false, nowMs = NOW, meetingUrl = MEETING_URL } = {}) {
    const calls = [];
    const calendarRows = [];
    const stored = new Map(records.map((record) => [record.id, structuredClone(record)]));
    const dates = structuredClone(rows);
    async function call(type, payload) {
        calls.push({ type, payload: payload === undefined ? undefined : structuredClone(payload) });
        if (faults[type]) throw faults[type];
    }
    const appointments = {
        async readDates() {
            await call('dates');
            const snapshot = structuredClone(dates);
            calendarRows.push(snapshot);
            return snapshot;
        },
        async reserveDate(payload, snapshot) {
            await call('reserve', payload);
            if (reserveDate) await reserveDate(payload, snapshot);
            dates.find((row) => row.date === payload.date).status = 'RESERVADO';
        },
    };
    if (syncAvailability) appointments.syncAvailability = async (payload) => {
        await call('sync', payload);
        const snapshot = structuredClone(dates);
        calendarRows.push(snapshot);
        return snapshot;
    };
    const handlers = createSupportBookingHandlers({
        now: () => nowMs,
        appointments, users,
        async companyCards(actingUser, options) {
            await call(options?.cached ? 'cards:cached' : 'cards:fresh', { user: actingUser, options });
            return { card: typeof card === 'function' ? card(options) : card };
        },
        bookingStore: {
            async read(id) { await call('read', id); return structuredClone(stored.get(id) || null); },
            async write(record) {
                await call(`write:${record.stage}`, record);
                if (writeRecord) await writeRecord(record);
                stored.set(record.id, structuredClone(record));
            },
            async remove(id) { await call('remove', id); stored.delete(id); },
            async pendingDates() {
                await call('pending');
                return new Set([...stored.values()].filter((record) => record.stage !== 'complete').map((record) => record.date));
            },
        },
        bookingNotifications: {
            async createEvent(payload) {
                await call('event', payload);
                return { eventId: 'graph-event-id', iCalUid: 'provider-calendar-uid@example.invalid', ...(meetingUrl ? { meetingUrl } : {}) };
            },
            async sendConfirmation(message) {
                await call(message.toRecipients[0].emailAddress.address === CONTACT ? 'notification' : 'mail', message);
            },
        },
    });
    return { calls, stored, handlers, calendarRows,
        async reserve(body = { date: WEEKDAYS[0], priceCents: 100000, requestId: ID, acknowledged: true }, actingUser = user) {
            const res = response(actingUser, directory);
            await handlers.reserve({ body }, res);
            return res;
        },
        async calendar() {
            const res = response(user, directory);
            await handlers.calendar({}, res);
            return res;
        } };
}

function assertError(res, status, code) {
    assert.equal(res.statusCode, status);
    assert.deepEqual(res.body, { error: `support_management.${code}` });
}

function deferredOperation() {
    let enter;
    let release;
    const started = new Promise((resolve) => { enter = resolve; });
    const pending = new Promise((resolve) => { release = resolve; });
    return { started, release, wait() { enter(); return pending; } };
}

function cardLifecycleFixture({ getCard, deleteCard, reserveDate, createEvent } = {}) {
    const user = { ...USER, role: 'Administrador' };
    const calls = [];
    const stored = new Map();
    const rows = WEEKDAYS.map((date) => ({ date, status: '' }));
    let record = { customerId: 'cus_synthetic', cardId: 'card_synthetic', verification: 'pagarme_verify_card_v1',
        workbookSynced: true, workbookReferencesSynced: true };
    const management = createSupportManagementHandlers({
        now: () => NOW, publicKey: 'pk_test_synthetic',
        users: { async clearCompanyCard(company) { calls.push(['clear', company]); } },
        cardStore: {
            async read(code) { assert.equal(code, companyCode(COMPANY_ID)); return structuredClone(record); },
            async write(code, value) { assert.equal(code, companyCode(COMPANY_ID)); record = structuredClone(value); },
        },
        pagarme: {
            async getCard() {
                calls.push(['getCard']);
                if (getCard) await getCard();
                return { ...MASKED_CARD, id: 'card_synthetic', status: 'active', expMonth: 12, expYear: 2030 };
            },
            async deleteCard() {
                calls.push(['deleteCard']);
                if (deleteCard) await deleteCard();
                return { deleted: true };
            },
        },
    });
    const bookings = createSupportBookingHandlers({
        now: () => NOW, companyCards: management.companyPayment, claimCompany: management.claimCompany,
        appointments: {
            async readDates() { calls.push(['dates']); return structuredClone(rows); },
            async reserveDate(payload) {
                calls.push(['reserve', payload.date]);
                if (reserveDate) await reserveDate(payload.date);
                rows.find((row) => row.date === payload.date).status = 'RESERVADO';
            },
        },
        bookingStore: {
            async read(id) { return stored.get(id); },
            async write(booking) { stored.set(booking.id, structuredClone(booking)); },
            async remove(id) { stored.delete(id); },
            async pendingDates() { return new Set([...stored.values()].filter((booking) => booking.stage !== 'complete').map((booking) => booking.date)); },
        },
        bookingNotifications: {
            async createEvent() {
                calls.push(['event']);
                if (createEvent) await createEvent();
                return { eventId: 'synthetic', iCalUid: 'synthetic@example.invalid', meetingUrl: MEETING_URL };
            },
            async sendConfirmation() { calls.push(['mail']); },
        },
    });
    return { calls, stored,
        async reserve(date = WEEKDAYS[0], requestId = ID) {
            const res = response(user);
            await bookings.reserve({ body: { date, priceCents: 100000, requestId, acknowledged: true } }, res);
            return res;
        },
        async delete() {
            const res = response(user);
            const expectedCardRevision = crypto.createHash('sha256').update('cus_synthetic:card_synthetic').digest('hex');
            await management.deleteCard({ body: { expectedCardRevision } }, res);
            return res;
        },
    };
}

test('card deletion cannot race a delayed active-card response into an unbacked reservation', async () => {
    const reading = deferredOperation();
    const app = cardLifecycleFixture({ getCard: () => reading.wait() });
    const booking = app.reserve();
    await reading.started;
    assertError(await app.delete(), 409, 'card_busy');
    assert.equal(app.calls.filter(([type]) => type === 'deleteCard' || type === 'clear').length, 0);
    reading.release();
    assert.equal((await booking).body.reserved, true);
    assert.deepEqual((await app.delete()).body, { card: null });
    assert.equal(app.calls.filter(([type]) => type === 'getCard').length, 1);
    assert.equal(app.calls.filter(([type]) => type === 'deleteCard').length, 1);
    assert.ok(app.calls.findIndex(([type]) => type === 'reserve') < app.calls.findIndex(([type]) => type === 'deleteCard'));
});

test('a reservation starting during card deletion makes no provider or booking writes and remains gated afterward', async () => {
    const deleting = deferredOperation();
    const app = cardLifecycleFixture({ deleteCard: () => deleting.wait() });
    const deletion = app.delete();
    await deleting.started;
    assertError(await app.reserve(), 409, 'card_busy');
    assert.equal(app.stored.size, 0);
    assert.equal(app.calls.filter(([type]) => ['getCard', 'reserve', 'event', 'mail'].includes(type)).length, 0);
    deleting.release();
    assert.deepEqual((await deletion).body, { card: null });
    assertError(await app.reserve(), 403, 'payment_method_required');
    assert.equal(app.stored.size, 0);
});

test('parallel same-company bookings retain their shared card guard until every workbook write completes', async () => {
    const first = deferredOperation();
    const second = deferredOperation();
    const app = cardLifecycleFixture({ reserveDate: (date) => date === WEEKDAYS[0] ? first.wait() : second.wait() });
    const bookingOne = app.reserve();
    await first.started;
    const bookingTwo = app.reserve(WEEKDAYS[1], OTHER_ID);
    await second.started;
    assertError(await app.delete(), 409, 'card_busy');
    first.release();
    assert.equal((await bookingOne).body.reserved, true);
    assertError(await app.delete(), 409, 'card_busy');
    second.release();
    assert.equal((await bookingTwo).body.reserved, true);
    assert.deepEqual((await app.delete()).body, { card: null });
    assert.equal(app.calls.filter(([type]) => type === 'getCard').length, 2);
    assert.equal(app.calls.filter(([type]) => type === 'reserve').length, 2);
});

test('the company card guard is released before booking notifications and after payment-check failures', async () => {
    const notifying = deferredOperation();
    const app = cardLifecycleFixture({ createEvent: () => notifying.wait() });
    const booking = app.reserve();
    await notifying.started;
    assert.deepEqual((await app.delete()).body, { card: null });
    notifying.release();
    assert.equal((await booking).body.reserved, true);
    const failed = cardLifecycleFixture({ getCard: async () => { throw new Error('Synthetic read failure'); } });
    assertError(await failed.reserve(), 503, 'reservation_failed');
    assert.deepEqual((await failed.delete()).body, { card: null });
});

test('calendar preserves past days and weekends as unavailable while tomorrow follows São Paulo time', () => {
    const rows = ['2026-11-09', '2026-11-10', '2026-11-11', '2026-11-14', '2026-11-15',
        '2026-11-16', '2026-02-30', 'not-a-date'].map((date) => ({ date, status: '' })).reverse();
    const original = structuredClone(rows);
    const result = calendarDays(rows, 1000, Date.parse('2026-11-10T01:30:00Z'));
    assert.deepEqual(result.map(({ date }) => date), ['2026-11-09', '2026-11-10', '2026-11-11', '2026-11-14', '2026-11-15', '2026-11-16']);
    assert.deepEqual(result.map(({ priceCents }) => priceCents), [null, 100000, 100000, null, null, 100000]);
    assert.deepEqual(result.map(({ available }) => available), [false, true, true, false, false, true]);
    assert.deepEqual(rows, original);
    assert.equal(calendarDays(rows, 1000, Date.parse('2026-11-14T02:30:00Z')).find(({ available }) => available).date, '2026-11-16');
});

test('calendar includes all seven days and eligible dates through the end of next month', () => {
    const dates = ['2026-10-31', '2026-11-01', '2026-11-09', '2026-11-10', '2026-11-14',
        '2026-11-15', '2026-11-20', '2026-12-08', '2026-12-09', '2026-12-10', '2026-12-31', '2027-01-01'];
    const result = calendarDays(dates.map((date) => ({ date, status: '' })), 1000, NOW);
    assert.deepEqual(result.map(({ date }) => date), dates.slice(1, -1));
    assert.deepEqual(result.filter(({ available }) => available).map(({ date }) => date),
        ['2026-11-10', '2026-12-08', '2026-12-09', '2026-12-10', '2026-12-31']);
    assert.ok(result.filter(({ available }) => !available).every(({ priceCents, pricePosition }) => priceCents === null && pricePosition === null));
});

test('national and Curitiba holidays remain visible and unavailable without consuming a price step', () => {
    for (const [nowMs, dates, available] of [
        [Date.parse('2026-09-01T15:00:00Z'), ['2026-09-07', '2026-09-08', '2026-09-09'], ['2026-09-09']],
        [Date.parse('2026-03-30T15:00:00Z'), ['2026-04-02', '2026-04-03', '2026-04-06'], ['2026-04-02', '2026-04-06']],
        [Date.parse('2026-06-01T15:00:00Z'), ['2026-06-03', '2026-06-04', '2026-06-05'], ['2026-06-03', '2026-06-05']],
    ]) {
        const result = calendarDays(dates.map((date) => ({ date, status: '' })), 1000, nowMs);
        assert.deepEqual(result.map(({ date }) => date), dates);
        assert.deepEqual(result.filter((day) => day.available).map(({ date }) => date), available);
        assert.ok(result.filter((day) => day.available).every(({ priceCents, pricePosition }) => priceCents === 100000 && pricePosition === 0));
    }
});

test('the first three available dates keep the reference rate, then prices decrease to 80 percent on the tenth', () => {
    const result = calendarDays(WEEKDAYS.map((date) => ({ date, status: '' })), 1000, NOW);
    assert.deepEqual(result.map(({ priceCents }) => priceCents), [100000, 100000, 100000, 97143, 94286,
        91429, 88571, 85714, 82857, 80000, 80000, 80000, 80000]);
    assert.deepEqual(result.map(({ pricePosition }) => pricePosition), [0, 0, 0, 1 / 7, 2 / 7,
        3 / 7, 4 / 7, 5 / 7, 6 / 7, 1, 1, 1, 1]);
    assert.equal(result[9].priceCents, result[0].priceCents * 0.8);
});

test('a company reference of 1250 decreases to 1000 at the tenth available date and stays there', () => {
    const result = calendarDays(WEEKDAYS.map((date) => ({ date, status: '' })), 1250, NOW);
    assert.deepEqual(result.map(({ priceCents }) => priceCents), [125000, 125000, 125000,
        121429, 117857, 114286, 110714, 107143, 103571, 100000, 100000, 100000, 100000]);
});

test('the server rejects a quote from the previous discount curve before reserving and accepts the current quote', async () => {
    const app = fixture({ user: { ...USER, referenceRate: 1250 } });
    const oldQuote = { date: WEEKDAYS[3], priceCents: 120536, requestId: ID, acknowledged: true };
    assertError(await app.reserve(oldQuote), 409, 'price_changed');
    assert.equal(app.calls.filter(({ type }) => ['reserve', 'event', 'mail', 'notification'].includes(type)).length, 0);
    assert.equal(app.stored.size, 0);
    assert.deepEqual((await app.reserve({ ...oldQuote, priceCents: 121429 })).body, { date: WEEKDAYS[3], reserved: true });
    assert.equal(app.calls.find(({ type }) => type === 'reserve').payload.hourlyRate, 1214.29);
});

test('all nonblank statuses and journal-pending dates are unavailable and do not consume discount steps', () => {
    const statuses = ['', 'RESERVADO', 'INDISPONÍVEL', 'BLOQUEADO', '  ', '', null, '0', ''];
    const rows = WEEKDAYS.slice(0, statuses.length).map((date, index) => ({ date, status: statuses[index] }));
    const result = calendarDays(rows, 1000, NOW, new Set([WEEKDAYS[5]]));
    assert.deepEqual(result.map(({ available }) => available), [true, false, false, false, true, false, true, false, true]);
    assert.deepEqual(result.map(({ priceCents }) => priceCents), [100000, null, null, null, 100000, null, 100000, null, 97143]);
    assert.deepEqual(result.filter(({ available }) => available).map(({ pricePosition }) => pricePosition), [0, 0, 0, 1 / 7]);
    assert.ok(result.filter(({ available }) => !available).every(({ pricePosition }) => pricePosition === null));
});

test('invalid reference prices fail closed instead of showing bookable dates', async () => {
    for (const rate of [undefined, null, '1000', 0, -10, NaN, Infinity, Number.MAX_SAFE_INTEGER]) {
        assert.throws(() => calendarDays([{ date: WEEKDAYS[0], status: '' }], rate, NOW), /Reference rate unavailable/u);
        assertError(await fixture({ user: { ...USER, referenceRate: rate } }).calendar(), 503, 'calendar_unavailable');
    }
    const result = calendarDays([{ date: WEEKDAYS[0], status: '' }], 1000.01, NOW);
    assert.equal(result[0].priceCents, 100001);
});

test('calendar read failures return a stable unavailable response without writes', async () => {
    for (const failure of ['dates', 'pending']) {
        const app = fixture({ faults: { [failure]: new Error('private service response') } });
        assertError(await app.calendar(), 503, 'calendar_unavailable');
        assert.ok(app.calls.every(({ type }) => ['cards:cached', 'dates', 'pending'].includes(type)));
    }
});

test('calendar without a company card returns administrator context and does not read or synchronize availability', async () => {
    const administrator = { ...USER, name: 'Ana', role: 'Administrador' };
    const directory = [administrator, { ...administrator }, { ...administrator, name: 'Beto' },
        { ...administrator, name: 'Cora' },
        { ...administrator, name: 'Colleague', role: 'Usuário' },
        { ...administrator, name: 'Other company', companyId: '99999999000199' }];
    const app = fixture({ card: null, directory, syncAvailability: true });
    const res = await app.calendar();
    assert.equal(res.statusCode, 200);
    assert.deepEqual(res.body, { days: [], canBook: false, paymentMethodRegistered: false, administratorName: 'Ana, Beto, Cora',
        userEmail: USER.email, today: '2026-11-09' });
    assert.deepEqual(app.calls.map(({ type }) => type), ['cards:cached']);
    assert.deepEqual(app.calls[0].payload.options, { cached: true });
    assert.deepEqual(app.calls[0].payload.user, USER);
});

test('calendar context shares masked payment state while the route returns only calendar fields', async () => {
    const app = fixture();
    const context = await app.handlers.getCalendarContext(USER, []);
    assert.deepEqual(context.card, { card: MASKED_CARD });
    assert.equal(context.calendar.paymentMethodRegistered, true);
    assert.equal(context.calendar.administratorName, 'não informado');
    assert.equal(context.calendar.userEmail, USER.email);
    assert.equal(context.calendar.today, '2026-11-09');
    assert.equal(context.calendar.days.length, WEEKDAYS.length);
    const res = await fixture().calendar();
    assert.deepEqual(res.body, context.calendar);
    assert.equal(Object.hasOwn(res.body, 'card'), false);
});

test('availability synchronization receives São Paulo today and replaces the fixture read-only fallback', async () => {
    const app = fixture({ syncAvailability: true, nowMs: Date.parse('2026-11-10T01:30:00Z') });
    assert.equal((await app.calendar()).statusCode, 200);
    assert.deepEqual(app.calls.map(({ type }) => type), ['cards:cached', 'pending', 'sync']);
    assert.deepEqual(app.calls.find(({ type }) => type === 'sync').payload, { today: '2026-11-09' });
    const failure = fixture({ syncAvailability: true, faults: { sync: new Error('private synchronization response') } });
    assertError(await failure.calendar(), 503, 'calendar_unavailable');
    assert.equal(failure.calls.some(({ type }) => type.startsWith('write:') || ['reserve', 'event', 'mail'].includes(type)), false);
});

test('reservation requires a fresh company card even after a cached calendar showed it registered', async () => {
    let card = MASKED_CARD;
    const app = fixture({ card: () => card, syncAvailability: true });
    assert.equal((await app.calendar()).body.paymentMethodRegistered, true);
    const previousCalls = app.calls.length;
    card = null;
    assertError(await app.reserve(), 403, 'payment_method_required');
    assert.deepEqual(app.calls.slice(previousCalls).map(({ type }) => type), ['read', 'cards:fresh']);
    assert.equal(app.calls.at(-1).payload.options, undefined);
    assert.equal(app.stored.size, 0);
    assert.equal(app.calls.some(({ type }) => ['reserve', 'event', 'mail'].includes(type) || type.startsWith('write:')), false);
});

test('payment lookup failures fail closed before reservation or availability effects', async () => {
    const app = fixture({ syncAvailability: true, faults: { 'cards:fresh': new Error('private provider response') } });
    assertError(await app.reserve(), 503, 'reservation_failed');
    assert.deepEqual(app.calls.map(({ type }) => type), ['read', 'cards:fresh']);
    assert.equal(app.stored.size, 0);
    const calendar = fixture({ faults: { 'cards:cached': new Error('private provider response') } });
    assertError(await calendar.calendar(), 503, 'calendar_unavailable');
    assert.deepEqual(calendar.calls.map(({ type }) => type), ['cards:cached']);
});

test('an uncontended reservation reuses its fresh authorized directory and exact availability snapshot', async () => {
    let directoryReads = 0;
    let receivedRows;
    const app = fixture({ syncAvailability: true,
        users: { async readUsers() { directoryReads++; throw new Error('Redundant directory read'); } },
        reserveDate: async (payload, snapshot) => { receivedRows = snapshot; },
    });
    assert.deepEqual((await app.reserve()).body, { date: WEEKDAYS[0], reserved: true });
    assert.equal(directoryReads, 0);
    assert.equal(app.calendarRows.length, 1);
    assert.equal(receivedRows, app.calendarRows[0], 'Only the exact just-validated snapshot may use the adapter fast path');
});

test('a queued reservation rechecks employee permission after waiting instead of trusting its old request snapshot', async () => {
    const writing = deferredOperation();
    let directoryReads = 0;
    const app = fixture({
        users: { async readUsers() { directoryReads++; return [{ ...USER, bookingAllowed: false }]; } },
        writeRecord: async record => { if (record.id === ID && record.stage === 'workbook') await writing.wait(); },
    });
    const first = app.reserve();
    await writing.started;
    const second = app.reserve({ date: WEEKDAYS[1], priceCents: 100000, requestId: OTHER_ID, acknowledged: true });
    writing.release();
    assert.equal((await first).body.reserved, true);
    assertError(await second, 403, 'booking_permission_required');
    assert.equal(directoryReads, 1);
    assert.equal(app.calls.filter(({ type }) => type === 'cards:fresh').length, 1);
    assert.equal(app.calls.filter(({ type }) => type === 'reserve').length, 1);
    assert.equal(app.stored.has(OTHER_ID), false);
});

test('reservation validates the request and identity before reading or writing state', async () => {
    const valid = { date: WEEKDAYS[0], priceCents: 100000, requestId: ID, acknowledged: true };
    for (const body of [null, {}, { ...valid, extra: 'ignored?' }, { ...valid, date: '2026-02-31' },
        { ...valid, date: '../2026-11-10' }, { ...valid, priceCents: 100000.5 }, { ...valid, priceCents: 0 },
        { ...valid, requestId: 'not-a-uuid' }, { ...valid, requestId: [ID] },
        { ...valid, acknowledged: undefined }, { ...valid, acknowledged: false },
        { ...valid, acknowledged: 'true' }, { ...valid, acknowledged: 1 }, { ...valid, acknowledged: null }]) {
        const app = fixture();
        assertError(await app.reserve(body), 400, 'invalid_request');
        assert.deepEqual(app.calls, []);
    }
    for (const user of [{ ...USER, companyId: 'bad' }, { ...USER, companyId: 12345678000195 },
        { ...USER, email: 'invalid' }, { ...USER, email: 'user@example.invalid\r\nInjected: x' }]) {
        const app = fixture();
        assertError(await app.reserve(valid, user), 400, 'invalid_request');
        assert.deepEqual(app.calls, []);
    }
});

test('reservation rechecks availability and the current price before its first write', async () => {
    const unavailable = fixture({ rows: [{ date: WEEKDAYS[0], status: 'RESERVADO' }] });
    assertError(await unavailable.reserve(), 409, 'date_unavailable');
    const pending = fixture({ records: [savedRecord({ id: OTHER_ID, stage: 'event' })] });
    assertError(await pending.reserve(), 409, 'date_unavailable');
    const stale = fixture({ user: { ...USER, referenceRate: 1200 } });
    assertError(await stale.reserve(), 409, 'price_changed');
    for (const app of [unavailable, pending, stale]) {
        assert.ok(app.calls.every(({ type }) => ['read', 'cards:fresh', 'dates', 'pending'].includes(type)));
    }
});

test('today, weekends, holidays, and dates after next month cannot be reserved', async () => {
    for (const [date, nowMs] of [['2026-11-09', NOW], ['2026-11-14', NOW], ['2026-11-20', NOW],
        ['2027-01-04', NOW], ['2026-09-08', Date.parse('2026-09-01T15:00:00Z')]]) {
        const app = fixture({ rows: [{ date, status: '' }], nowMs });
        assertError(await app.reserve({ date, priceCents: 100000, requestId: ID, acknowledged: true }), 409, 'date_unavailable');
        assert.ok(app.calls.every(({ type }) => ['read', 'cards:fresh', 'pending', 'dates'].includes(type)));
        assert.equal(app.stored.size, 0);
    }
});

test('eligible late-next-month dates can be reserved beyond thirty days, including a year boundary', async () => {
    for (const [today, date] of [['2026-11-10', '2026-12-31'], ['2026-12-15', '2027-01-29']]) {
        const app = fixture({ rows: [{ date, status: '' }], nowMs: Date.parse(`${today}T15:00:00Z`) });
        const result = await app.reserve({ date, priceCents: 100000, requestId: ID, acknowledged: true });
        assert.equal(result.statusCode, 200);
        assert.deepEqual(result.body, { date, reserved: true });
        assert.equal(app.calls.filter(({ type }) => type === 'reserve').length, 1);
    }
});

test('successful reservation creates one 09:00–10:00 Brasília Teams invitation and a separate user-only confirmation', async () => {
    const app = fixture();
    const res = await app.reserve();
    assert.equal(res.statusCode, 200);
    assert.deepEqual(res.body, { date: WEEKDAYS[0], reserved: true });
    assert.equal(app.calls[0].type, 'read');
    assert.equal(app.calls[1].type, 'cards:fresh');
    assert.deepEqual(new Set(app.calls.slice(2, 4).map(({ type }) => type)), new Set(['dates', 'pending']));
    assert.deepEqual(app.calls.slice(4).map(({ type }) => type), ['write:workbook',
        'reserve', 'write:event', 'event', 'write:confirmation', 'mail', 'write:notification', 'notification', 'write:complete']);
    assert.deepEqual(app.calls.find(({ type }) => type === 'reserve').payload, {
        date: WEEKDAYS[0], companyId: COMPANY_ID, name: USER.name, hourlyRate: 1000,
    });
    const event = app.calls.find(({ type }) => type === 'event').payload;
    assert.equal(event.subject, `Atendimento Consultivo: ${USER.name}`);
    assert.equal(event.transactionId, ID);
    assert.equal(event.isAllDay, false);
    assert.equal(event.showAs, 'busy');
    assert.deepEqual(event.start, { dateTime: '2026-11-10T09:00:00', timeZone: 'E. South America Standard Time' });
    assert.deepEqual(event.end, { dateTime: '2026-11-10T10:00:00', timeZone: 'E. South America Standard Time' });
    assert.equal(event.isOnlineMeeting, true);
    assert.equal(event.onlineMeetingProvider, 'teamsForBusiness');
    assert.deepEqual(event.attendees, [{ emailAddress: { address: USER.email.toLowerCase(), name: USER.name }, type: 'required' }]);
    const mail = app.calls.find(({ type }) => type === 'mail').payload;
    assert.deepEqual(mail.toRecipients, [{ emailAddress: { address: USER.email.toLowerCase() } }]);
    assert.equal(mail.ccRecipients, undefined);
    assert.equal(mail.bccRecipients, undefined);
    assert.equal(mail.subject, 'Atendimento Consultivo - Machado | Método Gerencial para Empresas');
    assert.match(mail.body.content, /^<p>Prezado\(a\) João,<\/p>/u);
    assert.ok(mail.body.content.includes('<p>Seu atendimento consultivo junto a nosso diretor está agendado para 10/11/2026 às 09:00 (horário de Brasília).</p>'));
    assert.ok(mail.body.content.includes('<p>Por favor confirme sua presença no invite do calendário.</p>'));
    assert.ok(mail.body.content.includes('>Clique aqui</a> para acessar o encontro via Microsoft Teams.</p>'));
    assert.ok(mail.body.content.includes('<p>Atenciosamente,</p>'));
    assert.ok(mail.body.content.includes('<img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg" width="600" />'));
    assert.equal(mail.attachments, undefined, 'Graph event invitations must not be duplicated by an ICS attachment');
    assert.deepEqual(app.calls.find(({ type }) => type === 'notification').payload, {
        subject: 'Atendimento Consultivo Agendado - João da Silva - 10/11/2026',
        body: { contentType: 'Text', content: '' },
        toRecipients: [{ emailAddress: { address: CONTACT } }],
    });
    assert.equal(app.calls.filter(({ type }) => type === 'notification').length, 1);
    assert.deepEqual(app.stored.get(ID), savedRecord({ stage: 'complete', acknowledgedAt: NOW, eventId: 'graph-event-id',
        iCalUid: 'provider-calendar-uid@example.invalid', meetingUrl: MEETING_URL }));
});

test('timed invites preserve the selected Brasília date across month, year, and leap-day boundaries', async (t) => {
    for (const date of ['2026-11-30', '2026-12-31', '2027-01-01', '2028-02-29']) {
        await t.test(date, async () => {
            const booking = savedRecord({ date, iCalUid: 'provider-calendar-uid@example.invalid', meetingUrl: MEETING_URL });
            const event = eventPayload(booking);
            assert.equal(event.isAllDay, false);
            assert.deepEqual(event.start, { dateTime: `${date}T09:00:00`, timeZone: 'E. South America Standard Time' });
            assert.deepEqual(event.end, { dateTime: `${date}T10:00:00`, timeZone: 'E. South America Standard Time' });
            const mail = confirmation(booking);
            assert.ok(mail.body.content.includes(`${date.split('-').reverse().join('/')} às 09:00 (horário de Brasília)`));
            assert.equal(mail.attachments, undefined);
        });
    }
});

test('the exact Teams join URL is stored once and linked safely without a second invitation', async () => {
    const app = fixture({ meetingUrl: MEETING_URL });
    assert.deepEqual((await app.reserve()).body, { date: WEEKDAYS[0], reserved: true });
    assert.equal(app.stored.get(ID).meetingUrl, MEETING_URL);
    assert.equal(app.stored.get(ID).acknowledgedAt, NOW);
    const mail = app.calls.find(({ type }) => type === 'mail').payload;
    assert.ok(mail.body.content.includes(`href="${MEETING_URL.replace(/&/gu, '&amp;')}"`));
    assert.ok(mail.body.content.includes(`href="${MEETING_URL.replace(/&/gu, '&amp;')}">Clique aqui</a> para acessar o encontro via Microsoft Teams.</p>`));
    assert.ok(!mail.body.content.includes(`>${MEETING_URL.replace(/&/gu, '&amp;')}</a>`));
    assert.equal(mail.attachments, undefined);
    assert.equal(app.calls.filter(({ type }) => type === 'mail').length, 1);
    assert.deepEqual((await app.reserve()).body, { date: WEEKDAYS[0], reserved: true });
    assert.equal(app.calls.filter(({ type }) => type === 'event').length, 1);
    assert.equal(app.calls.filter(({ type }) => type === 'mail').length, 1);
    assert.equal(app.calls.filter(({ type }) => type === 'notification').length, 1);
});

test('a created event without its Teams URL leaves notification pending and never sends incomplete confirmation or replays the invite', async () => {
    const app = fixture({ meetingUrl: null });
    assert.deepEqual((await app.reserve()).body, { date: WEEKDAYS[0], reserved: true, notificationPending: true });
    assert.equal(app.stored.get(ID).stage, 'confirmation');
    assert.equal(app.calls.filter(({ type }) => type === 'mail').length, 0);
    assert.deepEqual((await app.reserve()).body, { date: WEEKDAYS[0], reserved: true, notificationPending: true });
    assert.equal(app.calls.filter(({ type }) => type === 'event').length, 1);
    assert.equal(app.calls.filter(({ type }) => type === 'mail').length, 0);
});

test('notification HTML escapes the first name, full name and company without a calendar attachment', async () => {
    const user = { ...USER, name: 'João<A&B> da Silva', company: '<Empresa & Gestão>' };
    const app = fixture({ user });
    assert.equal((await app.reserve()).statusCode, 200);
    const mail = app.calls.find(({ type }) => type === 'mail').payload;
    assert.match(mail.body.content, /Prezado\(a\) João&lt;A&amp;B&gt;,/u);
    const event = app.calls.find(({ type }) => type === 'event').payload;
    assert.match(event.body.content, /João&lt;A&amp;B&gt; da Silva/u);
    assert.match(event.body.content, /&lt;Empresa &amp; Gestão&gt;/u);
    assert.equal(mail.attachments, undefined);
});

test('contact booking sends exactly one empty-body alert without a user confirmation or duplicate self-invite', async () => {
    const app = fixture({ user: { ...USER, email: CONTACT.toUpperCase() }, faults: { mail: new Error('Must not send') } });
    assert.deepEqual((await app.reserve()).body, { date: WEEKDAYS[0], reserved: true });
    assert.deepEqual(app.calls.find(({ type }) => type === 'event').payload.attendees, []);
    assert.equal(app.calls.filter(({ type }) => type === 'mail').length, 0);
    assert.deepEqual(app.calls.filter(({ type }) => type === 'notification').map(({ payload }) => payload), [{
        subject: 'Atendimento Consultivo Agendado - João da Silva - 10/11/2026',
        body: { contentType: 'Text', content: '' },
        toRecipients: [{ emailAddress: { address: CONTACT } }],
    }]);
    assert.equal(app.stored.get(ID).stage, 'complete');
    assert.equal(confirmation(savedRecord({ email: ` ${CONTACT.toUpperCase()} ` })), null);
    await app.reserve();
    assert.equal(app.calls.filter(({ type }) => type === 'event').length, 1);
    assert.equal(app.calls.filter(({ type }) => type === 'mail').length, 0);
    assert.equal(app.calls.filter(({ type }) => type === 'notification').length, 1);
});

test('a completed legacy request replay succeeds without payment lookup or repeated external writes', async () => {
    const app = fixture({ card: null, records: [savedRecord({ stage: 'complete', eventId: 'saved-event', iCalUid: 'saved-uid' })] });
    const res = await app.reserve();
    assert.deepEqual(res.body, { date: WEEKDAYS[0], reserved: true });
    assert.deepEqual(app.calls.map(({ type }) => type), ['read']);
});

test('unfinished journal replays never repeat an ambiguous external write', async () => {
    for (const stage of ['workbook', 'event', 'confirmation', 'notification']) {
        const app = fixture({ card: null, records: [savedRecord({ stage })] });
        const res = await app.reserve();
        if (stage === 'workbook') assertError(res, 409, 'reservation_review_required');
        else assert.deepEqual(res.body, { date: WEEKDAYS[0], reserved: true, notificationPending: true });
        assert.deepEqual(app.calls.map(({ type }) => type), ['read']);
    }
});

test('a request ID cannot be replayed for another date, quote, company, or email', async () => {
    for (const change of [{ date: WEEKDAYS[1] }, { priceCents: 1 }, { companyId: '99999999000199' }, { email: 'another@example.invalid' }]) {
        const app = fixture({ records: [savedRecord({ stage: 'complete', eventId: 'saved-event', iCalUid: 'saved-uid', ...change })] });
        assertError(await app.reserve(), 409, 'reservation_review_required');
        assert.deepEqual(app.calls.map(({ type }) => type), ['read']);
    }
});

test('read failures before the journal do not start a reservation', async () => {
    for (const stage of ['read', 'dates', 'pending']) {
        const app = fixture({ faults: { [stage]: new Error('private provider error') } });
        assertError(await app.reserve(), 503, 'reservation_failed');
        assert.ok(app.calls.every(({ type }) => ['read', 'cards:fresh', 'dates', 'pending'].includes(type)));
        assert.equal(app.stored.size, 0);
    }
});

test('failure to persist intent prevents the workbook mutation', async () => {
    const app = fixture({ faults: { 'write:workbook': new Error('disk unavailable') } });
    assertError(await app.reserve(), 409, 'reservation_review_required');
    assert.equal(app.calls.some(({ type }) => type === 'reserve'), false);
    assert.equal(app.stored.size, 0);
});

test('definite workbook rejection clears intent, while ambiguous failure stays blocked for review', async () => {
    const definite = Object.assign(new Error('date taken'), { definiteRejection: true, code: 'support_appointments.date_unavailable' });
    const app = fixture({ faults: { reserve: definite } });
    assertError(await app.reserve(), 409, 'date_unavailable');
    assert.equal(app.stored.size, 0);
    assert.equal(app.calls.filter(({ type }) => type === 'remove').length, 1);
    assert.equal(app.calls.some(({ type }) => type === 'event'), false);
    const uncertain = fixture({ faults: { reserve: new Error('connection lost after possible write') } });
    assertError(await uncertain.reserve(), 409, 'reservation_review_required');
    assert.equal(uncertain.stored.get(ID).stage, 'workbook');
    assertError(await uncertain.reserve(), 409, 'reservation_review_required');
    assert.equal(uncertain.calls.filter(({ type }) => type === 'reserve').length, 1);
});

test('a definite failure before the workbook mutation permits a safe retry', async () => {
    const faults = { reserve: Object.assign(new Error('schema read unavailable before mutation'), { definiteRejection: true }) };
    const app = fixture({ faults });
    assertError(await app.reserve(), 503, 'reservation_failed');
    assert.equal(app.stored.size, 0);
    assert.equal(app.calls.filter(({ type }) => type === 'remove').length, 1);
    delete faults.reserve;
    assert.deepEqual((await app.reserve()).body, { date: WEEKDAYS[0], reserved: true });
    assert.equal(app.calls.filter(({ type }) => type === 'event').length, 1);
});

test('post-reservation failures retain the booking and never resend on an idempotent replay', async (t) => {
    for (const [failure, remainingStage, eventCount, mailCount, notificationCount] of [
        ['write:event', 'workbook', 0, 0, 0], ['event', 'event', 1, 0, 0],
        ['write:confirmation', 'event', 1, 0, 0], ['mail', 'confirmation', 1, 1, 0],
        ['write:notification', 'confirmation', 1, 1, 0], ['notification', 'notification', 1, 1, 1],
        ['write:complete', 'notification', 1, 1, 1],
    ]) {
        await t.test(failure, async () => {
            const app = fixture({ faults: { [failure]: new Error('synthetic interruption') } });
            assert.deepEqual((await app.reserve()).body, { date: WEEKDAYS[0], reserved: true, notificationPending: true });
            assert.equal(app.stored.get(ID).stage, remainingStage);
            const replay = await app.reserve();
            if (remainingStage === 'workbook') assertError(replay, 409, 'reservation_review_required');
            else assert.deepEqual(replay.body, { date: WEEKDAYS[0], reserved: true, notificationPending: true });
            assert.equal(app.calls.filter(({ type }) => type === 'reserve').length, 1);
            assert.equal(app.calls.filter(({ type }) => type === 'event').length, eventCount);
            assert.equal(app.calls.filter(({ type }) => type === 'mail').length, mailCount);
            assert.equal(app.calls.filter(({ type }) => type === 'notification').length, notificationCount);
        });
    }
});

test('concurrent reservations for the same date or request ID make only one workbook write', async () => {
    let release;
    let entered;
    const gate = new Promise((resolve) => { release = resolve; });
    const started = new Promise((resolve) => { entered = resolve; });
    const app = fixture({ reserveDate: async () => { entered(); await gate; } });
    const first = app.reserve();
    await started;
    assertError(await app.reserve({ date: WEEKDAYS[0], priceCents: 100000, requestId: OTHER_ID, acknowledged: true }), 409, 'date_unavailable');
    assertError(await app.reserve({ date: WEEKDAYS[1], priceCents: 100000, requestId: ID, acknowledged: true }), 409, 'date_unavailable');
    release();
    assert.deepEqual((await first).body, { date: WEEKDAYS[0], reserved: true });
    assert.deepEqual((await app.reserve()).body, { date: WEEKDAYS[0], reserved: true });
    assert.equal(app.calls.filter(({ type }) => type === 'reserve').length, 1);
    assert.equal(app.calls.filter(({ type }) => type === 'event').length, 1);
});

test('simultaneous different-date claims refresh the second price after the first intent is durable', async () => {
    let release;
    let entered;
    const gate = new Promise((resolve) => { release = resolve; });
    const started = new Promise((resolve) => { entered = resolve; });
    const app = fixture({ writeRecord: async (value) => {
        if (value.id === ID && value.stage === 'workbook') { entered(); await gate; }
    } });
    const initial = await app.calendar();
    assert.deepEqual(initial.body.days.slice(0, 4).map(({ priceCents }) => priceCents), [100000, 100000, 100000, 97143]);
    const first = app.reserve();
    await started;
    const second = app.reserve({ date: WEEKDAYS[3], priceCents: 97143, requestId: OTHER_ID, acknowledged: true });
    await new Promise(setImmediate);
    assert.equal(app.calls.filter(({ type }) => type === 'read').length, 1);
    assert.equal(app.calls.filter(({ type }) => type === 'reserve').length, 0);
    release();
    assertError(await second, 409, 'price_changed');
    assert.deepEqual((await first).body, { date: WEEKDAYS[0], reserved: true });
    assert.equal(app.calls.filter(({ type }) => type === 'reserve').length, 1);
    assert.equal(app.stored.has(OTHER_ID), false);
    const refreshed = await app.reserve({ date: WEEKDAYS[3], priceCents: 100000, requestId: OTHER_ID, acknowledged: true });
    assert.deepEqual(refreshed.body, { date: WEEKDAYS[3], reserved: true });
    assert.equal(app.calls.filter(({ type }) => type === 'reserve').length, 2);
});


test('missing, false or malformed employee permission hides availability and blocks new reservations before payment checks', async () => {
    for (const bookingAllowed of [undefined, false, 'Sim', 1]) {
        const app = fixture({ user: { ...USER, bookingAllowed } });
        const calendar = await app.calendar();
        assert.equal(calendar.body.paymentMethodRegistered, true);
        assert.equal(calendar.body.canBook, false);
        assert.deepEqual(calendar.body.days, []);
        const before = app.calls.length;
        assertError(await app.reserve(), 403, 'booking_permission_required');
        assert.deepEqual(app.calls.slice(before).map(({ type }) => type), ['read']);
    }
});
