'use strict';

require('./deny-production-network');
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const { once } = require('node:events');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const crypto = require('node:crypto');
const { createSupportApp } = require('../support-app');
const { companyCode } = require('../domains/support-management');
const { createSupportCardStore } = require('../integrations/support-card-store');
const { openSupportSessionStore, createMemorySupportSessionStore } = require('../integrations/support-session-store');
const { createTestApp, startLoopback } = require('./app-test-support');

const admin = { company: 'Empresa A', name: 'Administradora teste', email: 'admin@example.test', password: 'test-password', role: 'Administrador', bookingAllowed: true };
const colleague = { ...admin, name: 'Colega teste', email: 'colleague@example.test', role: 'Usuário' };
const otherAdmin = { ...admin, company: 'Empresa B', email: 'other@example.test' };
const verification = 'pagarme_verify_card_v1';
const card = { id: 'card_synthetic', brand: 'Visa', lastFour: '0010', holderName: 'SYNTHETIC HOLDER', expMonth: 12, expYear: 2030, status: 'active' };
const billingAddress = { line_1: '1, Rua Teste, Centro', zip_code: '01001000', city: 'São Paulo', state: 'SP', country: 'BR' };
const cardRequest = { token: 'token_synthetic', billingAddress };
const cardRevision = (customerId = 'cus_synthetic', cardId = card.id) => crypto.createHash('sha256').update(`${customerId}:${cardId}`).digest('hex');
const deleteRequest = { expectedCardRevision: cardRevision() };

function assertCardResponse(actual, expected) {
    const { cardRevision: revision, ...details } = actual;
    assert.match(revision, /^[a-f0-9]{64}$/u);
    assert.deepEqual(details, expected);
}

function assertStoredRevision(actual, expected) {
    assert.match(actual.cardRevision, /^[a-f0-9]{64}$/u);
    assert.deepEqual(actual, { ...expected, cardRevision: actual.cardRevision });
}

async function fixture(t, overrides = {}) {
    let clock = 1_800_000_000_000;
    const state = { users: structuredClone([admin, colleague, otherAdmin]) };
    const records = new Map();
    const calls = [];
    const users = {
        async readUsers() { calls.push(['readUsers']); return state.users; },
        async writeCompanyCard(company, details, references) { calls.push(['writeCompanyCard', company, structuredClone(details), structuredClone(references)]); },
        async clearCompanyCard(company) { calls.push(['clearCompanyCard', company]); },
        async writeBookingPermissions(companyId, employees, directory) {
            calls.push(['writeBookingPermissions', companyId, structuredClone(employees)]);
            assert.equal(directory, state.users);
            for (const { email, allowed } of employees) state.users.find((person) => person.companyId === companyId && person.email.toLowerCase() === email).bookingAllowed = allowed;
        },
    };
    const cardStore = {
        async read(code) { calls.push(['read', code]); return structuredClone(records.get(code) || {}); },
        async write(code, record) { calls.push(['write', code, structuredClone(record)]); records.set(code, structuredClone(record)); },
    };
    const pagarme = {
        async findCompanyCustomer(code) { calls.push(['findCompanyCustomer', code]); return null; },
        async createCompanyCustomer(details) { calls.push(['createCompanyCustomer', details]); return { id: 'cus_synthetic', code: details.code }; },
        async createCard(customerId, details) { calls.push(['createCard', customerId, details]); return { ...card, verificationRequested: true }; },
        async getCard(customerId, cardId) { calls.push(['getCard', customerId, cardId]); return { ...card, holder: 'PRIVATE', secret: 'PRIVATE' }; },
        async deleteCard(customerId, cardId) { calls.push(['deleteCard', customerId, cardId]); return { deleted: true }; },
    };
    let app;
    const server = http.createServer((req, res) => app(req, res));
    server.listen(0, '127.0.0.1');
    await once(server, 'listening');
    async function close() {
        if (!server.listening) return;
        server.closeAllConnections();
        await new Promise((resolve) => server.close(resolve));
    }
    t.after(close);
    const origin = `http://127.0.0.1:${server.address().port}`;
    app = createSupportApp({ origin, users, cardStore, pagarme, publicKey: 'pk_test_synthetic', now: () => clock, ...overrides });
    async function request(route, { method = 'GET', token, body, headers = {} } = {}) {
        return fetch(`${origin}/support-management/api/${route}`, {
            method,
            headers: { ...(body === undefined ? {} : { 'Content-Type': 'application/json' }), ...(token ? { Authorization: `Bearer ${token}` } : {}), ...headers },
            ...(body === undefined ? {} : { body: JSON.stringify(body) }),
        });
    }
    async function login(user = admin) {
        const response = await request('login', { method: 'POST', body: { email: user.email, password: user.password } });
        assert.equal(response.status, 200);
        return (await response.json()).token;
    }
    return { origin, state, records, calls, users, cardStore, pagarme, request, login, close, advance(ms) { clock += ms; } };
}

function expectError(response, status, reason) {
    assert.equal(response.status, status);
    return response.json().then((body) => assert.deepEqual(body, { error: `support_management.${reason}` }));
}

function seedCompanyPayment(f) {
    const companyId = '12345678000190';
    for (const user of f.state.users.filter((user) => user.company === admin.company)) {
        user.companyId = companyId;
        user.referenceRate = 1000;
    }
    f.records.set(companyCode(companyId), { customerId: 'cus_synthetic', cardId: card.id, verification,
        workbookSynced: true, workbookReferencesSynced: true });
    return companyId;
}

test('calendar routes authorize freshly and verify the card before reserving without charging', async (t) => {
    const bookings = new Map();
    const dates = [{ date: '2027-02-01', status: '' }];
    const notifications = [];
    const f = await fixture(t, {
        appointments: {
            readDates: async () => structuredClone(dates),
            reserveDate: async () => { dates[0].status = 'RESERVADO'; },
        },
        bookingStore: {
            read: async (id) => bookings.get(id),
            write: async (record) => bookings.set(record.id, structuredClone(record)),
            remove: async (id) => bookings.delete(id),
            pendingDates: async () => new Set([...bookings.values()].filter((record) => record.stage !== 'complete').map((record) => record.date)),
        },
        bookingNotifications: {
            createEvent: async (event) => { notifications.push(event); return { eventId: 'syntheticEvent', iCalUid: 'syntheticUid', meetingUrl: 'https://teams.microsoft.com/meet/synthetic' }; },
            sendConfirmation: async (mail) => { notifications.push(mail); },
        },
    });
    seedCompanyPayment(f);
    await expectError(await f.request('calendar'), 401, 'session_expired');
    await expectError(await f.request('reservations', { method: 'POST', body: {} }), 401, 'session_expired');
    const token = await f.login();
    assert.equal((await (await f.request('calendar', { token })).json()).days[0].priceCents, 100000);
    f.state.users[0].referenceRate = 1200;
    assert.equal((await (await f.request('calendar', { token })).json()).days[0].priceCents, 120000);
    const reservation = { date: dates[0].date, priceCents: 120000, requestId: '12345678-1234-4234-8234-123456789abc' };
    await expectError(await f.request('reservations', { method: 'POST', token, body: reservation }), 400, 'invalid_request');
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 1);
    const getCard = f.pagarme.getCard;
    f.pagarme.getCard = async (...args) => ({ ...await getCard(...args), status: 'inactive' });
    await expectError(await f.request('reservations', { method: 'POST', token, body: { ...reservation, acknowledged: true } }), 403, 'payment_method_required');
    assert.equal(bookings.size, 0);
    assert.equal(notifications.length, 0);
    assert.equal(dates[0].status, '');
    f.pagarme.getCard = getCard;
    const response = await f.request('reservations', { method: 'POST', token,
        body: { ...reservation, acknowledged: true } });
    assert.deepEqual(await response.json(), { date: dates[0].date, reserved: true });
    assert.equal(response.status, 200);
    assert.equal(notifications.length, 2);
    assert.equal(bookings.get(reservation.requestId).acknowledgedAt, 1_800_000_000_000);
    assert.equal((await (await f.request('calendar', { token })).json()).days[0].available, false);
    assert.ok(f.calls.every(([type]) => ['readUsers', 'read', 'getCard'].includes(type)));
    f.state.users.shift();
    await expectError(await f.request('calendar', { token }), 401, 'session_expired');
});

test('preloaded login supplies safe initial context once and session refresh keeps authorization and expiry current', async (t) => {
    let calendarReads = 0;
    const f = await fixture(t, { preload: true,
        appointments: { readDates: async () => { calendarReads++; return [{ date: '2027-02-01', status: '' }]; } },
        bookingStore: { pendingDates: async () => new Set() },
    });
    seedCompanyPayment(f);
    const response = await f.request('login', { method: 'POST', body: { email: admin.email, password: admin.password } });
    assert.equal(response.status, 200);
    const body = await response.json();
    assert.deepEqual(Object.keys(body).sort(), ['calendar', 'card', 'config', 'expiresAt', 'permissions', 'token', 'user']);
    assert.deepEqual(body.config, { environment: 'sandbox', publicKey: 'pk_test_synthetic' });
    assertCardResponse(body.card, { card: { brand: card.brand, lastFour: card.lastFour, expMonth: card.expMonth, expYear: card.expYear } });
    assert.equal(body.calendar.paymentMethodRegistered, true);
    assert.equal(body.calendar.administratorName, admin.name);
    assert.equal(body.calendar.userEmail, admin.email);
    assert.equal(body.calendar.days[0].priceCents, 100000);
    assert.equal(body.expiresAt, 1_800_000_000_000 + 60 * 60 * 1000);
    assert.equal(calendarReads, 1);
    assert.equal(f.calls.filter(([operation]) => operation === 'readUsers').length, 1);
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 1);
    const serialized = JSON.stringify(body);
    for (const privateValue of [admin.password, otherAdmin.email, 'cus_synthetic', 'card_synthetic', verification, 'PRIVATE', 'referenceRate']) {
        assert.equal(serialized.includes(privateValue), false, privateValue);
    }
    f.advance(30_000);
    f.state.users[0].role = 'Usuário';
    const current = await (await f.request('session', { token: body.token })).json();
    assert.equal(current.user.role, 'Usuário');
    assert.equal(current.user.email, admin.email);
    assert.equal(current.expiresAt, body.expiresAt);
    assert.equal(current.calendar.administratorName, 'não informado');
    assert.equal(f.calls.filter(([operation]) => operation === 'readUsers').length, 2);
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 1);
    f.state.users.shift();
    await expectError(await f.request('session', { token: body.token }), 401, 'session_expired');
    assert.equal(calendarReads, 2);
});

test('preloaded context gates missing cards without reading the calendar or exposing other companies', async (t) => {
    let calendarReads = 0;
    const f = await fixture(t, { preload: true, appointments: { readDates: async () => { calendarReads++; return []; } },
        bookingStore: { pendingDates: async () => new Set() } });
    const companyId = seedCompanyPayment(f);
    f.records.delete(companyCode(companyId));
    const response = await f.request('login', { method: 'POST', body: { email: colleague.email, password: colleague.password } });
    const body = await response.json();
    assert.equal(response.status, 200);
    assert.deepEqual(body.card, { card: null });
    assert.equal(body.calendar.paymentMethodRegistered, false);
    assert.equal(body.calendar.administratorName, admin.name);
    assert.deepEqual(body.calendar.days, []);
    assert.equal(calendarReads, 0);
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 0);
    assert.equal(JSON.stringify(body).includes(otherAdmin.email), false);
});

test('a missing company card gates every role and permission using CNPJ rather than company labels', async (t) => {
    for (const role of ['Administrador', 'Usuário']) {
        for (const bookingAllowed of [true, false, undefined]) {
            await t.test(`${role}, bookingAllowed ${bookingAllowed}`, async (t) => {
                let calendarReads = 0;
                const f = await fixture(t, { preload: true,
                    appointments: { readDates: async () => { calendarReads++; return [{ date: '2027-02-01', status: '' }]; } },
                    bookingStore: { pendingDates: async () => new Set() } });
                const companyId = seedCompanyPayment(f);
                const record = f.records.get(companyCode(companyId));
                f.records.delete(companyCode(companyId));
                const subject = f.state.users[1];
                Object.assign(subject, { role, bookingAllowed, company: 'Empresa A, unidade consultiva' });
                f.state.users[0].company = 'Empresa A, cadastro principal';
                Object.assign(f.state.users[2], { company: subject.company, companyId: '98765432000110', name: 'Administradora de outra empresa' });
                f.records.set(companyCode(f.state.users[2].companyId), record);
                const expectedAdministrators = role === 'Administrador' ? `${admin.name}, ${subject.name}` : admin.name;
                const response = await f.request('login', { method: 'POST', body: { email: subject.email, password: subject.password } });
                assert.equal(response.status, 200);
                const initial = await response.json();
                const session = await (await f.request('session', { token: initial.token })).json();
                const calendar = await (await f.request('calendar', { token: initial.token })).json();
                for (const context of [initial.calendar, session.calendar, calendar]) {
                    assert.equal(context.paymentMethodRegistered, false);
                    assert.equal(context.canBook, false);
                    assert.equal(context.administratorName, expectedAdministrators);
                    assert.deepEqual(context.days, []);
                }
                assert.deepEqual(initial.card, { card: null });
                assert.deepEqual(session.card, { card: null });
                assert.equal(calendarReads, 0);
                assert.ok(f.calls.every(([operation]) => ['readUsers', 'read'].includes(operation)));
                assert.ok(f.calls.filter(([operation]) => operation === 'read').every(([, code]) => code === companyCode(companyId)));
            });
        }
    }
});

test('calendar payment readiness is cached for sixty seconds while explicit card reads are fresh', async (t) => {
    const f = await fixture(t, { appointments: { readDates: async () => [{ date: '2027-02-01', status: '' }] },
        bookingStore: { pendingDates: async () => new Set() } });
    seedCompanyPayment(f);
    const token = await f.login();
    const readCalendar = async () => (await f.request('calendar', { token })).json();
    assert.equal((await readCalendar()).paymentMethodRegistered, true);
    const getCard = f.pagarme.getCard;
    f.pagarme.getCard = async (...args) => ({ ...await getCard(...args), status: 'inactive' });
    assert.equal((await readCalendar()).paymentMethodRegistered, true);
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 1);
    assertCardResponse(await (await f.request('card', { token })).json(), { card: null, verificationRequired: true });
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 2);
    assert.equal((await readCalendar()).paymentMethodRegistered, false);
    f.pagarme.getCard = getCard;
    f.advance(59_999);
    assert.equal((await readCalendar()).paymentMethodRegistered, false);
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 2);
    f.advance(1);
    assert.equal((await readCalendar()).paymentMethodRegistered, true);
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 3);
});

test('saving and synchronizing a card invalidate readiness even when the provider references stay the same', async (t) => {
    const f = await fixture(t, { appointments: { readDates: async () => [{ date: '2027-02-01', status: '' }] },
        bookingStore: { pendingDates: async () => new Set() } });
    seedCompanyPayment(f);
    const token = await f.login();
    const providerReads = () => f.calls.filter(([operation]) => operation === 'getCard').length;
    assert.equal((await f.request('calendar', { token })).status, 200);
    assert.equal(providerReads(), 1);
    assert.equal((await f.request('card', { method: 'PUT', token, body: cardRequest })).status, 200);
    assert.equal((await f.request('calendar', { token })).status, 200);
    assert.equal(providerReads(), 2);
    assert.equal((await f.request('card/sync', { method: 'POST', token })).status, 200);
    assert.equal(providerReads(), 3);
    assert.equal((await f.request('calendar', { token })).status, 200);
    assert.equal(providerReads(), 4);
    assert.equal(f.calls.filter(([operation]) => operation === 'createCard').length, 1);
});

test('card deletion authorizes the current administrator and rejects body overrides before card access', async (t) => {
    const f = await fixture(t);
    seedCompanyPayment(f);
    const token = await f.login();
    const colleagueToken = await f.login(colleague);
    f.calls.length = 0;
    await expectError(await f.request('card', { method: 'DELETE' }), 401, 'session_expired');
    assert.equal(f.calls.length, 0);
    await expectError(await f.request('card', { method: 'DELETE', token: colleagueToken }), 403, 'administrator_required');
    for (const body of [undefined, {}, [], ['foreign'], { expectedCardRevision: 'bad' }, { expectedCardRevision: 'A'.repeat(64) },
        { ...deleteRequest, company: otherAdmin.company }, { ...deleteRequest, companyId: '98765432000110' },
        { ...deleteRequest, customerId: 'cus_foreign' }, { ...deleteRequest, cardId: 'card_foreign' }, { ...deleteRequest, token: 'token_foreign' }]) {
        await expectError(await f.request('card', { method: 'DELETE', token, body }), 400, 'invalid_request');
    }
    f.state.users[0].role = 'Usuário';
    await expectError(await f.request('card', { method: 'DELETE', token }), 403, 'administrator_required');
    assert.ok(f.calls.every(([operation]) => operation === 'readUsers'));
});

test('deletion uses trusted company references, persists each stage, clears CNPJ references and permits customer reuse', async (t) => {
    const f = await fixture(t, { appointments: { readDates: async () => [{ date: '2027-02-01', status: '' }] },
        bookingStore: { pendingDates: async () => new Set() } });
    const companyId = seedCompanyPayment(f);
    const code = companyCode(companyId);
    const previous = structuredClone(f.records.get(code));
    const foreign = { customerId: 'cus_foreign', cardId: 'card_foreign', verification };
    f.records.set(companyCode(otherAdmin.company), foreign);
    const token = await f.login();
    assert.equal((await (await f.request('calendar', { token })).json()).paymentMethodRegistered, true);
    f.calls.length = 0;
    const response = await f.request('card', { method: 'DELETE', token, body: deleteRequest });
    assert.equal(response.status, 200);
    assert.match(response.headers.get('content-type'), /application\/json/u);
    assert.deepEqual(await response.json(), { card: null });
    assert.deepEqual(f.calls.filter(([operation]) => ['write', 'deleteCard', 'clearCompanyCard'].includes(operation)), [
        ['write', code, { ...previous, pending: 'delete' }],
        ['deleteCard', previous.customerId, previous.cardId],
        ['write', code, { customerId: previous.customerId, cardId: previous.cardId, cardRevision: cardRevision(), pending: 'delete_workbook' }],
        ['clearCompanyCard', companyId],
        ['write', code, { customerId: previous.customerId }],
    ]);
    assert.deepEqual(f.records.get(companyCode(otherAdmin.company)), foreign);
    assert.deepEqual(f.records.get(code), { customerId: previous.customerId });
    assert.deepEqual(await (await f.request('card', { token })).json(), { card: null });
    assert.equal((await (await f.request('calendar', { token })).json()).paymentMethodRegistered, false);
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 0);
    assert.equal((await f.request('card', { method: 'PUT', token, body: cardRequest })).status, 200);
    assert.deepEqual(f.calls.find(([operation]) => operation === 'createCard'), ['createCard', previous.customerId, cardRequest]);
    assert.equal(f.calls.filter(([operation]) => operation === 'findCompanyCustomer' || operation === 'createCompanyCustomer').length, 0);
    assert.equal(f.calls.find(([operation]) => operation === 'writeCompanyCard')[1], companyId);
    assert.equal((await (await f.request('calendar', { token })).json()).paymentMethodRegistered, true);
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 1);
});

test('deletion shares the company lock and pending deletion closes cached calendar readiness immediately', async (t) => {
    let calendarReads = 0;
    const f = await fixture(t, { appointments: { readDates: async () => { calendarReads++; return [{ date: '2027-02-01', status: '' }]; } },
        bookingStore: { pendingDates: async () => new Set() } });
    const companyId = seedCompanyPayment(f);
    const token = await f.login();
    assert.equal((await (await f.request('calendar', { token })).json()).paymentMethodRegistered, true);
    let release;
    let entered;
    const started = new Promise((resolve) => { entered = resolve; });
    const gate = new Promise((resolve) => { release = resolve; });
    t.after(() => release());
    f.pagarme.deleteCard = async (...references) => {
        f.calls.push(['deleteCard', ...references]);
        entered();
        await gate;
        return { deleted: true };
    };
    const deleting = f.request('card', { method: 'DELETE', token, body: deleteRequest });
    await started;
    const mutationsBefore = f.calls.filter(([operation]) => ['write', 'deleteCard', 'clearCompanyCard', 'getCard'].includes(operation));
    for (const request of [
        ['card', { method: 'DELETE', token, body: deleteRequest }],
        ['card', { method: 'PUT', token, body: cardRequest }],
        ['card/sync', { method: 'POST', token }],
    ]) await expectError(await f.request(...request), 409, 'card_busy');
    assertCardResponse(await (await f.request('card', { token })).json(), { card: null, cardDeletionPending: true });
    const calendar = await (await f.request('calendar', { token })).json();
    assert.equal(calendar.paymentMethodRegistered, false);
    assert.deepEqual(calendar.days, []);
    assert.equal(calendarReads, 1);
    assert.deepEqual(f.calls.filter(([operation]) => ['write', 'deleteCard', 'clearCompanyCard', 'getCard'].includes(operation)), mutationsBefore);
    assert.equal(f.records.get(companyCode(companyId)).pending, 'delete');
    release();
    assert.deepEqual(await (await deleting).json(), { card: null });
});

test('definite deletion rejection restores the previous card while uncertain failures retain deletion intent', async (t) => {
    for (const definite of [false, true]) {
        await t.test(definite ? 'definite rejection' : 'uncertain failure', async (t) => {
            const f = await fixture(t);
            const companyId = seedCompanyPayment(f);
            const code = companyCode(companyId);
            const previous = structuredClone(f.records.get(code));
            const token = await f.login();
            f.pagarme.deleteCard = async (...references) => {
                f.calls.push(['deleteCard', ...references]);
                throw Object.assign(new Error('PRIVATE provider details'), definite ? { definiteRejection: true } : {});
            };
            await expectError(await f.request('card', { method: 'DELETE', token, body: deleteRequest }), 503, 'card_delete_failed');
            assert.deepEqual(f.records.get(code), definite ? previous : { ...previous, pending: 'delete' });
            assert.equal(f.calls.filter(([operation]) => operation === 'clearCompanyCard').length, 0);
            assert.equal(f.calls.filter(([operation]) => operation === 'deleteCard').length, 1);
            const current = await f.request('card', { token });
            if (definite) assert.equal((await current.json()).card.lastFour, card.lastFour);
            else {
                assertCardResponse(await current.json(), { card: null, cardDeletionPending: true });
                await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 409, 'card_review_required');
                await expectError(await f.request('card/sync', { method: 'POST', token }), 409, 'card_review_required');
            }
        });
    }
});

test('missing or unconfirmed provider deletion results never clear workbook references or claim success', async (t) => {
    for (const result of [undefined, null, {}, { deleted: false }, { deleted: 'true' }]) {
        await t.test(JSON.stringify(result) || 'undefined result', async (t) => {
            const f = await fixture(t);
            const companyId = seedCompanyPayment(f);
            const token = await f.login();
            f.pagarme.deleteCard = async (...references) => { f.calls.push(['deleteCard', ...references]); return result; };
            await expectError(await f.request('card', { method: 'DELETE', token, body: deleteRequest }), 503, 'card_delete_failed');
            assert.equal(f.records.get(companyCode(companyId)).pending, 'delete');
            assert.equal(f.calls.filter(([operation]) => operation === 'clearCompanyCard').length, 0);
            assert.equal(f.calls.filter(([operation]) => operation === 'deleteCard').length, 1);
        });
    }
});

test('pending deletion reconciles provider deleted status or 404 with GET and never repeats DELETE', async (t) => {
    for (const absent of ['deleted', '404']) {
        await t.test(absent, async (t) => {
            const f = await fixture(t);
            const companyId = seedCompanyPayment(f);
            const code = companyCode(companyId);
            f.records.set(code, { ...f.records.get(code), pending: 'delete' });
            const token = await f.login();
            f.pagarme.getCard = async (...references) => {
                f.calls.push(['getCard', ...references]);
                if (absent === '404') throw Object.assign(new Error('Card absent'), { code: 'pagarme.card_not_found' });
                return { ...card, status: 'deleted' };
            };
            assert.deepEqual(await (await f.request('card', { method: 'DELETE', token, body: deleteRequest })).json(), { card: null });
            assert.deepEqual(f.records.get(code), { customerId: 'cus_synthetic' });
            assert.deepEqual(f.calls.find(([operation]) => operation === 'getCard'), ['getCard', 'cus_synthetic', card.id]);
            assert.equal(f.calls.filter(([operation]) => operation === 'deleteCard').length, 0);
            assert.deepEqual(f.calls.find(([operation]) => operation === 'clearCompanyCard'), ['clearCompanyCard', companyId]);
        });
    }
});

test('pending deletion does not treat provider read errors as absence and retries DELETE only after finding the card', async (t) => {
    const f = await fixture(t);
    const companyId = seedCompanyPayment(f);
    const code = companyCode(companyId);
    const pending = { ...f.records.get(code), pending: 'delete' };
    f.records.set(code, pending);
    const token = await f.login();
    const getCard = f.pagarme.getCard;
    f.pagarme.getCard = async (...references) => { f.calls.push(['failedGetCard', ...references]); throw new Error('PRIVATE provider read failure'); };
    await expectError(await f.request('card', { method: 'DELETE', token, body: deleteRequest }), 503, 'card_delete_failed');
    assert.deepEqual(f.records.get(code), pending);
    assert.equal(f.calls.filter(([operation]) => operation === 'deleteCard' || operation === 'clearCompanyCard').length, 0);
    f.pagarme.getCard = getCard;
    assert.deepEqual(await (await f.request('card', { method: 'DELETE', token, body: deleteRequest })).json(), { card: null });
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 1);
    assert.equal(f.calls.filter(([operation]) => operation === 'deleteCard').length, 1);
    assert.ok(f.calls.findIndex(([operation]) => operation === 'getCard') < f.calls.findIndex(([operation]) => operation === 'deleteCard'));
});

test('partial workbook clearing remains pending and retries only workbook clearing for the trusted company', async (t) => {
    const f = await fixture(t);
    const companyId = seedCompanyPayment(f);
    const code = companyCode(companyId);
    const token = await f.login();
    const colleagues = [false, false];
    let attempts = 0;
    f.users.clearCompanyCard = async (id) => {
        f.calls.push(['clearCompanyCard', id]);
        attempts++;
        colleagues[0] = true;
        if (attempts === 1) throw new Error('PRIVATE partial workbook failure');
        colleagues[1] = true;
    };
    await expectError(await f.request('card', { method: 'DELETE', token, body: deleteRequest }), 503, 'card_delete_failed');
    assert.deepEqual(colleagues, [true, false]);
    assert.deepEqual(f.records.get(code), { customerId: 'cus_synthetic', cardId: card.id, cardRevision: cardRevision(), pending: 'delete_workbook' });
    assertCardResponse(await (await f.request('card', { token })).json(), { card: null, cardDeletionPending: true });
    await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 409, 'card_review_required');
    await expectError(await f.request('card/sync', { method: 'POST', token }), 409, 'card_review_required');
    assert.deepEqual(await (await f.request('card', { method: 'DELETE', token, body: deleteRequest })).json(), { card: null });
    assert.deepEqual(colleagues, [true, true]);
    assert.deepEqual(f.records.get(code), { customerId: 'cus_synthetic' });
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 0);
    assert.equal(f.calls.filter(([operation]) => operation === 'deleteCard').length, 1);
    assert.deepEqual(f.calls.filter(([operation]) => operation === 'clearCompanyCard'), [['clearCompanyCard', companyId], ['clearCompanyCard', companyId]]);
});

test('deletion persistence failures retain the last safe stage and recover without repeated wallet deletion', async (t) => {
    for (const stage of ['delete', 'delete_workbook', 'complete']) {
        await t.test(stage, async (t) => {
            const f = await fixture(t);
            const companyId = seedCompanyPayment(f);
            const code = companyCode(companyId);
            const previous = structuredClone(f.records.get(code));
            const token = await f.login();
            const write = f.cardStore.write;
            f.cardStore.write = async (key, record) => {
                if ((record.pending || 'complete') === stage) throw new Error('PRIVATE persistence failure');
                return write(key, record);
            };
            await expectError(await f.request('card', { method: 'DELETE', token, body: deleteRequest }), 503, 'card_delete_failed');
            assert.deepEqual(f.records.get(code), stage === 'delete' ? previous
                : stage === 'delete_workbook' ? { ...previous, pending: 'delete' }
                    : { customerId: 'cus_synthetic', cardId: card.id, cardRevision: cardRevision(), pending: 'delete_workbook' });
            assert.equal(f.calls.filter(([operation]) => operation === 'deleteCard').length, stage === 'delete' ? 0 : 1);
            assert.equal(f.calls.filter(([operation]) => operation === 'clearCompanyCard').length, stage === 'complete' ? 1 : 0);
            f.cardStore.write = write;
            if (stage === 'delete_workbook') f.pagarme.getCard = async (...references) => { f.calls.push(['getCard', ...references]); return { ...card, status: 'deleted' }; };
            assert.deepEqual(await (await f.request('card', { method: 'DELETE', token, body: deleteRequest })).json(), { card: null });
            assert.deepEqual(f.records.get(code), { customerId: 'cus_synthetic' });
            assert.equal(f.calls.filter(([operation]) => operation === 'deleteCard').length, 1);
            assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, stage === 'delete_workbook' ? 1 : 0);
        });
    }
});

test('repeating deletion without a registered card is idempotent and never invokes the provider', async (t) => {
    const f = await fixture(t);
    const companyId = seedCompanyPayment(f);
    f.records.set(companyCode(companyId), { customerId: 'cus_synthetic' });
    const token = await f.login();
    for (let attempt = 0; attempt < 2; attempt++) {
        assert.deepEqual(await (await f.request('card', { method: 'DELETE', token, body: deleteRequest })).json(), { card: null });
        assert.deepEqual(f.records.get(companyCode(companyId)), { customerId: 'cus_synthetic' });
    }
    assert.equal(f.calls.filter(([operation]) => ['getCard', 'deleteCard', 'findCompanyCustomer', 'createCompanyCustomer', 'createCard'].includes(operation)).length, 0);
});

test('a stale deletion retry cannot remove a freshly registered card, including reused wallet IDs', async (t) => {
    for (const replacementId of [card.id, 'card_replacement']) {
        await t.test(replacementId === card.id ? 'same wallet ID' : 'different wallet ID', async (t) => {
            const f = await fixture(t);
            const companyId = seedCompanyPayment(f);
            const code = companyCode(companyId);
            const token = await f.login();
            const before = await (await f.request('card', { token })).json();
            assert.equal(before.cardRevision, cardRevision());
            for (const sensitive of ['cus_synthetic', card.id]) assert.equal(JSON.stringify(before).includes(sensitive), false);
            const completedDeletion = await f.request('card', { method: 'DELETE', token, body: { expectedCardRevision: before.cardRevision } });
            assert.equal(completedDeletion.status, 200);
            await completedDeletion.text(); // The caller may lose this response after the server completes deletion.
            f.pagarme.createCard = async (customerId, details) => {
                f.calls.push(['createCard', customerId, details]);
                return { ...card, id: replacementId, verificationRequested: true };
            };
            f.pagarme.getCard = async (customerId, cardId) => {
                f.calls.push(['getCard', customerId, cardId]);
                return { ...card, id: cardId };
            };
            const registration = await f.request('card', { method: 'PUT', token, body: cardRequest });
            assert.equal(registration.status, 200);
            const replacement = await registration.json();
            assert.match(replacement.cardRevision, /^[a-f0-9]{64}$/u);
            assert.notEqual(replacement.cardRevision, before.cardRevision);
            const current = structuredClone(f.records.get(code));
            assert.equal(current.cardId, replacementId);
            assert.equal(current.cardRevision, replacement.cardRevision);
            f.calls.length = 0;
            await expectError(await f.request('card', { method: 'DELETE', token, body: { expectedCardRevision: before.cardRevision } }), 409, 'card_changed');
            assert.deepEqual(f.records.get(code), current);
            assert.ok(f.calls.every(([operation]) => operation === 'readUsers' || operation === 'read'));
            const remaining = await (await f.request('card', { token })).json();
            assert.equal(remaining.cardRevision, replacement.cardRevision);
            assert.equal(remaining.card.lastFour, card.lastFour);
        });
    }
});

test('sandbox config and login expose only public fields and issue an opaque token', async (t) => {
    const f = await fixture(t);
    const config = await f.request('config');
    assert.deepEqual(await config.json(), { environment: 'sandbox', publicKey: 'pk_test_synthetic' });
    assert.equal(config.headers.get('cache-control'), 'no-store');
    assert.match(config.headers.get('content-security-policy'), /frame-ancestors 'none'/);
    const response = await f.request('login', { method: 'POST', body: { email: ' ADMIN@EXAMPLE.TEST ', password: admin.password } });
    assert.equal(response.status, 200);
    const body = await response.json();
    assert.match(body.token, /^[A-Za-z0-9_-]{43}$/);
    assert.deepEqual(body.user, { name: admin.name, company: admin.company, role: admin.role, email: admin.email, bookingAllowed: true });
    assert.equal(body.expiresAt, 1_800_000_000_000 + 60 * 60 * 1000);
    assert.deepEqual(Object.keys(body).sort(), ['expiresAt', 'token', 'user']);
    const session = await f.request('session', { token: body.token });
    assert.deepEqual(await session.json(), { user: body.user, expiresAt: body.expiresAt });
});

test('CNPJ keeps sessions and company card identity stable across company name changes', async (t) => {
    const f = await fixture(t);
    const companyId = '12345678000195';
    f.state.users[0].companyId = companyId;
    f.state.users[1].companyId = companyId;
    const token = await f.login();
    const code = companyCode(companyId);
    const record = { customerId: 'cus_synthetic', cardId: card.id, verification, workbookSynced: true, workbookReferencesSynced: true };
    f.records.set(code, record);
    f.state.users[0].company = 'Empresa A renomeada';
    f.state.users[1].company = 'Outra apresentação da mesma razão social';

    assert.deepEqual(await (await f.request('session', { token })).json(), {
        user: { name: admin.name, company: f.state.users[0].company, role: admin.role, email: admin.email, bookingAllowed: true },
        expiresAt: 1_800_000_000_000 + 60 * 60 * 1000,
    });
    const colleagueToken = await f.login(colleague);
    for (const currentToken of [token, colleagueToken]) {
        assert.equal((await f.request('card', { token: currentToken })).status, 200);
    }
    assert.ok(f.calls.filter(([operation]) => operation === 'read').every(([, key]) => key === code));
    assert.deepEqual(f.records.get(code), record);
    assert.notEqual(code, companyCode(admin.company));
    f.state.users[0].companyId = '98765432000110';
    await expectError(await f.request('session', { token }), 401, 'session_expired');
});

test('workbook card references recover the customer without creating verification evidence', async (t) => {
    const f = await fixture(t);
    const companyId = '12345678000195';
    const references = { customerId: 'cus_synthetic', cardId: card.id, environment: 'sandbox' };
    f.state.users[0].companyId = companyId;
    f.users.readCompanyCardReferences = async (id) => {
        f.calls.push(['readCompanyCardReferences', id]);
        return references;
    };
    f.cardStore.findByReferences = async (received) => {
        f.calls.push(['findByReferences', received]);
        return null;
    };
    const token = await f.login();
    assertCardResponse(await (await f.request('card', { token })).json(), { card: null, verificationRequired: true });
    assert.deepEqual(f.records.get(companyCode(companyId)), { customerId: references.customerId, cardId: references.cardId });
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 0);
    await expectError(await f.request('card/sync', { method: 'POST', token, body: {} }), 409, 'card_verification_required');

    assert.equal((await f.request('card', { method: 'PUT', token, body: cardRequest })).status, 200);
    assert.equal(f.calls.filter(([operation]) => operation === 'createCompanyCustomer' || operation === 'findCompanyCustomer').length, 0);
    assert.equal(f.calls.filter(([operation]) => operation === 'createCard').length, 1);
    assert.deepEqual(f.calls.find(([operation]) => operation === 'readCompanyCardReferences'), ['readCompanyCardReferences', companyId]);
    assert.equal(f.calls.find(([operation]) => operation === 'writeCompanyCard')[1], companyId);
    assertStoredRevision(f.records.get(companyCode(companyId)), { customerId: references.customerId, cardId: references.cardId, verification, workbookSynced: true, workbookReferencesSynced: true });
});

test('CNPJ recovery preserves local verification only when both provider references match exactly', async (t) => {
    for (const changed of [{}, { customerId: 'cus_Other' }, { cardId: 'card_Other' }]) {
        await t.test(Object.keys(changed)[0] || 'exact match', async (t) => {
            const root = await fs.mkdtemp(path.join(os.tmpdir(), 'support-card-cnpj-recovery-'));
            t.after(() => fs.rm(root, { recursive: true, force: true }));
            const store = createSupportCardStore(root);
            const previous = { customerId: 'cus_synthetic', cardId: card.id, cardRevision: 'c'.repeat(64), verification, workbookSynced: true, workbookReferencesSynced: true };
            const legacyCode = companyCode(admin.company);
            await store.write(legacyCode, previous);
            const f = await fixture(t, { cardStore: store });
            const companyId = '12345678000195';
            f.state.users[0].companyId = companyId;
            const references = { customerId: previous.customerId, cardId: previous.cardId, environment: 'sandbox', ...changed };
            f.users.readCompanyCardReferences = async () => references;
            const token = await f.login();
            const response = await f.request('card', { token });
            assert.equal(response.status, 200);
            const migrated = await store.read(companyCode(companyId));
            assert.deepEqual(await store.read(legacyCode), previous);
            if (!Object.keys(changed).length) {
                assert.deepEqual(migrated, { customerId: previous.customerId, cardId: previous.cardId, cardRevision: previous.cardRevision, verification });
                const result = await response.json();
                assert.equal(result.workbookSyncPending, true);
                assert.equal(result.cardRevision, previous.cardRevision);
                assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 1);
            } else {
                assert.deepEqual(migrated, { customerId: references.customerId, cardId: references.cardId });
                assertCardResponse(await response.json(), { card: null, verificationRequired: true });
                assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 0);
            }
            assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard' || operation.startsWith('create')).length, 0);
        });
    }
});

test('CNPJ recovery retains pending card and workbook states instead of accepting uncertain operations', async (t) => {
    for (const pending of ['card', 'workbook']) {
        await t.test(pending, async (t) => {
            const root = await fs.mkdtemp(path.join(os.tmpdir(), 'support-card-cnpj-pending-'));
            t.after(() => fs.rm(root, { recursive: true, force: true }));
            const store = createSupportCardStore(root);
            const previous = { customerId: 'cus_synthetic', cardId: card.id, verification, pending };
            await store.write(companyCode(admin.company), previous);
            const f = await fixture(t, { cardStore: store });
            const companyId = '12345678000195';
            f.state.users[0].companyId = companyId;
            f.users.readCompanyCardReferences = async () => ({ customerId: previous.customerId, cardId: previous.cardId, environment: 'sandbox' });
            const token = await f.login();
            const response = await f.request('card', { token });
            if (pending === 'card') {
                await expectError(response, 409, 'card_review_required');
                assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 0);
            } else {
                assert.equal(response.status, 200);
                assert.equal((await response.json()).workbookSyncPending, true);
            }
            assert.deepEqual(await store.read(companyCode(companyId)), previous);
            assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard' || operation.startsWith('create')).length, 0);
        });
    }
});

test('CNPJ recovery never guesses a card from company labels when workbook references are missing', async (t) => {
    const f = await fixture(t);
    const companyId = '12345678000195';
    f.state.users[0].companyId = companyId;
    const previous = { customerId: 'cus_synthetic', cardId: card.id, verification };
    f.records.set(companyCode(admin.company), previous);
    f.users.readCompanyCardReferences = async () => null;
    f.cardStore.findByReferences = async () => { throw new Error('Lookup must not run without references'); };
    const token = await f.login();
    assert.deepEqual(await (await f.request('card', { token })).json(), { card: null });
    assert.equal(f.records.has(companyCode(companyId)), false);
    assert.deepEqual(f.records.get(companyCode(admin.company)), previous);
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard' || operation.startsWith('create') || operation === 'writeCompanyCard').length, 0);
});

test('ambiguous local or unavailable workbook references fail closed without provider calls', async (t) => {
    for (const variant of ['duplicate-local', 'conflicting-workbook']) {
        await t.test(variant, async (t) => {
            const root = await fs.mkdtemp(path.join(os.tmpdir(), 'support-card-cnpj-ambiguous-'));
            t.after(() => fs.rm(root, { recursive: true, force: true }));
            const store = createSupportCardStore(root);
            const previous = { customerId: 'cus_synthetic', cardId: card.id, verification };
            await store.write(companyCode(admin.company), previous);
            if (variant === 'duplicate-local') await store.write(companyCode(otherAdmin.company), previous);
            const f = await fixture(t, { cardStore: store });
            const companyId = '12345678000195';
            f.state.users[0].companyId = companyId;
            f.users.readCompanyCardReferences = async () => {
                if (variant === 'conflicting-workbook') throw new Error('PRIVATE ambiguous workbook references');
                return { customerId: previous.customerId, cardId: previous.cardId, environment: 'sandbox' };
            };
            const token = await f.login();
            await expectError(await f.request('card', { token }), 503, 'card_unavailable');
            await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 503, 'card_save_failed');
            assert.deepEqual(await store.read(companyCode(companyId)), {});
            assert.deepEqual(await store.read(companyCode(admin.company)), previous);
            assert.ok(f.calls.every(([operation]) => operation === 'readUsers'));
        });
    }
});

test('invalid credentials, missing users and ambiguous emails fail without provider calls', async (t) => {
    for (const variant of ['wrong-password', 'missing-user', 'duplicate', 'missing-company']) {
        await t.test(variant, async (t) => {
            const f = await fixture(t);
            if (variant === 'missing-user') f.state.users.shift();
            if (variant === 'duplicate') f.state.users.push({ ...admin });
            if (variant === 'missing-company') f.state.users[0].company = ' ';
            await expectError(await f.request('login', { method: 'POST', body: { email: admin.email, password: variant === 'wrong-password' ? 'wrong' : admin.password } }), 401, 'invalid_credentials');
            assert.ok(f.calls.every(([operation]) => operation === 'readUsers'));
        });
    }
});

test('a listed user with booking permission Não can log in without a login-status field', async (t) => {
    const f = await fixture(t);
    f.state.users[0].bookingAllowed = false;
    assert.equal(Object.hasOwn(f.state.users[0], 'status'), false);
    const response = await f.request('login', { method: 'POST', body: { email: admin.email, password: admin.password } });
    assert.equal(response.status, 200);
    const result = await response.json();
    assert.equal(result.user.bookingAllowed, false);
    assert.equal(Object.hasOwn(result.user, 'status'), false);
    assert.equal((await f.request('session', { token: result.token })).status, 200);
    assert.ok(f.calls.every(([operation]) => operation === 'readUsers'));
});

test('login throttles requests and opens again after its window', async (t) => {
    const f = await fixture(t);
    for (let attempt = 0; attempt < 10; attempt += 1) {
        assert.equal((await f.request('login', { method: 'POST', body: { email: admin.email, password: 'wrong' } })).status, 401);
    }
    await expectError(await f.request('login', { method: 'POST', body: { email: admin.email, password: admin.password } }), 429, 'try_later');
    f.advance(60_000);
    await f.login();
});

test('missing and tampered sessions fail before workbook access and expire at one hour', async (t) => {
    const f = await fixture(t);
    const token = await f.login();
    f.calls.length = 0;
    for (const badToken of [undefined, 'short', `${token[0] === 'a' ? 'b' : 'a'}${token.slice(1)}`]) {
        await expectError(await f.request('session', { token: badToken }), 401, 'session_expired');
    }
    assert.equal(f.calls.length, 0);
    f.advance(60 * 60 * 1000 - 1);
    assert.equal((await f.request('session', { token })).status, 200);
    f.advance(1);
    await expectError(await f.request('session', { token }), 401, 'session_expired');
});

test('a persisted session survives backend recreation and keeps its original one-hour deadline', async (t) => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'support-session-restart-'));
    t.after(() => fs.rm(directory, { recursive: true, force: true }));
    const issuedAt = 1_800_000_000_000;
    let clock = issuedAt;
    const first = await fixture(t, { sessionStore: await openSupportSessionStore(directory), now: () => clock });
    const token = await first.login();
    await first.close();
    const contents = await fs.readFile(path.join(directory, 'sessions.json'), 'utf8');
    assert.equal(contents.includes(token), false);
    assert.equal(contents.includes(admin.password), false);
    assert.equal(JSON.parse(contents).sessions[0].expiresAt, issuedAt + 60 * 60 * 1000);
    clock += 30 * 60 * 1000;
    const second = await fixture(t, { sessionStore: await openSupportSessionStore(directory), now: () => clock });
    assert.deepEqual(await (await second.request('session', { token })).json(), {
        user: { name: admin.name, company: admin.company, role: admin.role, email: admin.email, bookingAllowed: true },
        expiresAt: issuedAt + 60 * 60 * 1000,
    });
    clock = issuedAt + 60 * 60 * 1000 - 1;
    assert.equal((await second.request('session', { token })).status, 200);
    await second.close();
    const third = await fixture(t, { sessionStore: await openSupportSessionStore(directory), now: () => clock });
    assert.equal((await third.request('session', { token })).status, 200);
    clock += 1;
    await expectError(await third.request('session', { token }), 401, 'session_expired');
    await third.close();
    assert.deepEqual(JSON.parse(await fs.readFile(path.join(directory, 'sessions.json'), 'utf8')).sessions, []);
});

test('logout remains revoked after backend recreation', async (t) => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'support-session-logout-'));
    t.after(() => fs.rm(directory, { recursive: true, force: true }));
    const first = await fixture(t, { sessionStore: await openSupportSessionStore(directory) });
    const token = await first.login();
    assert.equal((await first.request('logout', { method: 'POST', token })).status, 204);
    await first.close();
    const second = await fixture(t, { sessionStore: await openSupportSessionStore(directory) });
    await expectError(await second.request('session', { token }), 401, 'session_expired');
    assert.equal(second.calls.length, 0);
    await second.close();
});

test('logout revokes only its bearer without workbook reads and remains available during a workbook outage', async (t) => {
    const f = await fixture(t);
    const token = await f.login();
    const otherToken = await f.login(colleague);
    let workbookReads = 0;
    f.users.readUsers = async () => { workbookReads++; throw new Error('Synthetic workbook outage'); };
    f.calls.length = 0;
    assert.equal((await f.request('logout', { method: 'POST', token })).status, 204);
    assert.equal((await f.request('logout', { method: 'POST', token })).status, 204);
    await expectError(await f.request('logout', { method: 'POST' }), 401, 'session_expired');
    await expectError(await f.request('logout', { method: 'POST', headers: { Authorization: 'Bearer malformed' } }), 401, 'session_expired');
    assert.equal(workbookReads, 0);
    assert.equal(f.calls.length, 0);
    await expectError(await f.request('session', { token }), 401, 'session_expired');
    assert.equal(workbookReads, 0);
    f.users.readUsers = async () => f.state.users;
    assert.equal((await f.request('session', { token: otherToken })).status, 200);
});

test('recreated backends still enforce fresh workbook presence, password, company and administrator access', async (t) => {
    for (const patch of [{ email: 'removed@example.test' }, { password: 'changed' }, { company: 'Empresa B' }, { role: 'Usuário' }]) {
        await t.test(Object.keys(patch)[0], async (t) => {
            const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'support-session-access-'));
            t.after(() => fs.rm(directory, { recursive: true, force: true }));
            const first = await fixture(t, { sessionStore: await openSupportSessionStore(directory) });
            const token = await first.login();
            await first.close();
            const second = await fixture(t, { sessionStore: await openSupportSessionStore(directory) });
            Object.assign(second.state.users[0], patch);
            if (patch.role) {
                const response = await second.request('session', { token });
                assert.equal(response.status, 200);
                assert.equal((await response.json()).user.role, 'Usuário');
                await expectError(await second.request('card', { method: 'PUT', token, body: cardRequest }), 403, 'administrator_required');
            } else {
                await expectError(await second.request('session', { token }), 401, 'session_expired');
                second.state.users[0] = { ...admin };
                await second.close();
                const third = await fixture(t, { sessionStore: await openSupportSessionStore(directory) });
                await expectError(await third.request('session', { token }), 401, 'session_expired');
                await third.close();
            }
            assert.ok(second.calls.every(([operation]) => operation === 'readUsers'));
            await second.close();
        });
    }
});

test('session persistence failures are sanitized instead of escaping async handlers or claiming logout', async (t) => {
    for (const operation of ['prune', 'get', 'set', 'delete']) {
        await t.test(operation, async (t) => {
            const store = createMemorySupportSessionStore();
            const f = await fixture(t, { sessionStore: store });
            const token = await f.login();
            t.mock.method(store, operation, async () => { throw new Error('PRIVATE session file failure'); });
            if (operation === 'set') {
                await expectError(await f.request('login', { method: 'POST', body: { email: admin.email, password: admin.password } }), 503, 'users_unavailable');
            } else if (operation === 'delete') {
                await expectError(await f.request('logout', { method: 'POST', token }), 503, 'session_unavailable');
                assert.equal((await f.request('session', { token })).status, 200);
            } else {
                await expectError(await f.request('session', { token }), 503, 'users_unavailable');
                if (operation === 'prune') {
                    await expectError(await f.request('login', { method: 'POST', body: { email: admin.email, password: admin.password } }), 503, 'users_unavailable');
                }
            }
        });
    }
});

test('current workbook presence, password and company changes revoke an existing session', async (t) => {
    for (const patch of [{ email: 'removed@example.test' }, { password: 'changed' }, { company: 'Empresa B' }]) {
        await t.test(Object.keys(patch)[0], async (t) => {
            const f = await fixture(t);
            const token = await f.login();
            Object.assign(f.state.users[0], patch);
            await expectError(await f.request('card', { token }), 401, 'session_expired');
            f.state.users[0] = { ...admin };
            await expectError(await f.request('session', { token }), 401, 'session_expired');
            assert.ok(f.calls.every(([operation]) => operation === 'readUsers'));
        });
    }
});

test('administrator demotion and ordinary users cannot store a card', async (t) => {
    const f = await fixture(t);
    const adminToken = await f.login();
    f.state.users[0].role = 'Usuário';
    const colleagueToken = await f.login(colleague);
    for (const token of [adminToken, colleagueToken]) {
        await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 403, 'administrator_required');
    }
    assert.ok(f.calls.every(([operation]) => operation === 'readUsers'));
});

test('company card is shared with colleagues and isolated from another company', async (t) => {
    const f = await fixture(t);
    const adminToken = await f.login();
    const colleagueToken = await f.login(colleague);
    const otherToken = await f.login(otherAdmin);
    const saved = await f.request('card', { method: 'PUT', token: adminToken, body: cardRequest });
    assert.equal(saved.status, 200);
    const publicCard = { brand: 'Visa', lastFour: '0010', expMonth: 12, expYear: 2030 };
    assertCardResponse(await saved.json(), { card: publicCard });
    assertStoredRevision(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', cardId: 'card_synthetic', verification, workbookSynced: true, workbookReferencesSynced: true });
    const colleagueCard = await f.request('card', { token: colleagueToken });
    assertCardResponse(await colleagueCard.json(), { card: publicCard });
    const otherCard = await f.request('card', { token: otherToken });
    assert.deepEqual(await otherCard.json(), { card: null });
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 1);
    assert.equal(companyCode('  EMPRESA   A '), companyCode(admin.company));
    assert.notEqual(companyCode(admin.company), companyCode(otherAdmin.company));
    assert.ok([...f.records.values()].every((record) => !JSON.stringify(record).includes('token_')));
    assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 1);
    assert.deepEqual(f.calls.find(([operation]) => operation === 'writeCompanyCard')[3], {
        customerId: 'cus_synthetic', cardId: 'card_synthetic', environment: 'sandbox',
    });
    const providerIndex = f.calls.findIndex(([operation]) => operation === 'createCard');
    const pendingIndex = f.calls.findIndex(([operation, , record]) => operation === 'write' && record.pending === 'workbook');
    const workbookIndex = f.calls.findIndex(([operation]) => operation === 'writeCompanyCard');
    assert.ok(providerIndex < pendingIndex && pendingIndex < workbookIndex);
});

test('confirmed provider cards remain readable when workbook synchronization fails, without provider replay', async (t) => {
    const f = await fixture(t);
    let writes = 0;
    f.users.writeCompanyCard = async () => { writes++; throw new Error('private workbook transport failure'); };
    const token = await f.login();
    const saved = await f.request('card', { method: 'PUT', token, body: cardRequest });
    assert.equal(saved.status, 200);
    const result = await saved.json();
    assert.equal(result.workbookSyncPending, true);
    assert.deepEqual(Object.keys(result.card).sort(), ['brand', 'expMonth', 'expYear', 'lastFour']);
    assertStoredRevision(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', cardId: card.id, verification, pending: 'workbook' });
    const read = await f.request('card', { token });
    assert.deepEqual(await read.json(), result);
    assert.equal(writes, 1, 'GET must not mutate the workbook');
    await expectError(await f.request('card', { method: 'PUT', token, body: { ...cardRequest, token: 'token_fresh' } }), 409, 'workbook_sync_required');
    assert.equal(f.calls.filter(([operation]) => operation === 'createCard').length, 1);
    assert.equal(writes, 1);
});

test('administrator synchronization retries only the workbook for its own confirmed card', async (t) => {
    const f = await fixture(t);
    f.records.set(companyCode(admin.company), { customerId: 'cus_synthetic', cardId: card.id, verification, pending: 'workbook' });
    const token = await f.login();
    const response = await f.request('card/sync', { method: 'POST', token });
    assert.equal(response.status, 200);
    assertCardResponse(await response.json(), { card: { brand: 'Visa', lastFour: '0010', expMonth: 12, expYear: 2030 } });
    assert.deepEqual(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', cardId: card.id, cardRevision: cardRevision(), verification, workbookSynced: true, workbookReferencesSynced: true });
    const writes = f.calls.filter(([operation]) => operation === 'writeCompanyCard');
    assert.equal(writes.length, 1);
    assert.equal(writes[0][1], admin.company);
    assert.equal(writes[0][2].holderName, card.holderName);
    assert.deepEqual(writes[0][3], { customerId: 'cus_synthetic', cardId: card.id, environment: 'sandbox' });
    assert.equal(f.calls.filter(([operation]) => operation === 'createCompanyCustomer' || operation === 'createCard').length, 0);
    const otherToken = await f.login(otherAdmin);
    await expectError(await f.request('card/sync', { method: 'POST', token: otherToken }), 404, 'card_unavailable');
    assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 1);
});

test('workbook sync requires current administrator access and rejects untrusted company or card overrides', async (t) => {
    const f = await fixture(t);
    const token = await f.login();
    f.records.set(companyCode(admin.company), { customerId: 'cus_synthetic', cardId: card.id, verification, pending: 'workbook' });
    await expectError(await f.request('card/sync', { method: 'POST', token, body: { company: otherAdmin.company, cardId: 'card_other' } }), 400, 'invalid_request');
    const colleagueToken = await f.login(colleague);
    await expectError(await f.request('card/sync', { method: 'POST', token: colleagueToken }), 403, 'administrator_required');
    const colleagueCard = await f.request('card', { token: colleagueToken });
    assert.equal((await colleagueCard.json()).workbookSyncPending, true);
    f.state.users[0].role = 'Usuário';
    await expectError(await f.request('card/sync', { method: 'POST', token }), 403, 'administrator_required');
    assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 0);
});

test('workbook sync failures preserve the confirmed card and can be retried without repeating provider creation', async (t) => {
    const f = await fixture(t);
    const record = { customerId: 'cus_synthetic', cardId: card.id, cardRevision: cardRevision(), verification, pending: 'workbook' };
    f.records.set(companyCode(admin.company), record);
    const token = await f.login();
    f.users.writeCompanyCard = async () => { throw new Error('private workbook failure'); };
    await expectError(await f.request('card/sync', { method: 'POST', token }), 503, 'workbook_sync_failed');
    assert.deepEqual(f.records.get(companyCode(admin.company)), record);
    const write = f.cardStore.write;
    f.cardStore.write = async (code, state) => {
        if (!state.pending) throw new Error('private persistence failure');
        return write(code, state);
    };
    f.users.writeCompanyCard = async () => {};
    await expectError(await f.request('card/sync', { method: 'POST', token }), 503, 'workbook_sync_failed');
    assert.deepEqual(f.records.get(companyCode(admin.company)), record);
    f.cardStore.write = write;
    assert.equal((await f.request('card/sync', { method: 'POST', token })).status, 200);
    assert.deepEqual(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', cardId: card.id, cardRevision: cardRevision(), verification, workbookSynced: true, workbookReferencesSynced: true });
    assert.equal(f.calls.filter(([operation]) => operation === 'createCompanyCustomer' || operation === 'createCard').length, 0);
});

test('verified cards with masked-only workbook sync backfill reusable references without provider creation', async (t) => {
    const f = await fixture(t);
    f.records.set(companyCode(admin.company), { customerId: 'cus_synthetic', cardId: card.id, verification, workbookSynced: true });
    const token = await f.login();
    const before = await (await f.request('card', { token })).json();
    assert.equal(before.workbookSyncPending, true);
    assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 0, 'GET never writes to the workbook');
    assert.equal((await f.request('card/sync', { method: 'POST', token })).status, 200);
    const write = f.calls.find(([operation]) => operation === 'writeCompanyCard');
    assert.deepEqual(write[3], { customerId: 'cus_synthetic', cardId: card.id, environment: 'sandbox' });
    assert.equal(f.records.get(companyCode(admin.company)).workbookReferencesSynced, true);
    const after = await (await f.request('card', { token })).json();
    assert.equal(after.workbookSyncPending, undefined);
    assert.equal(f.calls.filter(([operation]) => operation === 'createCompanyCustomer' || operation === 'createCard').length, 0);
});

test('legacy card references require fresh verification and cannot be promoted by workbook synchronization', async (t) => {
    const f = await fixture(t);
    const token = await f.login();
    for (const flags of [{}, { workbookSynced: true }, { workbookSynced: true, workbookReferencesSynced: true }, { pending: 'workbook' }]) {
        f.records.set(companyCode(admin.company), { customerId: 'cus_synthetic', cardId: 'card_legacy', ...flags });
        const before = await f.request('card', { token });
        assertCardResponse(await before.json(), { card: null, verificationRequired: true });
        await expectError(await f.request('card/sync', { method: 'POST', token }), 409, 'card_verification_required');
    }
    assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 0);
    assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 0, 'An active wallet response is not verification evidence');
    assert.equal((await f.request('card', { method: 'PUT', token, body: cardRequest })).status, 200);
    const after = await f.request('card', { token });
    assert.equal((await after.json()).workbookSyncPending, undefined);
    assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 1);
    for (const record of [{ pending: 'customer' }, { customerId: 'cus_synthetic', pending: 'card' }]) {
        f.records.set(companyCode(admin.company), record);
        await expectError(await f.request('card/sync', { method: 'POST', token }), 409, 'card_review_required');
    }
    assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 1);
});

test('fresh accepted verification registers a legacy card even when Pagar.me returns its existing wallet ID', async (t) => {
    for (const flags of [{}, { workbookSynced: true }, { pending: 'workbook' }]) {
        await t.test(JSON.stringify(flags), async (t) => {
            const f = await fixture(t);
            f.records.set(companyCode(admin.company), { customerId: 'cus_synthetic', cardId: card.id, ...flags });
            const token = await f.login();
            assertCardResponse(await (await f.request('card', { token })).json(), { card: null, verificationRequired: true });
            assert.equal(f.calls.filter(([operation]) => operation === 'getCard').length, 0, 'GET cannot grant verification to the old wallet record');
            const response = await f.request('card', { method: 'PUT', token, body: cardRequest });
            assert.equal(response.status, 200);
            assertCardResponse(await response.json(), { card: { brand: card.brand, lastFour: card.lastFour, expMonth: card.expMonth, expYear: card.expYear } });
            assertStoredRevision(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', cardId: card.id, verification, workbookSynced: true, workbookReferencesSynced: true });
            assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 1);
            assert.deepEqual(f.calls.find(([operation]) => operation === 'createCard'), ['createCard', 'cus_synthetic', cardRequest]);
            assert.equal(f.calls.filter(([operation]) => operation === 'findCompanyCustomer' || operation === 'createCompanyCustomer').length, 0);
            assert.equal((await (await f.request('card', { token })).json()).card.lastFour, card.lastFour);
        });
    }
});

test('legacy re-registration without accepted verification retains its references and blocks replay', async (t) => {
    for (const failure of ['missing evidence', 'ambiguous request']) {
        await t.test(failure, async (t) => {
            const f = await fixture(t);
            const existing = { customerId: 'cus_synthetic', cardId: card.id };
            f.records.set(companyCode(admin.company), existing);
            f.pagarme.createCard = async () => {
                if (failure === 'ambiguous request') throw new Error('private provider failure');
                return { ...card };
            };
            const token = await f.login();
            await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 503, 'card_save_failed');
            assert.deepEqual(f.records.get(companyCode(admin.company)), { ...existing, pending: 'card' });
            assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 0);
            await expectError(await f.request('card', { token }), 409, 'card_review_required');
            await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 409, 'card_review_required');
        });
    }
});

test('registration requires explicit accepted verification plus an active unexpired card', async (t) => {
    for (const patch of [
        { verificationRequested: undefined }, { verificationRequested: false },
        { status: undefined }, { status: 'deleted' }, { status: 'expired' },
        { expYear: 2026, expMonth: 12 }, { expYear: 26, expMonth: 12 },
    ]) {
        await t.test(JSON.stringify(patch), async (t) => {
            const f = await fixture(t);
            f.pagarme.createCard = async () => ({ ...card, verificationRequested: true, ...patch });
            const token = await f.login();
            await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 503, 'card_save_failed');
            assert.deepEqual(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', pending: 'card' });
            assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 0);
            await expectError(await f.request('card', { token }), 409, 'card_review_required');
            await expectError(await f.request('card/sync', { method: 'POST', token }), 409, 'card_review_required');
        });
    }
});

test('registered reads and workbook writes reject revoked or expired cards despite a verification marker', async (t) => {
    for (const patch of [{ status: 'deleted' }, { status: undefined }, { expYear: 2026, expMonth: 12 }]) {
        await t.test(JSON.stringify(patch), async (t) => {
            const f = await fixture(t);
            const existing = { customerId: 'cus_synthetic', cardId: 'card_previous', verification, pending: 'workbook' };
            f.records.set(companyCode(admin.company), existing);
            f.pagarme.getCard = async () => ({ ...card, ...patch });
            const token = await f.login();
            assertCardResponse(await (await f.request('card', { token })).json(), { card: null, verificationRequired: true });
            await expectError(await f.request('card/sync', { method: 'POST', token }), 409, 'card_verification_required');
            assert.deepEqual(f.records.get(companyCode(admin.company)), existing);
            assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 0);
            assert.equal((await f.request('card', { method: 'PUT', token, body: cardRequest })).status, 200, 'An unusable old card must not block a fresh verified replacement');
        });
    }
});

test('expiry checks use the injected UTC clock and accept the entire expiry month', async (t) => {
    const f = await fixture(t);
    f.pagarme.createCard = async () => ({ ...card, expMonth: 1, expYear: 27, verificationRequested: true });
    f.pagarme.getCard = async () => ({ ...card, expMonth: 1, expYear: 2027 });
    const token = await f.login();
    assert.equal((await f.request('card', { method: 'PUT', token, body: cardRequest })).status, 200);
    assert.ok((await (await f.request('card', { token })).json()).card);
    f.advance(Date.UTC(2027, 1, 1) - 1_800_000_000_000);
    const refreshedToken = await f.login();
    assertCardResponse(await (await f.request('card', { token: refreshedToken })).json(), { card: null, verificationRequired: true });
});

test('provider verification rejection never writes the workbook and preserves previous card and verification state', async (t) => {
    for (const previous of [{ customerId: 'cus_synthetic' }, { customerId: 'cus_synthetic', cardId: card.id }, { customerId: 'cus_synthetic', cardId: card.id, verification, workbookSynced: true, workbookReferencesSynced: true }]) {
        await t.test(previous.verification ? 'replacement' : previous.cardId ? 'legacy re-registration' : 'first card', async (t) => {
            const f = await fixture(t);
            f.records.set(companyCode(admin.company), previous);
            f.pagarme.createCard = async () => { throw Object.assign(new Error('private provider details'), { code: 'pagarme.card_verification_failed', definiteRejection: true }); };
            const token = await f.login();
            await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 422, 'card_verification_failed');
            assert.deepEqual(f.records.get(companyCode(admin.company)), previous);
            assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 0);
            assert.equal(Boolean((await (await f.request('card', { token })).json()).card), Boolean(previous.verification));
        });
    }
});

test('workbook synchronization and registration share the company operation lock', async (t) => {
    const f = await fixture(t);
    const token = await f.login();
    let release;
    let entered;
    const started = new Promise((resolve) => { entered = resolve; });
    const deferred = new Promise((resolve) => { release = resolve; });
    f.users.writeCompanyCard = async () => { entered(); await deferred; };
    const first = f.request('card', { method: 'PUT', token, body: cardRequest });
    await started;
    await expectError(await f.request('card/sync', { method: 'POST', token }), 409, 'card_busy');
    await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 409, 'card_busy');
    await expectError(await f.request('card', { method: 'DELETE', token, body: deleteRequest }), 409, 'card_busy');
    release();
    assert.equal((await first).status, 200);
    assert.equal(f.calls.filter(([operation]) => operation === 'createCard').length, 1);
    assert.equal(f.calls.filter(([operation]) => operation === 'deleteCard' || operation === 'clearCompanyCard').length, 0);
});

test('raw payment fields and malformed addresses are rejected before storage or provider access', async (t) => {
    const f = await fixture(t);
    const token = await f.login();
    for (const body of [
        { ...cardRequest, number: '4000000000000010' },
        { ...cardRequest, cvv: '123' },
        { ...cardRequest, company: otherAdmin.company },
        { ...cardRequest, customerId: 'cus_untrusted' },
        { ...cardRequest, cardId: 'card_untrusted' },
        { ...cardRequest, environment: 'production' },
        { ...cardRequest, references: { customerId: 'cus_untrusted', cardId: 'card_untrusted', environment: 'sandbox' } },
        { ...cardRequest, verification: 'pagarme_verify_card_v1' },
        { ...cardRequest, verificationRequested: true },
        { ...cardRequest, token: 'card_untrusted' },
        { ...cardRequest, billingAddress: { ...billingAddress, number: '4000000000000010' } },
        { ...cardRequest, billingAddress: { ...billingAddress, country: 'US' } },
        { ...cardRequest, billingAddress: { ...billingAddress, zip_code: 'invalid' } },
    ]) {
        f.calls.length = 0;
        await expectError(await f.request('card', { method: 'PUT', token, body }), 400, 'invalid_card');
        assert.deepEqual(f.calls.map(([operation]) => operation), ['readUsers']);
    }
});

test('ambiguous customer or card creation persists intent and blocks a duplicate replay', async (t) => {
    for (const operation of ['createCompanyCustomer', 'createCard']) {
        await t.test(operation, async (t) => {
            const f = await fixture(t);
            let attempts = 0;
            f.pagarme[operation] = async () => { attempts += 1; throw new Error('ambiguous provider timeout'); };
            const token = await f.login();
            await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 503, 'card_save_failed');
            assert.equal(f.records.get(companyCode(admin.company)).pending, operation === 'createCard' ? 'card' : 'customer');
            await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 409, 'card_review_required');
            await expectError(await f.request('card', { token }), 409, 'card_review_required');
            assert.equal(attempts, 1);
        });
    }
});

test('a definite provider rejection clears its intent and permits a fresh manual submission', async (t) => {
    for (const operation of ['createCompanyCustomer', 'createCard']) {
        await t.test(operation, async (t) => {
            const f = await fixture(t);
            let attempts = 0;
            const succeed = f.pagarme[operation];
            f.pagarme[operation] = async (...args) => {
                attempts += 1;
                if (attempts === 1) throw Object.assign(new Error('provider rejected request'), { definiteRejection: true });
                return succeed(...args);
            };
            const token = await f.login();
            await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 503, 'card_save_failed');
            assert.equal(attempts, 1, 'A rejected request must not retry automatically');
            assert.deepEqual(f.records.get(companyCode(admin.company)), operation === 'createCard' ? { customerId: 'cus_synthetic' } : {});
            assert.equal((await f.request('card', { token })).status, 200);
            assert.equal((await f.request('card', { method: 'PUT', token, body: { ...cardRequest, token: 'token_fresh' } })).status, 200);
            assert.equal(attempts, 2);
            assertStoredRevision(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', cardId: 'card_synthetic', verification, workbookSynced: true, workbookReferencesSynced: true });
        });
    }
});

test('a rejected replacement preserves the current card and a failed rollback remains blocked', async (t) => {
    const f = await fixture(t);
    const existing = { customerId: 'cus_existing', cardId: 'card_existing', verification, workbookSynced: true, workbookReferencesSynced: true };
    f.records.set(companyCode(admin.company), existing);
    let attempts = 0;
    f.pagarme.createCard = async () => {
        attempts += 1;
        throw Object.assign(new Error('provider rejected replacement'), { definiteRejection: true });
    };
    const token = await f.login();
    await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 503, 'card_save_failed');
    assert.deepEqual(f.records.get(companyCode(admin.company)), existing);
    assert.equal((await f.request('card', { token })).status, 200);
    const write = f.cardStore.write;
    f.cardStore.write = async (code, record) => {
        if (!record.pending) throw new Error('rollback could not persist');
        return write(code, record);
    };
    await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 503, 'card_save_failed');
    assert.deepEqual(f.records.get(companyCode(admin.company)), { ...existing, pending: 'card' });
    await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 409, 'card_review_required');
    assert.equal(attempts, 2);
});

test('failed persistence prevents a non-idempotent provider request', async (t) => {
    const f = await fixture(t);
    const token = await f.login();
    f.cardStore.write = async () => { throw new Error('disk full'); };
    await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 503, 'card_save_failed');
    assert.ok(!f.calls.some(([operation]) => operation === 'createCompanyCustomer' || operation === 'createCard'));
});

test('a simultaneous company save is refused while its first save completes', async (t) => {
    const f = await fixture(t);
    const token = await f.login();
    let release;
    let entered;
    const started = new Promise((resolve) => { entered = resolve; });
    const deferred = new Promise((resolve) => { release = resolve; });
    f.pagarme.createCard = async () => { entered(); return deferred; };
    const first = f.request('card', { method: 'PUT', token, body: cardRequest });
    await started;
    await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 409, 'card_busy');
    release({ ...card, verificationRequested: true });
    assert.equal((await first).status, 200);
    assertStoredRevision(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', cardId: 'card_synthetic', verification, workbookSynced: true, workbookReferencesSynced: true });
});

test('provider read failure is sanitized and logout revokes the bearer', async (t) => {
    const f = await fixture(t);
    const token = await f.login();
    f.records.set(companyCode(admin.company), { customerId: 'cus_synthetic', cardId: card.id, verification });
    f.pagarme.getCard = async () => { throw new Error('private upstream data'); };
    await expectError(await f.request('card', { token }), 503, 'card_unavailable');
    assert.equal((await f.request('logout', { method: 'POST', token })).status, 204);
    await expectError(await f.request('session', { token }), 401, 'session_expired');
});

test('workbook transport failure is sanitized during login and authorization', async (t) => {
    let broken = false;
    const f = await fixture(t, { users: { async readUsers() { if (broken) throw new Error('private upstream data'); return [admin]; } } });
    const token = await f.login();
    broken = true;
    await expectError(await f.request('session', { token }), 503, 'users_unavailable');
    await expectError(await f.request('login', { method: 'POST', body: { email: admin.email, password: admin.password } }), 503, 'users_unavailable');
});

test('sandbox enforces exact origin and host before dependencies and rejects non-loopback configuration', async (t) => {
    const f = await fixture(t);
    for (const headers of [{ Origin: 'https://attacker.invalid' }, { Origin: 'null' }]) {
        assert.equal((await f.request('login', { method: 'POST', headers, body: { email: admin.email, password: admin.password } })).status, 403);
    }
    // Fetch overrides Host; use the native HTTP client for DNS-rebinding checks.
    for (const host of ['attacker.invalid', 'localhost']) {
        const status = await new Promise((resolve, reject) => {
            const request = http.get(`${f.origin}/support-management/api/config`, { headers: { Host: host } }, (response) => {
                response.resume();
                resolve(response.statusCode);
            });
            request.once('error', reject);
        });
        assert.equal(status, 403);
    }
    assert.equal(f.calls.length, 0);
    assert.equal((await f.request('config', { headers: { Origin: f.origin } })).status, 200);
    for (const origin of ['http://localhost:4174', 'https://127.0.0.1:4174', 'https://example.test']) {
        assert.throws(() => createSupportApp({ origin }), /loopback/);
    }
});

test('sandbox rejects oversized or malformed JSON before workbook or provider access', async (t) => {
    const f = await fixture(t);
    for (const [body, status] of [['{malformed', 400], [JSON.stringify({ email: 'x'.repeat(9000) }), 413]]) {
        await expectError(await fetch(`${f.origin}/support-management/api/login`, {
            method: 'POST', headers: { 'Content-Type': 'application/json' }, body,
        }), status, 'invalid_request');
    }
    assert.equal(f.calls.length, 0);
});

test('an existing company customer is reused and a failed final save remains pending', async (t) => {
    const f = await fixture(t);
    f.pagarme.findCompanyCustomer = async (code) => ({ id: 'cus_existing', code });
    const write = f.cardStore.write;
    f.cardStore.write = async (code, record) => {
        if (record.cardId) throw new Error('disk full after provider success');
        return write(code, record);
    };
    const token = await f.login();
    await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 503, 'card_save_failed');
    assert.deepEqual(f.records.get(companyCode(admin.company)), { customerId: 'cus_existing', pending: 'card' });
    await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 409, 'card_review_required');
    assert.equal(f.calls.filter(([operation]) => operation === 'createCompanyCustomer').length, 0);
    assert.equal(f.calls.filter(([operation]) => operation === 'createCard').length, 1);
});

test('sandbox and production retain separate route tables', async (t) => {
    const f = await fixture(t);
    assert.equal((await fetch(`${f.origin}/plataforma_v2/login-FaceID`, { method: 'POST' })).status, 404);
    const { app, ledger } = createTestApp();
    const production = await startLoopback(app, t);
    assert.equal((await fetch(`${production.origin}/support-management/api/config`)).status, 404);
    assert.equal(ledger.length, 0);
});

test('card reference store survives reload, replaces records and refuses invalid keys or corrupt state', async (t) => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'support-card-store-test-'));
    t.after(() => fs.rm(directory, { recursive: true, force: true }));
    const code = companyCode(admin.company);
    const store = createSupportCardStore(directory);
    assert.deepEqual(await store.read(code), {});
    await store.write(code, { customerId: 'cus_synthetic', pending: 'card' });
    const reloaded = createSupportCardStore(directory);
    assert.deepEqual(await reloaded.read(code), { customerId: 'cus_synthetic', pending: 'card' });
    await reloaded.write(code, { customerId: 'cus_synthetic', cardId: 'card_synthetic' });
    assert.deepEqual(await store.read(code), { customerId: 'cus_synthetic', cardId: 'card_synthetic' });
    assert.deepEqual(await fs.readdir(directory), [`${code}.json`]);
    await assert.rejects(store.write('../escape', {}), /Invalid company code/);
    await assert.rejects(store.read('../escape'), /storage unavailable/);
    await fs.writeFile(path.join(directory, `${code}.json`), '{corrupt');
    await assert.rejects(reloaded.read(code), /storage unavailable/);
});


test('company permissions list only authenticated company employees and never private workbook fields', async (t) => {
    const f = await fixture(t);
    seedCompanyPayment(f);
    f.state.users.push({ ...f.state.users[1], name: 'Invalid email', email: 'invalid-email', bookingAllowed: false });
    const token = await f.login();
    const result = await (await f.request('permissions', { token })).json();
    assert.deepEqual(result.employees, [
        { name: admin.name, email: admin.email, allowed: true, active: true },
        { name: colleague.name, email: colleague.email, allowed: true, active: true },
        { name: 'Invalid email', email: 'invalid-email', allowed: false, active: false },
    ]);
    assert.ok(!JSON.stringify(result).includes(admin.password));
    const memberToken = await f.login(colleague);
    await expectError(await f.request('permissions', { token: memberToken }), 403, 'administrator_required');
});

test('administrator permission changes derive the company and reject malformed, foreign or ambiguous targets before writes', async (t) => {
    const f = await fixture(t);
    const companyId = seedCompanyPayment(f);
    const token = await f.login();
    const memberToken = await f.login(colleague);
    await expectError(await f.request('permissions', { method: 'PATCH', token: memberToken, body: { employees: [{ email: admin.email, allowed: true }, { email: colleague.email, allowed: true }] } }), 403, 'administrator_required');
    for (const body of [undefined, [], null, {}, { email: colleague.email, allowed: 'Sim' }, { email: 'bad', allowed: true }, { email: colleague.email, allowed: true, companyId },
        { employees: null }, { employees: {} }, { employees: [null] }, { employees: [[]] },
        { employees: [{ email: colleague.email, allowed: 'Sim' }] }, { employees: [{ email: 'bad', allowed: true }] },
        { employees: [{ email: colleague.email, allowed: true, row: 99 }] },
        { employees: [{ email: admin.email, allowed: true }, { email: admin.email.toUpperCase(), allowed: false }] },
        { employees: [], companyId }]) {
        await expectError(await f.request('permissions', { method: 'PATCH', token, body }), 400, 'invalid_request');
    }
    await expectError(await f.request('permissions', { method: 'PATCH', token, body: { employees: [{ email: admin.email, allowed: true }, { email: otherAdmin.email, allowed: true }] } }), 404, 'employee_not_found');
    f.state.users.push({ ...f.state.users[1] });
    await expectError(await f.request('permissions', { method: 'PATCH', token, body: { employees: [{ email: admin.email, allowed: true }, { email: colleague.email, allowed: true }] } }), 404, 'employee_not_found');
    assert.equal(f.calls.some(([kind]) => kind === 'writeBookingPermissions'), false);
});

test('Save validates the full current employee list and reuses only its freshly authorized directory', async (t) => {
    const f = await fixture(t);
    const companyId = seedCompanyPayment(f);
    f.state.users.push({ ...f.state.users[1], email: 'invalid-email' });
    const token = await f.login();
    for (const employees of [[], [{ email: colleague.email, allowed: false }], [
        { email: admin.email, allowed: true }, { email: colleague.email, allowed: true }, { email: 'foreign@example.test', allowed: true },
    ]]) {
        await expectError(await f.request('permissions', { method: 'PATCH', token, body: { employees } }), 404, 'employee_not_found');
    }
    assert.equal(f.calls.some(([kind]) => kind === 'writeBookingPermissions'), false);
    const before = f.calls.length;
    const employees = [{ email: admin.email.toUpperCase(), allowed: false }, { email: colleague.email, allowed: false }];
    const response = await f.request('permissions', { method: 'PATCH', token, body: { employees } });
    assert.equal(response.status, 200);
    assert.deepEqual(f.calls.slice(before).map(([kind]) => kind), ['readUsers', 'writeBookingPermissions']);
    assert.deepEqual(f.calls.at(-1), ['writeBookingPermissions', companyId, [
        { email: admin.email, allowed: false }, { email: colleague.email, allowed: false },
    ]]);
    assert.ok((await response.json()).employees.filter(({ active }) => active).every(({ allowed }) => allowed === false));
    f.state.users[0].role = 'Usuário';
    await expectError(await f.request('permissions', { method: 'PATCH', token, body: { employees } }), 403, 'administrator_required');
});

test('permission changes take effect for an existing session and revoke booking before provider or workbook effects', async (t) => {
    const f = await fixture(t, { appointments: { readDates: async () => [{ date: '2027-02-01', status: '' }] }, bookingStore: { pendingDates: async () => new Set(), read: async () => null } });
    const companyId = seedCompanyPayment(f);
    const token = await f.login();
    const memberToken = await f.login(colleague);
    assert.equal((await (await f.request('calendar', { token: memberToken })).json()).canBook, true);
    const result = await (await f.request('permissions', { method: 'PATCH', token, body: { employees: [{ email: admin.email, allowed: true }, { email: colleague.email, allowed: false }] } })).json();
    assert.equal(result.employees.find((person) => person.email === colleague.email).allowed, false);
    assert.deepEqual(f.calls.filter(([kind]) => kind === 'writeBookingPermissions'), [['writeBookingPermissions', companyId, [{ email: admin.email, allowed: true }, { email: colleague.email, allowed: false }]]]);
    const calendar = await (await f.request('calendar', { token: memberToken })).json();
    assert.equal(calendar.paymentMethodRegistered, true);
    assert.equal(calendar.canBook, false);
    assert.deepEqual(calendar.days, []);
    const before = f.calls.length;
    await expectError(await f.request('reservations', { method: 'POST', token: memberToken, body: { date: '2027-02-01', priceCents: 100000, requestId: 'd0b07b46-1311-41fb-984f-b338fd6b8a16', acknowledged: true } }), 403, 'booking_permission_required');
    assert.ok(f.calls.slice(before).every(([kind]) => kind === 'readUsers'));
    const session = await (await f.request('session', { token: memberToken })).json();
    assert.equal(session.user.bookingAllowed, false);
});

test('permissions cannot change during a company card operation and failed saves remain retryable', async (t) => {
    const f = await fixture(t);
    seedCompanyPayment(f);
    const token = await f.login();
    let release;
    let started;
    const entered = new Promise((resolve) => { started = resolve; });
    f.pagarme.deleteCard = async () => { started(); await new Promise((resolve) => { release = resolve; }); return { deleted: true }; };
    const deleting = f.request('card', { method: 'DELETE', token, body: deleteRequest });
    await entered;
    await expectError(await f.request('permissions', { method: 'PATCH', token, body: { employees: [{ email: admin.email, allowed: true }, { email: colleague.email, allowed: false }] } }), 409, 'permission_busy');
    release();
    assert.equal((await deleting).status, 200);
    f.users.writeBookingPermissions = async () => { throw new Error('private Graph response'); };
    await expectError(await f.request('permissions', { method: 'PATCH', token, body: { employees: [{ email: admin.email, allowed: true }, { email: colleague.email, allowed: false }] } }), 503, 'permission_save_failed');
    f.users.writeBookingPermissions = async () => {};
    assert.equal((await f.request('permissions', { method: 'PATCH', token, body: { employees: [{ email: admin.email, allowed: true }, { email: colleague.email, allowed: false }] } })).status, 200);
});

test('addressless registration still requires explicit Pagar.me verification', async (t) => {
    const f = await fixture(t);
    seedCompanyPayment(f);
    const token = await f.login();
    const result = await f.request('card', { method: 'PUT', token, body: { token: 'token_synthetic' } });
    assert.equal(result.status, 200);
    assert.ok((await result.json()).card);
    assert.equal(f.calls.find(([kind]) => kind === 'createCard')[2].billingAddress, undefined);
});
