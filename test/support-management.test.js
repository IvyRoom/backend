'use strict';

require('./deny-production-network');
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const { once } = require('node:events');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { createSupportApp } = require('../support-app');
const { companyCode } = require('../domains/support-management');
const { createSupportCardStore } = require('../integrations/support-card-store');
const { createTestApp, startLoopback } = require('./app-test-support');

const admin = { company: 'Empresa A', name: 'Administradora teste', email: 'admin@example.test', password: 'test-password', role: 'Administrador', status: 'Ativo' };
const colleague = { ...admin, name: 'Colega teste', email: 'colleague@example.test', role: 'Usuário' };
const otherAdmin = { ...admin, company: 'Empresa B', email: 'other@example.test' };
const verification = 'pagarme_verify_card_v1';
const card = { id: 'card_synthetic', brand: 'Visa', lastFour: '0010', holderName: 'SYNTHETIC HOLDER', expMonth: 12, expYear: 2030, status: 'active' };
const billingAddress = { line_1: '1, Rua Teste, Centro', zip_code: '01001000', city: 'São Paulo', state: 'SP', country: 'BR' };
const cardRequest = { token: 'token_synthetic', billingAddress };

async function fixture(t, overrides = {}) {
    let clock = 1_800_000_000_000;
    const state = { users: structuredClone([admin, colleague, otherAdmin]) };
    const records = new Map();
    const calls = [];
    const users = {
        async readUsers() { calls.push(['readUsers']); return state.users; },
        async writeCompanyCard(company, details, references) { calls.push(['writeCompanyCard', company, structuredClone(details), structuredClone(references)]); },
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
    };
    let app;
    const server = http.createServer((req, res) => app(req, res));
    server.listen(0, '127.0.0.1');
    await once(server, 'listening');
    t.after(async () => { server.closeAllConnections(); await new Promise((resolve) => server.close(resolve)); });
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
    return { origin, state, records, calls, users, cardStore, pagarme, request, login, advance(ms) { clock += ms; } };
}

function expectError(response, status, reason) {
    assert.equal(response.status, status);
    return response.json().then((body) => assert.deepEqual(body, { error: `support_management.${reason}` }));
}

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
    assert.deepEqual(body.user, { name: admin.name, company: admin.company, role: admin.role });
    assert.deepEqual(Object.keys(body).sort(), ['token', 'user']);
    const session = await f.request('session', { token: body.token });
    assert.deepEqual(await session.json(), { user: body.user });
});

test('invalid credentials, inactive users and ambiguous emails fail without provider calls', async (t) => {
    for (const variant of ['wrong-password', 'inactive', 'duplicate', 'missing-company']) {
        await t.test(variant, async (t) => {
            const f = await fixture(t);
            if (variant === 'inactive') f.state.users[0].status = 'Inativo';
            if (variant === 'duplicate') f.state.users.push({ ...admin });
            if (variant === 'missing-company') f.state.users[0].company = ' ';
            await expectError(await f.request('login', { method: 'POST', body: { email: admin.email, password: variant === 'wrong-password' ? 'wrong' : admin.password } }), 401, 'invalid_credentials');
            assert.ok(f.calls.every(([operation]) => operation === 'readUsers'));
        });
    }
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

test('missing and tampered sessions fail before workbook access and expire at four hours', async (t) => {
    const f = await fixture(t);
    const token = await f.login();
    f.calls.length = 0;
    for (const badToken of [undefined, 'short', `${token[0] === 'a' ? 'b' : 'a'}${token.slice(1)}`]) {
        await expectError(await f.request('session', { token: badToken }), 401, 'session_expired');
    }
    assert.equal(f.calls.length, 0);
    f.advance(4 * 60 * 60 * 1000 - 1);
    assert.equal((await f.request('session', { token })).status, 200);
    f.advance(1);
    await expectError(await f.request('session', { token }), 401, 'session_expired');
});

test('current workbook status, password and company changes revoke an existing session', async (t) => {
    for (const patch of [{ status: 'Inativo' }, { password: 'changed' }, { company: 'Empresa B' }]) {
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
    assert.deepEqual(await saved.json(), { card: publicCard });
    assert.deepEqual(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', cardId: 'card_synthetic', verification, workbookSynced: true, workbookReferencesSynced: true });
    const colleagueCard = await f.request('card', { token: colleagueToken });
    assert.deepEqual(await colleagueCard.json(), { card: publicCard });
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
    assert.deepEqual(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', cardId: card.id, verification, pending: 'workbook' });
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
    assert.deepEqual(await response.json(), { card: { brand: 'Visa', lastFour: '0010', expMonth: 12, expYear: 2030 } });
    assert.deepEqual(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', cardId: card.id, verification, workbookSynced: true, workbookReferencesSynced: true });
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
    const record = { customerId: 'cus_synthetic', cardId: card.id, verification, pending: 'workbook' };
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
    assert.deepEqual(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', cardId: card.id, verification, workbookSynced: true, workbookReferencesSynced: true });
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
        assert.deepEqual(await before.json(), { card: null, verificationRequired: true });
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

test('a reused legacy wallet card is not silently accepted as freshly verified', async (t) => {
    const f = await fixture(t);
    const existing = { customerId: 'cus_synthetic', cardId: card.id, workbookSynced: true };
    f.records.set(companyCode(admin.company), existing);
    const token = await f.login();
    await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 422, 'card_verification_unconfirmed');
    assert.deepEqual(f.records.get(companyCode(admin.company)), existing);
    assert.deepEqual(await (await f.request('card', { token })).json(), { card: null, verificationRequired: true });
    assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 0);
    assert.equal(f.calls.filter(([operation]) => operation === 'createCard').length, 1);
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
            assert.deepEqual(await (await f.request('card', { token })).json(), { card: null, verificationRequired: true });
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
    assert.deepEqual(await (await f.request('card', { token: refreshedToken })).json(), { card: null, verificationRequired: true });
});

test('provider verification rejection never writes the workbook and preserves a previously verified card', async (t) => {
    for (const previous of [{ customerId: 'cus_synthetic' }, { customerId: 'cus_synthetic', cardId: card.id, verification, workbookSynced: true, workbookReferencesSynced: true }]) {
        await t.test(previous.cardId ? 'replacement' : 'first card', async (t) => {
            const f = await fixture(t);
            f.records.set(companyCode(admin.company), previous);
            f.pagarme.createCard = async () => { throw Object.assign(new Error('private provider details'), { code: 'pagarme.card_verification_failed', definiteRejection: true }); };
            const token = await f.login();
            await expectError(await f.request('card', { method: 'PUT', token, body: cardRequest }), 422, 'card_verification_failed');
            assert.deepEqual(f.records.get(companyCode(admin.company)), previous);
            assert.equal(f.calls.filter(([operation]) => operation === 'writeCompanyCard').length, 0);
            assert.equal(Boolean((await (await f.request('card', { token })).json()).card), Boolean(previous.cardId));
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
    release();
    assert.equal((await first).status, 200);
    assert.equal(f.calls.filter(([operation]) => operation === 'createCard').length, 1);
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
            assert.deepEqual(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', cardId: 'card_synthetic', verification, workbookSynced: true, workbookReferencesSynced: true });
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
    assert.deepEqual(f.records.get(companyCode(admin.company)), { customerId: 'cus_synthetic', cardId: 'card_synthetic', verification, workbookSynced: true, workbookReferencesSynced: true });
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
