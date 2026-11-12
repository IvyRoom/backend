'use strict';

require('./deny-production-network');
const test = require('node:test');
const assert = require('node:assert/strict');
const { createSupportManagementHandlers } = require('../domains/support-management');

const user = { companyId: '12345678000195', company: 'Empresa Exemplo', name: 'Administradora Exemplo',
    email: 'admin@example.test', password: 'unchanged-workbook-password', role: 'Administrador', bookingAllowed: true };
const card = { id: 'card_Live1', brand: 'Visa', lastFour: '0010', expMonth: 12, expYear: 2030, status: 'active' };

function fixture(environment = 'production', references) {
    const calls = [];
    let record = {};
    const handlers = createSupportManagementHandlers({
        environment, publicKey: `pk_${environment === 'production' ? 'live' : 'test'}_synthetic`, now: () => Date.UTC(2026, 10, 12),
        users: {
            async readUsers() { return [user]; },
            async readCompanyCardReferences() { calls.push('references'); return references; },
            async writeCompanyCard(company, details, value) { calls.push(['workbook', company, value]); },
        },
        cardStore: {
            async read() { return structuredClone(record); },
            async write(code, value) { record = structuredClone(value); calls.push(['store', record]); },
            async findByReferences() { throw new Error('Foreign references must not be imported'); },
        },
        pagarme: {
            async findCompanyCustomer() { calls.push('findCustomer'); return null; },
            async createCompanyCustomer() { calls.push('createCustomer'); return { id: 'cus_Live1' }; },
            async createCard(customerId, payload) { calls.push(['createCard', customerId, payload]); return { ...card, verificationRequested: true }; },
            async getCard() { calls.push('getCard'); return card; },
        },
        async bootstrap() { throw new Error('Unavailable bootstrap must retain the configured mode'); },
    });
    function response() { return { locals: { supportUser: user, supportUsers: [user] }, statusCode: 200,
        status(value) { this.statusCode = value; return this; }, json(value) { this.body = value; return this; } }; }
    return { calls, handlers, response, record: () => record };
}

test('production public configuration and login keep their mode without altering workbook-password authentication', async () => {
    const f = fixture();
    const config = f.response();
    f.handlers.config({}, config);
    assert.deepEqual(config.body, { environment: 'production', publicKey: 'pk_live_synthetic' });
    const login = f.response();
    await f.handlers.login({ ip: '127.0.0.1', body: { email: user.email, password: user.password } }, login);
    assert.equal(login.statusCode, 200);
    assert.deepEqual(login.body.config, config.body);
    assert.equal(typeof login.body.token, 'string');
    assert.equal(login.body.user.password, undefined);
    assert.equal(f.calls.length, 0, 'Login fallback must not write passwords, references or provider state');
    assert.equal(user.password, 'unchanged-workbook-password');
});

test('neither mode can import the other environment wallet or its verification evidence', async () => {
    for (const environment of ['sandbox', 'production']) {
        const f = fixture(environment, { customerId: 'cus_Foreign1', cardId: 'card_Foreign1',
            environment: environment === 'production' ? 'sandbox' : 'production' });
        const result = await f.handlers.companyPayment(user);
        assert.deepEqual(result, { card: null });
        assert.deepEqual(f.record(), {});
        assert.deepEqual(f.calls, ['references']);
    }
});

test('production registration creates a verified live wallet and writes only production references', async () => {
    const f = fixture('production', { customerId: 'cus_Test1', cardId: 'card_Test1', environment: 'sandbox' });
    const response = f.response();
    await f.handlers.saveCard({ body: { token: 'token_Live1' } }, response);
    assert.equal(response.statusCode, 200);
    assert.equal(response.body.card.lastFour, '0010');
    assert.deepEqual(f.calls.find((call) => Array.isArray(call) && call[0] === 'createCard'),
        ['createCard', 'cus_Live1', { token: 'token_Live1', billingAddress: undefined }]);
    assert.deepEqual(f.calls.find((call) => Array.isArray(call) && call[0] === 'workbook'),
        ['workbook', user.companyId, { customerId: 'cus_Live1', cardId: 'card_Live1', environment: 'production' }]);
    assert.equal(f.record().verification, 'pagarme_verify_card_v1');
    assert.equal(f.record().pending, undefined);
});
