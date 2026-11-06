'use strict';

const assert = require('node:assert/strict');
const test = require('node:test');
const { createPagarmeAdapter } = require('../integrations/pagarme');

const keys = { secretKey: 'sk_test_synthetic', publicKey: 'pk_test_synthetic' };
const now = () => Date.UTC(2026, 10, 6, 12);
const customer = { id: 'cus_Company1', code: 'support-company1', live: false };
const card = {
    id: 'card_Synthetic1', brand: 'Visa', last_four_digits: '0010', exp_month: 12,
    exp_year: 2030, created_at: '2026-11-06T12:00:00Z', status: 'active', live: false,
    first_six_digits: '400000', holder_name: 'Test Holder', metadata: { private: 'hidden' },
};
const summary = {
    id: card.id, brand: 'Visa', lastFour: '0010', expMonth: 12, expYear: 2030,
    holderName: card.holder_name, createdAt: card.created_at, status: 'active',
};
const address = {
    line_1: '1, Rua de Teste, Centro', zip_code: '01001000', city: 'São Paulo', state: 'SP', country: 'BR',
};

function fixture(results) {
    const calls = [];
    const adapter = createPagarmeAdapter({
        ...keys,
        now,
        async fetchImpl(url, options) {
            calls.push({ url, options });
            assert.ok(results.length, 'Unexpected request');
            const result = results.shift();
            if (result instanceof Error) throw result;
            return { ok: true, async json() { return result; } };
        },
    });
    return { calls, adapter };
}

test('constructs without requests and refuses absent or production keys', () => {
    for (const config of [
        {}, { ...keys, secretKey: 'sk_production' }, { ...keys, publicKey: 'pk_production' },
        { ...keys, secretKey: 'sk_test_synthetic\n' }, { ...keys, publicKey: 'pk_test_' },
        { ...keys, now: 1234 },
    ]) {
        assert.throws(() => createPagarmeAdapter(config), { code: 'pagarme.invalid_configuration' });
    }
    assert.equal(fixture([]).calls.length, 0);
});

test('looks up exact company code using fixed base and sandbox Basic auth', async () => {
    const { adapter, calls } = fixture([{ data: [{ ...customer, code: 'unrelated' }, customer] }]);
    assert.deepEqual(await adapter.findCompanyCustomer(customer.code), { id: customer.id, code: customer.code });
    assert.equal(calls[0].url,
        'https://api.pagar.me/core/v5/customers?code=support-company1&size=100&page=1');
    assert.equal(calls[0].options.method, 'GET');
    assert.equal(calls[0].options.redirect, 'error');
    assert.equal(calls[0].options.headers.Authorization,
        `Basic ${Buffer.from(`${keys.secretKey}:`).toString('base64')}`);
    assert.ok(calls[0].options.signal instanceof AbortSignal);
    assert.equal(calls[0].options.signal.aborted, false);
});

test('absent customer returns null; duplicate codes across pages fail closed', async () => {
    const absent = fixture([{ data: [{ ...customer, code: 'other' }] }]);
    assert.equal(await absent.adapter.findCompanyCustomer(customer.code), null);
    const { adapter, calls } = fixture([
        { data: [customer], paging: { next: 'https://untrusted.invalid/do-not-follow' } },
        { data: [{ ...customer, id: 'cus_Duplicate' }] },
    ]);
    await assert.rejects(adapter.findCompanyCustomer(customer.code), { code: 'pagarme.ambiguous_customer' });
    assert.equal(calls.length, 2);
    assert.equal(calls[1].url,
        'https://api.pagar.me/core/v5/customers?code=support-company1&size=100&page=2');
});

test('creates company without email and projects the result', async () => {
    const { adapter, calls } = fixture([{ ...customer, email: 'private@example.test' }]);
    assert.deepEqual(await adapter.createCompanyCustomer({ code: customer.code, name: 'EMPRESA TESTE' }),
        { id: customer.id, code: customer.code });
    assert.equal(calls[0].options.method, 'POST');
    assert.deepEqual(JSON.parse(calls[0].options.body), { code: customer.code, name: 'EMPRESA TESTE' });
});

test('company creation rejects mismatched provider code', async () => {
    const { adapter } = fixture([{ ...customer, code: 'other-company' }]);
    await assert.rejects(adapter.createCompanyCustomer({ code: customer.code, name: 'EMPRESA TESTE' }),
        { code: 'pagarme.invalid_response' });
});

test('lists cards with pagination and excludes private provider fields', async () => {
    const { adapter, calls } = fixture([
        { data: [card], paging: { next: '/ignored' } },
        { data: [{ ...card, id: 'card_Synthetic2', status: 'expired' }] },
    ]);
    assert.deepEqual(await adapter.listCards(customer.id), [summary, { ...summary, id: 'card_Synthetic2', status: 'expired' }]);
    assert.equal(calls[1].url, `https://api.pagar.me/core/v5/customers/${customer.id}/cards?size=100&page=2`);
});

test('gets selected card and verifies its identity', async () => {
    const good = fixture([card]);
    assert.deepEqual(await good.adapter.getCard(customer.id, card.id), summary);
    assert.equal(good.calls[0].url, `https://api.pagar.me/core/v5/customers/${customer.id}/cards/${card.id}`);
    const bad = fixture([{ ...card, id: 'card_Other' }]);
    await assert.rejects(bad.adapter.getCard(customer.id, card.id), { code: 'pagarme.invalid_response' });
});

test('card metadata exposes only the holder needed for workbook sync and excludes PAN or CVV', async () => {
    const { adapter } = fixture([{ ...card, number: '4000000000000010', cvv: '123', billing_address: address }]);
    assert.deepEqual(await adapter.getCard(customer.id, card.id), summary);
    const missing = fixture([{ ...card, holder_name: undefined }]);
    assert.deepEqual(await missing.adapter.getCard(customer.id, card.id), { ...summary, holderName: undefined });
});

test('requests authorization before accepting an active tokenized card and projects only safe metadata', async () => {
    const { adapter, calls } = fixture([card]);
    assert.deepEqual(await adapter.createCard(customer.id, {
        token: 'token_Synthetic', billingAddress: { ...address, number: 'DO NOT FORWARD' },
    }), { ...summary, verificationRequested: true });
    assert.equal(calls[0].options.method, 'POST');
    assert.deepEqual(JSON.parse(calls[0].options.body), {
        token: 'token_Synthetic', billing_address: address, options: { verify_card: true },
    });
    assert.equal(calls.length, 1);
});

test('GET and list never infer verification from an active card or provider-supplied evidence fields', async () => {
    const providerCard = { ...card, verificationRequested: true, verified: true, options: { verify_card: true } };
    const { adapter } = fixture([providerCard, { data: [providerCard] }]);
    assert.deepEqual(await adapter.getCard(customer.id, card.id), summary);
    assert.deepEqual(await adapter.listCards(customer.id), [summary]);
});

test('successful creation without an active unexpired card remains ambiguous and is never accepted', async () => {
    for (const result of [
        { ...card, status: undefined }, { ...card, status: 'expired' }, { ...card, status: 'deleted' },
        { ...card, exp_year: 2025 }, { ...card, exp_year: 25 },
        { ...card, exp_year: 2026, exp_month: 10 }, { ...card, exp_year: 26, exp_month: 10 },
    ]) {
        const { adapter, calls } = fixture([result]);
        await assert.rejects(adapter.createCard(customer.id, { token: 'token_Synthetic', billingAddress: address }), error => {
            assert.equal(error.code, 'pagarme.invalid_response');
            assert.equal(error.message, 'Pagar.me sandbox operation failed');
            assert.equal(error.definiteRejection, undefined);
            return true;
        });
        assert.equal(calls.length, 1);
    }
});

test('expiration includes the entire current month and accepts documented two-digit years', async () => {
    for (const expYear of [2026, 26]) {
        const { adapter } = fixture([{ ...card, exp_year: expYear, exp_month: 11 }]);
        const accepted = await adapter.createCard(customer.id, { token: 'token_Synthetic', billingAddress: address });
        assert.equal(accepted.verificationRequested, true);
        assert.equal(accepted.expYear, expYear);
        assert.equal(accepted.expMonth, 11);
    }
});

test('rejects path injection and raw card payloads before any request', async () => {
    const { adapter, calls } = fixture([]);
    await assert.rejects(adapter.listCards('cus_Company1/../other'));
    await assert.rejects(adapter.findCompanyCustomer('company?other=true'));
    await assert.rejects(adapter.createCard(customer.id, { number: '4000000000000010', cvv: '123' }));
    await assert.rejects(adapter.createCompanyCustomer({ code: customer.code, name: '' }));
    assert.equal(calls.length, 0);
});

test('refuses live and malformed provider objects without exposing provider content', async () => {
    for (const result of [
        { ...card, live: true }, { ...card, live: 'false' }, { ...card, last_four_digits: 'private content' },
        { ...card, id: 'card_../../private' }, { ...card, status: 'unknown' },
        { ...card, holder_name: 'a'.repeat(65) }, { ...card, holder_name: 'Test\nHolder' }, { ...card, holder_name: 1234 },
    ]) {
        const { adapter } = fixture([result]);
        await assert.rejects(adapter.getCard(customer.id, card.id), error => {
            assert.equal(error.message, 'Pagar.me sandbox operation failed');
            return true;
        });
    }
    const malformed = fixture([{ data: null }]);
    await assert.rejects(malformed.adapter.listCards(customer.id));
});

test('does not retry an ambiguous write or expose the transport exception', async () => {
    const { adapter, calls } = fixture([new Error('private provider body and credentials')]);
    await assert.rejects(adapter.createCard(customer.id, { token: 'token_Synthetic', billingAddress: address }),
        { code: 'pagarme.request_failed', message: 'Pagar.me sandbox operation failed' });
    assert.equal(calls.length, 1);
});

test('HTTP failures are sanitized and made only once', async () => {
    let attempts = 0;
    const adapter = createPagarmeAdapter({
        ...keys,
        async fetchImpl() {
            attempts += 1;
            return { ok: false, async json() { throw new Error('Must not read failure body'); } };
        },
    });
    await assert.rejects(adapter.createCompanyCustomer({ code: customer.code, name: 'EMPRESA TESTE' }),
        { code: 'pagarme.request_failed' });
    assert.equal(attempts, 1);
});

test('non-card business rejections are not misreported as card verification failures', async () => {
    const adapter = createPagarmeAdapter({
        ...keys,
        async fetchImpl() {
            return { ok: false, status: 412, async json() { throw new Error('Must not read failure body'); } };
        },
    });
    await assert.rejects(adapter.createCompanyCustomer({ code: customer.code, name: 'EMPRESA TESTE' }), {
        code: 'pagarme.request_rejected', definiteRejection: true,
    });
});

test('only documented definite HTTP rejections permit clearing pending intent', async () => {
    for (const status of [400, 401, 403, 404, 412, 422, 429, 500, 502, 503]) {
        let attempts = 0;
        const adapter = createPagarmeAdapter({
            ...keys,
            async fetchImpl() {
                attempts += 1;
                return { ok: false, status, async json() { throw new Error('Do not read provider error bodies'); } };
            },
        });
        await assert.rejects(adapter.createCard(customer.id, { token: 'token_Synthetic', billingAddress: address }), error => {
            const definite = [400, 401, 403, 404, 412, 422].includes(status);
            assert.equal(error.definiteRejection === true, definite);
            assert.equal(error.code, status === 412 ? 'pagarme.card_verification_failed'
                : definite ? 'pagarme.request_rejected' : 'pagarme.request_failed');
            assert.equal(error.message, 'Pagar.me sandbox operation failed');
            assert.equal(error.cause, undefined);
            return true;
        });
        assert.equal(attempts, 1);
    }
});
