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

function fixture(results, options = {}) {
    const calls = [];
    const adapter = createPagarmeAdapter({
        ...keys,
        now,
        ...options,
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
        { ...keys, secretKey: 'sk_live_synthetic' }, { ...keys, publicKey: 'pk_live_synthetic' },
        { ...keys, secretKey: 'sk_test_synthetic\n' }, { ...keys, publicKey: 'pk_test_' },
        { ...keys, now: 1234 },
    ]) {
        assert.throws(() => createPagarmeAdapter(config), { code: 'pagarme.invalid_configuration' });
    }
    assert.equal(fixture([]).calls.length, 0);
});

test('production accepts documented sk_/pk_ keys and refuses test credentials without changing the default sandbox', () => {
    const live = { environment: 'production', secretKey: 'sk_productionSynthetic', publicKey: 'pk_productionSynthetic' };
    assert.equal(fixture([], live).calls.length, 0);
    assert.equal(fixture([], { ...live, secretKey: 'sk_live_synthetic', publicKey: 'pk_live_synthetic' }).calls.length, 0);
    for (const options of [
        { ...live, environment: 'unknown' }, { ...live, environment: 'sandbox' },
        { ...live, secretKey: keys.secretKey }, { ...live, publicKey: keys.publicKey },
        { ...keys, environment: 'production' }, { ...live, environment: undefined },
        { ...live, secretKey: 'sk_' }, { ...live, publicKey: 'pk_' },
    ]) assert.throws(() => createPagarmeAdapter(options), { code: 'pagarme.invalid_configuration' });
});

test('production wallet registration verifies the tokenized card and keeps secret authentication server-side', async () => {
    const live = { environment: 'production', secretKey: 'sk_productionSynthetic', publicKey: 'pk_productionSynthetic' };
    const { adapter, calls } = fixture([{ ...card, live: true }], live);
    assert.deepEqual(await adapter.createCard(customer.id, { token: 'token_Synthetic' }), { ...summary, verificationRequested: true });
    assert.equal(calls.length, 1);
    assert.equal(calls[0].url, `https://api.pagar.me/core/v5/customers/${customer.id}/cards`);
    assert.equal(calls[0].options.headers.Authorization, `Basic ${Buffer.from(`${live.secretKey}:`).toString('base64')}`);
    assert.deepEqual(JSON.parse(calls[0].options.body), { token: 'token_Synthetic', options: { verify_card: true } });
});

test('opposite or malformed environment flags fail closed at both list envelope and item boundaries', async () => {
    for (const environment of ['sandbox', 'production']) {
        const live = environment === 'production';
        const options = { environment, secretKey: `sk_${live ? 'live' : 'test'}_synthetic`, publicKey: `pk_${live ? 'live' : 'test'}_synthetic` };
        for (const flag of ['live', 'livemode']) {
            for (const value of [!live, String(live), null]) {
                const badEnvelope = fixture([{ data: [], [flag]: value }], options);
                await assert.rejects(badEnvelope.adapter.listCards(customer.id), { code: 'pagarme.request_failed' });
                const badItem = fixture([{ data: [{ ...card, live, [flag]: value }] }], options);
                await assert.rejects(badItem.adapter.listCards(customer.id), { code: 'pagarme.invalid_response' });
                const badCard = fixture([{ ...card, live, [flag]: value }], options);
                await assert.rejects(badCard.adapter.getCard(customer.id, card.id), { code: 'pagarme.request_failed' });
                assert.equal(badEnvelope.calls.length, 1);
                assert.equal(badItem.calls.length, 1);
                assert.equal(badCard.calls.length, 1);
            }
        }
        const omitted = { ...card };
        delete omitted.live;
        const noFlag = fixture([omitted], options);
        assert.deepEqual(await noFlag.adapter.getCard(customer.id, card.id), summary);
    }
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

test('deletes one selected sandbox wallet card and projects only confirmed deletion', async () => {
    const { adapter, calls } = fixture([{ ...card, status: 'deleted', number: '4000000000000010', cvv: '123' }]);
    assert.deepEqual(await adapter.deleteCard(customer.id, card.id), { deleted: true });
    assert.equal(calls.length, 1);
    assert.equal(calls[0].url, `https://api.pagar.me/core/v5/customers/${customer.id}/cards/${card.id}`);
    assert.equal(calls[0].options.method, 'DELETE');
    assert.equal(calls[0].options.redirect, 'error');
    assert.equal(calls[0].options.body, undefined);
    assert.equal(calls[0].options.headers['Content-Type'], undefined);
    assert.ok(calls[0].options.signal instanceof AbortSignal);
    const minimal = fixture([{ id: card.id, status: 'deleted', live: false }]);
    assert.deepEqual(await minimal.adapter.deleteCard(customer.id, card.id), { deleted: true });
});

test('documented card absence confirms deletion while GET exposes a fixed not-found code', async () => {
    for (const method of ['deleteCard', 'getCard']) {
        let attempts = 0;
        const adapter = createPagarmeAdapter({ ...keys, async fetchImpl() {
            attempts += 1;
            return { ok: false, status: 404, async json() { throw new Error('Must not read provider failure body'); } };
        } });
        if (method === 'deleteCard') assert.deepEqual(await adapter.deleteCard(customer.id, card.id), { deleted: true });
        else await assert.rejects(adapter.getCard(customer.id, card.id), {
            code: 'pagarme.card_not_found', message: 'Pagar.me operation failed', definiteRejection: true,
        });
        assert.equal(attempts, 1);
    }
});

test('deletion rejects unsafe identifiers before a request and never retries transport failures', async () => {
    const invalid = fixture([]);
    for (const [customerId, cardId] of [[null, card.id], ['cus_Company1/../other', card.id], [customer.id, 'card_../other'], [customer.id, null]]) {
        await assert.rejects(invalid.adapter.deleteCard(customerId, cardId), { code: 'pagarme.invalid_input' });
    }
    assert.equal(invalid.calls.length, 0);
    const failed = fixture([new Error('PRIVATE provider request body and credentials')]);
    await assert.rejects(failed.adapter.deleteCard(customer.id, card.id), (error) => {
        assert.equal(error.code, 'pagarme.request_failed');
        assert.equal(error.message, 'Pagar.me operation failed');
        assert.equal(error.definiteRejection, undefined);
        assert.equal(error.cause, undefined);
        return true;
    });
    assert.equal(failed.calls.length, 1);
});

test('only confirmed matching sandbox deletion bodies are accepted and malformed successes remain ambiguous', async () => {
    for (const value of [
        { id: 'card_Other', status: 'deleted' }, { id: card.id, status: 'active' },
        { id: card.id }, {}, null, { id: card.id, status: 'deleted', live: true },
        { id: card.id, status: 'deleted', live: 'false' },
    ]) {
        const { adapter, calls } = fixture([value]);
        await assert.rejects(adapter.deleteCard(customer.id, card.id), (error) => {
            assert.equal(error.message, 'Pagar.me operation failed');
            assert.equal(error.definiteRejection, undefined);
            assert.equal(error.cause, undefined);
            return true;
        });
        assert.equal(calls.length, 1);
    }
    let attempts = 0;
    const malformed = createPagarmeAdapter({ ...keys, async fetchImpl() {
        attempts += 1;
        return { ok: true, async json() { throw new Error('PRIVATE malformed response'); } };
    } });
    await assert.rejects(malformed.deleteCard(customer.id, card.id), { code: 'pagarme.request_failed' });
    assert.equal(attempts, 1);
});

test('deletion distinguishes definite rejection from ambiguous HTTP failures without treating either as absence', async () => {
    for (const status of [400, 401, 403, 412, 422, 429, 500, 502, 503]) {
        let attempts = 0;
        const adapter = createPagarmeAdapter({ ...keys, async fetchImpl() {
            attempts += 1;
            return { ok: false, status, async json() { throw new Error('Must not read provider failure body'); } };
        } });
        await assert.rejects(adapter.deleteCard(customer.id, card.id), (error) => {
            const definite = [400, 401, 403, 412, 422].includes(status);
            assert.equal(error.code, definite ? 'pagarme.request_rejected' : 'pagarme.request_failed');
            assert.equal(error.definiteRejection === true, definite);
            assert.equal(error.message, 'Pagar.me operation failed');
            assert.equal(error.cause, undefined);
            return true;
        });
        assert.equal(attempts, 1);
    }
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
            assert.equal(error.message, 'Pagar.me operation failed');
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
            assert.equal(error.message, 'Pagar.me operation failed');
            return true;
        });
    }
    const malformed = fixture([{ data: null }]);
    await assert.rejects(malformed.adapter.listCards(customer.id));
});

test('does not retry an ambiguous write or expose the transport exception', async () => {
    const { adapter, calls } = fixture([new Error('private provider body and credentials')]);
    await assert.rejects(adapter.createCard(customer.id, { token: 'token_Synthetic', billingAddress: address }),
        { code: 'pagarme.request_failed', message: 'Pagar.me operation failed' });
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
            assert.equal(error.message, 'Pagar.me operation failed');
            assert.equal(error.cause, undefined);
            return true;
        });
        assert.equal(attempts, 1);
    }
});


test('wallet card verification can omit the optional billing address without weakening verification', async () => {
    const { adapter, calls } = fixture([{ ...card, status: 'active' }]);
    const result = await adapter.createCard(customer.id, { token: 'token_Synthetic' });
    assert.equal(result.verificationRequested, true);
    assert.deepEqual(JSON.parse(calls[0].options.body), { token: 'token_Synthetic', options: { verify_card: true } });
});
