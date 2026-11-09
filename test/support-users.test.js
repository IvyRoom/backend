'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createSupportUsersAdapter } = require('../integrations/support-users');

const columns = [
    'CNPJ', 'RAZÃO SOCIAL', 'VALOR REFERÊNCIA', 'NOME COMPLETO', 'E-MAIL', 'SENHA',
    'TIPO DE ACESSO', 'STATUS', 'PAGARME CUSTOMER ID', 'PAGARME CARD ID', 'PAGARME AMBIENTE ',
].map((name, index) => ({ name, index }));
const companyId = '12345678000195';
const defaultRow = ['12.345.678/0001-95', 'Empresa teste', 1000, 'Pessoa teste', 'admin@example.test', 'private-password', 'Administrador', 'Ativo'];
const cardReferences = { customerId: 'cus_Test1', cardId: 'card_Test1', environment: 'sandbox' };

function fixture({ metadata = columns, range, values = [defaultRow], companyValues, referenceRows = {}, failUpdateAt } = {}) {
    const calls = [];
    const writes = [];
    const graphClient = {
        api(path) {
            calls.push(path);
            return {
                async get() {
                    if (path.endsWith('/columns?$select=index,name')) return { value: metadata };
                    if (path.endsWith('/range?$select=address,rowCount,columnCount')) return range || {
                        address: `USUÁRIOS!A3:K${values.length + 3}`, columnCount: 11, rowCount: values.length + 1,
                    };
                    if (/address='A4:A\d+'/u.test(path)) return { values: companyValues ?? values.map((row) => [row[0]]) };
                    const reference = /address='I(\d+):K\d+'/u.exec(path);
                    if (reference) return { values: referenceRows[reference[1]] ?? [['', '', '']] };
                    return { values };
                },
                async update(body) {
                    writes.push({ path, body: structuredClone(body) });
                    if (writes.length === failUpdateAt) throw new Error('synthetic patch failure');
                },
            };
        },
    };
    return { users: createSupportUsersAdapter({ graphClient }), calls, writes };
}

test('support users adapter construction does not access Graph', () => {
    const graphClient = new Proxy({}, { get() { throw new Error('Graph accessed'); } });
    const adapter = createSupportUsersAdapter({ graphClient });
    assert.equal(typeof adapter.readUsers, 'function');
    assert.equal(typeof adapter.readCompanyCardReferences, 'function');
});

test('support users reads identity, stable company and reference price without provider references', async () => {
    const { users, calls } = fixture();
    assert.deepEqual(await users.readUsers(), [{
        companyId, company: 'Empresa teste', referenceRate: 1000, name: 'Pessoa teste',
        email: 'admin@example.test', password: 'private-password', role: 'Administrador', status: 'Ativo',
    }]);
    assert.equal(calls.length, 3);
    assert.match(calls[0], /\/tables\/\{7C4EBF15-124A-4107-9867-F83E9C664B31\}\/columns\?\$select=index,name$/u);
    assert.match(calls[1], /\/range\?\$select=address,rowCount,columnCount$/u);
    assert.match(calls[2], /\/range\(address='A4:H4'\)\?\$select=values$/u);
    assert.ok(calls.every((path) => !path.includes('/rows') && !path.includes('/usedRange')));
});

test('support users rejects retired, partial, reordered and changed schemas before reading values', async () => {
    const retired = ['EMPRESA', 'NOME COMPLETO', 'PRIMEIRO NOME', 'E-MAIL', 'SENHA', 'TIPO DE ACESSO', 'STATUS LOGIN', 'PAGARME CUSTOMER ID', 'PAGARME CARD ID', 'AMBIENTE PAGARME']
        .map((name, index) => ({ name, index }));
    for (const metadata of [
        retired, columns.slice(0, 8), columns.slice(0, 10),
        columns.map((column, index) => index === 2 ? { ...column, name: 'CHANGED' } : column),
        columns.map((column, index) => index === 9 ? { ...column, index: 10 } : column),
        columns.map((column, index) => index === 9 ? { ...column, name: null } : column),
        [...columns].reverse(), [...columns, { name: 'EXTRA', index: 11 }],
    ]) {
        const { users, calls } = fixture({ metadata });
        await assert.rejects(users.readUsers(), /schema_mismatch/u);
        assert.equal(calls.length, 1);
    }
});

test('support users refuses shifted tables and inconsistent range metadata', async () => {
    for (const range of [
        { address: 'USUÁRIOS!A3:J4', columnCount: 11, rowCount: 2 },
        { address: 'USUÁRIOS!A2:K4', columnCount: 11, rowCount: 3 },
        { address: 'USUÁRIOS!A3:K4', columnCount: 11, rowCount: 8 },
        { address: 'USUÁRIOS!A3:K4', columnCount: 10, rowCount: 2 },
        { address: 'USUÁRIOS!B3:L4', columnCount: 11, rowCount: 2 },
    ]) {
        const { users, calls } = fixture({ range });
        await assert.rejects(users.readUsers(), /schema_mismatch/u);
        assert.equal(calls.length, 2);
    }
});

test('support users handles growth and empty tables without fetching provider references', async () => {
    const grown = fixture({ values: [defaultRow, defaultRow] });
    assert.equal((await grown.users.readUsers()).length, 2);
    assert.match(grown.calls[2], /A4:H5/u);
    const empty = fixture({ values: [] });
    assert.deepEqual(await empty.users.readUsers(), []);
    assert.equal(empty.calls.length, 2);
});

test('support users refuses row shape drift and invalid company identity or currency', async () => {
    for (const values of [[['one cell']], [new Array(11).fill('')]]) {
        await assert.rejects(fixture({ values }).users.readUsers(), /schema_mismatch/u);
    }
    for (const [index, value] of [[0, 'Company name'], [0, '123'], [0, 'letters12345678000195'], [2, 'R$ 1.000,00'], [2, 0], [2, Infinity]]) {
        const row = [...defaultRow];
        row[index] = value;
        await assert.rejects(fixture({ values: [row] }).users.readUsers(), /invalid_company_data/u);
    }
});

test('company card sync writes I:K only for matching CNPJ, regardless of company labels', async () => {
    const other = [...defaultRow];
    other[0] = '98.765.432/0001-10';
    const colleague = [...defaultRow];
    colleague[0] = companyId;
    colleague[1] = 'Renamed company';
    const { users, calls, writes } = fixture({ values: [defaultRow, other, colleague] });
    await users.writeCompanyCard(companyId, { number: '4000000000000010', cvv: '123', holderName: '=ignored' }, cardReferences);
    assert.equal(calls.length, 5);
    assert.match(calls[2], /range\(address='A4:A6'\)\?\$select=values$/u);
    assert.deepEqual(writes.map(({ path, body }) => ({ address: path.match(/address='([^']+)'/u)[1], body })), [
        { address: 'I4:K4', body: { values: [['cus_Test1', 'card_Test1', 'sandbox']] } },
        { address: 'I6:K6', body: { values: [['cus_Test1', 'card_Test1', 'sandbox']] } },
    ]);
    for (const excluded of ['4000000000000010', '123', '=ignored', 'private-password']) {
        assert.equal(JSON.stringify(writes).includes(excluded), false, excluded);
    }
});

test('company card sync validates schema and existence before any write', async () => {
    const changed = fixture({ metadata: columns.map((column, index) => index === 8 ? { ...column, name: 'SHIFTED' } : column) });
    await assert.rejects(changed.users.writeCompanyCard(companyId, null, cardReferences), /schema_mismatch/u);
    assert.equal(changed.writes.length, 0);
    const absent = fixture();
    await assert.rejects(absent.users.writeCompanyCard('98765432000110', null, cardReferences), /company_not_found/u);
    assert.equal(absent.writes.length, 0);
});

test('company card sync refuses invalid company and unsafe or non-sandbox references before Graph access', async () => {
    for (const references of [
        undefined, null, [], {}, { ...cardReferences, customerId: 1 }, { ...cardReferences, cardId: 1 },
        { ...cardReferences, customerId: 'cus_bad/id' }, { ...cardReferences, cardId: 'token_Test1' },
        { ...cardReferences, cardId: 'card_' }, { ...cardReferences, customerId: 'cus_' },
        { ...cardReferences, environment: 'production' }, { ...cardReferences, environment: undefined },
        { ...cardReferences, cvv: '123' }, { ...cardReferences, verification: 'pagarme_verify_card_v1' },
    ]) {
        const { users, calls } = fixture();
        await assert.rejects(users.writeCompanyCard(companyId, null, references), /invalid_card_references/u);
        assert.equal(calls.length, 0);
    }
    for (const invalid of ['', 'Empresa teste', '123']) {
        const { users, calls } = fixture();
        await assert.rejects(users.writeCompanyCard(invalid, null, cardReferences), /invalid_company/u);
        assert.equal(calls.length, 0);
    }
});

test('company reference lookup reads only matching references and accepts agreeing populated rows', async () => {
    const other = [...defaultRow];
    other[0] = '98.765.432/0001-10';
    const { users, calls } = fixture({
        values: [defaultRow, other, defaultRow, defaultRow],
        referenceRows: { 4: [['cus_Test1', 'card_Test1', 'sandbox']], 6: [['', '', '']], 7: [[' cus_Test1 ', 'card_Test1', ' sandbox ']] },
    });
    assert.deepEqual(await users.readCompanyCardReferences(companyId), cardReferences);
    assert.match(calls[2], /A4:A7/u);
    assert.deepEqual(calls.slice(3).map((path) => /address='([^']+)'/u.exec(path)[1]), ['I4:K4', 'I6:K6', 'I7:K7']);
    assert.ok(calls.every((path) => !/A4:H/u.test(path)));
    assert.equal(await fixture().users.readCompanyCardReferences(companyId), null);
});

test('company reference lookup rejects partial, conflicting and non-sandbox data', async () => {
    for (const referenceRows of [
        { 4: [['cus_Test1', '', 'sandbox']] },
        { 4: [['cus_Test1', 'card_Test1', 'production']] },
        { 4: [['cus_Test1', 'card_Test1', 'sandbox']], 5: [['cus_Test1', 'card_Test2', 'sandbox']] },
    ]) {
        const { users, writes } = fixture({ values: [defaultRow, defaultRow], referenceRows });
        await assert.rejects(users.readCompanyCardReferences(companyId), /(?:invalid|conflicting)_card_references/u);
        assert.equal(writes.length, 0);
    }
});

test('company card sync stops on a failed patch without retrying or appending', async () => {
    for (const failUpdateAt of [1, 2]) {
        const { users, writes } = fixture({ values: [defaultRow, defaultRow], failUpdateAt });
        await assert.rejects(users.writeCompanyCard(companyId, null, cardReferences), /synthetic patch failure/u);
        assert.equal(writes.length, failUpdateAt);
        assert.ok(writes.every(({ path }) => /range\(address='I[45]:K[45]'\)$/u.test(path)));
    }
});
