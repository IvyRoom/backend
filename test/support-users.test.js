'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createSupportUsersAdapter } = require('../integrations/support-users');

const columns = [
    'EMPRESA', 'NOME COMPLETO', 'PRIMEIRO NOME', 'E-MAIL', 'SENHA',
    'TIPO DE ACESSO', 'STATUS LOGIN', 'PAGARME CUSTOMER ID', 'PAGARME CARD ID', 'AMBIENTE PAGARME',
].map((name, index) => ({ name, index }));
const legacyNames = ['NÚMERO', 'TITULAR', 'MÊS', 'ANO', 'CVV', 'STATUS CARTÃO'];
const legacyColumns = [...columns.slice(0, 7), ...legacyNames.map((name, offset) => ({ name, index: offset + 7 }))];
const previousReferenceColumns = [...legacyColumns, ...columns.slice(7).map((column, offset) => ({ ...column, index: offset + 13 }))];

function fixture({ metadata = columns, range, values, failUpdateAt } = {}) {
    const calls = [];
    const writes = [];
    const graphClient = {
        api(path) {
            calls.push(path);
            return {
                async get() {
                    if (path.endsWith('/columns?$select=index,name')) return { value: metadata };
                    if (path.endsWith('/range?$select=address,rowCount,columnCount')) return range || {
                        address: 'USUÁRIOS!A3:J4', columnCount: metadata.length, rowCount: 2,
                    };
                    return { values: values || [['Empresa teste', 'Pessoa teste', 'Pessoa', 'admin@example.test', 'test-password', 'Administrador', 'Ativo']] };
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
    assert.equal(typeof createSupportUsersAdapter({ graphClient }).readUsers, 'function');
});

test('support users reads only identity and access cells from the reference-only table', async () => {
    const { users, calls } = fixture();
    assert.deepEqual(await users.readUsers(), [{
        company: 'Empresa teste', name: 'Pessoa teste', email: 'admin@example.test',
        password: 'test-password', role: 'Administrador', status: 'Ativo',
    }]);
    assert.equal(calls.length, 3);
    assert.match(calls[0], /\/tables\/\{7C4EBF15-124A-4107-9867-F83E9C664B31\}\/columns\?\$select=index,name$/u);
    assert.match(calls[1], /\/range\?\$select=address,rowCount,columnCount$/u);
    assert.match(calls[2], /\/range\(address='A4:G4'\)\?\$select=values$/u);
    assert.ok(calls.every((path) => !path.includes('/rows') && !path.includes('/usedRange')));
});

test('support users rejects retired raw-card schemas and all partial or changed headers before reading values', async () => {
    for (const metadata of [
        legacyColumns, previousReferenceColumns, columns.slice(0, 7), columns.slice(0, 8), columns.slice(0, 9),
        columns.map((column, index) => index === 4 ? { ...column, name: 'CHANGED' } : column),
        columns.map((column, index) => index === 8 ? { ...column, name: 'CHANGED' } : column),
        columns.map((column, index) => index === 9 ? { ...column, index: 10 } : column),
        [...columns].reverse(), [...columns, { name: 'EXTRA', index: 10 }],
    ]) {
        const { users, calls } = fixture({ metadata });
        await assert.rejects(users.readUsers(), /schema_mismatch/u);
        assert.equal(calls.length, 1);
    }
});

test('support users refuses shifted tables and inconsistent range metadata', async () => {
    for (const range of [
        { address: 'USUÁRIOS!A3:P4', columnCount: 10, rowCount: 2 },
        { address: 'USUÁRIOS!A3:M4', columnCount: 10, rowCount: 2 },
        { address: 'USUÁRIOS!A2:J4', columnCount: 10, rowCount: 3 },
        { address: 'USUÁRIOS!A3:J4', columnCount: 10, rowCount: 8 },
        { address: 'USUÁRIOS!A3:J4', columnCount: 9, rowCount: 2 },
        { address: 'USUÁRIOS!B3:K4', columnCount: 10, rowCount: 2 },
    ]) {
        const { users, calls } = fixture({ range });
        await assert.rejects(users.readUsers(), /schema_mismatch/u);
        assert.equal(calls.length, 2);
    }
});

test('support users handles table growth without fetching provider references', async () => {
    const row = ['Empresa teste', 'Pessoa teste', 'Pessoa', 'user@example.test', 'test-password', 'Usuário', 'Ativo'];
    const { users, calls } = fixture({
        range: { address: 'USUÁRIOS!A3:J5', columnCount: 10, rowCount: 3 }, values: [row, row],
    });
    assert.equal((await users.readUsers()).length, 2);
    assert.match(calls[2], /A4:G5/u);
});

test('support users refuses row width or count drift', async () => {
    for (const values of [[['only one cell']], [], [new Array(10).fill('')]]) {
        const { users } = fixture({ values });
        await assert.rejects(users.readUsers(), /schema_mismatch/u);
    }
});

test('support users returns no rows when the table has only its header', async () => {
    const { users, calls } = fixture({ range: { address: 'USUÁRIOS!A3:J3', columnCount: 10, rowCount: 1 } });
    assert.deepEqual(await users.readUsers(), []);
    assert.equal(calls.length, 2);
});

const cardReferences = { customerId: 'cus_Test1', cardId: 'card_Test1', environment: 'sandbox' };

test('company card sync writes only reusable references for every matching company colleague', async () => {
    const { users, calls, writes } = fixture({
        range: { address: 'USUÁRIOS!A3:J6', columnCount: 10, rowCount: 4 },
        values: [
            [' Empresa  teste ', 'Pessoa teste', 'Pessoa', 'admin@example.test', 'private-password', 'Administrador', 'Ativo'],
            ['Outra empresa', 'Outra pessoa', 'Outra', 'other@example.test', 'private-password', 'Administrador', 'Ativo'],
            ['EMPRESA TESTE', 'Colega teste', 'Colega', 'colleague@example.test', 'private-password', 'Usuário', 'Inativo'],
        ],
    });
    const card = { number: '4000000000000010', cvv: '123', holderName: '=ignored', lastFour: '0010', expMonth: 12, expYear: 2030 };
    await users.writeCompanyCard('empresa teste', card, cardReferences);
    assert.equal(calls.length, 5);
    assert.match(calls[2], /range\(address='A4:G6'\)\?\$select=values$/u);
    assert.deepEqual(writes.map(({ path, body }) => ({ address: path.match(/address='([^']+)'/u)[1], body })), [
        { address: 'H4:J4', body: { values: [['cus_Test1', 'card_Test1', 'sandbox']] } },
        { address: 'H6:J6', body: { values: [['cus_Test1', 'card_Test1', 'sandbox']] } },
    ]);
    const transmitted = JSON.stringify(writes);
    for (const excluded of ['4000000000000010', '123', '=ignored', '0010', '2030', 'private-password']) {
        assert.equal(transmitted.includes(excluded), false, excluded);
    }
    assert.ok(writes.every(({ path }) => !path.includes('/rows')));
});

test('company card sync revalidates schema and company existence before any write', async () => {
    const changed = fixture({ metadata: columns.map((column, index) => index === 7 ? { ...column, name: 'SHIFTED' } : column) });
    await assert.rejects(changed.users.writeCompanyCard('Empresa teste', null, cardReferences), /schema_mismatch/u);
    assert.equal(changed.writes.length, 0);
    for (const metadata of [legacyColumns, previousReferenceColumns]) {
        const retired = fixture({ metadata });
        await assert.rejects(retired.users.writeCompanyCard('Empresa teste', null, cardReferences), /schema_mismatch/u);
        assert.equal(retired.calls.length, 1);
        assert.equal(retired.writes.length, 0);
    }
    const absent = fixture();
    await assert.rejects(absent.users.writeCompanyCard('Missing company', null, cardReferences), /company_not_found/u);
    assert.equal(absent.writes.length, 0);
});

test('company card sync refuses absent, unsafe or non-sandbox references before Graph access', async () => {
    for (const references of [
        undefined, null, [], {}, { ...cardReferences, customerId: 1 }, { ...cardReferences, cardId: 1 },
        { ...cardReferences, customerId: 'cus_bad/id' }, { ...cardReferences, cardId: 'token_Test1' },
        { ...cardReferences, cardId: 'card_' }, { ...cardReferences, customerId: 'cus_' },
        { ...cardReferences, environment: 'production' }, { ...cardReferences, environment: undefined },
        { ...cardReferences, cvv: '123' }, { ...cardReferences, verification: 'pagarme_verify_card_v1' },
    ]) {
        const { users, calls } = fixture();
        await assert.rejects(users.writeCompanyCard('Empresa teste', null, references), /invalid_card_references/u);
        assert.equal(calls.length, 0);
    }
    const { users, calls } = fixture();
    await assert.rejects(users.writeCompanyCard('', null, cardReferences), /invalid_company/u);
    assert.equal(calls.length, 0);
});

test('company card sync needs no display metadata and can repeat reference updates idempotently', async () => {
    const { users, writes } = fixture();
    await users.writeCompanyCard('Empresa teste', undefined, cardReferences);
    await users.writeCompanyCard('Empresa teste', { holderName: 'Not stored' }, cardReferences);
    assert.equal(writes.length, 2);
    assert.deepEqual(writes[0], writes[1]);
});

test('company card sync stops at the first failed company patch without retrying or appending', async () => {
    const row = ['Empresa teste', 'Pessoa teste', 'Pessoa', 'admin@example.test', 'test-password', 'Administrador', 'Ativo'];
    for (const failUpdateAt of [1, 2]) {
        const { users, writes } = fixture({
            range: { address: 'USUÁRIOS!A3:J5', columnCount: 10, rowCount: 3 }, values: [row, row], failUpdateAt,
        });
        await assert.rejects(users.writeCompanyCard('Empresa teste', null, cardReferences), /synthetic patch failure/u);
        assert.equal(writes.length, failUpdateAt);
        assert.ok(writes.every(({ path }) => /range\(address='H[45]:J[45]'\)$/u.test(path)));
    }
});
