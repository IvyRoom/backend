'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createSupportUsersAdapter } = require('../integrations/support-users');

const columns = [
    'EMPRESA', 'NOME COMPLETO', 'PRIMEIRO NOME', 'E-MAIL', 'SENHA',
    'TIPO DE ACESSO', 'STATUS LOGIN', 'NÚMERO', 'TITULAR', 'MÊS', 'ANO', 'CVV', 'STATUS CARTÃO',
].map((name, index) => ({ name, index }));

function fixture({ metadata = columns, range, values, failUpdateAt } = {}) {
    const calls = [];
    const writes = [];
    const graphClient = {
        api(path) {
            calls.push(path);
            return {
                async get() {
                    if (path.endsWith('/columns?$select=index,name')) return { value: metadata };
                    if (path.endsWith('/range?$select=address,rowCount,columnCount')) return range || { address: 'USUÁRIOS!A3:M4', columnCount: 13, rowCount: 2 };
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

test('support users reads only identity and access cells from the verified table', async () => {
    const { users, calls } = fixture();
    assert.deepEqual(await users.readUsers(), [{
        company: 'Empresa teste', name: 'Pessoa teste', email: 'admin@example.test',
        password: 'test-password', role: 'Administrador', status: 'Ativo',
    }]);
    assert.equal(calls.length, 3);
    assert.match(calls[0], /\/tables\/\{7C4EBF15-124A-4107-9867-F83E9C664B31\}\/columns\?\$select=index,name$/);
    assert.match(calls[1], /\/range\?\$select=address,rowCount,columnCount$/);
    assert.match(calls[2], /\/range\(address='A4:G4'\)\?\$select=values$/);
    assert.ok(calls.every((path) => !path.includes('/rows') && !path.includes('/usedRange')));
});

test('support users refuses changed table headers before reading values', async () => {
    for (const metadata of [columns.slice(0, 7), columns.map((column, index) => index === 4 ? { ...column, name: 'CHANGED' } : column), [...columns].reverse()]) {
        const { users, calls } = fixture({ metadata });
        await assert.rejects(users.readUsers(), /schema_mismatch/);
        assert.equal(calls.length, 1);
    }
});

test('support users refuses shifted tables and inconsistent metadata', async () => {
    for (const range of [
        { address: 'USUÁRIOS!A3:N4', columnCount: 14, rowCount: 2 },
        { address: 'USUÁRIOS!A2:M4', columnCount: 13, rowCount: 3 },
        { address: 'USUÁRIOS!A3:M4', columnCount: 13, rowCount: 8 },
        { address: 'USUÁRIOS!A3:M4', columnCount: 12, rowCount: 2 },
    ]) {
        const { users, calls } = fixture({ range });
        await assert.rejects(users.readUsers(), /schema_mismatch/);
        assert.equal(calls.length, 2);
    }
});

test('support users handles table growth without fetching legacy payment cells', async () => {
    const row = ['Empresa teste', 'Pessoa teste', 'Pessoa', 'user@example.test', 'test-password', 'Usuário', 'Ativo'];
    const { users, calls } = fixture({
        range: { address: 'USUÁRIOS!A3:M5', columnCount: 13, rowCount: 3 },
        values: [row, row],
    });
    assert.equal((await users.readUsers()).length, 2);
    assert.match(calls[2], /A4:G5/);
});

test('support users refuses row width or count drift', async () => {
    for (const values of [[['only one cell']], [], [new Array(13).fill('')]]) {
        const { users } = fixture({ values });
        await assert.rejects(users.readUsers(), /schema_mismatch/);
    }
});

test('support users returns no rows when the table has only its header', async () => {
    const { users, calls } = fixture({ range: { address: 'USUÁRIOS!A3:M3', columnCount: 13, rowCount: 1 } });
    assert.deepEqual(await users.readUsers(), []);
    assert.equal(calls.length, 2);
});

const cardMetadata = { lastFour: '0010', holderName: 'Test Holder', expMonth: 12, expYear: 2030 };

test('company card sync writes masked metadata and status for every company colleague only', async () => {
    const { users, calls, writes } = fixture({
        range: { address: 'USUÁRIOS!A3:M6', columnCount: 13, rowCount: 4 },
        values: [
            [' Empresa  teste ', 'Pessoa teste', 'Pessoa', 'admin@example.test', 'private-password', 'Administrador', 'Ativo'],
            ['Outra empresa', 'Outra pessoa', 'Outra', 'other@example.test', 'private-password', 'Administrador', 'Ativo'],
            ['EMPRESA TESTE', 'Colega teste', 'Colega', 'colleague@example.test', 'private-password', 'Usuário', 'Inativo'],
        ],
    });
    await users.writeCompanyCard('empresa teste', { ...cardMetadata, number: '4000000000000010', cvv: '123' });
    assert.equal(calls.length, 7);
    assert.match(calls[2], /range\(address='A4:G6'\)\?\$select=values$/u);
    assert.deepEqual(writes.map(({ path, body }) => ({ address: path.match(/address='([^']+)'/u)[1], body })), [
        { address: 'H4:K4', body: { values: [['**** **** **** 0010', 'Test Holder', 12, 2030]] } },
        { address: 'M4', body: { values: [['Cadastrado']] } },
        { address: 'H6:K6', body: { values: [['**** **** **** 0010', 'Test Holder', 12, 2030]] } },
        { address: 'M6', body: { values: [['Cadastrado']] } },
    ]);
    const transmitted = JSON.stringify(writes);
    assert.equal(transmitted.includes('4000000000000010'), false);
    assert.equal(transmitted.includes('123'), false);
    assert.equal(transmitted.includes('private-password'), false);
    assert.ok(writes.every(({ path }) => !path.includes('/rows')));
});

test('company card sync revalidates schema and company existence before any write', async () => {
    const shifted = fixture({ metadata: columns.map((column, index) => index === 7 ? { ...column, name: 'SHIFTED' } : column) });
    await assert.rejects(shifted.users.writeCompanyCard('Empresa teste', cardMetadata), /schema_mismatch/u);
    assert.equal(shifted.writes.length, 0);
    const absent = fixture();
    await assert.rejects(absent.users.writeCompanyCard('Missing company', cardMetadata), /company_not_found/u);
    assert.equal(absent.writes.length, 0);
});

test('company card sync refuses malformed metadata and formula-leading holder names before Graph access', async () => {
    for (const changes of [
        { lastFour: '4000000000000010' }, { lastFour: '123' }, { holderName: '' },
        { holderName: '=formula' }, { holderName: '@formula' }, { holderName: '+formula' }, { holderName: '-formula' },
        { holderName: 'Test\nHolder' }, { holderName: 'a'.repeat(65) }, { expMonth: 0 }, { expMonth: 13 },
        { expMonth: 1.5 }, { expYear: -1 }, { expYear: 101 }, { expYear: '2030' }, { expYear: 10000 },
    ]) {
        const { users, calls } = fixture();
        await assert.rejects(users.writeCompanyCard('Empresa teste', { ...cardMetadata, ...changes }), /invalid_card_metadata/u);
        assert.equal(calls.length, 0);
    }
});

test('company card sync normalizes short provider years and can repeat the same updates idempotently', async () => {
    const { users, writes } = fixture();
    await users.writeCompanyCard('Empresa teste', { ...cardMetadata, expYear: 30 });
    await users.writeCompanyCard('Empresa teste', cardMetadata);
    assert.deepEqual(writes.slice(0, 2), writes.slice(2));
    assert.equal(writes[0].body.values[0][3], 2030);
});

test('company card sync stops at the first failed patch without retrying or appending', async () => {
    const { users, writes } = fixture({ failUpdateAt: 2 });
    await assert.rejects(users.writeCompanyCard('Empresa teste', cardMetadata), /synthetic patch failure/u);
    assert.equal(writes.length, 2);
    assert.ok(writes.every(({ path }) => !path.includes('/rows')));
});
