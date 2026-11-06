'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createSupportUsersAdapter } = require('../integrations/support-users');

const columns = [
    'EMPRESA', 'NOME COMPLETO', 'PRIMEIRO NOME', 'E-MAIL', 'SENHA',
    'TIPO DE ACESSO', 'STATUS LOGIN', 'NÚMERO', 'TITULAR', 'MÊS', 'ANO', 'CVV', 'STATUS CARTÃO',
].map((name, index) => ({ name, index }));

function fixture({ metadata = columns, range, values } = {}) {
    const calls = [];
    const responses = [
        { value: metadata },
        range || { address: 'USUÁRIOS!A3:M4', columnCount: 13, rowCount: 2 },
        { values: values || [['Empresa teste', 'Pessoa teste', 'Pessoa', 'admin@example.test', 'test-password', 'Administrador', 'Ativo']] },
    ];
    const graphClient = {
        api(path) {
            calls.push(path);
            return { get: async () => responses.shift() };
        },
    };
    return { users: createSupportUsersAdapter({ graphClient }), calls };
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
