'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createSupportUsersAdapter } = require('../integrations/support-users');

const columns = [
    'CNPJ', 'RAZÃO SOCIAL', 'VALOR REFERÊNCIA', 'NOME COMPLETO', 'E-MAIL', 'SENHA',
    'TIPO DE ACESSO', 'AUTORIZADO A CONTRATAR ATENDIMENTOS', 'PAGARME CUSTOMER ID', 'PAGARME CARD ID', 'PAGARME AMBIENTE ',
].map((name, index) => ({ name, index }));
const companyId = '12345678000195';
const defaultRow = ['12.345.678/0001-95', 'Empresa teste', 1000, 'Pessoa teste', 'admin@example.test', 'private-password', 'Administrador', 'Sim'];
const cardReferences = { customerId: 'cus_Test1', cardId: 'card_Test1', environment: 'sandbox' };

function fixture({ metadata = columns, range, values = [defaultRow], companyValues, referenceRows = {}, failUpdateAt, batchResponses, failPost } = {}) {
    metadata = structuredClone(metadata);
    values = structuredClone(values);
    const calls = [];
    const writes = [];
    const graphClient = {
        api(path) {
            calls.push(path);
            let options;
            return {
                middlewareOptions(value) { options = value; return this; },
                async get() {
                    if (path.endsWith('/columns?$select=index,name')) return { value: metadata };
                    if (path.endsWith('/range?$select=address,rowCount,columnCount')) return range || {
                        address: `USUÁRIOS!A3:${metadata.length === 12 ? 'L' : 'K'}${values.length + 3}`, columnCount: metadata.length, rowCount: values.length + 1,
                    };
                    if (/address='A4:A\d+'/u.test(path)) return { values: companyValues ?? values.map((row) => [row[0]]) };
                    const reference = /address='I(\d+):K\d+'/u.exec(path);
                    if (reference) return { values: referenceRows[reference[1]] ?? [['', '', '']] };
                    return { values };
                },
                async update(body) {
                    writes.push({ path, body: structuredClone(body), ...(options ? { options } : {}) });
                    if (writes.length === failUpdateAt) throw new Error('synthetic patch failure');
                    const cell = /address='H(\d+):H\d+'/u.exec(path);
                    if (cell) body.values.forEach((row, index) => { values[Number(cell[1]) - 4 + index][7] = row[0]; });
                },
                async post(body) {
                    writes.push({ path, body: structuredClone(body), options });
                    if (failPost) throw new Error('synthetic batch response lost');
                    const result = typeof batchResponses === 'function' ? batchResponses(body.requests) : batchResponses
                        || { responses: body.requests.map(({ id }) => ({ id, status: 200 })) };
                    for (const response of result.responses || []) if (response.status >= 200 && response.status < 300) {
                        const request = body.requests.find(({ id }) => id === response.id);
                        const cell = /address='H(\d+):H\d+'/u.exec(request?.url || '');
                        if (cell) request.body.values.forEach((row, index) => { values[Number(cell[1]) - 4 + index][7] = row[0]; });
                    }
                    return result;
                },
            };
        },
    };
    return { users: createSupportUsersAdapter({ graphClient }), calls, writes, values };
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
        email: 'admin@example.test', password: 'private-password', role: 'Administrador', status: 'Ativo', bookingAllowed: true,
    }]);
    assert.equal(calls.length, 3);
    assert.match(calls[0], /\/tables\/\{7C4EBF15-124A-4107-9867-F83E9C664B31\}\/columns\?\$select=index,name$/u);
    assert.match(calls[1], /\/range\?\$select=address,rowCount,columnCount$/u);
    assert.match(calls[2], /\/range\(address='A4:H4'\)\?\$select=values$/u);
    assert.ok(calls.every((path) => !path.includes('/rows') && !path.includes('/usedRange')));
});

test('metadata requests start together and user values wait for both validated results', async () => {
    const calls = [];
    let resolveColumns;
    let resolveRange;
    const pendingColumns = new Promise((resolve) => { resolveColumns = resolve; });
    const pendingRange = new Promise((resolve) => { resolveRange = resolve; });
    const users = createSupportUsersAdapter({ graphClient: { api(path) {
        calls.push(path);
        return { get() {
            if (path.endsWith('/columns?$select=index,name')) return pendingColumns;
            if (path.endsWith('/range?$select=address,rowCount,columnCount')) return pendingRange;
            return Promise.resolve({ values: [defaultRow] });
        } };
    } } });
    const pendingUsers = users.readUsers();
    assert.equal(calls.length, 2);
    assert.ok(calls.every((path) => !path.includes("range(address='")));
    resolveRange({ address: 'USUÁRIOS!A3:K4', columnCount: 11, rowCount: 2 });
    await new Promise(setImmediate);
    assert.equal(calls.length, 2);
    resolveColumns({ value: columns });
    assert.equal((await pendingUsers).length, 1);
    assert.equal(calls.length, 3);
    assert.match(calls[2], /range\(address='A4:H4'\)\?\$select=values$/u);
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
        assert.equal(calls.length, 2);
        assert.ok(calls.every((path) => !path.includes("range(address='")));
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
        assert.ok(calls.every((path) => !path.includes("range(address='")));
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

test('company card deletion clears only reusable references for all matching CNPJ colleagues', async () => {
    const other = [...defaultRow];
    other[0] = '98.765.432/0001-10';
    const colleague = [...defaultRow];
    colleague[0] = companyId;
    colleague[1] = 'Renamed company';
    const { users, calls, writes } = fixture({ values: [defaultRow, other, colleague] });
    assert.equal(await users.clearCompanyCard(companyId), undefined);
    assert.match(calls[2], /range\(address='A4:A6'\)\?\$select=values$/u);
    assert.deepEqual(writes.map(({ path, body }) => ({ address: path.match(/address='([^']+)'/u)[1], body })), [
        { address: 'I4:K4', body: { values: [['', '', '']] } },
        { address: 'I6:K6', body: { values: [['', '', '']] } },
    ]);
    assert.ok(calls.every((path) => !path.includes('/rows') && !path.includes('/usedRange') && !/A4:H/u.test(path)));
});

test('company card clearing validates schema, row shape and company existence before any write', async () => {
    for (const options of [
        { metadata: columns.map((column, index) => index === 8 ? { ...column, name: 'SHIFTED' } : column) },
        { range: { address: 'USUÁRIOS!A3:J4', columnCount: 10, rowCount: 2 } },
        { companyValues: [['12.345.678/0001-95', 'extra']] },
    ]) {
        const { users, writes } = fixture(options);
        await assert.rejects(users.clearCompanyCard(companyId), /schema_mismatch/u);
        assert.equal(writes.length, 0);
    }
    const absent = fixture();
    await assert.rejects(absent.users.clearCompanyCard('98765432000110'), /company_not_found/u);
    assert.equal(absent.writes.length, 0);
    for (const invalid of ['', 'Empresa teste', '123']) {
        const { users, calls } = fixture();
        await assert.rejects(users.clearCompanyCard(invalid), /invalid_company/u);
        assert.equal(calls.length, 0);
    }
});

test('company card clearing stops on the first failed patch without retrying or appending', async () => {
    for (const failUpdateAt of [1, 2]) {
        const { users, writes } = fixture({ values: [defaultRow, defaultRow], failUpdateAt });
        await assert.rejects(users.clearCompanyCard(companyId), /synthetic patch failure/u);
        assert.equal(writes.length, failUpdateAt);
        assert.ok(writes.every(({ path, body }) => /range\(address='I[45]:K[45]'\)$/u.test(path)
            && JSON.stringify(body) === JSON.stringify({ values: [['', '', '']] })));
    }
});

const selection = (values, allowed) => values.map((row, index) => ({ email: row[4], allowed: Array.isArray(allowed) ? allowed[index] : allowed }));
const employee = (email, permission = 'Sim', company = companyId) => [company, 'Empresa teste', 1000, email, email, 'private-password', 'Usuário', permission];

test('H permission defaults to Sim while explicit Não and invalid values deny booking without a login-status column', async () => {
    const values = ['Sim', 'Não', '', ' Sim ', 'unexpected'].map((permission, index) => employee(`person${index}@example.test`, permission));
    const f = fixture({ values });
    const directory = await f.users.readUsers();
    assert.deepEqual(directory.map(({ bookingAllowed }) => bookingAllowed), [true, false, true, true, false]);
    assert.ok(directory.every(({ status }) => status === 'Ativo'));
    assert.equal(f.calls.length, 3);
    assert.ok(f.calls.every(path => !/address='[IL]/u.test(path)));
    const old = columns.map((column, index) => index === 7 ? { ...column, name: 'STATUS' } : column);
    await assert.rejects(fixture({ metadata: old }).users.readUsers(), /schema_mismatch/u);
    await assert.rejects(fixture({ metadata: [...columns, { name: 'CONTRATAR ATENDIMENTOS', index: 11 }] }).users.readUsers(), /schema_mismatch/u);
});

test('Save reuses its trusted authorization snapshot, skips unchanged values and writes one exact H cell for one change', async () => {
    const values = [defaultRow, employee('colleague@example.test')];
    const f = fixture({ values });
    const directory = await f.users.readUsers();
    const reads = f.calls.length;
    await f.users.writeBookingPermissions(companyId, selection(values, true), directory);
    assert.equal(f.calls.length, reads);
    assert.equal(f.writes.length, 0);
    await f.users.writeBookingPermissions(companyId, selection(values, [true, false]), directory);
    assert.equal(f.calls.length, reads + 1);
    assert.match(f.writes[0].path, /range\(address='H5:H5'\)$/u);
    assert.deepEqual(f.writes[0].body, { values: [['Não']] });
    assert.equal(f.writes[0].options[0].maxRetries, 0);
    assert.equal(f.writes[0].options[0].shouldRetry(0, 0, {}, {}, { status: 503 }), false);
});

test('an unrecognized or modified browser-like snapshot cannot choose workbook rows and blank defaults are materialized', async () => {
    const values = [employee('a@example.test', ''), employee('other@example.test', 'Sim', '98765432000110')];
    const f = fixture({ values });
    await f.users.writeBookingPermissions(companyId, [{ email: values[0][4], allowed: true }], [{ email: values[0][4], row: 99 }]);
    assert.equal(f.calls.length, 4);
    assert.match(f.writes[0].path, /range\(address='H4:H4'\)$/u);
    assert.deepEqual(f.writes[0].body, { values: [['Sim']] });
    assert.equal(f.values[1][7], 'Sim');
});

test('several changed company cells share one strict Graph batch while foreign company rows and provider columns are untouched', async () => {
    const a = employee('a@example.test'), b = employee('b@example.test'), foreign = employee('outside@example.test', 'Sim', '98765432000110');
    const f = fixture({ values: [a, foreign, b] });
    const directory = await f.users.readUsers();
    await f.users.writeBookingPermissions(companyId, selection([a, b], false), directory);
    assert.equal(f.writes.length, 1);
    assert.equal(f.writes[0].path, '/$batch');
    assert.equal(f.writes[0].options[0].maxRetries, 0);
    const requests = f.writes[0].body.requests;
    assert.deepEqual(requests.map(({ url }) => /address='([^']+)'/u.exec(url)[1]), ['H4:H4', 'H6:H6']);
    assert.ok(requests.every(request => request.method === 'PATCH' && request.headers['Content-Type'] === 'application/json'));
    assert.deepEqual(requests.map(({ body }) => body), [{ values: [['Não']] }, { values: [['Não']] }]);
    assert.deepEqual(requests[1].dependsOn, ['1']);
    assert.ok(!JSON.stringify(requests).includes('private-password'));
    assert.equal(f.values[1][7], 'Sim');
});

test('adjacent changed employee rows are saved with one range update while unchanged cells break the range', async () => {
    const values = [employee('a@example.test'), employee('b@example.test'), employee('c@example.test')];
    const all = fixture({ values });
    await all.users.writeBookingPermissions(companyId, selection(values, false));
    assert.equal(all.writes.length, 1);
    assert.match(all.writes[0].path, /range\(address='H4:H6'\)$/u);
    assert.deepEqual(all.writes[0].body, { values: [['Não'], ['Não'], ['Não']] });
    const separated = fixture({ values });
    await separated.users.writeBookingPermissions(companyId, selection(values, [false, true, false]));
    assert.equal(separated.writes[0].path, '/$batch');
    assert.deepEqual(separated.writes[0].body.requests.map(({ url }) => /address='([^']+)'/u.exec(url)[1]), ['H4:H4', 'H6:H6']);
});

test('batch saves respect Graph twenty-request limits and explicit retries update only remaining cells after partial success', async () => {
    const values = Array.from({ length: 23 }, (_, index) => employee(`person${index}@example.test`));
    const interleaved = values.flatMap((row, index) => [row, employee(`outside${index}@example.test`, 'Sim', '98765432000110')]);
    const many = fixture({ values: interleaved });
    await many.users.writeBookingPermissions(companyId, selection(values, false));
    assert.deepEqual(many.writes.map(({ body }) => body.requests.length), [20, 3]);
    let attempt = 0;
    const partial = fixture({ values: [values[0], employee('outside@example.test', 'Sim', '98765432000110'), values[1]], batchResponses: requests => ({ responses: requests.map(({ id }, index) => ({ id, status: attempt++ === 0 && index === 0 ? 503 : 200 })) }) });
    await assert.rejects(partial.users.writeBookingPermissions(companyId, selection(values.slice(0, 2), false)), /permission_save_failed/u);
    assert.equal(partial.writes.length, 1);
    await partial.users.writeBookingPermissions(companyId, selection(values.slice(0, 2), false));
    assert.equal(partial.writes.length, 2);
    assert.match(partial.writes[1].path, /range\(address='H4:H4'\)$/u);
});

test('invalid batch acknowledgments never claim a successful Save and no SDK retry or automatic replay occurs', async () => {
    const values = [employee('a@example.test'), employee('b@example.test')];
    for (const result of [undefined, {}, { responses: [] }, { responses: [{ id: '1', status: 200 }, { id: '1', status: 200 }] },
        { responses: [{ id: '1', status: 200 }, { id: 'foreign', status: 200 }] },
        { responses: [{ id: '1', status: 200 }, { id: '2', status: 429 }] }]) {
        const f = fixture({ values: [values[0], employee('outside@example.test', 'Sim', '98765432000110'), values[1]], batchResponses: () => result });
        await assert.rejects(f.users.writeBookingPermissions(companyId, selection(values, false)));
        assert.equal(f.writes.length, 1);
    }
});

test('Save rejects partial lists, duplicate or foreign identities and malformed values before any mutation', async () => {
    const values = [employee('a@example.test'), employee('b@example.test')];
    for (const payload of [[{ email: values[0][4], allowed: false }], selection([values[0], values[0]], false),
        [{ email: 'outside@example.test', allowed: false }, { email: values[1][4], allowed: false }]]) {
        const f = fixture({ values });
        await assert.rejects(f.users.writeBookingPermissions(companyId, payload), /(?:employee_not_found|invalid_permission)/u);
        assert.equal(f.writes.length, 0);
    }
    for (const payload of [undefined, {}, [{ email: 'bad', allowed: true }], [{ email: values[0][4], allowed: 'Sim' }], [{ email: values[0][4], allowed: true, row: 88 }]]) {
        const f = fixture({ values });
        await assert.rejects(f.users.writeBookingPermissions(companyId, payload), /invalid_permission/u);
        assert.equal(f.calls.length, 0);
    }
    const duplicate = fixture({ values: [values[0], values[0]] });
    await assert.rejects(duplicate.users.writeBookingPermissions(companyId, selection([values[0]], false)), /employee_not_found/u);
    assert.equal(duplicate.writes.length, 0);
});
