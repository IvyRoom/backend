'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { RetryHandlerOptions } = require('@microsoft/microsoft-graph-client');
const { createSupportAppointmentsAdapter } = require('../integrations/support-appointments');

const columns = ['DATA', 'STATUS', 'OBSERVAÇÃO', 'CNPJ', 'NOME COMPLETO', 'VALOR POR HORA', 'DURAÇÃO (MIN)', 'VALOR COBRADO']
    .map((name, index) => ({ name, index }));
const reservation = { date: '2026-11-10', companyId: '12345678000195', name: 'Pessoa teste', hourlyRate: 975.25 };

function fixture({ metadata = columns, range, values = [[46336, ''], [46337, 'INDISPONÍVEL']], currentRow, failUpdate = false } = {}) {
    const calls = [];
    const writes = [];
    const middlewareOptions = [];
    const graphClient = {
        api(path) {
            calls.push(path);
            return {
                middlewareOptions(options) {
                    middlewareOptions.push({ path, options });
                    return this;
                },
                async get() {
                    if (path.endsWith('/columns?$select=index,name')) return { value: metadata };
                    if (path.endsWith('/range?$select=address,rowCount,columnCount')) return range || {
                        address: `ATENDIMENTOS!A3:H${values.length + 3}`, columnCount: 8, rowCount: values.length + 1,
                    };
                    const selected = /address='A(\d+):B(\d+)'/u.exec(path);
                    if (selected && selected[1] === selected[2]) return { values: [currentRow ?? values[Number(selected[1]) - 4]] };
                    return { values };
                },
                async update(body) {
                    writes.push({ path, body: structuredClone(body) });
                    if (failUpdate) throw new Error('synthetic uncertain patch failure');
                },
            };
        },
    };
    return { appointments: createSupportAppointmentsAdapter({ graphClient }), calls, writes, middlewareOptions };
}

function assertNoSdkRetry(middlewareOptions) {
    assert.equal(middlewareOptions.length, 1);
    const [option] = middlewareOptions[0].options;
    assert.ok(option instanceof RetryHandlerOptions);
    assert.equal(option.maxRetries, 0);
    assert.equal(option.delay, 0);
    assert.equal(option.shouldRetry(0, 0, {}, {}, { status: 503 }), false);
}

test('support appointments construction is import safe and never accesses Graph', () => {
    const graphClient = new Proxy({}, { get() { throw new Error('Graph accessed'); } });
    assert.equal(typeof createSupportAppointmentsAdapter({ graphClient }).readDates, 'function');
});

test('availability reads only dates and statuses, preserving all-day dates and treating unknown status as occupied', async () => {
    const { appointments, calls } = fixture({ values: [[46336, ''], [46337, ' indisponível '], [46338, 'Custom status']] });
    assert.deepEqual(await appointments.readDates(), [
        { date: '2026-11-10', status: '', excelRow: 4 },
        { date: '2026-11-11', status: 'INDISPONÍVEL', excelRow: 5 },
        { date: '2026-11-12', status: 'CUSTOM STATUS', excelRow: 6 },
    ]);
    assert.equal(calls.length, 3);
    assert.match(calls[0], /\/items\/01OSXVECSFKTNUUJS3TZDKCLDRXFZCMVYB\/workbook\//u);
    assert.match(calls[2], /range\(address='A4:B6'\)\?\$select=values$/u);
    assert.ok(calls.every((path) => !path.includes('/rows') && !path.includes('/usedRange')));
});

test('availability starts independent metadata reads together and waits for both before reading dates', async () => {
    let resolveColumns;
    let resolveRange;
    const pendingColumns = new Promise((resolve) => { resolveColumns = resolve; });
    const pendingRange = new Promise((resolve) => { resolveRange = resolve; });
    const calls = [];
    const appointments = createSupportAppointmentsAdapter({ graphClient: { api(path) {
        calls.push(path);
        return { get() {
            if (path.endsWith('/columns?$select=index,name')) return pendingColumns;
            if (path.endsWith('/range?$select=address,rowCount,columnCount')) return pendingRange;
            return Promise.resolve({ values: [[46336, '']] });
        } };
    } } });
    const pending = appointments.readDates();
    assert.equal(calls.length, 2);
    resolveRange({ address: 'ATENDIMENTOS!A3:H4', columnCount: 8, rowCount: 2 });
    await new Promise(setImmediate);
    assert.equal(calls.length, 2);
    resolveColumns({ value: columns });
    assert.deepEqual(await pending, [{ date: '2026-11-10', status: '', excelRow: 4 }]);
    assert.equal(calls.length, 3);
});

test('availability rejects shifted schemas and ranges before reading appointment values', async () => {
    for (const metadata of [
        columns.slice(0, 7), [...columns].reverse(), [...columns, { name: 'EXTRA', index: 8 }],
        columns.map((column, index) => index === 6 ? { ...column, name: 'CHANGED' } : column),
        columns.map((column, index) => index === 7 ? { ...column, index: 6 } : column),
    ]) {
        const { appointments, calls } = fixture({ metadata });
        await assert.rejects(appointments.readDates(), /schema_mismatch/u);
        assert.equal(calls.length, 2);
    }
    for (const range of [
        { address: 'ATENDIMENTOS!A2:H5', columnCount: 8, rowCount: 4 },
        { address: 'ATENDIMENTOS!B3:I5', columnCount: 8, rowCount: 3 },
        { address: 'ATENDIMENTOS!A3:H5', columnCount: 7, rowCount: 3 },
        { address: 'ATENDIMENTOS!A3:H5', columnCount: 8, rowCount: 4 },
        { address: 'OTHER!A3:H5', columnCount: 8, rowCount: 3 },
    ]) {
        const { appointments, calls } = fixture({ range });
        await assert.rejects(appointments.readDates(), /schema_mismatch/u);
        assert.equal(calls.length, 2);
    }
});

test('availability refuses malformed or duplicate dates and row shapes', async () => {
    for (const value of ['2026-11-10', '10/11/2026', '', null, 46336.5, 60, Infinity, 2958466]) {
        await assert.rejects(fixture({ values: [[value, ''], [46337, '']] }).appointments.readDates(), /schema_mismatch/u);
    }
    for (const values of [[[46336, ''], [46336, '']], [[46336]], [[46336, '', 'extra']]]) {
        await assert.rejects(fixture({ values }).appointments.readDates(), /schema_mismatch/u);
    }
});

test('availability handles a header-only table without fetching values', async () => {
    const { appointments, calls } = fixture({ values: [] });
    assert.deepEqual(await appointments.readDates(), []);
    assert.equal(calls.length, 2);
});

test('Excel date conversion preserves month boundaries and valid calendar limits', async () => {
    const { appointments } = fixture({ values: [[61, ''], [46357, ''], [2958465, '']] });
    assert.deepEqual((await appointments.readDates()).map(({ date }) => date), ['1900-03-01', '2026-12-01', '9999-12-31']);
});

test('reservation rechecks the selected date and writes a fixed row preserving notes, duration and charge formula', async () => {
    const { appointments, calls, writes, middlewareOptions } = fixture();
    assert.deepEqual(await appointments.reserveDate(reservation), { date: '2026-11-10', status: 'RESERVADO', excelRow: 4 });
    assert.equal(calls.length, 5);
    assert.match(calls[3], /range\(address='A4:B4'\)\?\$select=values$/u);
    assert.match(writes[0].path, /range\(address='A4:H4'\)$/u);
    assert.deepEqual(writes[0].body, { values: [[null, 'RESERVADO', null, '12.345.678/0001-95', 'Pessoa teste', 975.25, null, null]] });
    assert.equal(writes.length, 1);
    assertNoSdkRetry(middlewareOptions);
});

test('reservation refuses missing or occupied dates and concurrent edits before writing', async () => {
    for (const options of [
        { values: [[46337, '']] },
        { values: [[46336, 'RESERVADO']] },
        { values: [[46336, 'INDISPONÍVEL']] },
        { values: [[46336, 'Custom status']] },
        { currentRow: [46336, 'RESERVADO'] },
        { currentRow: [46337, ''] },
    ]) {
        const { appointments, writes } = fixture(options);
        await assert.rejects(appointments.reserveDate(reservation), (error) => {
            assert.equal(error.code, 'support_appointments.date_unavailable');
            assert.equal(error.definiteRejection, true);
            return true;
        });
        assert.equal(writes.length, 0);
    }
});

test('reservation refuses malformed dates, untrusted identities and unsafe values before Graph access', async () => {
    for (const changed of [
        { date: '2026-02-30' }, { date: '2026-11-10T00:00:00Z' }, { companyId: '12.345.678/0001-95' },
        { companyId: 'Company' }, { name: '' }, { name: ' =SUM(A1)' }, { name: '@formula' },
        { hourlyRate: '975.25' }, { hourlyRate: 0 }, { hourlyRate: NaN }, { hourlyRate: Infinity },
    ]) {
        const { appointments, calls } = fixture();
        await assert.rejects(appointments.reserveDate({ ...reservation, ...changed }), /invalid_reservation/u);
        assert.equal(calls.length, 0);
    }
});

test('reservation never retries or appends when a patch has uncertain success', async () => {
    const { appointments, writes, middlewareOptions } = fixture({ failUpdate: true });
    await assert.rejects(appointments.reserveDate(reservation), (error) => {
        assert.equal(error.message, 'synthetic uncertain patch failure');
        assert.equal(error.definiteRejection, undefined);
        return true;
    });
    assert.equal(writes.length, 1);
    assert.ok(writes.every(({ path }) => !path.includes('/rows')));
    assertNoSdkRetry(middlewareOptions);
});
