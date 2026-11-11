'use strict';

require('./deny-production-network');
const test = require('node:test');
const assert = require('node:assert/strict');
const { RetryHandlerOptions } = require('@microsoft/microsoft-graph-client');
const { createSupportAppointmentsAdapter } = require('../integrations/support-appointments');
const { addCalendarDays, appointmentAvailability } = require('../domains/support-availability');

const COLUMNS = ['DATA', 'STATUS', 'OBSERVAÇÃO', 'CNPJ', 'NOME COMPLETO', 'VALOR POR HORA', 'DURAÇÃO (MIN)', 'VALOR COBRADO'];
const FORMULA = '=[@[DURAÇÃO (MIN)]]/60*[@[VALOR POR HORA]]';
const MARKER = 'APP: INDISPONÍVEL';
const TODAY = '2026-11-09';
const WINDOW = Array.from({ length: 52 }, (_, index) => addCalendarDays(TODAY, index + 1));

function row(date, status = '', note = '', identity = ['', '', '', '']) {
    return [(Date.parse(`${date}T00:00:00Z`) - Date.UTC(1899, 11, 30)) / 86400000, status, note, ...identity, 0];
}

function fixture({ initial = WINDOW.map((date) => row(date)), intent = null, failAppend, failJournal, invalidFormula = false, snapshotOverride, beforeGet, beforeAppend } = {}) {
    const state = structuredClone(initial);
    const calls = [];
    const requests = [];
    const journal = {
        async read() { calls.push('journal.read'); return structuredClone(intent); },
        async write(value) {
            calls.push('journal.write');
            if (failJournal) throw new Error('synthetic journal failure');
            intent = structuredClone(value);
        },
        async clear() { calls.push('journal.clear'); intent = null; },
    };
    const graphClient = {
        api(path) {
            const request = { path, options: null, body: null };
            requests.push(request);
            return {
                middlewareOptions(options) { request.options = options; return this; },
                async get() {
                    calls.push(`get:${path}`);
                    if (beforeGet) await beforeGet(path, state);
                    if (path.endsWith('/columns?$select=index,name')) return { value: COLUMNS.map((name, index) => ({ name, index })) };
                    if (path.endsWith('/range?$select=address,rowCount,columnCount')) return {
                        address: `ATENDIMENTOS!A3:H${state.length + 3}`, columnCount: 8, rowCount: state.length + 1,
                    };
                    if (path.endsWith('/range?$select=address,rowCount,columnCount,values,formulas')) {
                        const response = {
                            address: `ATENDIMENTOS!A3:H${state.length + 3}`, columnCount: 8, rowCount: state.length + 1,
                            values: [COLUMNS, ...structuredClone(state)],
                            formulas: [COLUMNS, ...state.map((entry) => entry.map((value, index) => index === 7 ? (invalidFormula ? '' : FORMULA) : value))],
                        };
                        return snapshotOverride ? snapshotOverride(response) : response;
                    }
                    const range = /address='A(\d+):([BH])(\d+)'/u.exec(path);
                    assert.ok(range, path);
                    const cells = structuredClone(state.slice(Number(range[1]) - 4, Number(range[3]) - 3));
                    if (range[2] === 'B') return { values: cells.map((entry) => entry.slice(0, 2)) };
                    return { values: cells, formulas: cells.map((entry) => entry.map((value, index) => index === 7 ? (invalidFormula ? '' : FORMULA) : value)) };
                },
                async post(body) {
                    calls.push('append');
                    request.body = structuredClone(body);
                    assert.ok(intent, 'durable intent precedes Graph append');
                    if (beforeAppend) await beforeAppend();
                    if (failAppend === 'before') throw new Error('synthetic lost append response');
                    state.push(...body.values.map((cells) => cells.map((value, index) => index === 7 ? 0 : value)));
                    if (failAppend === 'after') throw new Error('synthetic lost append response');
                },
                async update(body) {
                    calls.push(`patch:${path}`);
                    request.body = structuredClone(body);
                    const range = /address='([AB])(\d+):([CH])\d+'/u.exec(path);
                    assert.ok(range, path);
                    const offset = range[1] === 'A' ? 0 : 1;
                    body.values.forEach((cells, rowIndex) => {
                        const target = state[Number(range[2]) - 4 + rowIndex];
                        cells.forEach((value, index) => { if (value !== null) target[offset + index] = value; });
                    });
                },
            };
        },
    };
    return {
        state, calls, requests, journal, graphClient,
        adapter: createSupportAppointmentsAdapter({ graphClient, availabilityJournal: journal }),
        reload: () => createSupportAppointmentsAdapter({ graphClient, availabilityJournal: journal }),
        intent: () => structuredClone(intent),
    };
}

function mutations(f) { return f.requests.filter(({ body }) => body); }
function noSdkRetry(request) {
    assert.equal(request.options.length, 1);
    const [option] = request.options;
    assert.ok(option instanceof RetryHandlerOptions);
    assert.equal(option.maxRetries, 0);
    assert.equal(option.shouldRetry(0, 0, {}, {}, { status: 504 }), false);
}

test('no-change sync verifies table location, headers, data and formulas with one Graph GET', async () => {
    const initial = WINDOW.map((date) => appointmentAvailability(date, TODAY) ? row(date) : row(date, 'INDISPONÍVEL', MARKER));
    const f = fixture({ initial, snapshotOverride(response) {
        response.values[0] = response.values[0].map((value) => ` ${value} `);
        return response;
    } });
    const result = await f.adapter.syncAvailability({ today: TODAY });
    assert.equal(f.requests.length, 1);
    assert.match(f.requests[0].path, /\/tables\/\{7C4EBF15-124A-4107-9867-F83E9C664B31\}\/range\?\$select=address,rowCount,columnCount,values,formulas$/u);
    assert.equal(mutations(f).length, 0);
    assert.deepEqual(result.map(({ excelRow }) => excelRow), Array.from({ length: WINDOW.length }, (_, index) => index + 4));
    assert.deepEqual(f.state, initial);
});

const booking = { date: '2026-11-10', companyId: '12345678000195', name: 'Pessoa teste', hourlyRate: 975.25 };
const managedRows = () => WINDOW.map((date) => appointmentAvailability(date, TODAY) ? row(date) : row(date, 'INDISPONÍVEL', MARKER));

test('a fresh availability snapshot skips duplicate table reads while privately retaining its selected row', async () => {
    const f = fixture({ initial: managedRows() });
    const rows = await f.adapter.syncAvailability({ today: TODAY });
    rows[0].excelRow = 999;
    rows[0].date = '2026-12-31';
    rows[0].status = 'RESERVADO';
    assert.deepEqual(await f.adapter.reserveDate(booking, rows), { date: booking.date, status: 'RESERVADO', excelRow: 4 });
    assert.equal(f.requests.length, 3, 'Only the existing snapshot, a fresh selected-row check and the reservation PATCH are needed');
    assert.match(f.requests[1].path, /range\(address='A4:B4'\)\?\$select=values$/u);
    assert.match(f.requests[2].path, /range\(address='A4:H4'\)$/u);
    assert.deepEqual(f.requests[2].body, { values: [[null, 'RESERVADO', null, '12.345.678/0001-95', 'Pessoa teste', 975.25, null, null]] });
    assert.equal(f.state[0][1], 'RESERVADO');
    assert.equal(mutations(f).length, 1);
    noSdkRetry(mutations(f)[0]);
});

test('snapshot reuse still rejects a selected date that was moved or occupied after the quote', async () => {
    for (const changed of [[46336, 'RESERVADO'], [46336, 'INDISPONÍVEL'], [46337, '']]) {
        const f = fixture({ initial: managedRows() });
        const rows = await f.adapter.syncAvailability({ today: TODAY });
        [f.state[0][0], f.state[0][1]] = changed;
        await assert.rejects(f.adapter.reserveDate(booking, rows), (error) => {
            assert.equal(error.code, 'support_appointments.date_unavailable');
            assert.equal(error.definiteRejection, true);
            return true;
        });
        assert.equal(f.requests.length, 2);
        assert.equal(mutations(f).length, 0);
    }
});

test('copied, foreign, superseded and already-used snapshots fall back to complete current validation', async () => {
    for (const source of ['copied', 'foreign', 'readDates', 'sync', 'consumed']) {
        const f = fixture({ initial: managedRows() });
        const rows = await f.adapter.syncAvailability({ today: TODAY });
        let adapter = f.adapter;
        let availability = rows;
        let input = booking;
        if (source === 'copied') availability = rows.map((entry) => ({ ...entry, excelRow: 999 }));
        if (source === 'foreign') adapter = f.reload();
        if (source === 'readDates') await f.adapter.readDates();
        if (source === 'sync') await f.adapter.syncAvailability({ today: TODAY });
        if (source === 'consumed') {
            await f.adapter.reserveDate(booking, rows);
            input = { ...booking, date: '2026-11-11' };
        }
        const before = f.requests.length;
        await adapter.reserveDate(input, availability);
        const requests = f.requests.slice(before);
        assert.equal(requests.length, 5, source);
        assert.ok(requests[0].path.endsWith('/columns?$select=index,name'), source);
        assert.ok(requests[1].path.endsWith('/range?$select=address,rowCount,columnCount'), source);
        assert.ok(requests.every(({ path }) => !path.includes('999')), source);
    }
});

test('single-response sync snapshot fails closed on header, range, shape, formula or date drift', async () => {
    const changes = [
        (response) => { response.values[0] = [...COLUMNS].reverse(); },
        (response) => { response.values[0] = COLUMNS.slice(0, 7); },
        (response) => { response.values[0] = [...COLUMNS.slice(0, 7), null]; },
        (response) => { response.address = 'ATENDIMENTOS!A2:H33'; },
        (response) => { response.address = 'ATENDIMENTOS!B3:I33'; },
        (response) => { response.columnCount = 7; },
        (response) => { response.rowCount += 1; },
        (response) => { response.values.pop(); },
        (response) => { response.values[1].pop(); },
        (response) => { response.formulas.pop(); },
        (response) => { response.formulas[0] = []; },
        (response) => { response.formulas[1][7] = ''; },
        (response) => { response.values[1][0] = '2026-11-10'; },
        (response) => { response.values[2][0] = response.values[1][0]; },
    ];
    for (const change of changes) {
        const f = fixture({ snapshotOverride(response) { change(response); return response; } });
        await assert.rejects(f.adapter.syncAvailability({ today: TODAY }), /schema_mismatch/u);
        assert.equal(f.requests.length, 1);
        assert.equal(mutations(f).length, 0);
        assert.equal(f.intent(), null);
    }
});

test('availability sync preserves manual blocks, notes, bookings and computed columns while managing its own vacant rows', async () => {
    const initial = WINDOW.map((date) => row(date));
    initial[0] = row('2026-11-10', 'RESERVADO', 'Manual booking', ['12.345.678/0001-95', 'Pessoa teste', 1000, 60]);
    initial[1] = row('2026-11-11', 'INDISPONÍVEL');
    initial[2] = row('2026-11-12', '', 'Manual observation');
    initial[3] = row('2026-11-13', 'INDISPONÍVEL', MARKER);
    initial[WINDOW.indexOf('2026-12-28')] = row('2026-12-28', 'INDISPONÍVEL');
    initial[WINDOW.indexOf('2026-12-29')] = row('2026-12-29', 'INDISPONÍVEL', 'Recesso');
    initial.push(row('2026-11-08'), row('2027-01-04'));
    const f = fixture({ initial });
    const result = await f.adapter.syncAvailability({ today: TODAY });
    assert.deepEqual(f.state.slice(0, 3), initial.slice(0, 3));
    assert.deepEqual(f.state[3].slice(1, 3), ['', '']);
    for (const date of ['2026-11-08', '2026-11-14', '2026-11-20', '2027-01-04']) {
        const entry = result.find((item) => item.date === date);
        assert.equal(entry.status, 'INDISPONÍVEL', date);
        assert.equal(f.state[entry.excelRow - 4][2], MARKER, date);
    }
    assert.ok(mutations(f).every((request) => /address='B\d+:C\d+'/u.test(request.path)));
    assert.equal(mutations(f).length, 1);
    assert.ok(mutations(f)[0].body.values.some((cells) => cells[0] === null && cells[1] === null));
    for (const request of mutations(f)) noSdkRetry(request);
    assert.equal(result.some((entry) => Object.hasOwn(entry, 'cells')), false);
    assert.deepEqual(f.state[WINDOW.indexOf('2026-12-28')], initial[WINDOW.indexOf('2026-12-28')]);
    assert.deepEqual(f.state[WINDOW.indexOf('2026-12-29')], initial[WINDOW.indexOf('2026-12-29')]);
    assert.ok(f.state.every((cells, index) => cells.slice(3).every((value, offset) => value === initial[index][offset + 3])));
});

test('sync appends missing dates once in one batch through the end of next month', async () => {
    const missing = new Set(['2026-11-10', '2026-11-14', '2026-11-20', '2026-12-31']);
    const f = fixture({ initial: WINDOW.filter((date) => !missing.has(date)).map((date) => row(date)) });
    const result = await f.adapter.syncAvailability({ today: TODAY });
    const appends = mutations(f).filter(({ path }) => path.endsWith('/rows/add'));
    assert.equal(appends.length, 1);
    assert.equal(appends[0].body.values.length, 4);
    noSdkRetry(appends[0]);
    assert.ok(f.calls.indexOf('journal.write') < f.calls.indexOf('append'));
    assert.equal(f.intent(), null);
    assert.deepEqual(new Set(result.map(({ date }) => date)), new Set(WINDOW));
    const appended = appends[0].body.values;
    assert.ok(appended.every((cells) => cells.length === 8 && cells.slice(3, 6).every((value) => value === '') && cells[6] === null && cells[7] === null));
    assert.deepEqual(appended.map((cells) => cells[1]), ['', 'INDISPONÍVEL', 'INDISPONÍVEL', '']);
    await f.adapter.syncAvailability({ today: TODAY });
    assert.equal(mutations(f).filter(({ path }) => path.endsWith('/rows/add')).length, 1);
});

test('November tenth sync fills every remaining date through December thirty-first, including beyond day thirty', async () => {
    const f = fixture({ initial: [] });
    const result = await f.adapter.syncAvailability({ today: '2026-11-10' });
    const appends = mutations(f).filter(({ path }) => path.endsWith('/rows/add'));
    assert.equal(appends.length, 1);
    assert.equal(appends[0].body.values.length, 51);
    assert.equal(result[0].date, '2026-11-11');
    assert.equal(result[result.length - 1].date, '2026-12-31');
    assert.equal(result.find(({ date }) => date === '2026-12-11').status, '');
    assert.equal(result.find(({ date }) => date === '2026-12-31').status, '');
    assert.equal(result.some(({ date }) => date <= '2026-11-10' || date >= '2027-01-01'), false);
    assert.equal(f.intent(), null);
});

test('sync reconciles larger sorted intents but refuses unsorted or oversized intents without writes', async () => {
    const initial = WINDOW.map((date) => appointmentAvailability(date, TODAY) ? row(date) : row(date, 'INDISPONÍVEL', MARKER));
    const reconciled = fixture({ initial, intent: { dates: WINDOW } });
    assert.equal((await reconciled.adapter.syncAvailability({ today: TODAY })).length, WINDOW.length);
    assert.equal(reconciled.intent(), null);
    assert.equal(mutations(reconciled).length, 0);
    for (const dates of [[WINDOW[1], WINDOW[0]], Array.from({ length: 63 }, (_, index) => addCalendarDays(TODAY, index + 1))]) {
        const f = fixture({ initial, intent: { dates } });
        await assert.rejects(f.adapter.syncAvailability({ today: TODAY }), /availability_review_required/u);
        assert.equal(mutations(f).length, 0);
    }
});

test('successful append with a lost response reconciles its durable intent after adapter recreation without replay', async () => {
    const f = fixture({ initial: WINDOW.slice(1).map((date) => row(date)), failAppend: 'after' });
    await assert.rejects(f.adapter.syncAvailability({ today: TODAY }), /lost append response/u);
    assert.deepEqual(f.intent(), { dates: ['2026-11-10'] });
    const result = await f.reload().syncAvailability({ today: TODAY });
    assert.equal(f.intent(), null);
    assert.equal(result.filter(({ date }) => date === '2026-11-10').length, 1);
    assert.equal(f.calls.filter((call) => call === 'append').length, 1);
});

test('ambiguous append with missing or partial dates blocks replay across restarts', async () => {
    const f = fixture({ initial: WINDOW.slice(1).map((date) => row(date)), failAppend: 'before' });
    await assert.rejects(f.adapter.syncAvailability({ today: TODAY }), /lost append response/u);
    await assert.rejects(f.reload().syncAvailability({ today: TODAY }), /availability_review_required/u);
    assert.equal(f.calls.filter((call) => call === 'append').length, 1);
    assert.deepEqual(f.intent(), { dates: ['2026-11-10'] });
    const partial = fixture({ initial: WINDOW.slice(1).map((date) => row(date)), intent: { dates: ['2026-11-10', '2026-11-11'] } });
    await assert.rejects(partial.adapter.syncAvailability({ today: TODAY }), /availability_review_required/u);
    assert.equal(mutations(partial).length, 0);
});

test('sync rechecks each row and preserves a manual booking made after its initial read', async () => {
    let changed = false;
    const f = fixture({ beforeGet(path, state) {
        if (!changed && /address='A8:H\d+'/u.test(path)) {
            changed = true;
            state[4][1] = 'RESERVADO';
            state[4][3] = '12.345.678/0001-95';
            state[4][4] = 'Pessoa teste';
        }
    } });
    const result = await f.adapter.syncAvailability({ today: TODAY });
    assert.equal(changed, true);
    assert.equal(result.find(({ date }) => date === '2026-11-14').status, 'RESERVADO');
    const patch = mutations(f).find(({ path }) => /address='B8:C\d+'/u.test(path));
    assert.deepEqual(patch.body.values[0], [null, null]);
});

test('sync requires a durable journal and exact fresh schema before all writes', async () => {
    const missingJournal = fixture();
    const adapter = createSupportAppointmentsAdapter({ graphClient: missingJournal.graphClient });
    await assert.rejects(adapter.syncAvailability({ today: TODAY }), /availability_journal_required/u);
    assert.equal(mutations(missingJournal).length, 0);
    for (const options of [
        { invalidFormula: true },
        { initial: [row('2026-11-10'), row('2026-11-10')] },
        { initial: WINDOW.slice(1).map((date) => row(date)), failJournal: true },
        { intent: { dates: ['2026-11-10', '2026-11-10'] } },
    ]) {
        const f = fixture(options);
        await assert.rejects(f.adapter.syncAvailability({ today: TODAY }));
        assert.equal(mutations(f).length, 0);
    }
});

test('sync and reservation share one in-process lock so append and row mutation cannot overlap', async () => {
    let release;
    let entered;
    const appended = new Promise((resolve) => { entered = resolve; });
    const gate = new Promise((resolve) => { release = resolve; });
    const f = fixture({ initial: WINDOW.slice(1).map((date) => row(date)), beforeAppend: async () => { entered(); await gate; } });
    const sync = f.adapter.syncAvailability({ today: TODAY });
    await appended;
    const reserve = f.adapter.reserveDate({ date: '2026-11-10', companyId: '12345678000195', name: 'Pessoa teste', hourlyRate: 1000 });
    const secondSync = f.adapter.syncAvailability({ today: TODAY });
    await Promise.resolve();
    assert.equal(mutations(f).length, 1);
    release();
    await Promise.all([sync, reserve, secondSync]);
    assert.equal(f.calls.filter((call) => call === 'append').length, 1);
    const reservation = mutations(f).find(({ path }) => /address='A\d+:H\d+'/u.test(path));
    assert.ok(reservation);
    assert.equal(f.state.find((cells) => cells[0] === row('2026-11-10')[0])[1], 'RESERVADO');
});
