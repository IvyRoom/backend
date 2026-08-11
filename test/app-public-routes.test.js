'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
    createTestApp,
    startLoopback,
} = require('./app-test-support');

const PLATFORM_ROWS = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows';

function platformRow(name, completedTopics, gradesAndCertificate = []) {
    const cells = new Array(22).fill(null);
    cells[0] = name;
    cells[8] = completedTopics;
    gradesAndCertificate.forEach((value, index) => { cells[index + 10] = value; });
    return { values: [cells] };
}

async function postJson(origin, path, body) {
    return fetch(`${origin}${path}`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(body),
    });
}

test('statusreport is public and projects inclusive numeric bounds into exact 14-value rows', async (t) => {
    const rows = [
        platformRow('Ignored', 0, Array.from({ length: 12 }, (_, index) => index)),
        platformRow('First', 7, Array.from({ length: 12 }, (_, index) => `a${index}`)),
        platformRow('Second', 8, Array.from({ length: 12 }, (_, index) => `b${index}`)),
        platformRow('Ignored after', 9, Array.from({ length: 12 }, (_, index) => `c${index}`)),
    ];
    const harness = createTestApp();
    harness.graphClient.enqueue('GET', PLATFORM_ROWS, { value: rows });
    const { origin } = await startLoopback(harness.app, t);

    const response = await postJson(origin, '/plataforma_v2/statusreport', {
        linha_inicial: 1,
        linha_final: 2,
    });

    assert.equal(response.status, 200);
    assert.match(response.headers.get('content-type'), /^application\/json; charset=utf-8$/);
    assert.deepEqual(await response.json(), {
        Dados_Extraídos_BD_Plataforma: [
            ['First', 7, ...Array.from({ length: 12 }, (_, index) => `a${index}`)],
            ['Second', 8, ...Array.from({ length: 12 }, (_, index) => `b${index}`)],
        ],
    });
    assert.equal(harness.graphClient.calls.length, 1);
});

test('statusreport preserves linha_final string concatenation before slice coercion', async (t) => {
    const rows = Array.from({ length: 6 }, (_, index) => platformRow(`Row ${index}`, index, new Array(12).fill(index)));
    const harness = createTestApp();
    harness.graphClient.enqueue('GET', PLATFORM_ROWS, { value: rows });
    const { origin } = await startLoopback(harness.app, t);

    const response = await postJson(origin, '/plataforma_v2/statusreport', {
        linha_inicial: 2,
        linha_final: '3',
    });
    const body = await response.json();

    assert.equal(response.status, 200);
    assert.deepEqual(body.Dados_Extraídos_BD_Plataforma.map((row) => row[0]), ['Row 2', 'Row 3', 'Row 4', 'Row 5']);
});

test('statusreport exhausts five read attempts before Erro_001', async (t) => {
    const harness = createTestApp();
    harness.graphClient.enqueue('GET', PLATFORM_ROWS, ...Array.from({ length: 5 }, () => new Error('read failed')));
    const { origin } = await startLoopback(harness.app, t);

    const response = await postJson(origin, '/plataforma_v2/statusreport', {
        linha_inicial: 0,
        linha_final: 0,
    });

    assert.equal(response.status, 500);
    assert.deepEqual(await response.json(), { error: 'Erro_001' });
    assert.equal(harness.graphClient.calls.length, 5);
    assert.deepEqual(harness.ledger.filter((entry) => entry.type === 'sleep').map((entry) => entry.delay), [500, 1000, 1500, 2000]);
});

function certificateWorkbook(id, accumulated, name = 'Titular') {
    const cells = new Array(22).fill(null);
    cells[0] = name;
    cells[20] = accumulated;
    cells[21] = id;
    return { value: [{ values: [cells] }] };
}

test('certificate validation normalizes IDs and preserves fraction-or-percent inclusive threshold behavior', async (t) => {
    const harness = createTestApp();
    harness.graphClient.enqueue(
        'GET',
        PLATFORM_ROWS,
        certificateWorkbook(' fMg-AbCd-1234 ', 0.704, 'Pessoa Um'),
        certificateWorkbook('FMG-FRAC-0070', 0.70, 'Pessoa Dois'),
        certificateWorkbook('FMG-WHOL-0070', 70, 'Pessoa Três'),
        certificateWorkbook('FMG-ROUND-FAIL', 69.999, 'Pessoa Quatro'),
        certificateWorkbook('FMG-FRAC-FAIL', 0.69999, 'Pessoa Cinco'),
        certificateWorkbook('FMG-NAN-0000', Number.NaN, 'Pessoa Seis'),
        certificateWorkbook('FMG-OTHER-0001', 100, 'Pessoa Sete'),
        certificateWorkbook('FMG-OTHER-0002', 100, 'Pessoa Oito'),
    );
    const { origin } = await startLoopback(harness.app, t);

    const normalized = await fetch(`${origin}/validacaocertificados/%20fmg-abcd-1234%20`);
    assert.deepEqual(await normalized.json(), {
        Certificado_Válido: true,
        Titular_NomeCompleto: 'Pessoa Um',
        Acumulado_Percentual: 70,
        Certificado_ID: 'FMG-ABCD-1234',
    });

    for (const id of ['FMG-FRAC-0070', 'FMG-WHOL-0070']) {
        const response = await fetch(`${origin}/validacaocertificados/${id}`);
        assert.equal(response.status, 200);
        assert.equal((await response.json()).Certificado_Válido, true);
    }

    for (const id of ['FMG-ROUND-FAIL', 'FMG-FRAC-FAIL', 'FMG-NAN-0000', 'MISSING']) {
        const response = await fetch(`${origin}/validacaocertificados/${id}`);
        assert.equal(response.status, 200);
        assert.deepEqual(await response.json(), { Certificado_Válido: false });
    }

    const normalizedEmpty = await fetch(`${origin}/validacaocertificados/%20%20`);
    assert.equal(normalizedEmpty.status, 200);
    assert.deepEqual(await normalizedEmpty.json(), { Certificado_Válido: false });
});

test('certificate validation missing segment remains default 404 without a workbook read', async (t) => {
    const harness = createTestApp();
    const { origin } = await startLoopback(harness.app, t);

    const response = await fetch(`${origin}/validacaocertificados/`);

    assert.equal(response.status, 404);
    assert.match(response.headers.get('content-type'), /^text\/html; charset=utf-8$/);
    assert.match(await response.text(), /Cannot GET \/validacaocertificados\//);
    assert.equal(harness.graphClient.calls.length, 0);
});

test('certificate validation exhausts five read attempts before Erro_001', async (t) => {
    const harness = createTestApp();
    harness.graphClient.enqueue('GET', PLATFORM_ROWS, ...Array.from({ length: 5 }, () => new Error('read failed')));
    const { origin } = await startLoopback(harness.app, t);

    const response = await fetch(`${origin}/validacaocertificados/FMG-ABCD-1234`);

    assert.equal(response.status, 500);
    assert.deepEqual(await response.json(), { error: 'Erro_001' });
    assert.equal(harness.graphClient.calls.length, 5);
    assert.deepEqual(harness.ledger.filter((entry) => entry.type === 'sleep').map((entry) => entry.delay), [500, 1000, 1500, 2000]);
});
