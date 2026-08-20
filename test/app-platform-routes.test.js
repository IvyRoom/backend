'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
    TEST_NOW_MS,
    createTestPlatformRowAuthorization,
    createInvalidPlatformRowHandles,
    createValueSequence,
    createTestApp,
    startLoopback,
} = require('./app-test-support');

const PLATFORM_TABLE = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}';
const PLATFORM_ROWS = `${PLATFORM_TABLE}/rows`;
const FEEDBACK_ROWS = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECXO7I5R6LKLXJD3VWXORUAF7J37/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/add';
const FACE_SESSIONS = '/detectLivenessWithVerify-sessions';
const FACE_RESULT = '/detectLivenessWithVerify-sessions/{sessionId}';
const RETRY_DELAYS = [500, 1_000, 1_500, 2_000];
const REFERENCE_PHOTO = Buffer.from([0, 1, 2, 127, 128, 254, 255]);

function platformItem(rowIndex) {
    return `${PLATFORM_ROWS}/itemAt(index=${rowIndex})`;
}

function referencePhoto(rowIndex) {
    return `/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/root:/2. ENTREGA/1. CONTROLAR PLATAFORMA/PG - FOTOS DE REFERÊNCIA/${rowIndex}.jpg:/content`;
}

function fiveFailures(label) {
    return Array.from({ length: 5 }, (_, index) => new Error(`${label} ${index + 1}`));
}

function retrySleeps(ledger) {
    return ledger
        .filter((event) => event.type === 'sleep')
        .map((event) => event.delay);
}

function externalOrder(ledger) {
    return ledger
        .filter((event) => event.type === 'external-call')
        .map(({ client, method, path }) => [client, method, path]);
}

function workbookRow(cells) {
    return { values: [cells] };
}

function makePlatformRow(values = {}) {
    const row = new Array(22).fill(null);
    for (const [index, value] of Object.entries(values)) row[Number(index)] = value;
    return row;
}

function registrationForm(handle, fileField = 'file') {
    const form = new FormData();
    if (handle !== undefined) form.append('IndexVerificado', String(handle));
    form.append(fileField, new Blob([REFERENCE_PHOTO], { type: 'image/jpeg' }), 'reference.jpg');
    return form;
}

async function launch(t, overrides = {}) {
    const harness = createTestApp(overrides);
    const loopback = await startLoopback(harness.app, t);
    return { ...harness, ...loopback };
}

function postJson(origin, path, body) {
    return fetch(`${origin}${path}`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(body),
    });
}

async function assertJsonResponse(response, status, expectedBody) {
    assert.equal(response.status, status);
    assert.equal(response.headers.get('content-type'), 'application/json; charset=utf-8');
    assert.deepEqual(await response.json(), expectedBody);
}

function faceSessionResponse(authToken, sessionId, extra = {}) {
    return {
        ...extra,
        body: { authToken, sessionId },
    };
}

function assertFaceSessionCall(call, expectedPhoto, expectedUuid) {
    assert.equal(call.client, 'face');
    assert.equal(call.method, 'POST');
    assert.equal(call.path, FACE_SESSIONS);
    assert.deepEqual(call.pathArguments, [FACE_SESSIONS]);
    assert.deepEqual(call.body, {
        contentType: 'multipart/form-data',
        body: [
            { name: 'VerifyImage', body: expectedPhoto },
            { name: 'livenessOperationMode', body: 'Passive' },
            { name: 'deviceCorrelationId', body: expectedUuid },
        ],
    });
}

test('all five protected routes reject every invalid handle class before handler calls', async (t) => {
    const harness = await launch(t);
    const invalidHandles = createInvalidPlatformRowHandles({
        rowIndex: 7,
        nowMs: TEST_NOW_MS,
    });
    const jsonRoutes = [
        '/plataforma_v2/FaceID',
        '/plataforma_v2/refresh',
        '/plataforma_v2/updates',
        '/plataforma_v2/processa-feedback',
    ];

    for (const [invalidClass, handle] of Object.entries(invalidHandles)) {
        await t.test(`CadastroFoto_e_FaceID rejects ${invalidClass}`, async () => {
            const response = invalidClass === 'numeric'
                ? await postJson(harness.origin, '/plataforma_v2/CadastroFoto_e_FaceID', {
                    IndexVerificado: handle,
                })
                : await fetch(`${harness.origin}/plataforma_v2/CadastroFoto_e_FaceID`, {
                    method: 'POST',
                    body: registrationForm(handle),
                });
            await assertJsonResponse(response, 401, {});
        });

        for (const path of jsonRoutes) {
            await t.test(`${path} rejects ${invalidClass}`, async () => {
                const response = await postJson(harness.origin, path, {
                    IndexVerificado: handle,
                });
                await assertJsonResponse(response, 401, {});
            });
        }
    }

    assert.deepEqual(harness.graphClient.calls, []);
    assert.deepEqual(harness.faceClient.calls, []);
    assert.deepEqual(harness.ledger, []);
});

test('CadastroFoto_e_FaceID runs Multer before authorization', async (t) => {
    const baseAuthorization = createTestPlatformRowAuthorization();
    let authorizationCalls = 0;
    const platformRowAuthorization = {
        ...baseAuthorization,
        authorize(req, res, next) {
            authorizationCalls += 1;
            return baseAuthorization.authorize(req, res, next);
        },
    };
    const harness = await launch(t, { platformRowAuthorization });
    harness.app.set('env', 'test');
    const malformedHandle = createInvalidPlatformRowHandles().malformed;
    const response = await fetch(`${harness.origin}/plataforma_v2/CadastroFoto_e_FaceID`, {
        method: 'POST',
        body: registrationForm(malformedHandle, 'unexpected-file-field'),
    });

    assert.equal(response.status, 500);
    assert.match(response.headers.get('content-type'), /^text\/html;/);
    assert.match(await response.text(), /<!DOCTYPE html>/i);
    assert.equal(authorizationCalls, 0);
    assert.deepEqual(harness.graphClient.calls, []);
    assert.deepEqual(harness.faceClient.calls, []);
});

test('login-FaceID preserves active, inactive, and invalid-credential responses', async (t) => {
    const harness = await launch(t);
    const activeRow = makePlatformRow({
        2: 'active@example.com',
        3: 123456,
        4: 'Pendente',
        5: 'Não',
        6: 45292.5,
        7: 'Ativo',
    });
    const inactiveRow = makePlatformRow({
        2: 'inactive@example.com',
        3: 'inactive-password',
        4: 'Concluído',
        5: 'Sim',
        6: 45292.5,
        7: 'Inativo',
    });
    const decoy = makePlatformRow({ 2: 'decoy@example.com', 3: 'decoy-password' });
    harness.graphClient.enqueue(
        'GET',
        PLATFORM_ROWS,
        { value: [workbookRow(decoy), workbookRow(decoy), workbookRow(activeRow)] },
        { value: [workbookRow(inactiveRow)] },
        { value: [workbookRow(decoy)] },
    );

    const activeResponse = await postJson(harness.origin, '/plataforma_v2/login-FaceID', {
        Usuário_Login: 'active@example.com',
        Usuário_Senha: '123456',
    });
    await assertJsonResponse(activeResponse, 200, {
        Usuário_Status_FaceID: 'Pendente',
        Usuário_Foto_Cadastrada: 'Não',
        Usuário_PrazoAcesso: '01/jan/2024',
        Usuário_Status_Login: 'Ativo',
        IndexVerificado: harness.platformRowAuthorization.createHandle(2),
    });

    const inactiveResponse = await postJson(harness.origin, '/plataforma_v2/login-FaceID', {
        Usuário_Login: 'inactive@example.com',
        Usuário_Senha: 'inactive-password',
    });
    await assertJsonResponse(inactiveResponse, 200, {
        Usuário_Status_FaceID: 'Concluído',
        Usuário_Foto_Cadastrada: 'Sim',
        Usuário_PrazoAcesso: '01/jan/2024',
        Usuário_Status_Login: 'Inativo',
    });

    const invalidResponse = await postJson(harness.origin, '/plataforma_v2/login-FaceID', {
        Usuário_Login: 'decoy@example.com',
        Usuário_Senha: 'wrong-password',
    });
    await assertJsonResponse(invalidResponse, 401, { error: 'credenciais_inválidas' });

    assert.equal(harness.graphClient.calls.length, 3);
    assert.ok(harness.graphClient.calls.every((call) => (
        call.method === 'GET' && call.path === PLATFORM_ROWS
    )));
    assert.deepEqual(retrySleeps(harness.ledger), []);
    harness.graphClient.assertExhausted();
});

test('login-FaceID exhausts five Graph reads before learning_platform.read_platform_data_failed', async (t) => {
    const harness = await launch(t);
    harness.graphClient.enqueue('GET', PLATFORM_ROWS, ...fiveFailures('login read'));

    const response = await postJson(harness.origin, '/plataforma_v2/login-FaceID', {
        Usuário_Login: 'user@example.com',
        Usuário_Senha: 'password',
    });

    await assertJsonResponse(response, 500, { error: 'learning_platform.read_platform_data_failed' });
    assert.equal(harness.graphClient.calls.length, 5);
    assert.deepEqual(retrySleeps(harness.ledger), RETRY_DELAYS);
    harness.graphClient.assertExhausted();
});

test('CadastroFoto_e_FaceID preserves upload, workbook, and Face call order and payloads', async (t) => {
    const rowIndex = 7;
    const uuid = '11111111-1111-4111-8111-111111111111';
    const harness = await launch(t, { uuid: createValueSequence([uuid]) });
    const handle = harness.platformRowAuthorization.createHandle(rowIndex);
    harness.graphClient.enqueue('PUT', referencePhoto(rowIndex), { uploaded: true });
    harness.graphClient.enqueue('UPDATE', platformItem(rowIndex), { updated: true });
    harness.faceClient.enqueue(
        'POST',
        FACE_SESSIONS,
        faceSessionResponse('registration-token', 'registration-session', { status: '500' }),
    );

    const response = await fetch(`${harness.origin}/plataforma_v2/CadastroFoto_e_FaceID`, {
        method: 'POST',
        body: registrationForm(handle),
    });

    await assertJsonResponse(response, 200, {
        Azure_Face_API_LivenessSession_authToken: 'registration-token',
        Azure_Face_API_LivenessSession_sessionID: 'registration-session',
    });
    assert.deepEqual(externalOrder(harness.ledger), [
        ['graph', 'PUT', referencePhoto(rowIndex)],
        ['graph', 'UPDATE', platformItem(rowIndex)],
        ['face', 'POST', FACE_SESSIONS],
    ]);
    assert.deepEqual(harness.graphClient.calls[0].body, REFERENCE_PHOTO);

    const registrationUpdate = harness.graphClient.calls[1].body.values[0];
    const expectedUpdate = new Array(22).fill(null);
    expectedUpdate[5] = 'Sim';
    assert.equal(registrationUpdate.length, 22);
    assert.deepEqual(registrationUpdate, expectedUpdate);
    assertFaceSessionCall(harness.faceClient.calls[0], REFERENCE_PHOTO, uuid);
    assert.deepEqual(retrySleeps(harness.ledger), []);
    harness.graphClient.assertExhausted();
    harness.faceClient.assertExhausted();
});

test('CadastroFoto_e_FaceID preserves each five-attempt error boundary and prior effects', async (t) => {
    await t.test('photo upload exhaustion returns learning_platform.upload_reference_photo_failed before later effects', async (t) => {
        const rowIndex = 8;
        const harness = await launch(t);
        const handle = harness.platformRowAuthorization.createHandle(rowIndex);
        harness.graphClient.enqueue('PUT', referencePhoto(rowIndex), ...fiveFailures('photo upload'));

        const response = await fetch(`${harness.origin}/plataforma_v2/CadastroFoto_e_FaceID`, {
            method: 'POST',
            body: registrationForm(handle),
        });

        await assertJsonResponse(response, 500, { error: 'learning_platform.upload_reference_photo_failed' });
        assert.equal(harness.graphClient.calls.length, 5);
        assert.ok(harness.graphClient.calls.every((call) => (
            call.method === 'PUT'
            && call.path === referencePhoto(rowIndex)
            && Buffer.compare(call.body, REFERENCE_PHOTO) === 0
        )));
        assert.deepEqual(harness.faceClient.calls, []);
        assert.deepEqual(retrySleeps(harness.ledger), RETRY_DELAYS);
        harness.graphClient.assertExhausted();
    });

    await t.test('workbook exhaustion returns learning_platform.update_reference_photo_registration_failed after the photo persists', async (t) => {
        const rowIndex = 9;
        const harness = await launch(t);
        const handle = harness.platformRowAuthorization.createHandle(rowIndex);
        harness.graphClient.enqueue('PUT', referencePhoto(rowIndex), { uploaded: true });
        harness.graphClient.enqueue('UPDATE', platformItem(rowIndex), ...fiveFailures('photo flag'));

        const response = await fetch(`${harness.origin}/plataforma_v2/CadastroFoto_e_FaceID`, {
            method: 'POST',
            body: registrationForm(handle),
        });

        await assertJsonResponse(response, 500, { error: 'learning_platform.update_reference_photo_registration_failed' });
        assert.deepEqual(externalOrder(harness.ledger).map((entry) => entry[1]), [
            'PUT', 'UPDATE', 'UPDATE', 'UPDATE', 'UPDATE', 'UPDATE',
        ]);
        assert.deepEqual(harness.faceClient.calls, []);
        assert.deepEqual(retrySleeps(harness.ledger), RETRY_DELAYS);
        harness.graphClient.assertExhausted();
    });

    await t.test('Face exhaustion returns learning_platform.create_face_liveness_session_failed after photo and flag and regenerates UUIDs', async (t) => {
        const rowIndex = 10;
        const uuids = Array.from(
            { length: 5 },
            (_, index) => `22222222-2222-4222-8222-22222222222${index}`,
        );
        const harness = await launch(t, { uuid: createValueSequence(uuids) });
        const handle = harness.platformRowAuthorization.createHandle(rowIndex);
        harness.graphClient.enqueue('PUT', referencePhoto(rowIndex), { uploaded: true });
        harness.graphClient.enqueue('UPDATE', platformItem(rowIndex), { updated: true });
        harness.faceClient.enqueue('POST', FACE_SESSIONS, ...fiveFailures('Face registration'));

        const response = await fetch(`${harness.origin}/plataforma_v2/CadastroFoto_e_FaceID`, {
            method: 'POST',
            body: registrationForm(handle),
        });

        await assertJsonResponse(response, 500, { error: 'learning_platform.create_face_liveness_session_failed' });
        assert.deepEqual(externalOrder(harness.ledger).slice(0, 2), [
            ['graph', 'PUT', referencePhoto(rowIndex)],
            ['graph', 'UPDATE', platformItem(rowIndex)],
        ]);
        assert.equal(harness.faceClient.calls.length, 5);
        assert.deepEqual(
            harness.faceClient.calls.map((call) => call.body.body[2].body),
            uuids,
        );
        assert.ok(harness.faceClient.calls.every((call) => (
            Buffer.compare(call.body.body[0].body, REFERENCE_PHOTO) === 0
        )));
        assert.deepEqual(retrySleeps(harness.ledger), RETRY_DELAYS);
        harness.graphClient.assertExhausted();
        harness.faceClient.assertExhausted();
    });
});

test('FaceID reads the verified-row photo before creating an exact Face session', async (t) => {
    const rowIndex = 11;
    const uuid = '33333333-3333-4333-8333-333333333333';
    const downloadedPhoto = Buffer.from('downloaded reference photo');
    const harness = await launch(t, { uuid: createValueSequence([uuid]) });
    const handle = harness.platformRowAuthorization.createHandle(rowIndex);
    harness.graphClient.enqueue('GET', referencePhoto(rowIndex), downloadedPhoto);
    harness.faceClient.enqueue(
        'POST',
        FACE_SESSIONS,
        faceSessionResponse('login-token', 'login-session', { status: '500' }),
    );

    const response = await postJson(harness.origin, '/plataforma_v2/FaceID', {
        IndexVerificado: handle,
    });

    await assertJsonResponse(response, 200, {
        Azure_Face_API_LivenessSession_authToken: 'login-token',
        Azure_Face_API_LivenessSession_sessionID: 'login-session',
    });
    assert.deepEqual(externalOrder(harness.ledger), [
        ['graph', 'GET', referencePhoto(rowIndex)],
        ['face', 'POST', FACE_SESSIONS],
    ]);
    assertFaceSessionCall(harness.faceClient.calls[0], downloadedPhoto, uuid);
    harness.graphClient.assertExhausted();
    harness.faceClient.assertExhausted();
});

test('FaceID preserves photo and Face five-attempt errors', async (t) => {
    await t.test('photo exhaustion returns learning_platform.read_reference_photo_failed without Face work', async (t) => {
        const rowIndex = 12;
        const harness = await launch(t);
        const handle = harness.platformRowAuthorization.createHandle(rowIndex);
        harness.graphClient.enqueue('GET', referencePhoto(rowIndex), ...fiveFailures('photo read'));

        const response = await postJson(harness.origin, '/plataforma_v2/FaceID', {
            IndexVerificado: handle,
        });

        await assertJsonResponse(response, 500, { error: 'learning_platform.read_reference_photo_failed' });
        assert.equal(harness.graphClient.calls.length, 5);
        assert.deepEqual(harness.faceClient.calls, []);
        assert.deepEqual(retrySleeps(harness.ledger), RETRY_DELAYS);
        harness.graphClient.assertExhausted();
    });

    await t.test('Face exhaustion returns learning_platform.create_face_liveness_session_failed after the read and regenerates UUIDs', async (t) => {
        const rowIndex = 13;
        const downloadedPhoto = Buffer.from('persisted photo');
        const uuids = Array.from(
            { length: 5 },
            (_, index) => `44444444-4444-4444-8444-44444444444${index}`,
        );
        const harness = await launch(t, { uuid: createValueSequence(uuids) });
        const handle = harness.platformRowAuthorization.createHandle(rowIndex);
        harness.graphClient.enqueue('GET', referencePhoto(rowIndex), downloadedPhoto);
        harness.faceClient.enqueue('POST', FACE_SESSIONS, ...fiveFailures('Face login'));

        const response = await postJson(harness.origin, '/plataforma_v2/FaceID', {
            IndexVerificado: handle,
        });

        await assertJsonResponse(response, 500, { error: 'learning_platform.create_face_liveness_session_failed' });
        assert.deepEqual(externalOrder(harness.ledger)[0], [
            'graph', 'GET', referencePhoto(rowIndex),
        ]);
        assert.equal(harness.faceClient.calls.length, 5);
        assert.deepEqual(
            harness.faceClient.calls.map((call) => call.body.body[2].body),
            uuids,
        );
        assert.ok(harness.faceClient.calls.every((call) => (
            Buffer.compare(call.body.body[0].body, downloadedPhoto) === 0
        )));
        assert.deepEqual(retrySleeps(harness.ledger), RETRY_DELAYS);
        harness.graphClient.assertExhausted();
        harness.faceClient.assertExhausted();
    });
});

test('FaceID_resultado is public and forwards the decoded session parameter exactly', async (t) => {
    const harness = await launch(t);
    const sessionId = 'session / with spaces';
    harness.faceClient.enqueue('GET', FACE_RESULT, {
        status: '503',
        body: {
            results: {
                attempts: [{
                    result: {
                        livenessDecision: 'realface',
                        verifyResult: {
                            matchConfidence: 0.9876,
                            isIdentical: true,
                        },
                    },
                }],
            },
        },
    });

    const response = await fetch(
        `${harness.origin}/plataforma_v2/FaceID_resultado/${encodeURIComponent(sessionId)}`,
    );

    await assertJsonResponse(response, 200, {
        Azure_Face_API_LivenessSession_LivenessDecision: 'realface',
        Azure_Face_API_LivenessSession_MatchConfidence: 0.9876,
        Azure_Face_API_LivenessSession_MatchDecision: true,
    });
    assert.deepEqual(harness.graphClient.calls, []);
    assert.equal(harness.faceClient.calls.length, 1);
    assert.deepEqual(harness.faceClient.calls[0].pathArguments, [FACE_RESULT, sessionId]);
    assert.deepEqual(harness.faceClient.calls[0].parameters, [sessionId]);
    harness.faceClient.assertExhausted();
});

test('FaceID_resultado exhausts five Face reads before learning_platform.read_face_liveness_result_failed', async (t) => {
    const harness = await launch(t);
    harness.faceClient.enqueue('GET', FACE_RESULT, ...fiveFailures('Face result'));

    const response = await fetch(`${harness.origin}/plataforma_v2/FaceID_resultado/session-123`);

    await assertJsonResponse(response, 500, { error: 'learning_platform.read_face_liveness_result_failed' });
    assert.equal(harness.faceClient.calls.length, 5);
    assert.ok(harness.faceClient.calls.every((call) => (
        call.method === 'GET'
        && call.path === FACE_RESULT
        && call.pathArguments[1] === 'session-123'
    )));
    assert.deepEqual(retrySleeps(harness.ledger), RETRY_DELAYS);
    harness.faceClient.assertExhausted();
});

test('refresh selects the verified row and returns all exact fields without a new handle', async (t) => {
    const rowIndex = 2;
    const harness = await launch(t);
    const handle = harness.platformRowAuthorization.createHandle(rowIndex);
    const row = makePlatformRow({
        0: 'Nome Completo',
        1: 'Nome',
        2: 'learner@example.com',
        6: 45292.5,
        7: 'Ativo',
        8: 27,
        10: 0.1,
        11: 0.2,
        12: 0.3,
        13: 0.4,
        14: 0.5,
        15: 0.6,
        16: 0.7,
        17: 0.8,
        18: 0.9,
        19: 1,
        20: 0.73,
        21: 'FMG-ABCD-EFGH',
    });
    const decoy = makePlatformRow({ 3: 'safe-password-shape' });
    harness.graphClient.enqueue('GET', PLATFORM_ROWS, {
        value: [workbookRow(decoy), workbookRow(decoy), workbookRow(row)],
    });

    const response = await postJson(harness.origin, '/plataforma_v2/refresh', {
        IndexVerificado: handle,
    });

    await assertJsonResponse(response, 200, {
        Usuário_NomeCompleto: 'Nome Completo',
        Usuário_PrimeiroNome: 'Nome',
        Usuário_Email: 'learner@example.com',
        Usuário_PrazoAcesso: '01/jan/2024',
        Usuário_Status_Login: 'Ativo',
        Usuário_Formação_NúmeroTópicosConcluídos: 27,
        Usuário_Formação_NotaMódulo1: 0.1,
        Usuário_Formação_NotaMódulo2: 0.2,
        Usuário_Formação_NotaMódulo3: 0.3,
        Usuário_Formação_NotaMódulo4: 0.4,
        Usuário_Formação_NotaMódulo5: 0.5,
        Usuário_Formação_NotaMódulo6: 0.6,
        Usuário_Formação_NotaMódulo7: 0.7,
        Usuário_Formação_NotaMódulo8: 0.8,
        Usuário_Formação_NotaMódulo9: 0.9,
        Usuário_Formação_NotaMódulo10: 1,
        Usuário_Formação_NotaAcumulado: 0.73,
        Usuário_Formação_CertificadoID: 'FMG-ABCD-EFGH',
    });
    assert.equal(harness.graphClient.calls.length, 1);
    assert.equal(harness.graphClient.calls[0].path, PLATFORM_ROWS);
    harness.graphClient.assertExhausted();
});

test('refresh exhausts five Graph reads before learning_platform.read_platform_data_failed', async (t) => {
    const harness = await launch(t);
    const handle = harness.platformRowAuthorization.createHandle(2);
    harness.graphClient.enqueue('GET', PLATFORM_ROWS, ...fiveFailures('refresh read'));

    const response = await postJson(harness.origin, '/plataforma_v2/refresh', {
        IndexVerificado: handle,
    });

    await assertJsonResponse(response, 500, { error: 'learning_platform.read_platform_data_failed' });
    assert.equal(harness.graphClient.calls.length, 5);
    assert.deepEqual(retrySleeps(harness.ledger), RETRY_DELAYS);
    harness.graphClient.assertExhausted();
});

test('updates preserves exact payloads and unvalidated JavaScript index semantics', async (t) => {
    const rowIndex = 14;
    const harness = await launch(t);
    const handle = harness.platformRowAuthorization.createHandle(rowIndex);
    harness.graphClient.enqueue('UPDATE', platformItem(rowIndex), ...Array.from({ length: 7 }, () => ({})));
    const requests = [
        {
            TipoAtualização: 'NúmeroTópicosConcluídos',
            NúmeroTópicosConcluídos: 3,
            NúmeroMódulo: 'n/a',
            NotaTeste: 'n/a',
        },
        {
            TipoAtualização: 'NúmeroTópicosConcluídos-e-NotaTeste',
            NúmeroTópicosConcluídos: 4,
            NúmeroMódulo: 1,
            NotaTeste: 0.75,
        },
        {
            TipoAtualização: 'NúmeroTópicosConcluídos-e-NotaTeste',
            NúmeroTópicosConcluídos: 8,
            NúmeroMódulo: 10,
            NotaTeste: 0.9,
        },
        {
            TipoAtualização: 'NúmeroTópicosConcluídos-e-NotaTeste',
            NúmeroTópicosConcluídos: 9,
            NúmeroMódulo: '1',
            NotaTeste: 0.1,
        },
        {
            TipoAtualização: 'NúmeroTópicosConcluídos-e-NotaTeste',
            NúmeroTópicosConcluídos: 10,
            NúmeroMódulo: 20,
            NotaTeste: 0.2,
        },
        {
            TipoAtualização: 'NúmeroTópicosConcluídos-e-NotaTeste',
            NúmeroTópicosConcluídos: 11,
            NúmeroMódulo: -10,
            NotaTeste: 0.3,
        },
        {
            TipoAtualização: 'tipo-não-reconhecido',
            NúmeroTópicosConcluídos: 12,
            NúmeroMódulo: 20,
            NotaTeste: 0.4,
        },
    ];

    for (const request of requests) {
        const response = await postJson(harness.origin, '/plataforma_v2/updates', {
            IndexVerificado: handle,
            ...request,
        });
        await assertJsonResponse(response, 200, {});
    }

    const expectedProgress = new Array(22).fill(null);
    expectedProgress[8] = 3;
    const expectedModuleOne = new Array(22).fill(null);
    expectedModuleOne[8] = 4;
    expectedModuleOne[10] = 0.75;
    const expectedModuleTen = new Array(22).fill(null);
    expectedModuleTen[8] = 8;
    expectedModuleTen[19] = 0.9;
    const expectedStringModule = new Array(22).fill(null);
    expectedStringModule[8] = 9;
    expectedStringModule[19] = 0.1;
    const expectedExpandedModule = new Array(22).fill(null);
    expectedExpandedModule[8] = 10;
    expectedExpandedModule[29] = 0.2;
    const expectedNegativeModule = new Array(22).fill(null);
    expectedNegativeModule[8] = 11;
    expectedNegativeModule[-1] = 0.3;
    const expectedIgnoredType = new Array(22).fill(null);
    expectedIgnoredType[8] = 12;
    const payloads = harness.graphClient.calls.map((call) => call.body.values[0]);

    assert.deepEqual(payloads, [
        expectedProgress,
        expectedModuleOne,
        expectedModuleTen,
        expectedStringModule,
        expectedExpandedModule,
        expectedNegativeModule,
        expectedIgnoredType,
    ]);
    assert.deepEqual(payloads.map((payload) => payload.length), [22, 22, 22, 22, 30, 22, 22]);
    assert.ok(harness.graphClient.calls.every((call) => (
        call.method === 'UPDATE' && call.path === platformItem(rowIndex)
    )));
    harness.graphClient.assertExhausted();
});

test('updates retries the same fixed payload five times before learning_platform.update_platform_data_failed', async (t) => {
    const rowIndex = 15;
    const harness = await launch(t);
    const handle = harness.platformRowAuthorization.createHandle(rowIndex);
    harness.graphClient.enqueue('UPDATE', platformItem(rowIndex), ...fiveFailures('progress update'));

    const response = await postJson(harness.origin, '/plataforma_v2/updates', {
        IndexVerificado: handle,
        TipoAtualização: 'NúmeroTópicosConcluídos-e-NotaTeste',
        NúmeroTópicosConcluídos: 9,
        NúmeroMódulo: 4,
        NotaTeste: 0.88,
    });

    await assertJsonResponse(response, 500, { error: 'learning_platform.update_platform_data_failed' });
    assert.equal(harness.graphClient.calls.length, 5);
    assert.ok(harness.graphClient.calls.every((call) => (
        call.method === 'UPDATE'
        && call.path === platformItem(rowIndex)
        && call.body.values[0].length === 22
    )));
    const [firstPayload] = harness.graphClient.calls.map((call) => call.body);
    assert.ok(harness.graphClient.calls.every((call) => (
        JSON.stringify(call.body) === JSON.stringify(firstPayload)
    )));
    assert.deepEqual(retrySleeps(harness.ledger), RETRY_DELAYS);
    harness.graphClient.assertExhausted();
});

function feedbackRequest(handle) {
    return {
        IndexVerificado: handle,
        NúmeroTópicosConcluídos: 31,
        Usuário_NomeCompleto: 'Client Supplied Name',
        Usuário_Email: 'client-supplied@example.com',
        Feedback_DataPreenchimento: '11/08/2026',
        NúmeroMódulo: 6,
        Feedback_TamanhoMódulo: 4,
        Feedback_QualidadeConteúdo: 5,
        Feedback_QualidadePlataforma: 3,
        Feedback_QualidadeMateriaisImpressos: 2,
        Feedback_Comentários: 'Comentário enviado pelo cliente',
    };
}

function expectedFeedbackProgress() {
    const values = new Array(22).fill(null);
    values[8] = 31;
    return { values: [values] };
}

function expectedFeedbackAppend() {
    return {
        values: [[
            'Client Supplied Name',
            'client-supplied@example.com',
            '11/08/2026',
            6,
            4,
            5,
            3,
            2,
            'Comentário enviado pelo cliente',
        ]],
    };
}

test('processa-feedback preserves update-before-append payloads and repeats the append', async (t) => {
    const rowIndex = 16;
    const harness = await launch(t);
    const handle = harness.platformRowAuthorization.createHandle(rowIndex);
    harness.graphClient.enqueue('UPDATE', platformItem(rowIndex), {}, {});
    harness.graphClient.enqueue('POST', FEEDBACK_ROWS, {}, {});

    for (let requestNumber = 0; requestNumber < 2; requestNumber += 1) {
        const response = await postJson(
            harness.origin,
            '/plataforma_v2/processa-feedback',
            feedbackRequest(handle),
        );
        await assertJsonResponse(response, 200, {});
    }

    assert.deepEqual(externalOrder(harness.ledger), [
        ['graph', 'UPDATE', platformItem(rowIndex)],
        ['graph', 'POST', FEEDBACK_ROWS],
        ['graph', 'UPDATE', platformItem(rowIndex)],
        ['graph', 'POST', FEEDBACK_ROWS],
    ]);
    assert.deepEqual(harness.graphClient.calls[0].body, expectedFeedbackProgress());
    assert.deepEqual(harness.graphClient.calls[1].body, expectedFeedbackAppend());
    assert.deepEqual(harness.graphClient.calls[2].body, expectedFeedbackProgress());
    assert.deepEqual(harness.graphClient.calls[3].body, expectedFeedbackAppend());
    assert.equal(harness.graphClient.calls[0].body.values[0].length, 22);
    assert.equal(harness.graphClient.calls[1].body.values[0].length, 9);
    harness.graphClient.assertExhausted();
});

test('processa-feedback preserves update and append five-attempt error boundaries', async (t) => {
    await t.test('progress exhaustion returns learning_platform.update_platform_data_failed without an append', async (t) => {
        const rowIndex = 17;
        const harness = await launch(t);
        const handle = harness.platformRowAuthorization.createHandle(rowIndex);
        harness.graphClient.enqueue('UPDATE', platformItem(rowIndex), ...fiveFailures('feedback progress'));

        const response = await postJson(
            harness.origin,
            '/plataforma_v2/processa-feedback',
            feedbackRequest(handle),
        );

        await assertJsonResponse(response, 500, { error: 'learning_platform.update_platform_data_failed' });
        assert.equal(harness.graphClient.calls.length, 5);
        assert.ok(harness.graphClient.calls.every((call) => call.method === 'UPDATE'));
        assert.deepEqual(retrySleeps(harness.ledger), RETRY_DELAYS);
        harness.graphClient.assertExhausted();
    });

    await t.test('append exhaustion returns learning_platform.append_feedback_failed after progress persists', async (t) => {
        const rowIndex = 18;
        const harness = await launch(t);
        const handle = harness.platformRowAuthorization.createHandle(rowIndex);
        harness.graphClient.enqueue('UPDATE', platformItem(rowIndex), { updated: true });
        harness.graphClient.enqueue('POST', FEEDBACK_ROWS, ...fiveFailures('feedback append'));

        const response = await postJson(
            harness.origin,
            '/plataforma_v2/processa-feedback',
            feedbackRequest(handle),
        );

        await assertJsonResponse(response, 500, { error: 'learning_platform.append_feedback_failed' });
        assert.deepEqual(externalOrder(harness.ledger).map((entry) => entry[1]), [
            'UPDATE', 'POST', 'POST', 'POST', 'POST', 'POST',
        ]);
        assert.deepEqual(harness.graphClient.calls[0].body, expectedFeedbackProgress());
        assert.ok(harness.graphClient.calls.slice(1).every((call) => (
            JSON.stringify(call.body) === JSON.stringify(expectedFeedbackAppend())
        )));
        assert.deepEqual(retrySleeps(harness.ledger), RETRY_DELAYS);
        harness.graphClient.assertExhausted();
    });
});
