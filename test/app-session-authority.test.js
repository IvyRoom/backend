'use strict';

process.env.NODE_ENV = 'test';

const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const {
    SESSION_COOKIE_NAME,
    SESSION_REQUEST_HEADER,
    SESSION_REQUEST_HEADER_VALUE,
} = require('../domains/session-authority/http');
const {
    authorityUnavailable,
    invalidAuthority,
} = require('../domains/session-authority/errors');
const {
    createTestApp,
    startLoopback,
} = require('./app-test-support');

const SYNTHETIC_TARGET_HOSTNAME = 'api.session-composition.test';
const SYNTHETIC_TARGET_ORIGIN = 'https://session-composition.test';
const PRESENTED_IDENTIFIER = Buffer.alloc(32, 7).toString('base64url');
const ISSUED_IDENTIFIER = Buffer.alloc(32, 11).toString('base64url');
const SERVER_TIME = new Date('2042-06-01T12:00:00.000Z');
const EXPIRES_AT = new Date('2042-06-01T12:01:00.000Z');

const LEGACY_ROUTE_TABLE = [
    ['post', '/landingpage/solicitacaoorcamento', ['processQuoteRequest']],
    ['post', '/conecta/processa-recomendacao', ['processConectaRecommendation']],
    ['post', '/clientes/processa-formulario', ['processClientIntake']],
    ['post', '/clientes/liberacao-acesso-plataforma', ['releasePlatformAccess']],
    ['post', '/plataforma_v2/login-FaceID', ['loginWithFaceId']],
    [
        'post',
        '/plataforma_v2/CadastroFoto_e_FaceID',
        ['multerMiddleware', 'authorize', 'registerPhotoAndFaceId'],
    ],
    ['post', '/plataforma_v2/FaceID', ['authorize', 'createFaceIdSession']],
    [
        'get',
        '/plataforma_v2/FaceID_resultado/:Azure_Face_API_LivenessSession_sessionID',
        ['getFaceIdResult'],
    ],
    ['post', '/plataforma_v2/refresh', ['authorize', 'refresh']],
    ['post', '/plataforma_v2/updates', ['authorize', 'updateProgress']],
    ['post', '/plataforma_v2/processa-feedback', ['authorize', 'processFeedback']],
    ['get', '/ezdrm-playready-authorization-url', ['getPlayReadyAuthorizationUrl']],
    ['post', '/plataforma_v2/statusreport', ['getStatusReport']],
    [
        'get',
        '/validacaocertificados/:Solicitante_CertificadoID',
        ['validateCertificate'],
    ],
];

const TARGET_ROUTE_ADDITIONS = [
    ['post', '/plataforma_v2/sessions/current/registration-enrollment'],
    ['post', '/plataforma_v2/sessions/current/face-completion'],
    ['get', '/plataforma_v2/sessions/current'],
    ['delete', '/plataforma_v2/sessions/current'],
    ['delete', '/plataforma_v2/sessions'],
];

function routeTable(app) {
    return app._router.stack.filter((layer) => layer.route).map(({ route }) => [
        Object.keys(route.methods).find((method) => route.methods[method]),
        route.path,
        route.stack.map((layer) => layer.name),
    ]);
}

function routeKey([method, path]) {
    return `${method.toUpperCase()} ${path}`;
}

function createForbiddenStore() {
    const accesses = [];
    const store = new Proxy(Object.create(null), {
        get(_target, property) {
            accesses.push(String(property));
            throw new Error(`Synthetic store must not be used: ${String(property)}`);
        },
    });
    return { accesses, store };
}

function createFakeAuthority({ controls = {}, operations = {} } = {}) {
    const calls = [];
    const runtimeControls = {
        durableStoreRequired: false,
        targetRoutesEnabled: false,
        targetSessionIssuanceEnabled: false,
        legacyLedgerSeedingEnabled: false,
        legacyCompatibilityEnforcementEnabled: false,
        subjectTargetAdoptionEnabled: false,
        protectedRoutesEnabled: false,
        ...controls,
    };
    const defaults = {
        authorizeCurrent: async () => ({
            session: { phase: 'registration-pending' },
            subjectId: 'synthetic-subject',
            platformRowIndex: 4,
        }),
        authorizeLegacy: async () => ({
            subjectId: 'synthetic-subject',
            platformRowIndex: 4,
        }),
        authorizeProtected: async () => ({
            subjectId: 'synthetic-subject',
            platformRowIndex: 4,
        }),
        completeFace: async () => ({ status: 200, body: { operation: 'complete-face' } }),
        createExistingPhotoChallenge: async () => ({
            status: 200,
            body: { operation: 'existing-photo-challenge' },
        }),
        createRegistrationChallenge: async () => ({
            status: 200,
            body: { operation: 'registration-challenge' },
        }),
        current: async () => ({ status: 200, body: { operation: 'current' } }),
        loginLegacyWithSeeding: async () => ({
            status: 200,
            body: { operation: 'legacy-login' },
        }),
        loginTarget: async () => ({ status: 200, body: { operation: 'target-login' } }),
        logout: async () => ({ status: 204 }),
        registrationEnrollment: async () => ({
            status: 200,
            body: { operation: 'registration-enrollment' },
        }),
        revokeAll: async () => ({ status: 204 }),
    };
    const authority = { runtimeControls };

    for (const [name, defaultOperation] of Object.entries(defaults)) {
        authority[name] = async (...args) => {
            calls.push({ name, args });
            const operation = operations[name] || defaultOperation;
            return operation(...args);
        };
    }

    return { authority, calls };
}

function createSyntheticGraphClient(result = { value: [] }) {
    const calls = [];
    return {
        calls,
        api(path) {
            return {
                async get() {
                    calls.push({ method: 'GET', path });
                    return result;
                },
                async post() {
                    throw new Error('Synthetic Graph POST was not expected');
                },
                async put() {
                    throw new Error('Synthetic Graph PUT was not expected');
                },
                async update() {
                    throw new Error('Synthetic Graph UPDATE was not expected');
                },
            };
        },
    };
}

function createSessionHarness(options = {}) {
    const fakeAuthority = createFakeAuthority(options.authority);
    const forbiddenStore = createForbiddenStore();
    const harness = createTestApp({
        graphClient: options.graphClient,
        platformRowAuthorization: options.platformRowAuthorization,
        dependencies: {
            sessionAuthority: {
                authority: fakeAuthority.authority,
                store: forbiddenStore.store,
                http: {
                    targetHostname: SYNTHETIC_TARGET_HOSTNAME,
                    targetOrigin: SYNTHETIC_TARGET_ORIGIN,
                },
            },
        },
    });

    return {
        ...harness,
        authority: fakeAuthority.authority,
        authorityCalls: fakeAuthority.calls,
        storeAccesses: forbiddenStore.accesses,
    };
}

async function launchSessionHarness(t, options = {}) {
    const harness = createSessionHarness(options);
    const loopback = await startLoopback(harness.app, t);
    return { ...harness, ...loopback };
}

function targetHeaders(headers = {}) {
    return Object.fromEntries(Object.entries({
        Host: SYNTHETIC_TARGET_HOSTNAME,
        Origin: SYNTHETIC_TARGET_ORIGIN,
        [SESSION_REQUEST_HEADER]: SESSION_REQUEST_HEADER_VALUE,
        ...headers,
    }).filter(([, value]) => value !== undefined));
}

function targetCookieHeaders(headers = {}) {
    return targetHeaders({
        Cookie: `${SESSION_COOKIE_NAME}=${PRESENTED_IDENTIFIER}`,
        ...headers,
    });
}

function requestLoopback(origin, path, {
    method = 'GET',
    headers = {},
    body,
} = {}) {
    const url = new URL(path, origin);
    return new Promise((resolve, reject) => {
        const request = http.request(url, {
            method,
            headers: { Connection: 'close', ...headers },
        }, (response) => {
            const chunks = [];
            response.on('data', (chunk) => chunks.push(chunk));
            response.on('end', () => {
                const responseBody = Buffer.concat(chunks);
                resolve({
                    status: response.statusCode,
                    headers: {
                        get(name) {
                            const value = response.headers[String(name).toLowerCase()];
                            if (Array.isArray(value)) return value.join(', ');
                            return value === undefined ? null : value;
                        },
                    },
                    async json() {
                        return JSON.parse(responseBody.toString('utf8'));
                    },
                    async text() {
                        return responseBody.toString('utf8');
                    },
                });
            });
        });
        request.on('error', reject);
        if (body !== undefined) request.write(body);
        request.end();
    });
}

async function readJson(response, status) {
    assert.equal(response.status, status);
    assert.equal(response.headers.get('content-type'), 'application/json; charset=utf-8');
    return response.json();
}

function assertTargetEnvelope(response) {
    assert.equal(response.headers.get('access-control-allow-origin'), SYNTHETIC_TARGET_ORIGIN);
    assert.equal(response.headers.get('access-control-allow-credentials'), 'true');
    assert.equal(response.headers.get('vary'), 'Origin, Cookie');
    assert.equal(response.headers.get('cache-control'), 'no-store');
    assert.equal(response.headers.get('pragma'), 'no-cache');
    assert.equal(response.headers.get('expires'), '0');
    assert.equal(response.headers.get('referrer-policy'), 'no-referrer');
    assert.equal(response.headers.get('etag'), null);
    assert.equal(response.headers.get('last-modified'), null);
}

test('no session configuration preserves all 14 legacy registrations and stacks', () => {
    const harness = createTestApp();

    assert.equal(routeTable(harness.app).length, 14);
    assert.deepEqual(routeTable(harness.app), LEGACY_ROUTE_TABLE);
    assert.equal(harness.graphClient.calls.length, 0);
    assert.equal(harness.faceClient.calls.length, 0);
});

test('qualified target composition adds only the five session registrations', () => {
    const harness = createSessionHarness({
        authority: {
            controls: {
                durableStoreRequired: true,
                targetRoutesEnabled: true,
                targetSessionIssuanceEnabled: true,
                legacyLedgerSeedingEnabled: true,
                legacyCompatibilityEnforcementEnabled: true,
                subjectTargetAdoptionEnabled: true,
                protectedRoutesEnabled: true,
            },
        },
    });
    const routes = routeTable(harness.app);
    const legacyKeys = new Set(LEGACY_ROUTE_TABLE.map(routeKey));
    const additions = routes.filter((route) => !legacyKeys.has(routeKey(route)));

    assert.equal(routes.length, 19);
    assert.deepEqual(additions.map(([method, path]) => [method, path]), TARGET_ROUTE_ADDITIONS);

    const routesByKey = new Map(routes.map((route) => [routeKey(route), route]));
    assert.deepEqual(
        routesByKey.get('POST /plataforma_v2/CadastroFoto_e_FaceID')[2],
        [
            'targetOnlyMiddleware',
            'modeAwareMiddleware',
            'modeAwareMiddleware',
            'registerPhotoAndFaceId',
        ],
    );
    assert.deepEqual(
        routesByKey.get('POST /plataforma_v2/FaceID')[2],
        ['modeAwareMiddleware', 'createFaceIdSession'],
    );
    for (const path of [
        '/plataforma_v2/refresh',
        '/plataforma_v2/updates',
        '/plataforma_v2/processa-feedback',
    ]) {
        assert.equal(routesByKey.get(`POST ${path}`)[2][0], 'authorizeProtectedLearning');
    }

    assert.deepEqual(harness.authorityCalls, []);
    assert.deepEqual(harness.storeAccesses, []);
    assert.deepEqual(harness.graphClient.calls, []);
    assert.deepEqual(harness.faceClient.calls, []);
});

test('durable-store latch keeps 14 routes while central authority owns legacy dispatch', async (t) => {
    const legacyAuthorization = { calls: 0 };
    const harness = await launchSessionHarness(t, {
        platformRowAuthorization: {
            authorize(_req, res) {
                legacyAuthorization.calls += 1;
                return res.status(418).json({ mode: 'unchecked-legacy' });
            },
            createHandle() {
                throw new Error('Unchecked legacy issuance was not expected');
            },
            inspectHandle() {
                throw new Error('Direct legacy inspection was not expected');
            },
        },
        authority: { controls: { durableStoreRequired: true } },
    });

    assert.equal(routeTable(harness.app).length, 14);

    const login = await fetch(`${harness.origin}/plataforma_v2/login-FaceID`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            Usuário_Login: 'synthetic-latched-user',
            Usuário_Senha: 'synthetic-latched-credential',
        }),
    });
    assert.deepEqual(await readJson(login, 200), { operation: 'legacy-login' });
    assert.equal(login.headers.get('cache-control'), null);
    assert.notEqual(login.headers.get('etag'), null);
    assert.deepEqual(harness.authorityCalls.map(({ name }) => name), ['loginLegacyWithSeeding']);
    assert.equal(harness.graphClient.calls.length, 0);

    const cookieAttempt = await requestLoopback(harness.origin, '/plataforma_v2/refresh', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            Cookie: `${SESSION_COOKIE_NAME}=${PRESENTED_IDENTIFIER}`,
        },
        body: '{}',
    });
    assert.deepEqual(await readJson(cookieAttempt, 401), {});
    assert.equal(legacyAuthorization.calls, 0);
    assert.deepEqual(harness.authorityCalls.map(({ name }) => name), ['loginLegacyWithSeeding']);
});

test('central legacy login preserves the established workbook-read failure envelope', async (t) => {
    const harness = await launchSessionHarness(t, {
        authority: {
            controls: {
                durableStoreRequired: true,
                legacyLedgerSeedingEnabled: true,
            },
            operations: {
                loginLegacyWithSeeding: async () => {
                    throw authorityUnavailable('legacy-platform-data-read-failed');
                },
            },
        },
    });

    const response = await fetch(`${harness.origin}/plataforma_v2/login-FaceID`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            Usuário_Login: 'synthetic-legacy-user',
            Usuário_Senha: 'synthetic-legacy-credential',
        }),
    });

    assert.deepEqual(await readJson(response, 500), {
        error: 'learning_platform.read_platform_data_failed',
    });
    assert.equal(response.headers.get('access-control-allow-origin'), '*');
    assert.equal(response.headers.get('cache-control'), null);
    assert.equal(response.headers.get('set-cookie'), null);
    assert.deepEqual(
        harness.authorityCalls.map(({ name }) => name),
        ['loginLegacyWithSeeding'],
    );
});

test('dormant target routing neither registers session routes nor parses activation attempts', async (t) => {
    const harness = await launchSessionHarness(t);
    assert.equal(routeTable(harness.app).length, 14);
    assert.deepEqual(
        routeTable(harness.app).map(([method, path]) => [method, path]),
        LEGACY_ROUTE_TABLE.map(([method, path]) => [method, path]),
    );

    const blocked = await fetch(`${harness.origin}/plataforma_v2/login-FaceID`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            [SESSION_REQUEST_HEADER]: SESSION_REQUEST_HEADER_VALUE,
        },
        body: '{',
    });
    assert.deepEqual(await readJson(blocked, 503), {});
    assert.equal(blocked.headers.get('cache-control'), 'no-store');

    const absent = await fetch(`${harness.origin}/plataforma_v2/sessions/current`, {
        headers: { Host: SYNTHETIC_TARGET_HOSTNAME },
    });
    assert.equal(absent.status, 404);
    assert.equal(absent.headers.get('access-control-allow-origin'), '*');
    assert.equal(absent.headers.get('cache-control'), null);
    await absent.text();

    assert.deepEqual(harness.authorityCalls, []);
    assert.deepEqual(harness.storeAccesses, []);
    assert.deepEqual(harness.graphClient.calls, []);
    assert.deepEqual(harness.faceClient.calls, []);
});

test('dual-mode login keeps legacy dispatch separate from target cookie transport', async (t) => {
    const graphClient = createSyntheticGraphClient();
    const harness = await launchSessionHarness(t, {
        graphClient,
        authority: {
            controls: { targetRoutesEnabled: true, protectedRoutesEnabled: true },
            operations: {
                loginTarget: async () => ({
                    status: 200,
                    body: { mode: 'target' },
                    issuance: {
                        identifier: ISSUED_IDENTIFIER,
                        expiresAt: EXPIRES_AT,
                        serverTime: SERVER_TIME,
                    },
                }),
            },
        },
    });

    const legacy = await fetch(`${harness.origin}/plataforma_v2/login-FaceID`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            Usuário_Login: 'legacy.user@example.test',
            Usuário_Senha: 'invented-legacy-password',
        }),
    });
    assert.deepEqual(await readJson(legacy, 401), { error: 'credenciais_inválidas' });
    assert.equal(legacy.headers.get('access-control-allow-origin'), '*');
    assert.equal(legacy.headers.get('access-control-allow-credentials'), null);
    assert.equal(graphClient.calls.length, 1);
    assert.deepEqual(harness.authorityCalls, []);

    const target = await requestLoopback(harness.origin, '/plataforma_v2/login-FaceID', {
        method: 'POST',
        headers: targetHeaders({ 'Content-Type': 'application/json' }),
        body: JSON.stringify({
            Usuário_Login: 'target.user@example.test',
            Usuário_Senha: 'invented-target-password',
        }),
    });
    assert.deepEqual(await readJson(target, 200), { mode: 'target' });
    assertTargetEnvelope(target);
    assert.equal(
        target.headers.get('set-cookie'),
        `${SESSION_COOKIE_NAME}=${ISSUED_IDENTIFIER}; Path=/; Secure; HttpOnly; SameSite=Strict; Max-Age=60; Expires=${EXPIRES_AT.toUTCString()}`,
    );
    assert.equal(graphClient.calls.length, 1);
    assert.deepEqual(harness.authorityCalls.map(({ name }) => name), ['loginTarget']);
    assert.deepEqual(harness.authorityCalls[0].args, [{
        login: 'target.user@example.test',
        password: 'invented-target-password',
        presentedIdentifier: null,
    }]);

    const wrongOrigin = await requestLoopback(harness.origin, '/plataforma_v2/login-FaceID', {
        method: 'POST',
        headers: targetHeaders({
            Origin: 'https://wrong-origin.example.test',
            'Content-Type': 'application/json',
        }),
        body: '{}',
    });
    assert.deepEqual(await readJson(wrongOrigin, 403), {});
    assertTargetEnvelope(wrongOrigin);
    assert.equal(wrongOrigin.headers.get('set-cookie'), null);
    assert.deepEqual(harness.authorityCalls.map(({ name }) => name), ['loginTarget']);

    const missingCookieHeader = await requestLoopback(
        harness.origin,
        '/plataforma_v2/sessions/current',
        { headers: { Host: SYNTHETIC_TARGET_HOSTNAME, Cookie: `${SESSION_COOKIE_NAME}=${PRESENTED_IDENTIFIER}` } },
    );
    assert.deepEqual(await readJson(missingCookieHeader, 403), {});
    assertTargetEnvelope(missingCookieHeader);

    const current = await requestLoopback(harness.origin, '/plataforma_v2/sessions/current', {
        headers: targetCookieHeaders({ Origin: undefined }),
    });
    assert.deepEqual(await readJson(current, 200), { operation: 'current' });
    assertTargetEnvelope(current);
    assert.deepEqual(harness.authorityCalls.map(({ name }) => name), ['loginTarget', 'current']);
    assert.deepEqual(harness.authorityCalls[1].args, [PRESENTED_IDENTIFIER]);

    assert.deepEqual(harness.storeAccesses, []);
    assert.deepEqual(harness.faceClient.calls, []);
});

test('current logout failures never mutate the cookie while missing or malformed cookies remain idempotent', async (t) => {
    for (const reason of ['target-routes-disabled', 'session-store-unavailable']) {
        await t.test(reason, async (subtest) => {
            const harness = await launchSessionHarness(subtest, {
                authority: {
                    controls: { targetRoutesEnabled: true },
                    operations: {
                        logout: async (identifier) => {
                            if (identifier === undefined) return { status: 204 };
                            throw authorityUnavailable(reason);
                        },
                    },
                },
            });

            const failed = await requestLoopback(
                harness.origin,
                '/plataforma_v2/sessions/current',
                {
                    method: 'DELETE',
                    headers: targetCookieHeaders(),
                },
            );
            assert.deepEqual(await readJson(failed, 503), {});
            assertTargetEnvelope(failed);
            assert.equal(failed.headers.get('set-cookie'), null);

            for (const cookieHeader of [undefined, `${SESSION_COOKIE_NAME}=malformed`]) {
                const response = await requestLoopback(
                    harness.origin,
                    '/plataforma_v2/sessions/current',
                    {
                        method: 'DELETE',
                        headers: targetHeaders({ Cookie: cookieHeader }),
                    },
                );
                assert.equal(response.status, 204);
                assert.equal(await response.text(), '');
                assertTargetEnvelope(response);
                assert.equal(response.headers.get('set-cookie'), null);
            }

            assert.deepEqual(
                harness.authorityCalls.filter(({ name }) => name === 'logout').map(({ args }) => args),
                [[PRESENTED_IDENTIFIER], [undefined], [undefined]],
            );
            assert.deepEqual(harness.storeAccesses, []);
        });
    }
});

test('target registration authorization runs before Multer can inspect the upload', async (t) => {
    const harness = await launchSessionHarness(t, {
        authority: {
            controls: { targetRoutesEnabled: true },
            operations: {
                authorizeCurrent: async () => {
                    throw invalidAuthority('synthetic-preauthorization-rejection');
                },
            },
        },
    });
    const multipartBoundary = 'synthetic-session-composition-boundary';
    const multipartBody = Buffer.from([
        `--${multipartBoundary}`,
        'Content-Disposition: form-data; name="unexpected-file-field"; filename="synthetic-reference.jpg"',
        'Content-Type: image/jpeg',
        '',
        'invented reference photo',
        `--${multipartBoundary}--`,
        '',
    ].join('\r\n'));

    const response = await requestLoopback(
        harness.origin,
        '/plataforma_v2/CadastroFoto_e_FaceID',
        {
            method: 'POST',
            headers: targetCookieHeaders({
                'Content-Type': `multipart/form-data; boundary=${multipartBoundary}`,
                'Content-Length': multipartBody.length,
            }),
            body: multipartBody,
        },
    );

    assert.deepEqual(await readJson(response, 401), {});
    assertTargetEnvelope(response);
    assert.deepEqual(harness.authorityCalls.map(({ name }) => name), ['authorizeCurrent']);
    assert.deepEqual(harness.authorityCalls[0].args, [
        PRESENTED_IDENTIFIER,
        {
            allowedPhases: ['registration-pending', 'face-pending'],
            revalidate: true,
        },
    ]);
    assert.deepEqual(harness.storeAccesses, []);
    assert.deepEqual(harness.graphClient.calls, []);
    assert.deepEqual(harness.faceClient.calls, []);
});

test('an active registration Face challenge returns 409 before Multer parses a retry', async (t) => {
    const harness = await launchSessionHarness(t, {
        authority: {
            controls: { targetRoutesEnabled: true },
            operations: {
                authorizeCurrent: async () => ({ session: { phase: 'face-pending' } }),
            },
        },
    });
    const multipartBoundary = 'synthetic-active-face-boundary';
    const multipartBody = Buffer.from([
        `--${multipartBoundary}`,
        'Content-Disposition: form-data; name="unexpected-file-field"; filename="synthetic-reference.jpg"',
        'Content-Type: image/jpeg',
        '',
        'invented reference photo',
        `--${multipartBoundary}--`,
        '',
    ].join('\r\n'));

    const response = await requestLoopback(
        harness.origin,
        '/plataforma_v2/CadastroFoto_e_FaceID',
        {
            method: 'POST',
            headers: targetCookieHeaders({
                'Content-Type': `multipart/form-data; boundary=${multipartBoundary}`,
                'Content-Length': multipartBody.length,
            }),
            body: multipartBody,
        },
    );

    assert.deepEqual(await readJson(response, 409), {});
    assertTargetEnvelope(response);
    assert.deepEqual(harness.authorityCalls.map(({ name }) => name), ['authorizeCurrent']);
    assert.deepEqual(harness.storeAccesses, []);
    assert.deepEqual(harness.graphClient.calls, []);
    assert.deepEqual(harness.faceClient.calls, []);
});

test('protected learning middleware follows its independent target-mode gate', async (t) => {
    function rejectingLegacyAuthorization(counter) {
        return {
            authorize(_req, res) {
                counter.calls += 1;
                return res.status(418).json({ mode: 'legacy-authorization' });
            },
            createHandle() {
                throw new Error('Legacy handle issuance was not expected');
            },
            inspectHandle() {
                throw new Error('Legacy handle inspection was not expected');
            },
        };
    }

    const dormantLegacyAuthorization = { calls: 0 };
    const dormant = await launchSessionHarness(t, {
        platformRowAuthorization: rejectingLegacyAuthorization(dormantLegacyAuthorization),
        authority: {
            controls: { targetRoutesEnabled: true, protectedRoutesEnabled: false },
            operations: {
                authorizeProtected: async () => {
                    throw new Error('Dormant protected authority must not run');
                },
            },
        },
    });
    const legacyMode = await requestLoopback(dormant.origin, '/plataforma_v2/refresh', {
        method: 'POST',
        headers: targetCookieHeaders({ 'Content-Type': 'application/json' }),
        body: '{}',
    });
    assert.deepEqual(await readJson(legacyMode, 418), { mode: 'legacy-authorization' });
    assert.equal(legacyMode.headers.get('access-control-allow-origin'), '*');
    assert.equal(dormantLegacyAuthorization.calls, 1);
    assert.deepEqual(dormant.authorityCalls, []);

    const activeLegacyAuthorization = { calls: 0 };
    const active = await launchSessionHarness(t, {
        platformRowAuthorization: rejectingLegacyAuthorization(activeLegacyAuthorization),
        authority: {
            controls: { targetRoutesEnabled: true, protectedRoutesEnabled: true },
            operations: {
                authorizeProtected: async () => {
                    throw invalidAuthority('synthetic-protected-rejection');
                },
            },
        },
    });
    const targetMode = await requestLoopback(active.origin, '/plataforma_v2/refresh', {
        method: 'POST',
        headers: targetCookieHeaders({ 'Content-Type': 'application/json' }),
        body: JSON.stringify({ IndexVerificado: 'invented-legacy-handle' }),
    });
    assert.deepEqual(await readJson(targetMode, 401), {});
    assertTargetEnvelope(targetMode);
    assert.equal(activeLegacyAuthorization.calls, 0);
    assert.deepEqual(active.authorityCalls.map(({ name }) => name), ['authorizeProtected']);
    assert.deepEqual(active.authorityCalls[0].args, [PRESENTED_IDENTIFIER]);

    for (const harness of [dormant, active]) {
        assert.deepEqual(harness.storeAccesses, []);
        assert.deepEqual(harness.graphClient.calls, []);
        assert.deepEqual(harness.faceClient.calls, []);
    }
});
