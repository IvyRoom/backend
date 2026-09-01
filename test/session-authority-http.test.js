'use strict';

process.env.NODE_ENV = 'test';

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const {
    ACTUAL_VARY,
    ALLOWED_REQUEST_HEADERS,
    PREFLIGHT_VARY,
    SESSION_COOKIE_NAME,
    SESSION_COOKIE_STATES,
    SESSION_REQUEST_HEADER,
    SESSION_REQUEST_HEADER_VALUE,
    TARGET_ORIGIN,
    TARGET_REQUEST_KINDS,
    applySessionResponseHardening,
    createTargetHttpBoundary,
    createTargetRequestClassifier,
    formatSessionIssuanceCookie,
    parseSessionCookieHeader,
    readSessionCookie,
} = require('../domains/session-authority/http');
const {
    SESSION_API_HOSTNAME,
    SESSION_FRONTEND_ORIGIN,
} = require('../domains/session-authority/constants');

const SYNTHETIC_TARGET_HOSTNAME = 'api.session-authority.test';
const SYNTHETIC_TARGET_ORIGIN = 'https://session-authority.test';

function createIdentifier() {
    return crypto.randomBytes(32).toString('base64url');
}

function digest(value) {
    return crypto.createHash('sha256').update(value).digest('hex');
}

function request({
    method = 'GET',
    path = '/',
    headers = {},
} = {}) {
    return {
        method,
        path,
        headers: Object.fromEntries(
            Object.entries(headers).map(([name, value]) => [name.toLowerCase(), value]),
        ),
    };
}

function createResponse(initialHeaders = {}) {
    const headers = new Map();

    const res = {
        body: undefined,
        ended: false,
        locals: {},
        statusCode: 200,
        setHeader(name, value) {
            headers.set(String(name).toLowerCase(), value);
            return this;
        },
        getHeader(name) {
            return headers.get(String(name).toLowerCase());
        },
        removeHeader(name) {
            headers.delete(String(name).toLowerCase());
        },
        status(statusCode) {
            this.statusCode = statusCode;
            return this;
        },
        json(body) {
            this.body = body;
            this.setHeader('Content-Type', 'application/json; charset=utf-8');
            this.setHeader('ETag', 'synthetic-validator');
            this.ended = true;
            return this;
        },
        end() {
            this.ended = true;
            return this;
        },
        writeHead(statusCode, statusMessage, suppliedHeaders) {
            this.statusCode = statusCode;
            const headerObject = typeof statusMessage === 'string'
                ? suppliedHeaders
                : statusMessage;
            if (Array.isArray(headerObject)) {
                for (let index = 0; index < headerObject.length; index += 2) {
                    this.setHeader(headerObject[index], headerObject[index + 1]);
                }
            } else {
                for (const [name, value] of Object.entries(headerObject || {})) {
                    this.setHeader(name, value);
                }
            }
            return this;
        },
    };

    for (const [name, value] of Object.entries(initialHeaders)) res.setHeader(name, value);
    return res;
}

function assertHardened(res) {
    assert.equal(res.getHeader('Cache-Control'), 'no-store');
    assert.equal(res.getHeader('Pragma'), 'no-cache');
    assert.equal(res.getHeader('Expires'), '0');
    assert.equal(res.getHeader('Referrer-Policy'), 'no-referrer');
    assert.equal(res.getHeader('ETag'), undefined);
    assert.equal(res.getHeader('Last-Modified'), undefined);
    assert.equal(res.getHeader('Set-Cookie'), undefined);
}

function invokeBoundary(boundary, req, res = createResponse()) {
    let nextCalls = 0;
    boundary(req, res, () => {
        nextCalls += 1;
    });
    return { res, nextCalls };
}

test('cookie transport uses the exact partitioned host, names, and origin', () => {
    assert.equal(SESSION_COOKIE_NAME, '__Host-machado-session');
    assert.equal(SESSION_REQUEST_HEADER, 'x-machado-session-request');
    assert.equal(SESSION_REQUEST_HEADER_VALUE, '1');
    assert.equal(TARGET_ORIGIN, SESSION_FRONTEND_ORIGIN);
    assert.equal(TARGET_ORIGIN, 'https://machadogestao.com');
    assert.equal(SESSION_API_HOSTNAME, 'plataforma-backend-v3.azurewebsites.net');
});

test('session cookie parsing returns only canonical generated identifiers', async (t) => {
    const identifier = createIdentifier();

    await t.test('missing and unrelated cookies remain absent', () => {
        assert.deepEqual(parseSessionCookieHeader(undefined), {
            state: SESSION_COOKIE_STATES.missing,
        });
        assert.deepEqual(parseSessionCookieHeader('theme=dark; locale=pt-BR'), {
            state: SESSION_COOKIE_STATES.missing,
        });
    });

    await t.test('the exact cookie is read without percent decoding', () => {
        const parsed = parseSessionCookieHeader(
            `theme=dark; ${SESSION_COOKIE_NAME}=${identifier}; locale=pt-BR`,
        );
        assert.equal(parsed.state, SESSION_COOKIE_STATES.present);
        assert.equal(digest(parsed.value), digest(identifier));

        const fromRequest = readSessionCookie(request({
            headers: { Cookie: `${SESSION_COOKIE_NAME}=${identifier}` },
        }));
        assert.equal(fromRequest.state, SESSION_COOKIE_STATES.present);
        assert.equal(digest(fromRequest.value), digest(identifier));
    });

    await t.test('duplicates, arrays, and noncanonical values fail closed', () => {
        const malformedHeaders = [
            `${SESSION_COOKIE_NAME}=not-valid`,
            `${SESSION_COOKIE_NAME}=%${identifier.charCodeAt(0).toString(16)}${identifier.slice(1)}`,
            `${SESSION_COOKIE_NAME} =${identifier}`,
            `${SESSION_COOKIE_NAME}=${identifier}; ${SESSION_COOKIE_NAME}=${identifier}`,
            [SESSION_COOKIE_NAME, identifier],
        ];

        for (const header of malformedHeaders) {
            const parsed = parseSessionCookieHeader(header);
            assert.equal(parsed.state, SESSION_COOKIE_STATES.malformed);
            assert.equal(Object.hasOwn(parsed, 'value'), false);
        }
    });

    await t.test('cookie names are case-sensitive', () => {
        const parsed = parseSessionCookieHeader(
            `${SESSION_COOKIE_NAME.toLowerCase()}=${identifier}`,
        );
        assert.equal(parsed.state, SESSION_COOKIE_STATES.missing);
    });
});

test('issuance cookies have host-only partitioned flags and authoritative cleanup hints', () => {
    const identifier = createIdentifier();
    const now = new Date('2030-01-02T03:04:05.250Z');
    const expiresAt = new Date('2030-01-02T07:04:05.750Z');
    const cookie = formatSessionIssuanceCookie({ identifier, expiresAt, now });
    const segments = cookie.split('; ');
    const identifierPrefix = `${SESSION_COOKIE_NAME}=`;

    assert.equal(segments[0].startsWith(identifierPrefix), true);
    assert.equal(digest(segments[0].slice(identifierPrefix.length)), digest(identifier));
    assert.deepEqual(segments.slice(1), [
        'Path=/',
        'Secure',
        'HttpOnly',
        'SameSite=None',
        'Partitioned',
        'Max-Age=14401',
        `Expires=${expiresAt.toUTCString()}`,
    ]);
    assert.equal(segments.filter((segment) => segment === 'Partitioned').length, 1);
    assert.equal(cookie.includes('Domain='), false);
    assert.equal(cookie.includes('Max-Age=0'), false);

    assert.throws(
        () => formatSessionIssuanceCookie({ identifier, expiresAt: now, now }),
        /future authoritative expiry/u,
    );
    assert.throws(
        () => formatSessionIssuanceCookie({ identifier: 'not-valid', expiresAt, now }),
        /canonical transport encoding/u,
    );
});

test('session hardening removes and prevents response validators', () => {
    const res = createResponse({
        ETag: 'old-validator',
        'Last-Modified': 'Wed, 01 Jan 2030 00:00:00 GMT',
    });

    assert.equal(applySessionResponseHardening(res), res);
    res.setHeader('ETag', 'late-validator');
    res.setHeader('Last-Modified', 'Wed, 02 Jan 2030 00:00:00 GMT');
    res.writeHead(200, {
        ETag: 'write-head-validator',
        'Last-Modified': 'Wed, 03 Jan 2030 00:00:00 GMT',
        'X-Synthetic': 'preserved',
    });
    res.writeHead(200, [
        'ETag',
        'array-validator',
        'Last-Modified',
        'Wed, 04 Jan 2030 00:00:00 GMT',
        'X-Synthetic-Array',
        'preserved',
    ]);

    assertHardened(res);
    assert.equal(res.getHeader('X-Synthetic'), 'preserved');
    assert.equal(res.getHeader('X-Synthetic-Array'), 'preserved');
    assert.equal(applySessionResponseHardening(res), res);
});

test('target classifier respects disabled gates and dual-mode intent', async (t) => {
    const disabled = createTargetRequestClassifier();
    const enabled = createTargetRequestClassifier({ targetRoutesEnabled: true });
    const identifier = createIdentifier();

    await t.test('new routes remain ordinary unmatched requests while disabled', () => {
        const classification = disabled(request({
            method: 'GET',
            path: '/plataforma_v2/sessions/current',
        }));
        assert.equal(classification.kind, TARGET_REQUEST_KINDS.passThrough);
        assert.equal(classification.isTarget, false);
        assert.equal(classification.blocked, false);
    });

    await t.test('a header-only dual-mode activation attempt is blocked', () => {
        const classification = disabled(request({
            method: 'POST',
            path: '/plataforma_v2/login-FaceID',
            headers: { 'X-Machado-Session-Request': '1' },
        }));
        assert.equal(classification.kind, TARGET_REQUEST_KINDS.blockedTarget);
        assert.equal(classification.reason, 'target-routes-disabled');
    });

    await t.test('legacy dual-mode calls remain pass-through', () => {
        const classification = enabled(request({
            method: 'POST',
            path: '/plataforma_v2/login-FaceID',
        }));
        assert.equal(classification.kind, TARGET_REQUEST_KINDS.passThrough);
    });

    await t.test('header, target cookie, and custom-header preflight select target mode', () => {
        const cases = [
            request({
                method: 'POST',
                path: '/plataforma_v2/login-FaceID',
                headers: { 'X-Machado-Session-Request': '1' },
            }),
            request({
                method: 'POST',
                path: '/plataforma_v2/CadastroFoto_e_FaceID',
                headers: { Cookie: `${SESSION_COOKIE_NAME}=${identifier}` },
            }),
            request({
                method: 'OPTIONS',
                path: '/plataforma_v2/FaceID',
                headers: {
                    'Access-Control-Request-Method': 'POST',
                    'Access-Control-Request-Headers': 'content-type, X-Machado-Session-Request',
                },
            }),
        ];

        for (const req of cases) {
            const classification = enabled(req);
            assert.equal(classification.kind, TARGET_REQUEST_KINDS.target);
            assert.equal(classification.routeType, 'dual-mode');
            assert.equal(classification.methodAllowed, true);
            assert.equal(Object.hasOwn(classification, 'cookie'), false);
        }
    });

    await t.test('all five registrations are exact target method/path roles', () => {
        const cases = [
            ['POST', '/plataforma_v2/sessions/current/registration-enrollment'],
            ['POST', '/plataforma_v2/sessions/current/face-completion'],
            ['GET', '/plataforma_v2/sessions/current'],
            ['DELETE', '/plataforma_v2/sessions/current'],
            ['DELETE', '/plataforma_v2/sessions'],
        ];

        for (const [method, path] of cases) {
            const classification = enabled(request({ method, path }));
            assert.equal(classification.kind, TARGET_REQUEST_KINDS.target);
            assert.equal(classification.routeType, 'session');
            assert.equal(classification.methodAllowed, true);
        }

        const wrongMethod = enabled(request({
            method: 'PATCH',
            path: '/plataforma_v2/sessions/current',
        }));
        assert.equal(wrongMethod.kind, TARGET_REQUEST_KINDS.target);
        assert.equal(wrongMethod.methodAllowed, false);
    });

    await t.test('accepted Express path variants cannot bypass target transport checks', () => {
        const classification = enabled(request({
            method: 'get',
            path: '/PLATAFORMA_V2/SESSIONS/CURRENT/?ignored=true',
        }));
        assert.equal(classification.kind, TARGET_REQUEST_KINDS.target);
        assert.equal(classification.methodAllowed, true);
    });
});

test('protected classification is separately gated and configurable', () => {
    const defaultDisabled = createTargetRequestClassifier({ targetRoutesEnabled: true });
    assert.equal(defaultDisabled(request({
        method: 'POST',
        path: '/plataforma_v2/refresh',
    })).kind, TARGET_REQUEST_KINDS.passThrough);

    const enabled = createTargetRequestClassifier({
        protectedRoutesEnabled: true,
        protectedRoutes: [{ method: 'PATCH', path: '/synthetic/protected-operation' }],
    });

    assert.equal(enabled(request({
        method: 'PATCH',
        path: '/synthetic/protected-operation',
    })).kind, TARGET_REQUEST_KINDS.passThrough);

    const classification = enabled(request({
        method: 'PATCH',
        path: '/synthetic/protected-operation',
        headers: { 'X-Machado-Session-Request': '1' },
    }));
    assert.equal(classification.kind, TARGET_REQUEST_KINDS.target);
    assert.equal(classification.routeType, 'protected');
    assert.deepEqual(classification.allowedMethods, ['PATCH']);

    assert.throws(
        () => createTargetRequestClassifier({ targetRoutesEnabled: 'true' }),
        /must be a boolean/u,
    );
});

test('target boundary passes public and legacy requests through without mutation', () => {
    const boundary = createTargetHttpBoundary({
        targetRoutesEnabled: true,
        targetHostname: SYNTHETIC_TARGET_HOSTNAME,
        targetOrigin: SYNTHETIC_TARGET_ORIGIN,
    });
    const res = createResponse({ ETag: 'legacy-validator' });
    const result = invokeBoundary(boundary, request({
        method: 'POST',
        path: '/landingpage/solicitacaoorcamento',
    }), res);

    assert.equal(result.nextCalls, 1);
    assert.equal(res.getHeader('ETag'), 'legacy-validator');
    assert.equal(res.getHeader('Cache-Control'), undefined);
    assert.deepEqual(res.locals, {});
});

test('disabled dual-mode attempts return hardened 503 before body or domain work', () => {
    const boundary = createTargetHttpBoundary({
        targetHostname: SYNTHETIC_TARGET_HOSTNAME,
        targetOrigin: SYNTHETIC_TARGET_ORIGIN,
    });
    const result = invokeBoundary(boundary, request({
        method: 'POST',
        path: '/plataforma_v2/login-FaceID',
        headers: { 'X-Machado-Session-Request': '1' },
    }));

    assert.equal(result.nextCalls, 0);
    assert.equal(result.res.statusCode, 503);
    assert.deepEqual(result.res.body, {});
    assertHardened(result.res);
    assert.equal(result.res.getHeader('Access-Control-Allow-Origin'), undefined);
});

test('target actual responses use exact CORS, cache, Origin, Host, and header boundaries', async (t) => {
    const boundary = createTargetHttpBoundary({
        targetRoutesEnabled: true,
        targetHostname: SYNTHETIC_TARGET_HOSTNAME,
        targetOrigin: SYNTHETIC_TARGET_ORIGIN,
    });
    const identifier = createIdentifier();

    await t.test('safe cookie-free status reaches authority with hardened exact headers', () => {
        const result = invokeBoundary(boundary, request({
            method: 'GET',
            path: '/plataforma_v2/sessions/current',
            headers: { Host: SYNTHETIC_TARGET_HOSTNAME },
        }));

        assert.equal(result.nextCalls, 1);
        assert.equal(result.res.getHeader('Access-Control-Allow-Origin'), SYNTHETIC_TARGET_ORIGIN);
        assert.equal(result.res.getHeader('Access-Control-Allow-Credentials'), 'true');
        assert.equal(result.res.getHeader('Vary'), ACTUAL_VARY);
        assertHardened(result.res);
        assert.equal(result.res.locals.sessionAuthorityTransport.routeType, 'session');
    });

    await t.test('unsafe exact-origin target requests reach authority', () => {
        const result = invokeBoundary(boundary, request({
            method: 'POST',
            path: '/plataforma_v2/sessions/current/registration-enrollment',
            headers: {
                Host: SYNTHETIC_TARGET_HOSTNAME,
                Origin: SYNTHETIC_TARGET_ORIGIN,
            },
        }));
        assert.equal(result.nextCalls, 1);
        assert.equal(result.res.ended, false);
    });

    await t.test('a known target path with the wrong method is hardened and rejected', () => {
        const result = invokeBoundary(boundary, request({
            method: 'PATCH',
            path: '/plataforma_v2/sessions/current',
            headers: { Host: SYNTHETIC_TARGET_HOSTNAME },
        }));
        assert.equal(result.nextCalls, 0);
        assert.equal(result.res.statusCode, 403);
        assertHardened(result.res);
    });

    await t.test('missing, null, and wrong unsafe origins reject before authority', () => {
        for (const origin of [undefined, 'null', 'https://wrong-origin.test']) {
            const headers = { Host: SYNTHETIC_TARGET_HOSTNAME };
            if (origin !== undefined) headers.Origin = origin;
            const result = invokeBoundary(boundary, request({
                method: 'POST',
                path: '/plataforma_v2/sessions/current/registration-enrollment',
                headers,
            }));
            assert.equal(result.nextCalls, 0);
            assert.equal(result.res.statusCode, 403);
            assertHardened(result.res);
        }
    });

    await t.test('missing, synthetic, and unrelated App Service hosts fail closed', () => {
        for (const host of [undefined, `${SYNTHETIC_TARGET_HOSTNAME}:443`, 'backend.azurewebsites.net']) {
            const headers = {};
            if (host !== undefined) headers.Host = host;
            const result = invokeBoundary(boundary, request({
                method: 'GET',
                path: '/plataforma_v2/sessions/current',
                headers,
            }));
            assert.equal(result.nextCalls, 0);
            assert.equal(result.res.statusCode, 403);
            assertHardened(result.res);
        }
    });

    await t.test('a presented target cookie requires the exact custom header', () => {
        for (const headerValue of [undefined, '0', ['1']]) {
            const headers = {
                Host: SYNTHETIC_TARGET_HOSTNAME,
                Cookie: `${SESSION_COOKIE_NAME}=${identifier}`,
            };
            if (headerValue !== undefined) headers['X-Machado-Session-Request'] = headerValue;
            const result = invokeBoundary(boundary, request({
                method: 'GET',
                path: '/plataforma_v2/sessions/current',
                headers,
            }));
            assert.equal(result.nextCalls, 0);
            assert.equal(result.res.statusCode, 403);
            assertHardened(result.res);
        }

        const accepted = invokeBoundary(boundary, request({
            method: 'GET',
            path: '/plataforma_v2/sessions/current',
            headers: {
                Host: SYNTHETIC_TARGET_HOSTNAME,
                Cookie: `${SESSION_COOKIE_NAME}=${identifier}`,
                'X-Machado-Session-Request': '1',
            },
        }));
        assert.equal(accepted.nextCalls, 1);
    });
});

test('target preflights terminate before parsing with exact route-specific policy', async (t) => {
    const boundary = createTargetHttpBoundary({
        targetRoutesEnabled: true,
        targetHostname: SYNTHETIC_TARGET_HOSTNAME,
        targetOrigin: SYNTHETIC_TARGET_ORIGIN,
    });

    function preflight(overrides = {}) {
        return request({
            method: 'OPTIONS',
            path: overrides.path || '/plataforma_v2/sessions/current',
            headers: {
                Host: SYNTHETIC_TARGET_HOSTNAME,
                Origin: SYNTHETIC_TARGET_ORIGIN,
                'Access-Control-Request-Method': overrides.method || 'DELETE',
                'Access-Control-Request-Headers': overrides.requestHeaders
                    || 'X-Machado-Session-Request, Content-Type',
                ...(overrides.headers || {}),
            },
        });
    }

    await t.test('valid preflight emits the exact boundary and ends with 204', () => {
        const result = invokeBoundary(boundary, preflight());
        assert.equal(result.nextCalls, 0);
        assert.equal(result.res.statusCode, 204);
        assert.equal(result.res.ended, true);
        assert.equal(result.res.getHeader('Access-Control-Allow-Origin'), SYNTHETIC_TARGET_ORIGIN);
        assert.equal(result.res.getHeader('Access-Control-Allow-Credentials'), 'true');
        assert.equal(result.res.getHeader('Access-Control-Allow-Methods'), 'GET, DELETE');
        assert.equal(result.res.getHeader('Access-Control-Allow-Headers'), ALLOWED_REQUEST_HEADERS);
        assert.equal(result.res.getHeader('Vary'), PREFLIGHT_VARY);
        assertHardened(result.res);
    });

    await t.test('wrong method, Origin, Host, missing header, and extra header reject', () => {
        const cases = [
            preflight({ method: 'PATCH' }),
            preflight({ headers: { Origin: 'https://wrong-origin.test' } }),
            preflight({ headers: { Host: 'backend.azurewebsites.net' } }),
            preflight({ requestHeaders: 'Content-Type' }),
            preflight({ requestHeaders: 'X-Machado-Session-Request, Authorization' }),
        ];

        for (const req of cases) {
            const result = invokeBoundary(boundary, req);
            assert.equal(result.nextCalls, 0);
            assert.equal(result.res.statusCode, 403);
            assertHardened(result.res);
        }
    });
});
