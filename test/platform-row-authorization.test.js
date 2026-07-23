'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createHmac } = require('node:crypto');
const {
    PLATFORM_ROW_AUTHORIZATION_DURATION_SECONDS,
    decodePlatformRowAuthorizationKey,
    createPlatformRowAuthorizationHandle,
    createPlatformRowAuthorizer,
    readPlatformRowIndex,
} = require('../platform-row-authorization');

const NOW_MS = 1_800_000_000_000;
const NOW_SECONDS = NOW_MS / 1000;
const SIGNING_KEY = Buffer.from(Array.from({ length: 32 }, (_, index) => index));
const WRONG_SIGNING_KEY = Buffer.from(Array.from({ length: 32 }, (_, index) => index + 1));
const INVALID_HANDLE_PATTERN = /Invalid platform row authorization handle/;

function createSignedHandle(payload, signingKey = SIGNING_KEY, json = JSON.stringify(payload)) {
    const payloadSegment = Buffer.from(json, 'utf8').toString('base64url');
    const signatureSegment = createHmac('sha256', signingKey)
        .update(payloadSegment, 'ascii')
        .digest('base64url');

    return `${payloadSegment}.${signatureSegment}`;
}

function validPayload(overrides = {}) {
    return {
        v: 1,
        purpose: 'platform-row-authorization',
        rowIndex: 42,
        iat: NOW_SECONDS,
        exp: NOW_SECONDS + PLATFORM_ROW_AUTHORIZATION_DURATION_SECONDS,
        ...overrides,
    };
}

function changeCharacter(value, index = 0) {
    const replacement = value[index] === 'A' ? 'B' : 'A';
    return `${value.slice(0, index)}${replacement}${value.slice(index + 1)}`;
}

test('creates and reads a four-hour platform row authorization handle', () => {
    assert.equal(PLATFORM_ROW_AUTHORIZATION_DURATION_SECONDS, 14_400);

    const handle = createPlatformRowAuthorizationHandle(42, SIGNING_KEY, NOW_MS);
    const [payloadSegment, signatureSegment] = handle.split('.');
    const payload = JSON.parse(Buffer.from(payloadSegment, 'base64url').toString('utf8'));

    assert.match(payloadSegment, /^[A-Za-z0-9_-]+$/);
    assert.match(signatureSegment, /^[A-Za-z0-9_-]+$/);
    assert.deepEqual(payload, validPayload());
    assert.equal(readPlatformRowIndex(handle, SIGNING_KEY, NOW_MS), 42);
});

test('accepts zero and the largest safe row index', () => {
    for (const rowIndex of [0, Number.MAX_SAFE_INTEGER]) {
        const handle = createPlatformRowAuthorizationHandle(rowIndex, SIGNING_KEY, NOW_MS);

        assert.equal(readPlatformRowIndex(handle, SIGNING_KEY, NOW_MS), rowIndex);
    }
});

test('expires exactly at the four-hour boundary', () => {
    const handle = createPlatformRowAuthorizationHandle(7, SIGNING_KEY, NOW_MS);
    const expirationMs = NOW_MS + PLATFORM_ROW_AUTHORIZATION_DURATION_SECONDS * 1000;

    assert.equal(readPlatformRowIndex(handle, SIGNING_KEY, expirationMs - 1), 7);
    assert.throws(
        () => readPlatformRowIndex(handle, SIGNING_KEY, expirationMs),
        INVALID_HANDLE_PATTERN,
    );
});

test('rejects tampering with either signed segment', () => {
    const handle = createPlatformRowAuthorizationHandle(42, SIGNING_KEY, NOW_MS);
    const [payloadSegment, signatureSegment] = handle.split('.');

    assert.throws(
        () => readPlatformRowIndex(
            `${changeCharacter(payloadSegment)}.${signatureSegment}`,
            SIGNING_KEY,
            NOW_MS,
        ),
        INVALID_HANDLE_PATTERN,
    );
    assert.throws(
        () => readPlatformRowIndex(
            `${payloadSegment}.${changeCharacter(signatureSegment)}`,
            SIGNING_KEY,
            NOW_MS,
        ),
        INVALID_HANDLE_PATTERN,
    );
});

test('rejects a valid handle checked with the wrong key', () => {
    const handle = createPlatformRowAuthorizationHandle(42, SIGNING_KEY, NOW_MS);

    assert.throws(
        () => readPlatformRowIndex(handle, WRONG_SIGNING_KEY, NOW_MS),
        INVALID_HANDLE_PATTERN,
    );
});

test('rejects non-string, missing, and structurally malformed handles', () => {
    for (const handle of [undefined, null, 123, {}, '', '.', 'payload', 'a.b.c', '.signature']) {
        assert.throws(
            () => readPlatformRowIndex(handle, SIGNING_KEY, NOW_MS),
            INVALID_HANDLE_PATTERN,
        );
    }

    assert.throws(
        () => readPlatformRowIndex('not+base64url.signature', SIGNING_KEY, NOW_MS),
        INVALID_HANDLE_PATTERN,
    );
});

test('rejects padded and non-canonical base64url segments', () => {
    const handle = createPlatformRowAuthorizationHandle(42, SIGNING_KEY, NOW_MS);
    const [payloadSegment, signatureSegment] = handle.split('.');
    const lastSignatureCharacter = signatureSegment.at(-1);
    const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
    const nonCanonicalLastCharacter = alphabet[alphabet.indexOf(lastSignatureCharacter) + 1];

    assert.throws(
        () => readPlatformRowIndex(`${payloadSegment}=.${signatureSegment}`, SIGNING_KEY, NOW_MS),
        INVALID_HANDLE_PATTERN,
    );
    assert.throws(
        () => readPlatformRowIndex(`${payloadSegment}.${signatureSegment}=`, SIGNING_KEY, NOW_MS),
        INVALID_HANDLE_PATTERN,
    );
    assert.throws(
        () => readPlatformRowIndex(
            `${payloadSegment}.${signatureSegment.slice(0, -1)}${nonCanonicalLastCharacter}`,
            SIGNING_KEY,
            NOW_MS,
        ),
        INVALID_HANDLE_PATTERN,
    );
});

test('rejects non-canonical JSON even when it has a valid signature', () => {
    const payload = validPayload();
    const reorderedPayload = {
        purpose: payload.purpose,
        v: payload.v,
        rowIndex: payload.rowIndex,
        iat: payload.iat,
        exp: payload.exp,
    };

    assert.throws(
        () => readPlatformRowIndex(
            createSignedHandle(payload, SIGNING_KEY, JSON.stringify(payload, null, 2)),
            SIGNING_KEY,
            NOW_MS,
        ),
        INVALID_HANDLE_PATTERN,
    );
    assert.throws(
        () => readPlatformRowIndex(createSignedHandle(reorderedPayload), SIGNING_KEY, NOW_MS),
        INVALID_HANDLE_PATTERN,
    );
    assert.throws(
        () => readPlatformRowIndex(
            createSignedHandle(payload, SIGNING_KEY, '{'),
            SIGNING_KEY,
            NOW_MS,
        ),
        INVALID_HANDLE_PATTERN,
    );
});

test('rejects invalid payload schemas even when correctly signed', () => {
    const missingExpiration = validPayload();
    delete missingExpiration.exp;

    const cases = [
        null,
        [],
        42,
        missingExpiration,
        { ...validPayload(), extra: true },
        validPayload({ v: 2 }),
        validPayload({ v: '1' }),
        validPayload({ purpose: 'platform-session' }),
    ];

    for (const payload of cases) {
        assert.throws(
            () => readPlatformRowIndex(createSignedHandle(payload), SIGNING_KEY, NOW_MS),
            INVALID_HANDLE_PATTERN,
        );
    }
});

test('rejects invalid row indexes when creating or reading handles', () => {
    const invalidRowIndexes = [-1, 1.5, Number.MAX_SAFE_INTEGER + 1, '42', null];

    for (const rowIndex of invalidRowIndexes) {
        assert.throws(
            () => createPlatformRowAuthorizationHandle(rowIndex, SIGNING_KEY, NOW_MS),
            /nonnegative safe integer/,
        );
        assert.throws(
            () => readPlatformRowIndex(
                createSignedHandle(validPayload({ rowIndex })),
                SIGNING_KEY,
                NOW_MS,
            ),
            INVALID_HANDLE_PATTERN,
        );
    }
});

test('rejects invalid issue and expiration times even when correctly signed', () => {
    const cases = [
        validPayload({ iat: NOW_SECONDS + 1, exp: NOW_SECONDS + 1 + 14_400 }),
        validPayload({ iat: NOW_SECONDS - 14_400, exp: NOW_SECONDS }),
        validPayload({ exp: NOW_SECONDS + 14_399 }),
        validPayload({ iat: NOW_SECONDS + 0.5, exp: NOW_SECONDS + 14_400.5 }),
        validPayload({ iat: '1800000000' }),
        validPayload({ exp: Number.MAX_SAFE_INTEGER + 1 }),
        validPayload({ iat: -14_400, exp: 0 }),
    ];

    for (const payload of cases) {
        assert.throws(
            () => readPlatformRowIndex(createSignedHandle(payload), SIGNING_KEY, NOW_MS),
            INVALID_HANDLE_PATTERN,
        );
    }
});

test('rejects invalid current-time arguments', () => {
    const handle = createPlatformRowAuthorizationHandle(42, SIGNING_KEY, NOW_MS);

    for (const nowMs of [-1, 1.5, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1, '1800000000000']) {
        assert.throws(
            () => createPlatformRowAuthorizationHandle(42, SIGNING_KEY, nowMs),
            /nonnegative safe integer/,
        );
        assert.throws(
            () => readPlatformRowIndex(handle, SIGNING_KEY, nowMs),
            /nonnegative safe integer/,
        );
    }
});

test('decodes only a canonical base64 key of exactly 32 bytes', () => {
    const canonicalBase64 = SIGNING_KEY.toString('base64');
    const base64UrlKey = Buffer.alloc(32, 0xff).toString('base64').replace(/\//g, '_');

    assert.deepEqual(decodePlatformRowAuthorizationKey(canonicalBase64), SIGNING_KEY);

    const invalidKeys = [
        undefined,
        null,
        123,
        '',
        ` ${canonicalBase64}`,
        canonicalBase64.slice(0, -1),
        `${canonicalBase64}=`,
        base64UrlKey,
        Buffer.alloc(31).toString('base64'),
        Buffer.alloc(33).toString('base64'),
        'not base64',
    ];

    for (const encodedKey of invalidKeys) {
        assert.throws(() => decodePlatformRowAuthorizationKey(encodedKey));
    }
});

test('requires decoded 32-byte Buffer signing keys', () => {
    const handle = createPlatformRowAuthorizationHandle(42, SIGNING_KEY, NOW_MS);
    const invalidKeys = [
        SIGNING_KEY.toString('base64'),
        new Uint8Array(SIGNING_KEY),
        Buffer.alloc(31),
        Buffer.alloc(33),
        undefined,
    ];

    for (const signingKey of invalidKeys) {
        assert.throws(
            () => createPlatformRowAuthorizationHandle(42, signingKey, NOW_MS),
            /32-byte Buffer/,
        );
        assert.throws(
            () => readPlatformRowIndex(handle, signingKey, NOW_MS),
            /32-byte Buffer/,
        );
    }
});

test('authorizes a valid legacy wire handle and exposes only its verified row index', () => {
    const handle = createPlatformRowAuthorizationHandle(42, SIGNING_KEY);
    const authorizePlatformRow = createPlatformRowAuthorizer(SIGNING_KEY);
    const req = { body: { IndexVerificado: handle } };
    const res = { locals: {} };
    let nextCalls = 0;

    authorizePlatformRow(req, res, () => { nextCalls += 1; });

    assert.equal(nextCalls, 1);
    assert.equal(res.locals.platformRowIndex, 42);
});

test('rejects an unverified legacy wire value before the downstream handler', () => {
    const authorizePlatformRow = createPlatformRowAuthorizer(SIGNING_KEY);

    for (const IndexVerificado of [undefined, 42, '42', 'tampered.handle']) {
        const req = { body: { IndexVerificado } };
        const res = {
            locals: {},
            statusCode: null,
            body: null,
            status(code) {
                this.statusCode = code;
                return this;
            },
            json(body) {
                this.body = body;
                return this;
            },
        };
        let nextCalls = 0;

        authorizePlatformRow(req, res, () => { nextCalls += 1; });

        assert.equal(nextCalls, 0);
        assert.equal(res.statusCode, 401);
        assert.deepEqual(res.body, {});
        assert.equal(res.locals.platformRowIndex, undefined);
    }
});
