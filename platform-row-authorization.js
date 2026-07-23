'use strict';

const { createHmac, timingSafeEqual } = require('node:crypto');

const PLATFORM_ROW_AUTHORIZATION_DURATION_SECONDS = 4 * 60 * 60;
const PLATFORM_ROW_AUTHORIZATION_PURPOSE = 'platform-row-authorization';
const SIGNING_KEY_LENGTH_BYTES = 32;
const SIGNATURE_LENGTH_BYTES = 32;
const PAYLOAD_KEYS = ['v', 'purpose', 'rowIndex', 'iat', 'exp'];
const BASE64URL_PATTERN = /^[A-Za-z0-9_-]+$/;

function decodePlatformRowAuthorizationKey(base64) {
    if (typeof base64 !== 'string') {
        throw new TypeError('Platform row authorization key must be canonical base64');
    }

    const signingKey = Buffer.from(base64, 'base64');

    if (
        signingKey.length !== SIGNING_KEY_LENGTH_BYTES ||
        signingKey.toString('base64') !== base64
    ) {
        throw new TypeError('Platform row authorization key must be canonical base64 for exactly 32 bytes');
    }

    return signingKey;
}

function createPlatformRowAuthorizationHandle(rowIndex, signingKey, nowMs = Date.now()) {
    validateRowIndex(rowIndex);
    validateSigningKey(signingKey);

    const issuedAt = toEpochSeconds(nowMs);
    const payload = {
        v: 1,
        purpose: PLATFORM_ROW_AUTHORIZATION_PURPOSE,
        rowIndex,
        iat: issuedAt,
        exp: issuedAt + PLATFORM_ROW_AUTHORIZATION_DURATION_SECONDS,
    };
    const payloadSegment = Buffer.from(JSON.stringify(payload), 'utf8').toString('base64url');
    const signatureSegment = signPayloadSegment(payloadSegment, signingKey).toString('base64url');

    return `${payloadSegment}.${signatureSegment}`;
}

function readPlatformRowIndex(handle, signingKey, nowMs = Date.now()) {
    validateSigningKey(signingKey);
    const now = toEpochSeconds(nowMs);

    if (typeof handle !== 'string') {
        throw invalidHandle();
    }

    const segments = handle.split('.');

    if (segments.length !== 2) {
        throw invalidHandle();
    }

    const [payloadSegment, signatureSegment] = segments;
    const payloadBuffer = decodeCanonicalBase64Url(payloadSegment);
    const signature = decodeCanonicalBase64Url(signatureSegment);

    if (signature.length !== SIGNATURE_LENGTH_BYTES) {
        throw invalidHandle();
    }

    const expectedSignature = signPayloadSegment(payloadSegment, signingKey);

    if (!timingSafeEqual(signature, expectedSignature)) {
        throw invalidHandle();
    }

    return readValidatedRowIndex(payloadBuffer, now);
}

function createPlatformRowAuthorizer(signingKey) {
    validateSigningKey(signingKey);

    return function authorizePlatformRow(req, res, next) {
        try {
            const platformRowHandle = req.body && req.body.IndexVerificado;
            res.locals.platformRowIndex = readPlatformRowIndex(platformRowHandle, signingKey);
            return next();
        } catch {
            return res.status(401).json({});
        }
    };
}

function validateSigningKey(signingKey) {
    if (!Buffer.isBuffer(signingKey) || signingKey.length !== SIGNING_KEY_LENGTH_BYTES) {
        throw new TypeError('Platform row authorization signing key must be a 32-byte Buffer');
    }
}

function validateRowIndex(rowIndex) {
    if (!Number.isSafeInteger(rowIndex) || rowIndex < 0) {
        throw new TypeError('Platform row index must be a nonnegative safe integer');
    }
}

function toEpochSeconds(nowMs) {
    if (!Number.isSafeInteger(nowMs) || nowMs < 0) {
        throw new TypeError('Current time must be a nonnegative safe integer in milliseconds');
    }

    return Math.floor(nowMs / 1000);
}

function signPayloadSegment(payloadSegment, signingKey) {
    return createHmac('sha256', signingKey).update(payloadSegment, 'ascii').digest();
}

function decodeCanonicalBase64Url(segment) {
    if (typeof segment !== 'string' || !BASE64URL_PATTERN.test(segment)) {
        throw invalidHandle();
    }

    const decoded = Buffer.from(segment, 'base64url');

    if (decoded.toString('base64url') !== segment) {
        throw invalidHandle();
    }

    return decoded;
}

function readValidatedRowIndex(payloadBuffer, now) {
    let payload;

    try {
        payload = JSON.parse(payloadBuffer.toString('utf8'));
    } catch {
        throw invalidHandle();
    }

    if (payload === null || typeof payload !== 'object' || Array.isArray(payload)) {
        throw invalidHandle();
    }

    const keys = Object.keys(payload);

    if (keys.length !== PAYLOAD_KEYS.length || keys.some((key, index) => key !== PAYLOAD_KEYS[index])) {
        throw invalidHandle();
    }

    if (
        payload.v !== 1 ||
        payload.purpose !== PLATFORM_ROW_AUTHORIZATION_PURPOSE ||
        !Number.isSafeInteger(payload.iat) ||
        payload.iat < 0 ||
        !Number.isSafeInteger(payload.exp) ||
        payload.exp < 0
    ) {
        throw invalidHandle();
    }

    try {
        validateRowIndex(payload.rowIndex);
    } catch {
        throw invalidHandle();
    }

    const canonicalPayload = Buffer.from(JSON.stringify({
        v: payload.v,
        purpose: payload.purpose,
        rowIndex: payload.rowIndex,
        iat: payload.iat,
        exp: payload.exp,
    }), 'utf8');

    if (!canonicalPayload.equals(payloadBuffer)) {
        throw invalidHandle();
    }

    if (
        payload.exp - payload.iat !== PLATFORM_ROW_AUTHORIZATION_DURATION_SECONDS ||
        payload.iat > now ||
        payload.exp <= now
    ) {
        throw invalidHandle();
    }

    return payload.rowIndex;
}

function invalidHandle() {
    return new Error('Invalid platform row authorization handle');
}

module.exports = {
    PLATFORM_ROW_AUTHORIZATION_DURATION_SECONDS,
    createPlatformRowAuthorizer,
    decodePlatformRowAuthorizationKey,
    createPlatformRowAuthorizationHandle,
    readPlatformRowIndex,
};
