'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { randomBytes } = require('node:crypto');
const { ELIGIBILITY_REVALIDATION_MS } = require('../domains/session-authority/constants');
const { ERROR_CLASSES } = require('../domains/session-authority/errors');
const {
    CRYPTOGRAPHIC_PURPOSES,
    OPAQUE_IDENTIFIER_BYTES,
    assertDistinctPurposeKeys,
    createCredentialFingerprint,
    createLoginLookup,
    createOpaqueIdentifier,
    createVerifier,
    decryptPrivateValue,
    encryptPrivateValue,
    parseOpaqueIdentifier,
} = require('../domains/session-authority/cryptography');
const {
    excelSerialToCivilDate,
    excelSerialToEntitlementExpiry,
    normalizeEligibility,
    readCredentialAccount,
    readFaceRequirement,
    readMappedAccount,
} = require('../domains/session-authority/account-authority');

const EXCEL_EPOCH_UTC_MS = Date.UTC(1899, 11, 31);
const MILLISECONDS_PER_DAY = 24 * 60 * 60 * 1000;

function keyDescriptor(label = 'synthetic') {
    return { keyId: `${label}-${randomBytes(8).toString('hex')}`, key: randomBytes(32) };
}

function expectAuthorityError(errorClass, reason) {
    return (error) => {
        assert.equal(error.errorClass, errorClass);
        assert.equal(error.reason, reason);
        return true;
    };
}

function excelSerial(year, month, day) {
    const realDays = (Date.UTC(year, month - 1, day) - EXCEL_EPOCH_UTC_MS)
        / MILLISECONDS_PER_DAY;
    return realDays >= 60 ? realDays + 1 : realDays;
}

function syntheticAccountRow(overrides = {}) {
    return [
        'Synthetic Learner',
        'Synthetic',
        `learner-${randomBytes(8).toString('hex')}@example.invalid`,
        randomBytes(24).toString('base64url'),
        'Ativo',
        'Sim',
        excelSerial(2026, 9, 30),
        'Ativo',
        ...new Array(14).fill(0),
    ].map((value, index) => (
        Object.prototype.hasOwnProperty.call(overrides, index) ? overrides[index] : value
    ));
}

test('opaque identifiers contain 32 random bytes in canonical unpadded base64url', () => {
    const identifier = createOpaqueIdentifier();
    const parsed = parseOpaqueIdentifier(identifier);

    assert.equal(parsed.length, OPAQUE_IDENTIFIER_BYTES);
    assert.equal(identifier, parsed.toString('base64url'));
    assert.match(identifier, /^[A-Za-z0-9_-]{43}$/);
    assert.equal(identifier.includes('='), false);
    assert.throws(() => parseOpaqueIdentifier(`${identifier}=`), /Invalid opaque/);
    assert.throws(() => parseOpaqueIdentifier(identifier.slice(1)), /Invalid opaque/);
    assert.throws(() => createOpaqueIdentifier(() => randomBytes(31)), /invalid opaque/);
});

test('purpose-bound HMAC helpers return only keyed SHA-256 digests', () => {
    const identifier = createOpaqueIdentifier();
    const verifierKey = keyDescriptor('target');
    const legacyKey = keyDescriptor('legacy');
    const lookupKey = keyDescriptor('lookup');
    const fingerprintKey = keyDescriptor('fingerprint');
    const legacyHandle = randomBytes(48).toString('base64url');
    const exactLogin = ` exact-${randomBytes(8).toString('hex')}@example.invalid `;
    const exactCredential = ` ${randomBytes(24).toString('base64url')} `;

    const verifier = createVerifier(identifier, verifierKey);
    const verifierRepeat = createVerifier(identifier, verifierKey);
    const verifierFromParsedBytes = createVerifier(
        parseOpaqueIdentifier(identifier),
        verifierKey,
        CRYPTOGRAPHIC_PURPOSES.targetSession,
    );
    const legacyVerifier = createVerifier(
        legacyHandle,
        legacyKey,
        CRYPTOGRAPHIC_PURPOSES.legacyCompatibilityVerifier,
    );
    const lookup = createLoginLookup(exactLogin, lookupKey);
    const fingerprint = createCredentialFingerprint(exactCredential, fingerprintKey);

    assert.equal(verifier.keyId, verifierKey.keyId);
    assert.equal(verifier.verifier.length, 32);
    assert.deepEqual(verifier.verifier, verifierRepeat.verifier);
    assert.deepEqual(verifier.verifier, verifierFromParsedBytes.verifier);
    assert.equal(legacyVerifier.verifier.length, 32);
    assert.equal(lookup.token.length, 32);
    assert.equal(fingerprint.fingerprint.length, 32);
    assert.notDeepEqual(verifier.verifier, legacyVerifier.verifier);
    assert.equal(JSON.stringify({ verifier, legacyVerifier, lookup, fingerprint }).includes(identifier), false);
    assert.equal(JSON.stringify({ verifier, legacyVerifier, lookup, fingerprint }).includes(exactCredential), false);
});

test('AES-256-GCM preserves exact private values and binds purpose plus key ID', () => {
    const loginKey = keyDescriptor('login-encryption');
    const faceKey = keyDescriptor('face-encryption');
    const exactLogin = ` exact-${randomBytes(8).toString('hex')}@example.invalid `;
    const privateFaceReference = randomBytes(32).toString('base64url');

    const encryptedLogin = encryptPrivateValue(
        exactLogin,
        loginKey,
        CRYPTOGRAPHIC_PURPOSES.exactLoginEncryption,
    );
    const encryptedLoginAgain = encryptPrivateValue(
        exactLogin,
        loginKey,
        CRYPTOGRAPHIC_PURPOSES.exactLoginEncryption,
    );
    const encryptedFaceReference = encryptPrivateValue(
        privateFaceReference,
        faceKey,
        CRYPTOGRAPHIC_PURPOSES.faceReferenceEncryption,
    );

    assert.equal(
        decryptPrivateValue(
            encryptedLogin.ciphertext,
            loginKey,
            CRYPTOGRAPHIC_PURPOSES.exactLoginEncryption,
        ),
        exactLogin,
    );
    assert.equal(
        decryptPrivateValue(
            encryptedFaceReference.ciphertext,
            faceKey,
            CRYPTOGRAPHIC_PURPOSES.faceReferenceEncryption,
        ),
        privateFaceReference,
    );
    assert.notDeepEqual(encryptedLogin.ciphertext, encryptedLoginAgain.ciphertext);
    assert.equal(JSON.stringify(encryptedLogin).includes(exactLogin), false);
    assert.throws(
        () => decryptPrivateValue(
            encryptedLogin.ciphertext,
            loginKey,
            CRYPTOGRAPHIC_PURPOSES.faceReferenceEncryption,
        ),
        /could not be authenticated/,
    );

    const tampered = Buffer.from(encryptedFaceReference.ciphertext);
    tampered[tampered.length - 1] ^= 1;
    assert.throws(
        () => decryptPrivateValue(
            tampered,
            faceKey,
            CRYPTOGRAPHIC_PURPOSES.faceReferenceEncryption,
        ),
        /could not be authenticated/,
    );
});

test('all configured purposes require distinct key IDs and key material', () => {
    const keys = Object.fromEntries(
        Object.values(CRYPTOGRAPHIC_PURPOSES).map((purpose) => [purpose, keyDescriptor(purpose)]),
    );
    assert.equal(assertDistinctPurposeKeys(keys), true);

    const repeatedKeyId = {
        ...keys,
        [CRYPTOGRAPHIC_PURPOSES.faceReferenceEncryption]: {
            ...keys[CRYPTOGRAPHIC_PURPOSES.faceReferenceEncryption],
            keyId: keys[CRYPTOGRAPHIC_PURPOSES.exactLoginEncryption].keyId,
        },
    };
    assert.throws(() => assertDistinctPurposeKeys(repeatedKeyId), /key IDs must be distinct/);

    const repeatedKey = {
        ...keys,
        [CRYPTOGRAPHIC_PURPOSES.faceReferenceEncryption]: {
            ...keys[CRYPTOGRAPHIC_PURPOSES.faceReferenceEncryption],
            key: keys[CRYPTOGRAPHIC_PURPOSES.exactLoginEncryption].key,
        },
    };
    assert.throws(() => assertDistinctPurposeKeys(repeatedKey), /keys must be distinct/);
});

test('credential matching is exact and untrimmed and captures fresh Face policy', () => {
    const fingerprintKey = keyDescriptor('credential');
    const matchingRow = syntheticAccountRow();
    matchingRow[2] = ` exact-${randomBytes(8).toString('hex')}@example.invalid `;
    matchingRow[3] = ` ${randomBytes(24).toString('base64url')} `;
    const rows = [syntheticAccountRow(), matchingRow, syntheticAccountRow()];

    const account = readCredentialAccount(rows, {
        credentialFingerprintKey: fingerprintKey,
        login: matchingRow[2],
        password: matchingRow[3],
    });

    assert.equal(account.rowIndex, 1);
    assert.equal(account.exactLogin, matchingRow[2]);
    assert.equal(account.faceAuthRequired, true);
    assert.equal(account.photoRegistrationStatus, 'Sim');
    assert.deepEqual(
        account.credentialFingerprint,
        createCredentialFingerprint(matchingRow[3], fingerprintKey),
    );
    assert.equal(JSON.stringify(account).includes(matchingRow[2]), false);
    assert.throws(
        () => readCredentialAccount(rows, {
            credentialFingerprintKey: fingerprintKey,
            login: matchingRow[2].trim(),
            password: matchingRow[3],
        }),
        expectAuthorityError(ERROR_CLASSES.invalid, 'invalid-credentials'),
    );
    assert.throws(
        () => readCredentialAccount(rows, {
            credentialFingerprintKey: fingerprintKey,
            login: matchingRow[2],
            password: matchingRow[3].trim(),
        }),
        expectAuthorityError(ERROR_CLASSES.invalid, 'invalid-credentials'),
    );
});

test('credential lookup distinguishes duplicate and malformed account data', () => {
    const fingerprintKey = keyDescriptor('credential');
    const matchingRow = syntheticAccountRow();
    const options = {
        credentialFingerprintKey: fingerprintKey,
        login: matchingRow[2],
        password: matchingRow[3],
    };

    const conflictingDuplicate = [...matchingRow];
    conflictingDuplicate[3] = randomBytes(24).toString('base64url');
    assert.throws(
        () => readCredentialAccount([matchingRow, conflictingDuplicate], options),
        expectAuthorityError(ERROR_CLASSES.unavailable, 'duplicate-credential-match'),
    );

    const malformedRow = syntheticAccountRow({ 3: null });
    assert.throws(
        () => readCredentialAccount([matchingRow, malformedRow], options),
        expectAuthorityError(ERROR_CLASSES.unavailable, 'malformed-account-row'),
    );
});

test('mapped-account lookup follows an exact login across row movement without rereading Face policy', () => {
    const fingerprintKey = keyDescriptor('credential');
    const mappedRow = syntheticAccountRow();
    const original = readMappedAccount([mappedRow, syntheticAccountRow()], mappedRow[2], {
        credentialFingerprintKey: fingerprintKey,
    });

    mappedRow[4] = 'unexpected-later-edit';
    const moved = readMappedAccount([syntheticAccountRow(), syntheticAccountRow(), mappedRow], mappedRow[2], {
        credentialFingerprintKey: fingerprintKey,
    });

    assert.equal(original.rowIndex, 0);
    assert.equal(moved.rowIndex, 2);
    assert.equal(moved.exactLogin, mappedRow[2]);
    assert.deepEqual(moved.credentialFingerprint, original.credentialFingerprint);
    assert.equal(Object.prototype.hasOwnProperty.call(moved, 'faceAuthRequired'), false);

    assert.throws(
        () => readMappedAccount([], mappedRow[2], { credentialFingerprintKey: fingerprintKey }),
        expectAuthorityError(ERROR_CLASSES.unavailable, 'missing-account-mapping'),
    );
    assert.throws(
        () => readMappedAccount([mappedRow, [...mappedRow]], mappedRow[2], {
            credentialFingerprintKey: fingerprintKey,
        }),
        expectAuthorityError(ERROR_CLASSES.unavailable, 'duplicate-account-mapping'),
    );
});

test('Face policy and account status values are exact and case-sensitive', () => {
    assert.equal(readFaceRequirement('Ativo'), true);
    assert.equal(readFaceRequirement('Inativo'), false);

    for (const invalidValue of [undefined, null, '', 'ativo', 'Inativo ', true]) {
        assert.throws(
            () => readFaceRequirement(invalidValue),
            expectAuthorityError(ERROR_CLASSES.unavailable, 'invalid-face-policy'),
        );
    }

    const accessDateSerial = excelSerial(2026, 9, 30);
    for (const accountStatus of ['ativo', 'Ativo ', 'Inativo', '', undefined]) {
        assert.throws(
            () => normalizeEligibility(
                { accountStatus, accessDateSerial },
                Date.parse('2026-08-31T12:00:00.000Z'),
            ),
            expectAuthorityError(ERROR_CLASSES.forbidden, 'account-inactive'),
        );
    }
});

test('Excel serial access dates expire at the exclusive next São Paulo civil day', () => {
    const currentSerial = excelSerial(2026, 8, 31);
    assert.deepEqual(excelSerialToCivilDate(currentSerial), { year: 2026, month: 8, day: 31 });
    assert.equal(
        excelSerialToEntitlementExpiry(currentSerial).toISOString(),
        '2026-09-01T03:00:00.000Z',
    );

    const beforeHistoricalMidnightTransition = excelSerial(2018, 11, 3);
    assert.equal(
        excelSerialToEntitlementExpiry(beforeHistoricalMidnightTransition).toISOString(),
        '2018-11-04T03:00:00.000Z',
    );

    for (const invalidValue of [NaN, Infinity, '45500', 0, 60, 2_958_465]) {
        assert.throws(() => excelSerialToEntitlementExpiry(invalidValue), /Excel access date is invalid/);
    }
});

test('eligibility observations use backend time and revalidate no later than five minutes', () => {
    const observedAt = Date.parse('2026-08-31T12:00:00.000Z');
    const account = {
        accountStatus: 'Ativo',
        accessDateSerial: excelSerial(2026, 8, 31),
    };
    const observation = normalizeEligibility(account, observedAt);

    assert.equal(observation.eligibilityState, 'eligible');
    assert.equal(observation.eligibilityObservedAt.getTime(), observedAt);
    assert.equal(
        observation.eligibilityRevalidateAt.getTime(),
        observedAt + ELIGIBILITY_REVALIDATION_MS,
    );
    assert.equal(observation.entitlementExpiresAt.toISOString(), '2026-09-01T03:00:00.000Z');

    const nearExpiry = Date.parse('2026-09-01T02:58:00.000Z');
    assert.equal(
        normalizeEligibility(account, nearExpiry).eligibilityRevalidateAt.toISOString(),
        '2026-09-01T03:00:00.000Z',
    );
    assert.throws(
        () => normalizeEligibility(account, Date.parse('2026-09-01T03:00:00.000Z')),
        expectAuthorityError(ERROR_CLASSES.forbidden, 'entitlement-expired'),
    );
    assert.throws(
        () => normalizeEligibility({ ...account, accessDateSerial: NaN }, observedAt),
        expectAuthorityError(ERROR_CLASSES.forbidden, 'invalid-entitlement'),
    );
});
