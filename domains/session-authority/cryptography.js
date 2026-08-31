'use strict';

const {
    createCipheriv,
    createDecipheriv,
    createHmac,
    randomBytes,
    timingSafeEqual,
} = require('node:crypto');

const OPAQUE_IDENTIFIER_BYTES = 32;
const OPAQUE_IDENTIFIER_LENGTH = 43;
const AES_KEY_BYTES = 32;
const AES_GCM_IV_BYTES = 12;
const AES_GCM_TAG_BYTES = 16;
const PRIVATE_VALUE_ENVELOPE_VERSION = 1;

const CRYPTOGRAPHIC_PURPOSES = Object.freeze({
    targetSession: 'target-session-verifier',
    targetSessionVerifier: 'target-session-verifier',
    legacyCompatibility: 'legacy-compatibility-verifier',
    legacyCompatibilityVerifier: 'legacy-compatibility-verifier',
    loginLookup: 'login-lookup',
    credentialFingerprint: 'credential-fingerprint',
    accountMapping: 'exact-login-encryption',
    exactLoginEncryption: 'exact-login-encryption',
    faceChallenge: 'face-reference-encryption',
    faceReferenceEncryption: 'face-reference-encryption',
});

const HMAC_PURPOSES = new Set([
    CRYPTOGRAPHIC_PURPOSES.targetSessionVerifier,
    CRYPTOGRAPHIC_PURPOSES.legacyCompatibilityVerifier,
    CRYPTOGRAPHIC_PURPOSES.loginLookup,
    CRYPTOGRAPHIC_PURPOSES.credentialFingerprint,
]);

const ENCRYPTION_PURPOSES = new Set([
    CRYPTOGRAPHIC_PURPOSES.exactLoginEncryption,
    CRYPTOGRAPHIC_PURPOSES.faceReferenceEncryption,
]);

function invalidIdentifier() {
    return new TypeError('Invalid opaque session identifier');
}

function asBuffer(value, name) {
    if (!Buffer.isBuffer(value) && !(value instanceof Uint8Array)) {
        throw new TypeError(`${name} must be bytes`);
    }

    return Buffer.from(value);
}

function readKeyDescriptor(keyDescriptor) {
    if (!keyDescriptor || typeof keyDescriptor !== 'object') {
        throw new TypeError('A cryptographic key descriptor is required');
    }

    const { keyId } = keyDescriptor;
    if (
        typeof keyId !== 'string'
        || !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/.test(keyId)
    ) {
        throw new TypeError('The cryptographic key ID is invalid');
    }

    const key = asBuffer(keyDescriptor.key, 'The cryptographic key');
    if (key.length !== AES_KEY_BYTES) {
        throw new TypeError('The cryptographic key must contain exactly 32 bytes');
    }

    return { keyId, key };
}

function readPrivateString(value) {
    if (typeof value !== 'string' || value.length === 0) {
        throw new TypeError('The private authority value is invalid');
    }

    return value;
}

function readHmacInput(value) {
    if (typeof value === 'string') {
        if (value.length === 0) throw new TypeError('The verifier input is invalid');
        return Buffer.from(value, 'utf8');
    }

    const input = asBuffer(value, 'The verifier input');
    if (input.length === 0) throw new TypeError('The verifier input is invalid');
    return input;
}

function purposeFrame(purpose, keyId) {
    return Buffer.from(`machado-session-authority\0v1\0${purpose}\0${keyId}\0`, 'utf8');
}

function encryptionFrame(purpose, keyId, associatedData) {
    const frame = purposeFrame(purpose, keyId);
    if (associatedData === undefined) return frame;
    const binding = asBuffer(associatedData, 'The private authority binding');
    if (binding.length === 0) throw new TypeError('The private authority binding is invalid');
    const length = Buffer.allocUnsafe(4);
    length.writeUInt32BE(binding.length);
    return Buffer.concat([
        frame,
        Buffer.from('bound\0v1\0', 'utf8'),
        length,
        binding,
    ]);
}

function requirePurpose(purpose, allowedPurposes) {
    if (!allowedPurposes.has(purpose)) {
        throw new TypeError('The cryptographic purpose is invalid');
    }

    return purpose;
}

function createOpaqueIdentifier(randomBytesSource = randomBytes) {
    if (typeof randomBytesSource !== 'function') {
        throw new TypeError('The random byte source is invalid');
    }

    const identifierBytes = asBuffer(
        randomBytesSource(OPAQUE_IDENTIFIER_BYTES),
        'The generated opaque identifier',
    );
    if (identifierBytes.length !== OPAQUE_IDENTIFIER_BYTES) {
        throw new TypeError('The random byte source returned an invalid opaque identifier');
    }

    return identifierBytes.toString('base64url');
}

function parseOpaqueIdentifier(value) {
    if (
        typeof value !== 'string'
        || value.length !== OPAQUE_IDENTIFIER_LENGTH
        || !/^[A-Za-z0-9_-]+$/.test(value)
    ) {
        throw invalidIdentifier();
    }

    const identifierBytes = Buffer.from(value, 'base64url');
    const canonicalValue = identifierBytes.toString('base64url');
    if (
        identifierBytes.length !== OPAQUE_IDENTIFIER_BYTES
        || canonicalValue.length !== value.length
        || !timingSafeEqual(Buffer.from(canonicalValue, 'ascii'), Buffer.from(value, 'ascii'))
    ) {
        throw invalidIdentifier();
    }

    return identifierBytes;
}

function createPurposeHmac(value, keyDescriptor, purpose) {
    requirePurpose(purpose, HMAC_PURPOSES);
    const { keyId, key } = readKeyDescriptor(keyDescriptor);

    const digest = createHmac('sha256', key)
        .update(purposeFrame(purpose, keyId))
        .update(readHmacInput(value))
        .digest();

    return { keyId, digest };
}

function createVerifier(
    value,
    keyDescriptor,
    purpose = CRYPTOGRAPHIC_PURPOSES.targetSessionVerifier,
) {
    let input = value;
    if (purpose === CRYPTOGRAPHIC_PURPOSES.targetSessionVerifier) {
        if (typeof value === 'string') {
            input = parseOpaqueIdentifier(value);
        } else {
            input = asBuffer(value, 'The opaque session identifier');
            if (input.length !== OPAQUE_IDENTIFIER_BYTES) throw invalidIdentifier();
        }
    }
    const { keyId, digest } = createPurposeHmac(input, keyDescriptor, purpose);

    return { keyId, verifier: digest };
}

function createLoginLookup(exactLogin, keyDescriptor) {
    const { keyId, digest } = createPurposeHmac(
        readPrivateString(exactLogin),
        keyDescriptor,
        CRYPTOGRAPHIC_PURPOSES.loginLookup,
    );

    return { keyId, token: digest };
}

function createCredentialFingerprint(exactCredential, keyDescriptor) {
    const { keyId, digest } = createPurposeHmac(
        readPrivateString(exactCredential),
        keyDescriptor,
        CRYPTOGRAPHIC_PURPOSES.credentialFingerprint,
    );

    return { keyId, fingerprint: digest };
}

function encryptPrivateValue(
    value,
    keyDescriptor,
    purpose,
    randomBytesSource = randomBytes,
    associatedData,
) {
    requirePurpose(purpose, ENCRYPTION_PURPOSES);
    const privateValue = readPrivateString(value);
    const { keyId, key } = readKeyDescriptor(keyDescriptor);
    if (typeof randomBytesSource !== 'function') {
        throw new TypeError('The random byte source is invalid');
    }

    const iv = asBuffer(randomBytesSource(AES_GCM_IV_BYTES), 'The generated encryption IV');
    if (iv.length !== AES_GCM_IV_BYTES) {
        throw new TypeError('The random byte source returned an invalid encryption IV');
    }

    const cipher = createCipheriv('aes-256-gcm', key, iv, { authTagLength: AES_GCM_TAG_BYTES });
    cipher.setAAD(encryptionFrame(purpose, keyId, associatedData));
    const encrypted = Buffer.concat([
        cipher.update(privateValue, 'utf8'),
        cipher.final(),
    ]);
    const authenticationTag = cipher.getAuthTag();
    const ciphertext = Buffer.concat([
        Buffer.from([PRIVATE_VALUE_ENVELOPE_VERSION]),
        iv,
        authenticationTag,
        encrypted,
    ]);

    return { keyId, ciphertext };
}

function decryptPrivateValue(ciphertext, keyDescriptor, purpose, associatedData) {
    requirePurpose(purpose, ENCRYPTION_PURPOSES);
    const encryptedEnvelope = asBuffer(ciphertext, 'The encrypted private authority value');
    const { keyId, key } = readKeyDescriptor(keyDescriptor);
    const minimumLength = 1 + AES_GCM_IV_BYTES + AES_GCM_TAG_BYTES + 1;

    if (
        encryptedEnvelope.length < minimumLength
        || encryptedEnvelope[0] !== PRIVATE_VALUE_ENVELOPE_VERSION
    ) {
        throw new TypeError('The encrypted private authority value is invalid');
    }

    const ivStart = 1;
    const tagStart = ivStart + AES_GCM_IV_BYTES;
    const encryptedStart = tagStart + AES_GCM_TAG_BYTES;
    const iv = encryptedEnvelope.subarray(ivStart, tagStart);
    const authenticationTag = encryptedEnvelope.subarray(tagStart, encryptedStart);
    const encrypted = encryptedEnvelope.subarray(encryptedStart);

    try {
        const decipher = createDecipheriv('aes-256-gcm', key, iv, {
            authTagLength: AES_GCM_TAG_BYTES,
        });
        decipher.setAAD(encryptionFrame(purpose, keyId, associatedData));
        decipher.setAuthTag(authenticationTag);
        return Buffer.concat([decipher.update(encrypted), decipher.final()]).toString('utf8');
    } catch (_error) {
        throw new TypeError('The encrypted private authority value could not be authenticated');
    }
}

function assertDistinctPurposeKeys(purposeKeyDescriptors) {
    if (!purposeKeyDescriptors || typeof purposeKeyDescriptors !== 'object') {
        throw new TypeError('Purpose-key configuration is required');
    }

    const seenKeyIds = new Set();
    const seenKeys = [];
    for (const purpose of new Set(Object.values(CRYPTOGRAPHIC_PURPOSES))) {
        const descriptor = readKeyDescriptor(purposeKeyDescriptors[purpose]);
        if (seenKeyIds.has(descriptor.keyId)) {
            throw new TypeError('Cryptographic key IDs must be distinct by purpose');
        }
        if (seenKeys.some((key) => timingSafeEqual(key, descriptor.key))) {
            throw new TypeError('Cryptographic keys must be distinct by purpose');
        }

        seenKeyIds.add(descriptor.keyId);
        seenKeys.push(descriptor.key);
    }

    return true;
}

module.exports = {
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
};
