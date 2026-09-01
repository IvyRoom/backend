'use strict';

const { createHash, createHmac, timingSafeEqual } = require('node:crypto');
const { CRYPTOGRAPHIC_PURPOSES } = require('./cryptography');

const LOGIN_LOOKUP_KEY_BINDING_LABEL = 'session-authority/login-lookup-key-binding/v1';
const ACCOUNT_MAPPING_KEY_BINDING_LABEL = 'session-authority/account-mapping-key-binding/v1';
const AUTHORITY_KEYSET_LEAF_LABEL = Buffer.from(
    'machado-session-authority\0keyset-leaf\0v1\0',
    'utf8',
);
const AUTHORITY_KEYSET_LABEL = Buffer.from(
    'machado-session-authority\0keyset\0v1\0',
    'utf8',
);
const LEGACY_SIGNING_KEY_BINDING_LABEL = Buffer.from(
    'machado-session-authority\0legacy-signing-key-binding\0v1\0',
    'utf8',
);

const AUTHORITY_KEYSET_PURPOSES = Object.freeze([
    Object.freeze({
        keyName: 'targetVerifier',
        purpose: CRYPTOGRAPHIC_PURPOSES.targetSessionVerifier,
        rotatable: true,
    }),
    Object.freeze({
        keyName: 'legacyCompatibility',
        purpose: CRYPTOGRAPHIC_PURPOSES.legacyCompatibilityVerifier,
        rotatable: true,
    }),
    Object.freeze({
        keyName: 'loginLookup',
        purpose: CRYPTOGRAPHIC_PURPOSES.loginLookup,
        rotatable: false,
    }),
    Object.freeze({
        keyName: 'credentialFingerprint',
        purpose: CRYPTOGRAPHIC_PURPOSES.credentialFingerprint,
        rotatable: true,
    }),
    Object.freeze({
        keyName: 'accountMappingEncryption',
        purpose: CRYPTOGRAPHIC_PURPOSES.accountMapping,
        rotatable: false,
    }),
    Object.freeze({
        keyName: 'faceChallengeEncryption',
        purpose: CRYPTOGRAPHIC_PURPOSES.faceChallenge,
        rotatable: true,
    }),
]);

const SQL_OPTION_LIMITS = Object.freeze({
    connectionTimeout: 30_000,
    requestTimeout: 30_000,
    poolIdleTimeout: 300_000,
    poolMax: 10,
});

const CONTROL_ENVIRONMENT_NAMES = Object.freeze({
    targetRoutesEnabled: 'SESSION_AUTHORITY_TARGET_ROUTES_ENABLED',
    targetSessionIssuanceEnabled: 'SESSION_AUTHORITY_TARGET_ISSUANCE_ENABLED',
    legacyLedgerSeedingEnabled: 'SESSION_AUTHORITY_LEGACY_SEEDING_ENABLED',
    legacyCompatibilityEnforcementEnabled: 'SESSION_AUTHORITY_LEGACY_ENFORCEMENT_ENABLED',
    subjectTargetAdoptionEnabled: 'SESSION_AUTHORITY_SUBJECT_ADOPTION_ENABLED',
    protectedRoutesEnabled: 'SESSION_AUTHORITY_PROTECTED_ROUTES_ENABLED',
});

const KEY_ENVIRONMENT_NAMES = Object.freeze({
    targetVerifier: Object.freeze({
        key: 'SESSION_AUTHORITY_TARGET_VERIFIER_KEY_BASE64',
        keyId: 'SESSION_AUTHORITY_TARGET_VERIFIER_KEY_ID',
    }),
    legacyCompatibility: Object.freeze({
        key: 'SESSION_AUTHORITY_LEGACY_COMPATIBILITY_KEY_BASE64',
        keyId: 'SESSION_AUTHORITY_LEGACY_COMPATIBILITY_KEY_ID',
    }),
    loginLookup: Object.freeze({
        key: 'SESSION_AUTHORITY_LOGIN_LOOKUP_KEY_BASE64',
        keyId: 'SESSION_AUTHORITY_LOGIN_LOOKUP_KEY_ID',
    }),
    credentialFingerprint: Object.freeze({
        key: 'SESSION_AUTHORITY_CREDENTIAL_FINGERPRINT_KEY_BASE64',
        keyId: 'SESSION_AUTHORITY_CREDENTIAL_FINGERPRINT_KEY_ID',
    }),
    accountMappingEncryption: Object.freeze({
        key: 'SESSION_AUTHORITY_ACCOUNT_MAPPING_KEY_BASE64',
        keyId: 'SESSION_AUTHORITY_ACCOUNT_MAPPING_KEY_ID',
    }),
    faceChallengeEncryption: Object.freeze({
        key: 'SESSION_AUTHORITY_FACE_CHALLENGE_KEY_BASE64',
        keyId: 'SESSION_AUTHORITY_FACE_CHALLENGE_KEY_ID',
    }),
});

function readSessionAuthorityConfiguration(environment = {}) {
    const rolloutControls = Object.fromEntries(Object.entries(CONTROL_ENVIRONMENT_NAMES)
        .map(([name, environmentName]) => [name, readBoolean(environment, environmentName)]));
    const durableStoreRequired = readBoolean(
        environment,
        'SESSION_AUTHORITY_DURABLE_STORE_REQUIRED',
    );
    const runtimeControls = {
        durableStoreRequired,
        ...rolloutControls,
    };
    const topologyQualified = readBoolean(
        environment,
        'SESSION_AUTHORITY_PARTITIONED_COOKIE_TOPOLOGY_QUALIFIED',
    );

    validateControlDependencies(runtimeControls, topologyQualified);
    if (!durableStoreRequired) {
        return Object.freeze({
            enabled: false,
            runtimeControls: Object.freeze(runtimeControls),
            topologyQualified,
        });
    }

    const connectionString = requirePrivateString(
        environment.SESSION_AUTHORITY_SQL_CONNECTION_STRING,
        'SESSION_AUTHORITY_SQL_CONNECTION_STRING',
    );
    const expectedAuthorityGeneration = readPositiveInteger(
        environment,
        'SESSION_AUTHORITY_EXPECTED_GENERATION',
    );
    const keys = Object.fromEntries(Object.entries(KEY_ENVIRONMENT_NAMES).map(([name, names]) => [
        name,
        readKeyDescriptor(environment, names),
    ]));
    requireDistinctKeys(keys);
    const legacySigningKey = readCanonicalKey(
        environment,
        'PLATFORM_ROW_AUTHORIZATION_KEY_BASE64',
    );
    const legacySigningKeyId = readKeyId(
        environment,
        'SESSION_AUTHORITY_LEGACY_SIGNING_KEY_ID',
    );
    assertSeparateLegacySigningKey(keys, legacySigningKey, legacySigningKeyId);
    const legacySigningKeyBinding = createLegacySigningKeyBinding({
        keyId: legacySigningKeyId,
        key: legacySigningKey,
    });
    const loginLookupKeyBinding = createLoginLookupKeyBinding(keys.loginLookup);
    const authorityKeysetBinding = createAuthorityKeysetBinding(keys);
    const accountMappingKeyBinding = createAccountMappingKeyBinding(
        keys.accountMappingEncryption,
    );

    const sqlOptions = Object.freeze({
        connectionTimeout: readBoundedPositiveInteger(
            environment,
            'SESSION_AUTHORITY_SQL_CONNECTION_TIMEOUT_MS',
            5_000,
            SQL_OPTION_LIMITS.connectionTimeout,
        ),
        requestTimeout: readBoundedPositiveInteger(
            environment,
            'SESSION_AUTHORITY_SQL_REQUEST_TIMEOUT_MS',
            5_000,
            SQL_OPTION_LIMITS.requestTimeout,
        ),
        pool: Object.freeze({
            idleTimeoutMillis: readBoundedPositiveInteger(
                environment,
                'SESSION_AUTHORITY_SQL_POOL_IDLE_TIMEOUT_MS',
                30_000,
                SQL_OPTION_LIMITS.poolIdleTimeout,
            ),
            max: readBoundedPositiveInteger(
                environment,
                'SESSION_AUTHORITY_SQL_POOL_MAX',
                5,
                SQL_OPTION_LIMITS.poolMax,
            ),
            min: 0,
        }),
    });

    return Object.freeze({
        connectionString,
        enabled: true,
        expectedAuthorityGeneration,
        keys: Object.freeze(keys),
        accountMappingKeyBinding,
        authorityKeysetBinding,
        legacySigningKeyBinding,
        loginLookupKeyBinding,
        runtimeControls: Object.freeze(runtimeControls),
        sqlOptions,
        topologyQualified,
    });
}

function createLegacySigningKeyBinding({ keyId, key }) {
    if (!Buffer.isBuffer(key) || key.length !== 32) {
        throw new TypeError('The legacy signing key must contain exactly 32 bytes');
    }
    const keyIdBytes = Buffer.from(keyId, 'utf8');
    return Object.freeze({
        keyId,
        commitment: createHmac('sha256', key)
            .update(Buffer.concat([
                LEGACY_SIGNING_KEY_BINDING_LABEL,
                unsigned16(keyIdBytes.length),
                keyIdBytes,
            ]))
            .digest(),
    });
}

function createLoginLookupKeyBinding({ keyId, key }) {
    return Object.freeze({
        keyId,
        commitment: createHmac('sha256', key)
            .update(LOGIN_LOOKUP_KEY_BINDING_LABEL, 'utf8')
            .update('\0', 'utf8')
            .update(keyId, 'utf8')
            .digest(),
    });
}

function createAccountMappingKeyBinding({ keyId, key }) {
    return Object.freeze({
        keyId,
        commitment: createHmac('sha256', key)
            .update(ACCOUNT_MAPPING_KEY_BINDING_LABEL, 'utf8')
            .update('\0', 'utf8')
            .update(keyId, 'utf8')
            .digest(),
    });
}

function createAuthorityKeysetBinding(keys) {
    const purposes = {};

    for (const { keyName, purpose, rotatable } of AUTHORITY_KEYSET_PURPOSES) {
        const descriptor = keys[keyName];
        if (!descriptor) throw new TypeError(`Missing session-authority key: ${keyName}`);
        const purposeBytes = Buffer.from(purpose, 'utf8');
        const keyIdBytes = Buffer.from(descriptor.keyId, 'utf8');
        const leafFrame = Buffer.concat([
            AUTHORITY_KEYSET_LEAF_LABEL,
            unsigned16(purposeBytes.length),
            purposeBytes,
            unsigned16(keyIdBytes.length),
            keyIdBytes,
        ]);
        const commitment = createHmac('sha256', descriptor.key).update(leafFrame).digest();
        const binding = Object.freeze({
            commitment,
            keyId: descriptor.keyId,
            purpose,
            rotatable,
        });
        purposes[keyName] = binding;
    }

    const frozenPurposes = Object.freeze(purposes);
    return Object.freeze({
        commitment: createAuthorityKeysetAggregate(frozenPurposes),
        purposes: frozenPurposes,
    });
}

function createAuthorityKeysetAggregate(purposes) {
    if (!purposes || typeof purposes !== 'object' || Array.isArray(purposes)) {
        throw new TypeError('Session-authority purpose bindings are required');
    }
    const aggregateFrames = [
        AUTHORITY_KEYSET_LABEL,
        unsigned16(AUTHORITY_KEYSET_PURPOSES.length),
    ];
    for (const { keyName, purpose } of AUTHORITY_KEYSET_PURPOSES) {
        const binding = purposes[keyName];
        if (
            !binding
            || typeof binding.keyId !== 'string'
            || !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/u.test(binding.keyId)
            || !Buffer.isBuffer(binding.commitment)
            || binding.commitment.length !== 32
        ) throw new TypeError(`Invalid session-authority purpose binding: ${keyName}`);
        const purposeBytes = Buffer.from(purpose, 'utf8');
        const keyIdBytes = Buffer.from(binding.keyId, 'utf8');
        aggregateFrames.push(
            unsigned16(purposeBytes.length),
            purposeBytes,
            unsigned16(keyIdBytes.length),
            keyIdBytes,
            unsigned16(binding.commitment.length),
            binding.commitment,
        );
    }
    return createHash('sha256').update(Buffer.concat(aggregateFrames)).digest();
}

function unsigned16(value) {
    if (!Number.isSafeInteger(value) || value < 0 || value > 0xffff) {
        throw new TypeError('Session-authority keyset frame is too large');
    }
    const encoded = Buffer.allocUnsafe(2);
    encoded.writeUInt16BE(value);
    return encoded;
}

function readBoolean(environment, name) {
    const value = environment[name];
    if (value === undefined || value === '') return false;
    if (value === 'true') return true;
    if (value === 'false') return false;
    throw new TypeError(`${name} must be exactly true or false`);
}

function readPositiveInteger(environment, name, defaultValue) {
    const value = environment[name];
    if (value === undefined || value === '') {
        if (defaultValue === undefined) throw new TypeError(`${name} is required`);
        return defaultValue;
    }
    if (!/^[1-9][0-9]*$/u.test(value)) throw new TypeError(`${name} must be a positive integer`);
    const parsed = Number(value);
    if (!Number.isSafeInteger(parsed)) throw new TypeError(`${name} must be a safe positive integer`);
    return parsed;
}

function readBoundedPositiveInteger(environment, name, defaultValue, maximum) {
    const parsed = readPositiveInteger(environment, name, defaultValue);
    if (parsed > maximum) {
        throw new TypeError(`${name} must not exceed ${maximum}`);
    }
    return parsed;
}

function requirePrivateString(value, name) {
    if (typeof value !== 'string' || value.length === 0) {
        throw new TypeError(`${name} is required when session authority is enabled`);
    }
    return value;
}

function readKeyDescriptor(environment, names) {
    const keyId = readKeyId(environment, names.keyId);
    const key = readCanonicalKey(environment, names.key);
    return Object.freeze({ keyId, key });
}

function readKeyId(environment, name) {
    const keyId = requirePrivateString(environment[name], name);
    if (!/^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/u.test(keyId)) {
        throw new TypeError(`${name} is invalid`);
    }
    return keyId;
}

function readCanonicalKey(environment, name) {
    const encoded = requirePrivateString(environment[name], name);
    const key = Buffer.from(encoded, 'base64');
    if (key.length !== 32 || key.toString('base64') !== encoded) {
        throw new TypeError(`${name} must be canonical Base64 for exactly 32 bytes`);
    }
    return key;
}

function requireDistinctKeys(keys) {
    const descriptors = Object.values(keys);
    for (let index = 0; index < descriptors.length; index += 1) {
        for (let comparison = index + 1; comparison < descriptors.length; comparison += 1) {
            if (descriptors[index].keyId === descriptors[comparison].keyId) {
                throw new TypeError('Session-authority key IDs must be distinct by purpose');
            }
            if (timingSafeEqual(descriptors[index].key, descriptors[comparison].key)) {
                throw new TypeError('Session-authority keys must be distinct by purpose');
            }
        }
    }
}

function assertSeparateLegacySigningKey(keys, legacySigningKey, legacySigningKeyId) {
    if (!Buffer.isBuffer(legacySigningKey) || legacySigningKey.length !== 32) {
        throw new TypeError('The legacy signing key must contain exactly 32 bytes');
    }
    if (
        legacySigningKeyId !== undefined
        && Object.values(keys).some(({ keyId }) => keyId === legacySigningKeyId)
    ) {
        throw new TypeError('Session-authority key IDs must not reuse the legacy signing key ID');
    }
    if (Object.values(keys).some(({ key }) => timingSafeEqual(key, legacySigningKey))) {
        throw new TypeError('Session-authority keys must not reuse the legacy signing key');
    }
    return true;
}

function validateControlDependencies(controls, topologyQualified) {
    if (
        Object.entries(controls).some(([name, value]) => (
            name !== 'durableStoreRequired' && value === true
        ))
        && !controls.durableStoreRequired
    ) {
        throw new TypeError('Every session-authority rollout control requires the durable-store latch');
    }
    if (controls.targetSessionIssuanceEnabled && !controls.targetRoutesEnabled) {
        throw new TypeError('Target issuance requires the target route gate');
    }
    if (controls.targetSessionIssuanceEnabled && !controls.subjectTargetAdoptionEnabled) {
        throw new TypeError('Target issuance requires the subject-adoption gate');
    }
    if (controls.subjectTargetAdoptionEnabled && !controls.targetSessionIssuanceEnabled) {
        throw new TypeError('Subject adoption requires target issuance');
    }
    if (controls.targetSessionIssuanceEnabled && !controls.protectedRoutesEnabled) {
        throw new TypeError('Target issuance requires coordinated protected-route adoption');
    }
    if (
        controls.targetSessionIssuanceEnabled
        && !controls.legacyCompatibilityEnforcementEnabled
    ) {
        throw new TypeError('Target issuance requires irreversible legacy-ledger enforcement');
    }
    if (controls.protectedRoutesEnabled && !controls.targetRoutesEnabled) {
        throw new TypeError('Protected-route adoption requires the target route gate');
    }
    if (
        controls.protectedRoutesEnabled
        && (!controls.targetSessionIssuanceEnabled || !controls.subjectTargetAdoptionEnabled)
    ) {
        throw new TypeError('Protected-route adoption requires target issuance and subject adoption');
    }
    if (
        controls.legacyCompatibilityEnforcementEnabled
        && !controls.legacyLedgerSeedingEnabled
    ) {
        throw new TypeError('Legacy enforcement requires ledger seeding');
    }
    if (
        (controls.targetSessionIssuanceEnabled
            || controls.subjectTargetAdoptionEnabled
            || controls.protectedRoutesEnabled)
        && !topologyQualified
    ) {
        throw new TypeError('Target activation requires qualified partitioned-cookie topology');
    }
}

module.exports = {
    ACCOUNT_MAPPING_KEY_BINDING_LABEL,
    CONTROL_ENVIRONMENT_NAMES,
    KEY_ENVIRONMENT_NAMES,
    LEGACY_SIGNING_KEY_BINDING_LABEL,
    LOGIN_LOOKUP_KEY_BINDING_LABEL,
    AUTHORITY_KEYSET_PURPOSES,
    SQL_OPTION_LIMITS,
    assertSeparateLegacySigningKey,
    createAuthorityKeysetAggregate,
    createAuthorityKeysetBinding,
    createAccountMappingKeyBinding,
    createLegacySigningKeyBinding,
    createLoginLookupKeyBinding,
    readSessionAuthorityConfiguration,
};
