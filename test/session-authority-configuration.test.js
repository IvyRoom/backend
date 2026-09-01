'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
    CONTROL_ENVIRONMENT_NAMES,
    KEY_ENVIRONMENT_NAMES,
    SQL_OPTION_LIMITS,
    assertSeparateLegacySigningKey,
    createAuthorityKeysetAggregate,
    readSessionAuthorityConfiguration,
} = require('../domains/session-authority/configuration');

function qualifiedEnvironment(overrides = {}) {
    const environment = {
        SESSION_AUTHORITY_DURABLE_STORE_REQUIRED: 'true',
        SESSION_AUTHORITY_EXPECTED_GENERATION: '1',
        SESSION_AUTHORITY_LEGACY_SIGNING_KEY_ID: 'synthetic-legacy-signing-key',
        SESSION_AUTHORITY_SQL_CONNECTION_STRING: 'Server=sql.example.test;Database=authority;Encrypt=true',
        SESSION_AUTHORITY_TARGET_ROUTES_ENABLED: 'true',
        PLATFORM_ROW_AUTHORIZATION_KEY_BASE64: Buffer.alloc(32, 99).toString('base64'),
    };
    let fill = 1;
    for (const names of Object.values(KEY_ENVIRONMENT_NAMES)) {
        environment[names.keyId] = `synthetic-key-${fill}`;
        environment[names.key] = Buffer.alloc(32, fill).toString('base64');
        fill += 1;
    }
    return { ...environment, ...overrides };
}

test('session authority is absent by default and needs no store or key material', () => {
    const configuration = readSessionAuthorityConfiguration({});
    assert.equal(configuration.enabled, false);
    assert.equal(configuration.topologyQualified, false);
    assert.deepEqual(configuration.runtimeControls, {
        durableStoreRequired: false,
        ...Object.fromEntries(Object.keys(CONTROL_ENVIRONMENT_NAMES).map((name) => [name, false])),
    });
    assert.equal('connectionString' in configuration, false);
    assert.equal('keys' in configuration, false);
    assert.equal('accountMappingKeyBinding' in configuration, false);
    assert.equal('authorityKeysetBinding' in configuration, false);
    assert.equal('loginLookupKeyBinding' in configuration, false);
    assert.equal('legacySigningKeyBinding' in configuration, false);
});

test('a dormant route-only configuration remains topology-independent but requires durable authority', () => {
    const configuration = readSessionAuthorityConfiguration(qualifiedEnvironment());
    assert.equal(configuration.enabled, true);
    assert.equal(configuration.expectedAuthorityGeneration, 1);
    assert.equal(configuration.runtimeControls.targetRoutesEnabled, true);
    assert.equal(configuration.runtimeControls.targetSessionIssuanceEnabled, false);
    assert.equal(configuration.topologyQualified, false);
    assert.equal(
        configuration.loginLookupKeyBinding.keyId,
        configuration.keys.loginLookup.keyId,
    );
    assert.equal(configuration.loginLookupKeyBinding.commitment.length, 32);
    assert.equal(configuration.authorityKeysetBinding.commitment.length, 32);
    assert.equal(
        configuration.legacySigningKeyBinding.keyId,
        'synthetic-legacy-signing-key',
    );
    assert.equal(configuration.legacySigningKeyBinding.commitment.length, 32);
    assert.equal(configuration.accountMappingKeyBinding.keyId, 'synthetic-key-5');
    assert.equal(configuration.accountMappingKeyBinding.commitment.length, 32);
    assert.notDeepEqual(
        configuration.accountMappingKeyBinding.commitment,
        configuration.authorityKeysetBinding.purposes.accountMappingEncryption.commitment,
    );
    assert.notDeepEqual(
        configuration.loginLookupKeyBinding.commitment,
        configuration.authorityKeysetBinding.purposes.loginLookup.commitment,
    );
    assert.deepEqual(
        Object.keys(configuration.authorityKeysetBinding.purposes),
        Object.keys(KEY_ENVIRONMENT_NAMES),
    );
    assert.deepEqual(configuration.sqlOptions, {
        connectionTimeout: 5_000,
        requestTimeout: 5_000,
        pool: { idleTimeoutMillis: 30_000, max: 5, min: 0 },
    });
});

test('legacy signing binding commits independently to key ID and material', () => {
    const original = readSessionAuthorityConfiguration(qualifiedEnvironment());
    const changedId = readSessionAuthorityConfiguration(qualifiedEnvironment({
        SESSION_AUTHORITY_LEGACY_SIGNING_KEY_ID: 'changed-legacy-signing-key',
    }));
    const changedMaterial = readSessionAuthorityConfiguration(qualifiedEnvironment({
        PLATFORM_ROW_AUTHORIZATION_KEY_BASE64: Buffer.alloc(32, 98).toString('base64'),
    }));

    assert.equal(
        original.legacySigningKeyBinding.commitment.toString('hex'),
        'c431acf119b5adef5013fba37e08e004d3e519a537bf78427518a6b5dd5a2176',
    );
    assert.notDeepEqual(
        changedId.legacySigningKeyBinding.commitment,
        original.legacySigningKeyBinding.commitment,
    );
    assert.notDeepEqual(
        changedMaterial.legacySigningKeyBinding.commitment,
        original.legacySigningKeyBinding.commitment,
    );
    assert.equal(
        original.legacySigningKeyBinding.commitment.includes(
            Buffer.from(qualifiedEnvironment().PLATFORM_ROW_AUTHORIZATION_KEY_BASE64, 'base64'),
        ),
        false,
    );
});

test('canonical keyset commitments bind every purpose, key ID, and secret byte', () => {
    const originalEnvironment = qualifiedEnvironment();
    const original = readSessionAuthorityConfiguration(originalEnvironment);
    assert.equal(
        original.authorityKeysetBinding.commitment.toString('hex'),
        '246e739fe0a51a7d4eb078586f672674a87914d6ec2acc73f0b03024af72e69c',
    );
    assert.deepEqual(
        createAuthorityKeysetAggregate(original.authorityKeysetBinding.purposes),
        original.authorityKeysetBinding.commitment,
    );
    assert.throws(
        () => createAuthorityKeysetAggregate({}),
        /Invalid session-authority purpose binding/,
    );

    for (const [index, [keyName, names]] of Object.entries(KEY_ENVIRONMENT_NAMES).entries()) {
        const changedId = readSessionAuthorityConfiguration(qualifiedEnvironment({
            [names.keyId]: `changed-${keyName}`,
        }));
        assert.notDeepEqual(
            changedId.authorityKeysetBinding.commitment,
            original.authorityKeysetBinding.commitment,
        );

        const changedMaterial = readSessionAuthorityConfiguration(qualifiedEnvironment({
            [names.key]: Buffer.alloc(32, 40 + index).toString('base64'),
        }));
        assert.notDeepEqual(
            changedMaterial.authorityKeysetBinding.commitment,
            original.authorityKeysetBinding.commitment,
        );
        assert.equal(
            original.authorityKeysetBinding.purposes[keyName].commitment.includes(
                original.keys[keyName].key,
            ),
            false,
        );
    }

    const mappingChanged = readSessionAuthorityConfiguration(qualifiedEnvironment({
        SESSION_AUTHORITY_ACCOUNT_MAPPING_KEY_BASE64: Buffer.alloc(32, 91).toString('base64'),
    }));
    assert.notDeepEqual(
        mappingChanged.accountMappingKeyBinding.commitment,
        original.accountMappingKeyBinding.commitment,
    );
});

test('login-lookup binding commits to both key ID and material without exposing the key', () => {
    const firstEnvironment = qualifiedEnvironment();
    const differentMaterialEnvironment = qualifiedEnvironment({
        SESSION_AUTHORITY_LOGIN_LOOKUP_KEY_BASE64: Buffer.alloc(32, 42).toString('base64'),
    });
    const caseVariantIdEnvironment = qualifiedEnvironment({
        SESSION_AUTHORITY_LOGIN_LOOKUP_KEY_ID: 'Synthetic-Key-3',
    });
    const first = readSessionAuthorityConfiguration(firstEnvironment);
    const differentMaterial = readSessionAuthorityConfiguration(differentMaterialEnvironment);
    const caseVariantId = readSessionAuthorityConfiguration(caseVariantIdEnvironment);

    assert.equal(
        first.loginLookupKeyBinding.keyId,
        differentMaterial.loginLookupKeyBinding.keyId,
    );
    assert.notDeepEqual(
        first.loginLookupKeyBinding.commitment,
        differentMaterial.loginLookupKeyBinding.commitment,
    );
    assert.notDeepEqual(
        first.loginLookupKeyBinding.commitment,
        caseVariantId.loginLookupKeyBinding.commitment,
    );
    assert.equal(
        first.loginLookupKeyBinding.commitment.includes(first.keys.loginLookup.key),
        false,
    );
});

test('the durable-store latch composes authority even before a rollout permission is enabled', () => {
    const configuration = readSessionAuthorityConfiguration(qualifiedEnvironment({
        SESSION_AUTHORITY_TARGET_ROUTES_ENABLED: 'false',
    }));
    assert.equal(configuration.enabled, true);
    assert.equal(configuration.runtimeControls.durableStoreRequired, true);
    assert.equal(Object.values(configuration.runtimeControls).filter(Boolean).length, 1);
});

test('no rollout control can bypass the durable-store latch', () => {
    assert.throws(() => readSessionAuthorityConfiguration({
        SESSION_AUTHORITY_TARGET_ROUTES_ENABLED: 'true',
    }), /requires the durable-store latch/);
});

test('target issuance requires coordinated routes, adoption, protected routes, and partitioned-cookie proof', () => {
    const base = qualifiedEnvironment({
        SESSION_AUTHORITY_TARGET_ISSUANCE_ENABLED: 'true',
    });
    assert.throws(() => readSessionAuthorityConfiguration(base), /subject-adoption gate/);

    base.SESSION_AUTHORITY_SUBJECT_ADOPTION_ENABLED = 'true';
    assert.throws(() => readSessionAuthorityConfiguration(base), /protected-route adoption/);

    base.SESSION_AUTHORITY_PROTECTED_ROUTES_ENABLED = 'true';
    assert.throws(() => readSessionAuthorityConfiguration(base), /legacy-ledger enforcement/);

    base.SESSION_AUTHORITY_LEGACY_ENFORCEMENT_ENABLED = 'true';
    assert.throws(() => readSessionAuthorityConfiguration(base), /requires ledger seeding/);

    base.SESSION_AUTHORITY_LEGACY_SEEDING_ENABLED = 'true';
    assert.throws(() => readSessionAuthorityConfiguration(base), /qualified partitioned-cookie topology/);

    base.SESSION_AUTHORITY_FIRST_PARTY_TOPOLOGY_QUALIFIED = 'true';
    assert.throws(() => readSessionAuthorityConfiguration(base), /qualified partitioned-cookie topology/);
    delete base.SESSION_AUTHORITY_FIRST_PARTY_TOPOLOGY_QUALIFIED;

    base.SESSION_AUTHORITY_PARTITIONED_COOKIE_TOPOLOGY_QUALIFIED = 'true';
    const configuration = readSessionAuthorityConfiguration(base);
    assert.equal(configuration.runtimeControls.targetSessionIssuanceEnabled, true);
    assert.equal(configuration.runtimeControls.subjectTargetAdoptionEnabled, true);
    assert.equal(configuration.runtimeControls.protectedRoutesEnabled, true);
});

test('legacy enforcement cannot bypass continuous ledger seeding configuration', () => {
    assert.throws(() => readSessionAuthorityConfiguration(qualifiedEnvironment({
        SESSION_AUTHORITY_LEGACY_ENFORCEMENT_ENABLED: 'true',
    })), /requires ledger seeding/);
});

test('configuration rejects ambiguous booleans, malformed keys, shared keys, and invalid limits', () => {
    assert.throws(
        () => readSessionAuthorityConfiguration({ SESSION_AUTHORITY_TARGET_ROUTES_ENABLED: 'TRUE' }),
        /exactly true or false/,
    );
    assert.throws(() => readSessionAuthorityConfiguration(qualifiedEnvironment({
        SESSION_AUTHORITY_TARGET_VERIFIER_KEY_BASE64: Buffer.alloc(31).toString('base64'),
    })), /exactly 32 bytes/);
    assert.throws(() => readSessionAuthorityConfiguration(qualifiedEnvironment({
        SESSION_AUTHORITY_LEGACY_SIGNING_KEY_ID: '',
    })), /LEGACY_SIGNING_KEY_ID is required/);
    assert.throws(() => readSessionAuthorityConfiguration(qualifiedEnvironment({
        SESSION_AUTHORITY_LEGACY_SIGNING_KEY_ID: 'synthetic-key-1',
    })), /must not reuse the legacy signing key ID/);

    const shared = qualifiedEnvironment();
    shared.SESSION_AUTHORITY_LEGACY_COMPATIBILITY_KEY_BASE64 = shared.SESSION_AUTHORITY_TARGET_VERIFIER_KEY_BASE64;
    assert.throws(() => readSessionAuthorityConfiguration(shared), /distinct by purpose/);

    assert.throws(() => readSessionAuthorityConfiguration(qualifiedEnvironment({
        SESSION_AUTHORITY_SQL_POOL_MAX: '0',
    })), /positive integer/);
    assert.throws(() => readSessionAuthorityConfiguration(qualifiedEnvironment({
        SESSION_AUTHORITY_EXPECTED_GENERATION: '',
    })), /EXPECTED_GENERATION is required/);

    const boundedOptions = [
        ['SESSION_AUTHORITY_SQL_CONNECTION_TIMEOUT_MS', SQL_OPTION_LIMITS.connectionTimeout],
        ['SESSION_AUTHORITY_SQL_REQUEST_TIMEOUT_MS', SQL_OPTION_LIMITS.requestTimeout],
        ['SESSION_AUTHORITY_SQL_POOL_IDLE_TIMEOUT_MS', SQL_OPTION_LIMITS.poolIdleTimeout],
        ['SESSION_AUTHORITY_SQL_POOL_MAX', SQL_OPTION_LIMITS.poolMax],
    ];
    for (const [name, maximum] of boundedOptions) {
        const accepted = readSessionAuthorityConfiguration(qualifiedEnvironment({
            [name]: String(maximum),
        }));
        assert.equal(
            name === 'SESSION_AUTHORITY_SQL_POOL_MAX'
                ? accepted.sqlOptions.pool.max
                : name === 'SESSION_AUTHORITY_SQL_POOL_IDLE_TIMEOUT_MS'
                    ? accepted.sqlOptions.pool.idleTimeoutMillis
                    : name === 'SESSION_AUTHORITY_SQL_CONNECTION_TIMEOUT_MS'
                        ? accepted.sqlOptions.connectionTimeout
                        : accepted.sqlOptions.requestTimeout,
            maximum,
        );
        assert.throws(() => readSessionAuthorityConfiguration(qualifiedEnvironment({
            [name]: String(maximum + 1),
        })), new RegExp(`must not exceed ${maximum}`));
    }
});

test('every session-authority purpose stays separate from the legacy signing key', () => {
    const configuration = readSessionAuthorityConfiguration(qualifiedEnvironment());
    assert.equal(assertSeparateLegacySigningKey(
        configuration.keys,
        Buffer.alloc(32, 99),
    ), true);
    assert.throws(() => assertSeparateLegacySigningKey(
        configuration.keys,
        configuration.keys.legacyCompatibility.key,
    ), /must not reuse the legacy signing key/);
    assert.throws(() => readSessionAuthorityConfiguration(qualifiedEnvironment({
        PLATFORM_ROW_AUTHORIZATION_KEY_BASE64:
            configuration.keys.legacyCompatibility.key.toString('base64'),
    })), /must not reuse the legacy signing key/);
});
