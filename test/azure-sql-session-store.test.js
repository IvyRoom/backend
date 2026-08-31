'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const {
    createAzureSqlSessionStore: createAzureSqlSessionStoreFactoryRaw,
} = require('../integrations/azure-sql-session-store');
const {
    createTestSessionAuthorityBacking,
    createTestSessionAuthorityStore,
} = require('./session-authority-store-support');

const SERVER_TIME = new Date('2042-06-01T12:00:00.000Z');
const EXPECTED_AUTHORITY_GENERATION = 3;
const LOGIN_LOOKUP_KEY_ID = 'synthetic-lookup-key';
const LOGIN_LOOKUP_KEY_COMMITMENT = Buffer.alloc(32, 91);
const keyBinding = (keyId, byte) => Object.freeze({
    keyId,
    commitment: Buffer.alloc(32, byte),
});
const ACCOUNT_MAPPING_KEY_BINDING = keyBinding('synthetic-account-mapping-key', 92);
const LEGACY_SIGNING_KEY_BINDING = keyBinding('synthetic-legacy-signing-key', 93);
const AUTHORITY_KEYSET_BINDING = Object.freeze({
    commitment: Buffer.alloc(32, 94),
    purposes: Object.freeze({
        targetVerifier: keyBinding('synthetic-target-verifier-key', 95),
        legacyCompatibility: keyBinding('synthetic-legacy-compatibility-key', 96),
        loginLookup: Object.freeze({
            keyId: LOGIN_LOOKUP_KEY_ID,
            commitment: Buffer.alloc(32, 99),
        }),
        credentialFingerprint: keyBinding('synthetic-credential-fingerprint-key', 97),
        accountMappingEncryption: keyBinding(ACCOUNT_MAPPING_KEY_BINDING.keyId, 100),
        faceChallengeEncryption: keyBinding('synthetic-face-challenge-key', 98),
    }),
});
const STORE_METHODS = Object.freeze([
    'admitUnboundLegacyIssuance',
    'authorizeLegacy',
    'bindFaceChallengeAndRotate',
    'bindLegacy',
    'completeFaceFailure',
    'completeFaceSuccessAndRotate',
    'createOrLoadSubject',
    'disableLegacyAuthority',
    'heartbeatLegacySeedingContinuity',
    'initializeLoginLookupKey',
    'inspectLoginPredecessor',
    'issueSession',
    'logout',
    'markFaceFlowReconciliation',
    'readControl',
    'readFaceFlow',
    'readSession',
    'readSubjectByLookup',
    'remapSubjectLogin',
    'reserveFaceFlow',
    'revokeAll',
    'revokeForCredentialChange',
    'revokeForIneligibility',
    'revokeSubject',
    'rotateSession',
    'transitionControl',
    'updateEligibility',
]);

function createAzureSqlSessionStore(input) {
    if (!input || Object.hasOwn(input, 'expectedAuthorityGeneration')) {
        return createAzureSqlSessionStoreFactory(input);
    }
    return createAzureSqlSessionStoreFactory({
        ...input,
        expectedAuthorityGeneration: EXPECTED_AUTHORITY_GENERATION,
    });
}

function createAzureSqlSessionStoreFactory(input) {
    return createAzureSqlSessionStoreFactoryRaw(input && {
        loginLookupKeyId: LOGIN_LOOKUP_KEY_ID,
        loginLookupKeyCommitment: LOGIN_LOOKUP_KEY_COMMITMENT,
        accountMappingKeyBinding: ACCOUNT_MAPPING_KEY_BINDING,
        authorityKeysetBinding: AUTHORITY_KEYSET_BINDING,
        legacySigningKeyBinding: LEGACY_SIGNING_KEY_BINDING,
        ...input,
    });
}

test('adapter construction is inert and exports the complete authority surface', () => {
    const fake = createFakeSql();
    const store = createAzureSqlSessionStore({
        sql: fake.sql,
        connectionString: 'synthetic-only',
        options: { pool: { max: 4 } },
    });

    assert.equal(fake.state.poolConstructions, 0);
    assert.equal(fake.state.connects, 0);
    for (const method of STORE_METHODS) assert.equal(typeof store[method], 'function', method);
    assert.equal(typeof store.close, 'function');
});

test('factory requires an injected driver and exactly one connection configuration', () => {
    const fake = createFakeSql();
    assert.throws(() => createAzureSqlSessionStore(), /injected SQL driver/);
    assert.throws(() => createAzureSqlSessionStoreFactory({
        sql: fake.sql,
        connectionString: 'synthetic',
        options: {},
    }), /Expected authority generation/);
    assert.throws(() => createAzureSqlSessionStore({ sql: fake.sql }), /connection string/);
    assert.throws(() => createAzureSqlSessionStore({
        sql: fake.sql,
        connectionString: 'synthetic',
        options: null,
    }), /SQL options/);
    assert.throws(() => createAzureSqlSessionStoreFactory({
        sql: fake.sql,
        connectionString: 'synthetic',
        expectedAuthorityGeneration: EXPECTED_AUTHORITY_GENERATION,
        legacySigningKeyBinding: {
            keyId: `${LEGACY_SIGNING_KEY_BINDING.keyId} `,
            commitment: LEGACY_SIGNING_KEY_BINDING.commitment,
        },
    }), /Legacy signing key binding/);
});

test('configured authority generation fences reads and rolls back stale-instance mutations', async () => {
    const readFake = createFakeSql([
        result([controlRow({ authorityGeneration: 4 })]),
    ]);
    const staleReader = createAzureSqlSessionStoreFactory({
        sql: readFake.sql,
        connectionString: 'synthetic-only',
        expectedAuthorityGeneration: 3,
    });
    await assert.rejects(
        staleReader.readControl(),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'authority-generation-mismatch',
    );

    const transitionFake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({ authorityGeneration: 4 })]),
    ]);
    const staleControlClient = createAzureSqlSessionStoreFactory({
        sql: transitionFake.sql,
        connectionString: 'synthetic-only',
        expectedAuthorityGeneration: 3,
    });
    await assert.rejects(
        staleControlClient.transitionControl({
            expectedVersion: 3,
            changes: { targetRoutesEnabled: true },
        }),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'authority-generation-mismatch',
    );
    assert.equal(transitionFake.state.commits, 0);
    assert.equal(transitionFake.state.rollbacks, 1);
    assert.equal(
        transitionFake.state.queries.some(({ statement }) => statement.includes('transition-control')),
        false,
    );

    const mutationFake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({ authorityGeneration: 4 })]),
    ]);
    const staleWriter = createAzureSqlSessionStoreFactory({
        sql: mutationFake.sql,
        connectionString: 'synthetic-only',
        expectedAuthorityGeneration: 3,
    });
    await assert.rejects(
        staleWriter.updateEligibility({
            subjectId: subjectRow().subjectId,
            observationStartedAt: SERVER_TIME,
            expectedCredentialVersion: 2,
            expectedCredentialFingerprintKeyId: 'synthetic-fingerprint-key',
            expectedCredentialFingerprint: Buffer.alloc(32, 2),
            expectedControlVersion: 3,
            rowHint: 8,
            entitlementExpiresAt: new Date('2042-06-02T03:00:00.000Z'),
            eligibilityRevalidateAt: new Date('2042-06-01T12:05:00.000Z'),
        }),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'authority-generation-mismatch',
    );
    assert.equal(mutationFake.state.commits, 0);
    assert.equal(mutationFake.state.rollbacks, 1);
    assert.equal(mutationFake.state.queries.some(({ statement }) => (
        statement.includes('update-eligibility */')
    )), false);
});

test('memory store generation fencing requires a new explicitly configured instance after advancement', async () => {
    const backing = createTestSessionAuthorityBacking();
    const oldInstance = createTestSessionAuthorityStore({
        testOnly: true,
        backing,
        expectedAuthorityGeneration: 1,
        now: () => SERVER_TIME,
    });
    await oldInstance.transitionControl({
        expectedVersion: 1,
        changes: {
            incidentCode: 'synthetic-recovery',
            incidentRecordedAt: SERVER_TIME,
            incidentState: 'suspended',
        },
    });
    const upgradingInstance = createTestSessionAuthorityStore({
        testOnly: true,
        backing,
        expectedAuthorityGeneration: 2,
        now: () => SERVER_TIME,
    });
    Object.assign(backing.state.control, {
        authorityGeneration: 2,
        globalSessionEpoch: 2,
        incidentState: 'recovering',
    });
    await upgradingInstance.transitionControl({
        expectedVersion: 2,
        changes: {
            incidentState: 'normal',
        },
    });
    await assert.rejects(
        oldInstance.readControl(),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'authority-generation-mismatch',
    );

    const recoveredInstance = createTestSessionAuthorityStore({
        testOnly: true,
        backing,
        expectedAuthorityGeneration: 2,
        now: () => SERVER_TIME,
    });
    const recovered = await recoveredInstance.readControl();
    assert.equal(recovered.control.authorityGeneration, 2);
});

test('N+1 control client may cross the generation fence only through an atomic incident resume', async () => {
    const cases = [
        {
            current: { incidentState: 'normal' },
            changes: { authorityGeneration: 4, globalSessionEpoch: 3, incidentState: 'normal' },
        },
        {
            current: { incidentState: 'suspended' },
            changes: { authorityGeneration: 4, globalSessionEpoch: 3, incidentState: 'suspended' },
        },
        {
            current: { incidentState: 'suspended' },
            changes: { authorityGeneration: 4, incidentState: 'normal' },
        },
        {
            current: { incidentState: 'suspended' },
            changes: { globalSessionEpoch: 3, incidentState: 'normal' },
        },
    ];

    for (const entry of cases) {
        const fake = createFakeSql([
            result([{ serverTime: SERVER_TIME }]),
            result([controlRow(entry.current)]),
        ]);
        const store = createAzureSqlSessionStoreFactory({
            sql: fake.sql,
            connectionString: 'synthetic-only',
            expectedAuthorityGeneration: 4,
        });
        await assert.rejects(
            store.transitionControl({ expectedVersion: 3, changes: entry.changes }),
            (error) => error.errorClass === 'authority-unavailable'
                && error.reason === 'authority-generation-mismatch',
        );
        assert.equal(fake.state.rollbacks, 1);
        assert.equal(
            fake.state.queries.some(({ statement }) => statement.includes('transition-control')),
            false,
        );
    }
});

test('recovering is reachable only through a configured key replacement transaction', async () => {
    for (const incidentState of ['normal', 'suspended']) {
        const rejected = createFakeSql([
            result([controlRow({ incidentState })]),
            result([{ serverTime: SERVER_TIME }]),
        ]);
        await assert.rejects(
            createAzureSqlSessionStore({
                sql: rejected.sql,
                connectionString: 'synthetic-only',
            }).transitionControl({
                expectedVersion: 3,
                changes: { incidentState: 'recovering' },
            }),
            (error) => error.errorClass === 'forbidden-authority'
                && error.reason === 'recovering-requires-key-recovery',
        );
        assert.equal(rejected.state.queries.some(({ statement }) => (
            statement.includes('transition-control */')
        )), false);
    }

    const suspensionTime = new Date('2042-06-01T09:00:00.000Z');
    const replacement = {
        ...AUTHORITY_KEYSET_BINDING,
        commitment: Buffer.alloc(32, 111),
        purposes: {
            ...AUTHORITY_KEYSET_BINDING.purposes,
            targetVerifier: keyBinding('synthetic-target-verifier-key-v2', 112),
        },
    };
    const fake = createFakeSql([
        result([controlRow({
            authorityGeneration: 3,
            incidentState: 'suspended',
            incidentRecordedAt: suspensionTime,
            authorityKeysetAggregateMatches: 0,
            targetVerifierKeyMatches: 0,
        })]),
        result([{ serverTime: SERVER_TIME }]),
        result([], [1]),
        result([], [1]),
        result([], [1]),
        result([controlRow({
            authorityGeneration: 4,
            controlVersion: 4,
            globalSessionEpoch: 3,
            incidentState: 'recovering',
            incidentRecordedAt: suspensionTime,
            targetVerifierKeyIncidentAt: suspensionTime,
        })]),
        result([controlRow({
            authorityGeneration: 4,
            controlVersion: 4,
            globalSessionEpoch: 3,
            incidentState: 'recovering',
            incidentRecordedAt: suspensionTime,
            targetVerifierKeyIncidentAt: suspensionTime,
        })]),
        result([{ serverTime: SERVER_TIME }]),
        result([{ liveFaceAuthorityExists: 0 }]),
        result([], [1]),
        result([controlRow({
            authorityGeneration: 4,
            controlVersion: 5,
            globalSessionEpoch: 3,
            incidentState: 'normal',
            incidentCode: 'synthetic-incident',
            incidentRecordedAt: suspensionTime,
            targetVerifierKeyIncidentAt: suspensionTime,
        })]),
    ]);
    const replacementStore = createAzureSqlSessionStoreFactory({
        sql: fake.sql,
        connectionString: 'synthetic-only',
        expectedAuthorityGeneration: 4,
        authorityKeysetBinding: replacement,
    });
    const recovering = await replacementStore.transitionControl({
        expectedVersion: 3,
        changes: {
            authorityGeneration: 4,
            globalSessionEpoch: 3,
            incidentState: 'recovering',
        },
    });
    assert.equal(recovering.control.incidentState, 'recovering');
    const recoveryMutation = fake.state.queries.find(({ statement }) => (
        statement.includes('transition-control */')
    ));
    assert.deepEqual(recoveryMutation.parameters.change_targetVerifierKeyIncidentAt, suspensionTime);
    assert.ok(fake.state.queries.some(({ statement }) => (
        statement.includes('key-recovery:quarantine-unresolved-flows')
    )));
    const revocation = fake.state.queries.find(({ statement }) => (
        statement.includes('key-recovery:revoke-unresolved-flow-sessions')
    ));
    assert.match(revocation.statement, /revocation_reason = 'key-recovery'/);
    const resumed = await replacementStore.transitionControl({
        expectedVersion: 4,
        changes: { incidentState: 'normal' },
    });
    assert.equal(resumed.control.incidentState, 'normal');
    assert.equal(resumed.control.authorityGeneration, 4);
});

test('non-Face key recovery quarantines pending flows so the new generation can resume', async () => {
    const backing = createTestSessionAuthorityBacking();
    const replacement = {
        ...AUTHORITY_KEYSET_BINDING,
        commitment: Buffer.alloc(32, 121),
        purposes: {
            ...AUTHORITY_KEYSET_BINDING.purposes,
            credentialFingerprint: keyBinding('synthetic-credential-fingerprint-key-v2', 122),
        },
    };
    Object.assign(backing.state.control, {
        version: 3,
        authorityGeneration: 3,
        globalSessionEpoch: 2,
        loginLookupKeyId: LOGIN_LOOKUP_KEY_ID,
        loginLookupKeyCommitment: Buffer.from(LOGIN_LOOKUP_KEY_COMMITMENT),
        accountMappingKeyBinding: ACCOUNT_MAPPING_KEY_BINDING,
        authorityKeysetBinding: AUTHORITY_KEYSET_BINDING,
        legacySigningKeyBinding: LEGACY_SIGNING_KEY_BINDING,
        incidentState: 'suspended',
        incidentRecordedAt: SERVER_TIME,
        incidentCode: 'synthetic-key-recovery',
    });
    const sessionId = '00000000-0000-4000-8000-000000000171';
    const flowId = '00000000-0000-4000-8000-000000000172';
    const session = sessionRow({ sessionId, phase: 'face-pending' });
    const flow = flowRow({ flowId, currentSessionId: sessionId, challengeSessionId: sessionId });
    backing.state.sessions.set(sessionId, session);
    backing.state.flows.set(sessionId, flow);
    const store = createTestSessionAuthorityStore({
        testOnly: true,
        backing,
        now: () => SERVER_TIME,
        expectedAuthorityGeneration: 4,
        loginLookupKeyId: LOGIN_LOOKUP_KEY_ID,
        loginLookupKeyCommitment: LOGIN_LOOKUP_KEY_COMMITMENT,
        accountMappingKeyBinding: ACCOUNT_MAPPING_KEY_BINDING,
        authorityKeysetBinding: replacement,
        legacySigningKeyBinding: LEGACY_SIGNING_KEY_BINDING,
    });

    const recovering = await store.transitionControl({
        expectedVersion: 3,
        changes: {
            authorityGeneration: 4,
            globalSessionEpoch: 3,
            incidentState: 'recovering',
        },
    });
    assert.equal(recovering.control.incidentState, 'recovering');
    assert.equal(flow.challengeState, 'reconciliation-required');
    assert.equal(session.phase, 'revoked');
    assert.equal(session.revocationReason, 'key-recovery');

    const resumed = await store.transitionControl({
        expectedVersion: 4,
        changes: { incidentState: 'normal' },
    });
    assert.equal(resumed.control.incidentState, 'normal');
});

test('delayed legacy-in-scope recovery anchors new key evidence to the suspension', async () => {
    const backing = createTestSessionAuthorityBacking();
    const suspensionTime = new Date('2042-06-01T09:00:00.000Z');
    const seedingStart = new Date('2042-05-31T20:00:00.000Z');
    const replacement = {
        ...AUTHORITY_KEYSET_BINDING,
        commitment: Buffer.alloc(32, 123),
        purposes: {
            ...AUTHORITY_KEYSET_BINDING.purposes,
            legacyCompatibility: keyBinding('synthetic-legacy-compatibility-key-v2', 124),
        },
    };
    Object.assign(backing.state.control, {
        version: 3,
        authorityGeneration: 3,
        globalSessionEpoch: 2,
        loginLookupKeyId: LOGIN_LOOKUP_KEY_ID,
        loginLookupKeyCommitment: Buffer.from(LOGIN_LOOKUP_KEY_COMMITMENT),
        accountMappingKeyBinding: ACCOUNT_MAPPING_KEY_BINDING,
        authorityKeysetBinding: AUTHORITY_KEYSET_BINDING,
        legacySigningKeyBinding: LEGACY_SIGNING_KEY_BINDING,
        legacyLedgerSeedingStartedAt: seedingStart,
        seedingStartedAt: seedingStart,
        incidentState: 'suspended',
        incidentRecordedAt: suspensionTime,
        incidentCode: 'synthetic-key-recovery',
    });
    const store = createTestSessionAuthorityStore({
        testOnly: true,
        backing,
        now: () => SERVER_TIME,
        expectedAuthorityGeneration: 4,
        loginLookupKeyId: LOGIN_LOOKUP_KEY_ID,
        loginLookupKeyCommitment: LOGIN_LOOKUP_KEY_COMMITMENT,
        accountMappingKeyBinding: ACCOUNT_MAPPING_KEY_BINDING,
        authorityKeysetBinding: replacement,
        legacySigningKeyBinding: LEGACY_SIGNING_KEY_BINDING,
    });

    const recovering = await store.transitionControl({
        expectedVersion: 3,
        changes: {
            authorityGeneration: 4,
            globalSessionEpoch: 3,
            incidentState: 'recovering',
        },
    });
    assert.deepEqual(recovering.control.legacyVerifierKeyIncidentAt, suspensionTime);
    assert.deepEqual(recovering.control.legacyStopIssuanceAt, SERVER_TIME);
    assert.deepEqual(recovering.control.legacyAcceptanceDisabledAt, SERVER_TIME);

    const resumed = await store.transitionControl({
        expectedVersion: 4,
        changes: { incidentState: 'normal' },
    });
    assert.equal(resumed.control.incidentState, 'normal');
});

test('readControl connects lazily, reuses the pool, and maps database control values', async () => {
    const fake = createFakeSql([
        result([controlRow()]),
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({ controlVersion: 9 })]),
        result([{ serverTime: SERVER_TIME }]),
    ]);
    const store = createAzureSqlSessionStore({
        sql: fake.sql,
        connectionString: 'synthetic-only',
        options: {
            connectionTimeout: 4_000,
            options: { useUTC: false },
            requestTimeout: 6_000,
            pool: { idleTimeoutMillis: 8_000, max: 5, min: 0 },
        },
    });

    const first = await store.readControl();
    const second = await store.readControl();

    assert.equal(fake.state.poolConstructions, 1);
    assert.equal(fake.state.connects, 1);
    assert.deepEqual(fake.state.constructorConfigurations[0], {
        connectionTimeout: 4_000,
        database: 'inert',
        options: { encrypt: true, trustServerCertificate: false, useUTC: true },
        pool: { idleTimeoutMillis: 8_000, max: 5, min: 0 },
        requestTimeout: 6_000,
        server: 'synthetic.invalid',
    });
    assert.equal(first.control.version, 3);
    assert.equal(second.control.version, 9);
    assert.deepEqual(first.serverTime, SERVER_TIME);
    assert.match(fake.state.queries[1].statement, /SYSUTCDATETIME\(\)/);
    assert.doesNotMatch(fake.state.queries[0].statement, /synthetic-only/);
});

test('case-varied lookup key ownership fails before subject access or mutation', async () => {
    const configuredKeyId = 'Synthetic-Lookup-Key';
    const fake = createFakeSql([
        result([controlRow({ loginLookupKeyMatches: 0 })]),
    ]);
    const store = createAzureSqlSessionStoreFactoryRaw({
        sql: fake.sql,
        connectionString: 'synthetic-only',
        expectedAuthorityGeneration: EXPECTED_AUTHORITY_GENERATION,
        loginLookupKeyId: configuredKeyId,
        loginLookupKeyCommitment: LOGIN_LOOKUP_KEY_COMMITMENT,
        accountMappingKeyBinding: ACCOUNT_MAPPING_KEY_BINDING,
        legacySigningKeyBinding: LEGACY_SIGNING_KEY_BINDING,
        authorityKeysetBinding: {
            ...AUTHORITY_KEYSET_BINDING,
            purposes: {
                ...AUTHORITY_KEYSET_BINDING.purposes,
                loginLookup: {
                    keyId: configuredKeyId,
                    commitment: LOGIN_LOOKUP_KEY_COMMITMENT,
                },
            },
        },
    });

    await assert.rejects(
        store.readSubjectByLookup({
            loginLookupKeyId: configuredKeyId,
            loginLookupToken: Buffer.alloc(32, 93),
        }),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'login-lookup-key-mismatch',
    );
    assert.equal(fake.state.rollbacks, 1);
    assert.equal(fake.state.queries.length, 1);
    assert.match(fake.state.queries[0].statement, /COLLATE Latin1_General_100_BIN2/);
    assert.equal(
        fake.state.queries.some(({ statement }) => statement.includes('read-subject-by-lookup')),
        false,
    );
});

test('legacy signing ownership is independently fenced by exact ID and material', async () => {
    for (const binding of [
        { keyId: LEGACY_SIGNING_KEY_BINDING.keyId, commitment: Buffer.alloc(32, 121) },
        { keyId: 'synthetic-legacy-signing-key-v2', commitment: LEGACY_SIGNING_KEY_BINDING.commitment },
    ]) {
        const fake = createFakeSql([
            result([controlRow({ legacySigningKeyMatches: 0 })]),
        ]);
        const store = createAzureSqlSessionStoreFactory({
            sql: fake.sql,
            connectionString: 'synthetic-only',
            expectedAuthorityGeneration: EXPECTED_AUTHORITY_GENERATION,
            legacySigningKeyBinding: binding,
        });
        await assert.rejects(
            store.readControl(),
            (error) => error.errorClass === 'authority-unavailable'
                && error.reason === 'authority-keyset-mismatch',
        );
        assert.equal(fake.state.commits, 0);
        assert.equal(fake.state.rollbacks, 1);
    }
});

test('lookup key binding is explicit, empty-only, immutable, and commitment-private', async () => {
    const uninitialized = controlRow({
        loginLookupKeyInitialized: 0,
        loginLookupKeyMatches: 0,
    });
    const initialized = controlRow({ controlVersion: 4 });
    const fake = createFakeSql([
        result([uninitialized]),
        result([{ authorityDataExists: 0 }]),
        result([{ serverTime: SERVER_TIME }]),
        result([], [1]),
        result([initialized]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    const output = await store.initializeLoginLookupKey({
        loginLookupKeyId: LOGIN_LOOKUP_KEY_ID,
        loginLookupKeyCommitment: LOGIN_LOOKUP_KEY_COMMITMENT,
    });
    assert.equal(output.idempotent, false);
    assert.equal(output.control.version, 4);
    assert.equal(Object.hasOwn(output.control, 'loginLookupKeyCommitment'), false);
    const mutation = fake.state.queries.find(({ statement }) => (
        statement.includes('initialize-login-lookup-key */')
    ));
    assert.ok(mutation);
    assert.match(mutation.statement, /login_lookup_key_id IS NULL/);
    assert.match(mutation.statement, /login_lookup_key_commitment IS NULL/);

    const blocked = createFakeSql([
        result([uninitialized]),
        result([{ authorityDataExists: 1 }]),
    ]);
    await assert.rejects(
        createAzureSqlSessionStore({
            sql: blocked.sql,
            connectionString: 'synthetic-only',
        }).initializeLoginLookupKey({
            loginLookupKeyId: LOGIN_LOOKUP_KEY_ID,
            loginLookupKeyCommitment: LOGIN_LOOKUP_KEY_COMMITMENT,
        }),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'login-lookup-key-initialization-blocked',
    );
    assert.equal(blocked.state.queries.some(({ statement }) => (
        statement.includes('initialize-login-lookup-key */')
    )), false);

    const raced = createFakeSql([
        result([uninitialized]),
        result([{ authorityDataExists: 0 }]),
        result([{ serverTime: SERVER_TIME }]),
        result([], [0]),
    ]);
    await assert.rejects(
        createAzureSqlSessionStore({
            sql: raced.sql,
            connectionString: 'synthetic-only',
        }).initializeLoginLookupKey({
            loginLookupKeyId: LOGIN_LOOKUP_KEY_ID,
            loginLookupKeyCommitment: LOGIN_LOOKUP_KEY_COMMITMENT,
        }),
        (error) => error.errorClass === 'authority-conflict'
            && error.reason === 'login-lookup-key-initialization-race',
    );

    const nullFence = createFakeSql([result([uninitialized])]);
    await assert.rejects(
        createAzureSqlSessionStore({
            sql: nullFence.sql,
            connectionString: 'synthetic-only',
        }).readControl(),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'login-lookup-key-uninitialized',
    );
});

test('legacy seeding heartbeat resets joins and failures without changing policy version', async () => {
    const ownerId = '00000000-0000-4000-8000-000000000091';
    const secondTime = new Date(SERVER_TIME.getTime() + 10_000);
    const recoveryTime = new Date(SERVER_TIME.getTime() + 20_000);
    const firstControl = controlRow({
        legacyLedgerSeedingEnabled: 1,
        seedingStartedAt: new Date('2042-06-01T11:00:00.000Z'),
        seedingHeartbeatOwnerId: null,
        seedingHeartbeatAt: null,
        seedingLeaseExpiresAt: null,
    });
    const claimed = controlRow({
        legacyLedgerSeedingEnabled: 1,
        seedingContinuityVersion: 2,
        seedingStartedAt: SERVER_TIME,
        seedingHeartbeatOwnerId: ownerId,
        seedingHeartbeatAt: SERVER_TIME,
        seedingLeaseExpiresAt: new Date(SERVER_TIME.getTime() + 120_000),
    });
    const recovered = controlRow({
        legacyLedgerSeedingEnabled: 1,
        seedingContinuityVersion: 3,
        seedingStartedAt: recoveryTime,
        seedingHeartbeatOwnerId: ownerId,
        seedingHeartbeatAt: recoveryTime,
        seedingLeaseExpiresAt: new Date(recoveryTime.getTime() + 120_000),
    });
    const fake = createFakeSql([
        result([firstControl]),
        result([{ serverTime: SERVER_TIME }]),
        result([], [1]),
        result([claimed]),
        result([claimed]),
        result([]),
        result([{ serverTime: secondTime }]),
        result([claimed]),
        result([{ serverTime: recoveryTime }]),
        result([], [1]),
        result([recovered]),
    ], {
        commitOutcomes: [undefined, new Error('synthetic unknown commit'), undefined],
    });
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    const first = await store.heartbeatLegacySeedingContinuity({ ownerId });
    assert.equal(first.reset, true);
    assert.equal(first.control.version, 3);
    assert.equal(first.control.seedingContinuityVersion, 2);
    await assert.rejects(
        store.readSubjectByLookup({
            loginLookupKeyId: LOGIN_LOOKUP_KEY_ID,
            loginLookupToken: Buffer.alloc(32, 72),
        }),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'transaction-outcome-unknown',
    );
    const afterFailure = await store.heartbeatLegacySeedingContinuity({ ownerId });
    assert.equal(afterFailure.reset, true);
    assert.deepEqual(afterFailure.control.seedingStartedAt, recoveryTime);
    assert.equal(afterFailure.control.version, 3);

    const heartbeatMutations = fake.state.queries.filter(({ statement }) => (
        statement.includes('heartbeat-legacy-seeding-continuity */')
    ));
    assert.equal(heartbeatMutations.length, 2);
    assert.equal(heartbeatMutations.every(({ parameters }) => parameters.reset === true), true);
    assert.equal(heartbeatMutations.every(({ statement }) => (
        !/control_version\s*=\s*control_version\s*\+/u.test(statement)
    )), true);
    for (const mutation of heartbeatMutations) {
        assert.deepEqual(mutation.parameterTypes.serverTime, { name: 'DateTime2', scale: 7 });
        assert.deepEqual(mutation.parameterTypes.leaseExpiresAt, { name: 'DateTime2', scale: 7 });
        assert.deepEqual(mutation.parameterTypes.ownerId, { name: 'UniqueIdentifier' });
        assert.deepEqual(mutation.parameterTypes.controlVersion, { name: 'BigInt' });
        assert.deepEqual(mutation.parameterTypes.reset, { name: 'Bit' });
    }
    const fencedControlRead = fake.state.queries.find(({ statement }) => (
        statement.includes('select-control')
    ));
    assert.deepEqual(fencedControlRead.parameterTypes.loginLookupKeyId, {
        length: 128,
        name: 'VarChar',
    });
    assert.deepEqual(fencedControlRead.parameterTypes.loginLookupKeyCommitment, {
        length: 32,
        name: 'Binary',
    });
});

test('test stores reset continuity on process joins, stale leases, and local outages', async () => {
    let clock = new Date('2042-06-01T08:00:00.000Z');
    const backing = createTestSessionAuthorityBacking();
    const failureA = {};
    const createStore = (failure = {}) => createTestSessionAuthorityStore({
        testOnly: true,
        backing,
        failure,
        now: () => clock,
    });
    const storeA = createStore(failureA);
    const storeB = createStore();
    const started = await storeA.transitionControl({
        expectedVersion: 1,
        changes: { legacyLedgerSeedingEnabled: true },
    });
    const policyVersion = started.control.version;
    const ownerA = '00000000-0000-4000-8000-000000000093';
    const ownerB = '00000000-0000-4000-8000-000000000094';

    const first = await storeA.heartbeatLegacySeedingContinuity({ ownerId: ownerA });
    assert.equal(first.reset, true);
    clock = new Date(clock.getTime() + 30_000);
    const renewal = await storeA.heartbeatLegacySeedingContinuity({ ownerId: ownerA });
    assert.equal(renewal.reset, false);
    assert.equal(renewal.control.version, policyVersion);

    clock = new Date(clock.getTime() + 1_000);
    const join = await storeB.heartbeatLegacySeedingContinuity({ ownerId: ownerB });
    assert.equal(join.reset, true);
    assert.deepEqual(join.control.seedingStartedAt, clock);
    assert.equal(join.control.version, policyVersion);

    clock = new Date(clock.getTime() + 120_000);
    const stale = await storeB.heartbeatLegacySeedingContinuity({ ownerId: ownerB });
    assert.equal(stale.reset, true);
    assert.deepEqual(stale.control.seedingStartedAt, clock);

    failureA.unavailable = true;
    await assert.rejects(
        storeA.readControl(),
        (error) => error.reason === 'session-store-unavailable',
    );
    failureA.unavailable = false;
    clock = new Date(clock.getTime() + 1_000);
    const recovered = await storeA.heartbeatLegacySeedingContinuity({ ownerId: ownerA });
    assert.equal(recovered.reset, true);
    assert.deepEqual(recovered.control.seedingStartedAt, clock);
});

test('an incident breaks ledger continuity until a fresh normal heartbeat restarts the horizon', async () => {
    let clock = new Date(SERVER_TIME);
    const backing = createTestSessionAuthorityBacking();
    const store = createTestSessionAuthorityStore({
        testOnly: true,
        backing,
        now: () => clock,
    });
    await store.transitionControl({
        expectedVersion: 1,
        changes: { legacyLedgerSeedingEnabled: true },
    });
    const ownerId = '00000000-0000-4000-8000-0000000000aa';
    await store.heartbeatLegacySeedingContinuity({ ownerId });
    clock = new Date(clock.getTime() + 3 * 60 * 60 * 1000);
    await store.heartbeatLegacySeedingContinuity({ ownerId });
    const priorContinuityVersion = backing.state.control.seedingContinuityVersion;
    const suspended = await store.transitionControl({
        expectedVersion: 2,
        changes: {
            incidentState: 'suspended',
            incidentRecordedAt: clock,
            incidentCode: 'synthetic-short-incident',
        },
    });
    assert.deepEqual(suspended.control.seedingStartedAt, clock);
    assert.equal(
        suspended.control.seedingContinuityVersion,
        priorContinuityVersion + 1,
    );
    assert.equal(suspended.control.seedingHeartbeatOwnerId, null);
    assert.equal(suspended.control.seedingHeartbeatAt, null);
    assert.equal(suspended.control.seedingLeaseExpiresAt, null);

    backing.state.control.incidentState = 'recovering';
    backing.state.control.incidentState = 'normal';
    await assert.rejects(
        store.transitionControl({
            expectedVersion: 3,
            changes: { seedingQualifiedAt: clock },
        }),
        (error) => error.errorClass === 'forbidden-authority'
            && error.reason === 'legacy-ledger-not-qualified',
    );
});

test('SQL incident entry atomically invalidates a live seeding lease without waiting for heartbeat', async () => {
    const horizonStart = new Date(SERVER_TIME.getTime() - 3 * 60 * 60 * 1000);
    const ownerId = '00000000-0000-4000-8000-0000000000ab';
    const current = controlRow({
        legacyLedgerSeedingEnabled: 1,
        legacyLedgerSeedingStartedAt: horizonStart,
        seedingStartedAt: horizonStart,
        seedingContinuityVersion: 7,
        seedingHeartbeatOwnerId: ownerId,
        seedingHeartbeatAt: new Date(SERVER_TIME.getTime() - 1_000),
        seedingLeaseExpiresAt: new Date(SERVER_TIME.getTime() + 119_000),
    });
    const updated = controlRow({
        ...current,
        controlVersion: 4,
        incidentState: 'suspended',
        incidentRecordedAt: SERVER_TIME,
        incidentCode: 'synthetic-short-incident',
        seedingStartedAt: SERVER_TIME,
        seedingContinuityVersion: 8,
        seedingHeartbeatOwnerId: null,
        seedingHeartbeatAt: null,
        seedingLeaseExpiresAt: null,
    });
    const fake = createFakeSql([
        result([current]),
        result([{ serverTime: SERVER_TIME }]),
        result([], [1]),
        result([updated]),
    ]);
    await createAzureSqlSessionStore({
        sql: fake.sql,
        connectionString: 'synthetic-only',
    }).transitionControl({
        expectedVersion: 3,
        changes: {
            incidentState: 'suspended',
            incidentRecordedAt: SERVER_TIME,
            incidentCode: 'synthetic-short-incident',
        },
    });
    const mutation = fake.state.queries.find(({ statement }) => (
        statement.includes('transition-control */')
    ));
    assert.equal(mutation.parameters.change_seedingContinuityVersion, 8);
    assert.deepEqual(mutation.parameters.change_seedingStartedAt, SERVER_TIME);
    assert.equal(mutation.parameters.change_seedingHeartbeatOwnerId, null);
    assert.equal(mutation.parameters.change_seedingHeartbeatAt, null);
    assert.equal(mutation.parameters.change_seedingLeaseExpiresAt, null);
});

test('pool error events are handled privately and make later operations fail closed', async () => {
    const fake = createFakeSql([
        result([controlRow()]),
        result([{ serverTime: SERVER_TIME }]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });
    await store.readControl();

    assert.equal(typeof fake.state.poolErrorListener, 'function');
    fake.state.poolErrorListener(new Error('connection secret must stay private'));

    await assert.rejects(store.readControl(), (error) => {
        assert.equal(error.errorClass, 'authority-unavailable');
        assert.equal(error.reason, 'session-store-unavailable');
        assert.doesNotMatch(error.message, /connection|secret|private/);
        return true;
    });
    assert.equal(fake.state.queries.length, 2);
});

test('control transitions are serializable, parameterized, and committed once', async () => {
    const incidentCode = 'synthetic-control-incident';
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
        result([], [1]),
        result([controlRow({
            controlVersion: 4,
            incidentState: 'suspended',
            incidentRecordedAt: SERVER_TIME,
            incidentCode,
        })]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    const output = await store.transitionControl({
        expectedVersion: 3,
        changes: {
            incidentState: 'suspended',
            incidentRecordedAt: SERVER_TIME,
            incidentCode,
        },
    });

    assert.equal(output.control.version, 4);
    assert.deepEqual(fake.state.begins, ['SERIALIZABLE']);
    assert.equal(fake.state.commits, 1);
    assert.equal(fake.state.rollbacks, 0);
    const mutation = fake.state.queries.find(({ statement }) => statement.includes('transition-control'));
    assert.ok(mutation);
    assert.match(mutation.statement, /incident_state = @change_incidentState/);
    assert.match(mutation.statement, /control_version = control_version \+ 1/);
    assert.doesNotMatch(mutation.statement, /authority_generation = authority_generation \+ 1/);
    assert.equal(mutation.parameters.change_incidentState, 'suspended');
    assert.equal(mutation.parameters.change_incidentCode, incidentCode);
    assert.equal(mutation.parameters.expectedVersion, 3);
});

test('compare-and-replace misses roll back and preserve the domain conflict', async () => {
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    await assert.rejects(
        store.transitionControl({ expectedVersion: 8, changes: { targetRoutesEnabled: true } }),
        (error) => {
            assert.equal(error.name, 'SessionAuthorityError');
            assert.equal(error.errorClass, 'authority-conflict');
            assert.equal(error.reason, 'control-compare-and-replace');
            assert.doesNotMatch(error.message, /8|targetRoutesEnabled/);
            return true;
        },
    );
    assert.equal(fake.state.commits, 0);
    assert.equal(fake.state.rollbacks, 1);
});

test('incident codes reject free text and secret-like values before SQL', async () => {
    for (const incidentCode of [
        'Contains Uppercase',
        'credential=private',
        '-missing-prefix',
        'x'.repeat(129),
    ]) {
        const fake = createFakeSql();
        const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });
        await assert.rejects(
            store.transitionControl({ expectedVersion: 3, changes: { incidentCode } }),
            /privacy-safe machine value/,
        );
        assert.equal(fake.state.poolConstructions, 0);
        assert.equal(fake.state.connects, 0);
    }
});

test('control epochs stay monotonic and ledger qualification requires a completed four-hour horizon', async () => {
    const regressed = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
    ]);
    const regressedStore = createAzureSqlSessionStore({
        sql: regressed.sql,
        connectionString: 'synthetic-only',
    });
    await assert.rejects(
        regressedStore.transitionControl({ expectedVersion: 3, changes: { authorityGeneration: 2 } }),
        (error) => error.errorClass === 'forbidden-authority'
            && error.reason === 'irreversible-authorityGeneration',
    );
    assert.equal(regressed.state.rollbacks, 1);

    const premature = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({
            legacyLedgerSeedingEnabled: 1,
            legacyLedgerSeedingStartedAt: new Date('2042-06-01T08:00:00.001Z'),
            seedingStartedAt: new Date('2042-06-01T08:00:00.001Z'),
        })]),
    ]);
    const prematureStore = createAzureSqlSessionStore({
        sql: premature.sql,
        connectionString: 'synthetic-only',
    });
    await assert.rejects(
        prematureStore.transitionControl({
            expectedVersion: 3,
            changes: { seedingQualifiedAt: new Date('2042-06-01T11:59:59.999Z') },
        }),
        (error) => error.errorClass === 'forbidden-authority'
            && error.reason === 'legacy-ledger-not-qualified',
    );
    assert.equal(premature.state.rollbacks, 1);

    const restarted = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({
            legacyLedgerSeedingEnabled: 1,
            legacyLedgerSeedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
            seedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
            seedingQualifiedAt: new Date('2042-06-01T08:00:00.000Z'),
        })]),
        result([], [1]),
        result([controlRow({
            controlVersion: 4,
            legacyLedgerSeedingEnabled: 1,
            legacyLedgerSeedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
            seedingStartedAt: SERVER_TIME,
            seedingQualifiedAt: new Date('2042-06-01T08:00:00.000Z'),
        })]),
    ]);
    const restartedStore = createAzureSqlSessionStore({
        sql: restarted.sql,
        connectionString: 'synthetic-only',
    });
    const restart = await restartedStore.transitionControl({
        expectedVersion: 3,
        changes: { seedingStartedAt: new Date('2042-06-01T10:00:00.000Z') },
    });
    assert.equal(restart.control.version, 4);
    assert.equal(restarted.state.commits, 1);
    const restartMutation = restarted.state.queries.find(({ statement }) => (
        statement.includes('transition-control')
    ));
    assert.deepEqual(restartMutation.parameters.change_seedingStartedAt, SERVER_TIME);
});

test('target controls require qualified irreversible legacy enforcement and continued seeding', async () => {
    for (const { changes, expectedReason } of [
        {
            changes: {
                targetRoutesEnabled: true,
                targetSessionIssuanceEnabled: true,
                subjectTargetAdoptionEnabled: true,
            },
            expectedReason: 'legacy-enforcement-required-before-target',
        },
        {
            changes: { subjectTargetAdoptionEnabled: true },
            expectedReason: 'target-activation-pair-required',
        },
        {
            changes: { dualStackStartedAt: SERVER_TIME },
            expectedReason: 'target-activation-evidence-integrity',
        },
    ]) {
        const fake = createFakeSql([
            result([{ serverTime: SERVER_TIME }]),
            result([controlRow()]),
        ]);
        const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });
        await assert.rejects(
            store.transitionControl({ expectedVersion: 3, changes }),
            (error) => error.errorClass === 'forbidden-authority'
                && error.reason === expectedReason,
        );
        assert.equal(fake.state.rollbacks, 1);
    }

    const seedingStopped = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({
            legacyCompatibilityEnforcementEnabled: 1,
            legacyCompatibilityEnforcedAt: new Date('2042-06-01T08:00:00.000Z'),
            legacyLedgerSeedingEnabled: 1,
            seedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
            seedingQualifiedAt: new Date('2042-06-01T08:00:00.000Z'),
        })]),
    ]);
    const store = createAzureSqlSessionStore({
        sql: seedingStopped.sql,
        connectionString: 'synthetic-only',
    });
    await assert.rejects(
        store.transitionControl({
            expectedVersion: 3,
            changes: { legacyLedgerSeedingEnabled: false },
        }),
        (error) => error.errorClass === 'forbidden-authority'
            && error.reason === 'legacy-seeding-required-during-issuance',
    );
    assert.equal(seedingStopped.state.rollbacks, 1);
});

test('target activation is paired and SQL-stamps the fixed seven-day dual-stack clock', async () => {
    const qualified = {
        legacyCompatibilityEnforcementEnabled: 1,
        legacyCompatibilityEnforcedAt: new Date('2042-06-01T08:00:00.000Z'),
        legacyLedgerSeedingEnabled: 1,
        legacyLedgerSeedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
        seedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
        seedingQualifiedAt: new Date('2042-06-01T08:00:00.000Z'),
    };
    const rejected = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow(qualified)]),
    ]);
    await assert.rejects(
        createAzureSqlSessionStore({
            sql: rejected.sql,
            connectionString: 'synthetic-only',
        }).transitionControl({
            expectedVersion: 3,
            changes: { targetSessionIssuanceEnabled: true },
        }),
        (error) => error.errorClass === 'forbidden-authority'
            && error.reason === 'target-activation-pair-required',
    );

    const activatedRow = controlRow({
        ...qualified,
        controlVersion: 4,
        targetRoutesEnabled: 1,
        targetSessionIssuanceEnabled: 1,
        subjectTargetAdoptionEnabled: 1,
        targetSessionIssuanceStartedAt: SERVER_TIME,
        subjectTargetAdoptionStartedAt: SERVER_TIME,
        dualStackStartedAt: SERVER_TIME,
        hardSunsetAt: new Date(SERVER_TIME.getTime() + 7 * 24 * 60 * 60 * 1000),
    });
    const accepted = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow(qualified)]),
        result([], [1]),
        result([activatedRow]),
    ]);
    const resultValue = await createAzureSqlSessionStore({
        sql: accepted.sql,
        connectionString: 'synthetic-only',
    }).transitionControl({
        expectedVersion: 3,
        changes: {
            targetRoutesEnabled: true,
            targetSessionIssuanceEnabled: true,
            subjectTargetAdoptionEnabled: true,
            dualStackStartedAt: new Date('2042-01-01T00:00:00.000Z'),
            hardSunsetAt: new Date('2042-01-08T00:00:00.000Z'),
        },
    });
    assert.equal(resultValue.control.targetSessionIssuanceEnabled, true);
    const mutation = accepted.state.queries.find(({ statement }) => (
        statement.includes('transition-control */')
    ));
    assert.deepEqual(mutation.parameters.change_targetSessionIssuanceStartedAt, SERVER_TIME);
    assert.deepEqual(mutation.parameters.change_subjectTargetAdoptionStartedAt, SERVER_TIME);
    assert.deepEqual(mutation.parameters.change_dualStackStartedAt, SERVER_TIME);
    assert.deepEqual(
        mutation.parameters.change_hardSunsetAt,
        new Date(SERVER_TIME.getTime() + 7 * 24 * 60 * 60 * 1000),
    );
});

test('target controls remain operable after the legacy hard sunset', async () => {
    const postSunsetControl = {
        legacyCompatibilityEnforcementEnabled: 1,
        legacyCompatibilityEnforcedAt: new Date('2042-05-25T08:00:00.000Z'),
        legacyLedgerSeedingEnabled: 1,
        legacyLedgerSeedingStartedAt: new Date('2042-05-25T04:00:00.000Z'),
        seedingStartedAt: new Date('2042-05-25T04:00:00.000Z'),
        seedingQualifiedAt: new Date('2042-05-25T08:00:00.000Z'),
        subjectTargetAdoptionEnabled: 1,
        subjectTargetAdoptionStartedAt: new Date('2042-05-25T12:00:00.000Z'),
        targetRoutesEnabled: 1,
        targetSessionIssuanceEnabled: 1,
        targetSessionIssuanceStartedAt: new Date('2042-05-25T12:00:00.000Z'),
        dualStackStartedAt: new Date('2042-05-25T12:00:00.000Z'),
        hardSunsetAt: new Date('2042-06-01T12:00:00.000Z'),
        legacyIssuanceEnabled: 0,
        legacyStopIssuanceAt: new Date('2042-06-01T08:00:00.000Z'),
        legacyAcceptanceEnabled: 0,
        legacyAcceptanceDisabledAt: new Date('2042-06-01T12:00:00.000Z'),
    };
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow(postSunsetControl)]),
        result([], [1]),
        result([controlRow({ ...postSunsetControl, controlVersion: 4 })]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    const output = await store.transitionControl({
        expectedVersion: 3,
        changes: { targetRoutesEnabled: true },
    });
    assert.equal(output.control.targetRoutesEnabled, true);
    assert.equal(fake.state.commits, 1);
});

test('session issuance rechecks the active target window against SQL transaction time', async () => {
    const subjectId = '00000000-0000-4000-8000-000000000053';
    const baseControl = {
        legacyCompatibilityEnforcementEnabled: 1,
        legacyCompatibilityEnforcedAt: new Date('2042-05-25T12:00:00.000Z'),
        legacyLedgerSeedingEnabled: 1,
        legacyLedgerSeedingStartedAt: new Date('2042-05-25T08:00:00.000Z'),
        seedingStartedAt: new Date('2042-05-25T08:00:00.000Z'),
        seedingQualifiedAt: new Date('2042-05-25T12:00:00.000Z'),
        subjectTargetAdoptionEnabled: 1,
        subjectTargetAdoptionStartedAt: new Date('2042-05-25T12:00:00.000Z'),
        targetRoutesEnabled: 1,
        targetSessionIssuanceEnabled: 1,
        targetSessionIssuanceStartedAt: new Date('2042-05-25T12:00:00.000Z'),
    };
    const cases = [
        {
            expectedReason: 'target-session-window-inactive',
            window: {
                targetSessionIssuanceStartedAt: new Date('2042-06-01T12:00:00.001Z'),
                subjectTargetAdoptionStartedAt: new Date('2042-06-01T12:00:00.001Z'),
                dualStackStartedAt: new Date('2042-06-01T12:00:00.001Z'),
                hardSunsetAt: new Date('2042-06-08T12:00:00.001Z'),
            },
        },
    ];

    for (const [index, entry] of cases.entries()) {
        const fake = createFakeSql([
            result([{ serverTime: SERVER_TIME }]),
            result([subjectRow({ subjectId })]),
            result([controlRow({ ...baseControl, ...entry.window })]),
        ]);
        const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });
        await assert.rejects(
            store.issueSession({
                sessionId: `00000000-0000-4000-8000-00000000015${index}`,
                subjectId,
                expectedCredentialVersion: 2,
                expectedCredentialFingerprintKeyId: 'synthetic-fingerprint-key',
                expectedCredentialFingerprint: Buffer.alloc(32, 2),
                verifierKeyId: 'synthetic-target-key',
                verifier: Buffer.alloc(32, 20 + index),
                phase: 'authenticated',
                faceRequired: false,
                registrationRequired: false,
            }),
            (error) => error.errorClass === 'authority-unavailable'
                && error.reason === entry.expectedReason,
        );
        assert.equal(fake.state.rollbacks, 1);
        assert.equal(fake.state.queries.length, 3);
    }
});

test('target session issuance continues after the legacy hard sunset', async () => {
    const subjectId = '00000000-0000-4000-8000-000000000054';
    const sessionId = '00000000-0000-4000-8000-000000000158';
    const postSunsetControl = {
        legacyCompatibilityEnforcementEnabled: 1,
        legacyCompatibilityEnforcedAt: new Date('2042-05-25T08:00:00.000Z'),
        legacyLedgerSeedingEnabled: 1,
        legacyLedgerSeedingStartedAt: new Date('2042-05-25T04:00:00.000Z'),
        seedingStartedAt: new Date('2042-05-25T04:00:00.000Z'),
        seedingQualifiedAt: new Date('2042-05-25T08:00:00.000Z'),
        subjectTargetAdoptionEnabled: 1,
        subjectTargetAdoptionStartedAt: new Date('2042-05-25T12:00:00.000Z'),
        targetRoutesEnabled: 1,
        targetSessionIssuanceEnabled: 1,
        targetSessionIssuanceStartedAt: new Date('2042-05-25T12:00:00.000Z'),
        dualStackStartedAt: new Date('2042-05-25T12:00:00.000Z'),
        hardSunsetAt: new Date('2042-06-01T12:00:00.000Z'),
        legacyIssuanceEnabled: 0,
        legacyAcceptanceEnabled: 0,
    };
    const subject = subjectRow({
        subjectId,
        legacyAuthorityDisabledAt: new Date('2042-05-25T12:00:00.000Z'),
    });
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([subject]),
        result([controlRow(postSunsetControl)]),
        result([], [1]),
        result([sessionRow({ subjectId, sessionId })]),
        result([subject]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    const output = await store.issueSession({
        sessionId,
        subjectId,
        expectedCredentialVersion: 2,
        expectedCredentialFingerprintKeyId: 'synthetic-fingerprint-key',
        expectedCredentialFingerprint: Buffer.alloc(32, 2),
        verifierKeyId: 'synthetic-target-key',
        verifier: Buffer.alloc(32, 10),
        phase: 'authenticated',
        faceRequired: false,
        registrationRequired: false,
    });
    assert.equal(output.session.sessionId, sessionId);
    assert.equal(fake.state.commits, 1);
});

test('session issuance locks the subject and rejects a changed credential snapshot before minting', async () => {
    const subjectId = subjectRow().subjectId;
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([subjectRow({ subjectId, credentialVersion: 3 })]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    await assert.rejects(
        store.issueSession({
            sessionId: '00000000-0000-4000-8000-000000000159',
            subjectId,
            expectedCredentialVersion: 2,
            expectedCredentialFingerprintKeyId: 'synthetic-fingerprint-key',
            expectedCredentialFingerprint: Buffer.alloc(32, 2),
            verifierKeyId: 'synthetic-target-key',
            verifier: Buffer.alloc(32, 9),
            phase: 'authenticated',
            faceRequired: false,
            registrationRequired: false,
        }),
        (error) => error.errorClass === 'authority-conflict'
            && error.reason === 'subject-credential-compare-and-replace',
    );
    assert.equal(fake.state.commits, 0);
    assert.equal(fake.state.rollbacks, 1);
    const subjectRead = fake.state.queries.find(({ statement }) => statement.includes('select-subject'));
    assert.match(subjectRead.statement, /WITH \(UPDLOCK, HOLDLOCK\)/);
    assert.equal(fake.state.queries.some(({ statement }) => statement.includes('insert-session')), false);
});

test('unbound legacy issuance admission shares control while locking its lookup range', async () => {
    const issuedAt = new Date('2042-06-01T11:00:00.000Z');
    const expiresAt = new Date('2042-06-01T15:00:00.000Z');
    const lookupToken = Buffer.alloc(32, 11);
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
        result([]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    const output = await store.admitUnboundLegacyIssuance({
        loginLookupKeyId: 'synthetic-lookup-key',
        loginLookupToken: lookupToken,
        issuedAt,
        expiresAt,
    });

    assert.equal(output.admitted, true);
    assert.equal(output.control.version, 3);
    assert.deepEqual(output.serverTime, SERVER_TIME);
    assert.equal(fake.state.commits, 1);
    assert.deepEqual(fake.state.begins, ['SERIALIZABLE']);
    const controlQuery = fake.state.queries.find(({ statement }) => (
        statement.includes('select-control')
    ));
    const lookupQuery = fake.state.queries.find(({ statement }) => (
        statement.includes('admit-unbound-legacy-issuance')
    ));
    assert.match(controlQuery.statement, /WITH \(HOLDLOCK\)/);
    assert.doesNotMatch(controlQuery.statement, /UPDLOCK/);
    assert.match(lookupQuery.statement, /WITH \(UPDLOCK, HOLDLOCK\)/);
    assert.equal(lookupQuery.parameters.loginLookupKeyId, 'synthetic-lookup-key');
    assert.deepEqual(lookupQuery.parameters.loginLookupToken, lookupToken);
    assert.doesNotMatch(lookupQuery.statement, /synthetic-lookup-key/);
});

test('legacy bind and unbound admission stop at the SQL four-hour final-aging boundary', async () => {
    const hardSunsetAt = new Date('2042-06-01T16:00:00.000Z');
    const issuedAt = new Date('2042-06-01T11:00:00.000Z');
    const expiresAt = new Date('2042-06-01T15:00:00.000Z');
    const cases = [
        { expectedQueries: 3, invoke: (store) => store.admitUnboundLegacyIssuance({
            loginLookupKeyId: 'synthetic-lookup-key',
            loginLookupToken: Buffer.alloc(32, 14),
            issuedAt,
            expiresAt,
        }), control: {} },
        { expectedQueries: 4, invoke: (store) => store.bindLegacy({
            legacyCompatibilityId: '00000000-0000-4000-8000-000000000350',
            subjectId: subjectRow().subjectId,
            ...subjectCredentialExpectation(),
            verifierKeyId: 'synthetic-legacy-key',
            verifier: Buffer.alloc(32, 15),
            issuedAt,
            expiresAt,
        }), control: {
            legacyCompatibilityEnforcementEnabled: 1,
            legacyLedgerSeedingEnabled: 1,
        } },
    ];

    for (const entry of cases) {
        const fake = createFakeSql([
            result([{ serverTime: SERVER_TIME }]),
            result([controlRow({
                hardSunsetAt,
                ...entry.control,
            })]),
        ]);
        const store = createAzureSqlSessionStore({
            sql: fake.sql,
            connectionString: 'synthetic-only',
        });
        await assert.rejects(
            entry.invoke(store),
            (error) => error.errorClass === 'authority-conflict'
                && error.reason === 'legacy-issuance-disabled',
        );
        assert.equal(fake.state.commits, 0);
        assert.equal(fake.state.rollbacks, 1);
        assert.equal(fake.state.queries.length, entry.expectedQueries);
    }
});

test('legacy bind and unbound admission reject clock-skewed handles extending past sunset', async () => {
    const hardSunsetAt = new Date('2042-06-01T17:00:00.000Z');
    const issuedAt = new Date('2042-06-01T13:00:00.001Z');
    const expiresAt = new Date('2042-06-01T17:00:00.001Z');
    const cases = [
        {
            expectedClass: 'authority-unavailable',
            expectedReason: 'legacy-seeding-gate-mismatch',
            expectedQueries: 3,
            invoke: (store) => store.admitUnboundLegacyIssuance({
            loginLookupKeyId: 'synthetic-lookup-key',
            loginLookupToken: Buffer.alloc(32, 16),
            issuedAt,
            expiresAt,
        }), control: {} },
        {
            expectedClass: 'authority-conflict',
            expectedReason: 'legacy-issuance-disabled',
            expectedQueries: 4,
            invoke: (store) => store.bindLegacy({
            legacyCompatibilityId: '00000000-0000-4000-8000-000000000351',
            subjectId: subjectRow().subjectId,
            ...subjectCredentialExpectation(),
            verifierKeyId: 'synthetic-legacy-key',
            verifier: Buffer.alloc(32, 17),
            issuedAt,
            expiresAt,
        }), control: {
            legacyCompatibilityEnforcementEnabled: 1,
            legacyLedgerSeedingEnabled: 1,
        } },
    ];

    for (const entry of cases) {
        const fake = createFakeSql([
            result([{ serverTime: SERVER_TIME }]),
            result([controlRow({
                hardSunsetAt,
                ...entry.control,
            })]),
        ]);
        const store = createAzureSqlSessionStore({
            sql: fake.sql,
            connectionString: 'synthetic-only',
        });
        await assert.rejects(
            entry.invoke(store),
            (error) => error.errorClass === entry.expectedClass
                && error.reason === entry.expectedReason,
        );
        assert.equal(fake.state.rollbacks, 1);
        assert.equal(fake.state.queries.length, entry.expectedQueries);
    }
});

test('future-issued exact-sunset legacy metadata is rejected before final aging', async () => {
    const serverTime = new Date('2042-06-01T11:59:59.999Z');
    const issuedAt = new Date('2042-06-01T12:00:00.000Z');
    const fake = createFakeSql([
        result([{ serverTime }]),
        result([controlRow()]),
        result([]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });
    await assert.rejects(
        store.admitUnboundLegacyIssuance({
            loginLookupKeyId: 'synthetic-lookup-key',
            loginLookupToken: Buffer.alloc(32, 18),
            issuedAt,
            expiresAt: new Date(issuedAt.getTime() + 4 * 60 * 60 * 1000),
        }),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'legacy-issuance-time-invalid',
    );
});

test('unbound issuance admission closes on rollout races and irreversible subject adoption', async () => {
    const issuedAt = new Date('2042-06-01T11:00:00.000Z');
    const expiresAt = new Date('2042-06-01T15:00:00.000Z');
    const input = {
        loginLookupKeyId: 'synthetic-lookup-key',
        loginLookupToken: Buffer.alloc(32, 12),
        issuedAt,
        expiresAt,
    };
    const transitioned = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({
            legacyCompatibilityEnforcementEnabled: 1,
            legacyLedgerSeedingEnabled: 1,
        })]),
    ]);
    const transitionedStore = createAzureSqlSessionStore({
        sql: transitioned.sql,
        connectionString: 'synthetic-only',
    });
    await assert.rejects(
        transitionedStore.admitUnboundLegacyIssuance(input),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'legacy-seeding-gate-mismatch',
    );
    assert.equal(transitioned.state.rollbacks, 1);
    assert.equal(transitioned.state.queries.length, 3);

    const adopted = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
        result([subjectRow({ legacyAuthorityDisabledAt: SERVER_TIME })]),
    ]);
    const adoptedStore = createAzureSqlSessionStore({
        sql: adopted.sql,
        connectionString: 'synthetic-only',
    });
    await assert.rejects(
        adoptedStore.admitUnboundLegacyIssuance(input),
        (error) => error.errorClass === 'authority-conflict'
            && error.reason === 'target-authority-established',
    );
    assert.equal(adopted.state.rollbacks, 1);
});

test('memory-store ordering never admits unbound issuance after seeding begins', async () => {
    const backing = createTestSessionAuthorityBacking();
    const storeA = createTestSessionAuthorityStore({ testOnly: true, backing, now: () => SERVER_TIME });
    const storeB = createTestSessionAuthorityStore({ testOnly: true, backing, now: () => SERVER_TIME });
    const transition = storeA.transitionControl({
        expectedVersion: 1,
        changes: {
            legacyLedgerSeedingEnabled: true,
            legacyLedgerSeedingStartedAt: SERVER_TIME,
            seedingStartedAt: SERVER_TIME,
        },
    });
    const admission = storeB.admitUnboundLegacyIssuance({
        loginLookupKeyId: 'synthetic-lookup-key',
        loginLookupToken: Buffer.alloc(32, 13),
        issuedAt: new Date('2042-06-01T11:00:00.000Z'),
        expiresAt: new Date('2042-06-01T15:00:00.000Z'),
    });

    const [transitionResult, admissionResult] = await Promise.allSettled([transition, admission]);
    assert.equal(transitionResult.status, 'fulfilled');
    assert.equal(admissionResult.status, 'rejected');
    assert.equal(admissionResult.reason.errorClass, 'authority-unavailable');
    assert.equal(admissionResult.reason.reason, 'legacy-seeding-gate-mismatch');
});

test('legacy acceptance requires four-hour final aging and stop issuance cannot miss sunset', async () => {
    const immediate = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
    ]);
    const immediateStore = createAzureSqlSessionStore({
        sql: immediate.sql,
        connectionString: 'synthetic-only',
    });
    await assert.rejects(
        immediateStore.transitionControl({
            expectedVersion: 3,
            changes: {
                legacyAcceptanceDisabledAt: SERVER_TIME,
                legacyAcceptanceEnabled: false,
                legacyIssuanceEnabled: false,
                legacyStopIssuanceAt: SERVER_TIME,
            },
        }),
        (error) => error.errorClass === 'forbidden-authority'
            && error.reason === 'legacy-final-aging-incomplete',
    );
    assert.equal(immediate.state.rollbacks, 1);

    const lateStop = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({
            legacyCompatibilityEnforcementEnabled: 1,
            legacyCompatibilityEnforcedAt: new Date('2042-05-25T08:00:00.000Z'),
            legacyLedgerSeedingEnabled: 1,
            legacyLedgerSeedingStartedAt: new Date('2042-05-25T04:00:00.000Z'),
            seedingStartedAt: new Date('2042-05-25T04:00:00.000Z'),
            seedingQualifiedAt: new Date('2042-05-25T08:00:00.000Z'),
            dualStackStartedAt: new Date('2042-05-25T15:00:00.000Z'),
            hardSunsetAt: new Date('2042-06-01T15:00:00.000Z'),
        })]),
    ]);
    const lateStopStore = createAzureSqlSessionStore({
        sql: lateStop.sql,
        connectionString: 'synthetic-only',
    });
    await assert.rejects(
        lateStopStore.transitionControl({
            expectedVersion: 3,
            changes: { legacyIssuanceEnabled: false, legacyStopIssuanceAt: SERVER_TIME },
        }),
        (error) => error.errorClass === 'forbidden-authority'
            && error.reason === 'legacy-stop-too-late',
    );
    assert.equal(lateStop.state.rollbacks, 1);
});

test('legacy binding cannot issue at or after the hard sunset even if the flag was missed', async () => {
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({
            hardSunsetAt: SERVER_TIME,
            legacyLedgerSeedingEnabled: 1,
            legacyIssuanceEnabled: 1,
        })]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    await assert.rejects(
        store.bindLegacy({
            legacyCompatibilityId: '00000000-0000-4000-8000-000000000300',
            subjectId: '00000000-0000-4000-8000-000000000042',
            ...subjectCredentialExpectation(),
            verifierKeyId: 'synthetic-legacy-key',
            verifier: Buffer.alloc(32, 6),
            issuedAt: new Date('2042-06-01T12:00:00.000Z'),
            expiresAt: new Date('2042-06-01T16:00:00.000Z'),
        }),
        (error) => error.errorClass === 'authority-conflict'
            && error.reason === 'legacy-issuance-disabled',
    );
    assert.equal(fake.state.rollbacks, 1);
    assert.equal(fake.state.queries.length, 4);
});

test('legacy authorization returns an explicit unbound result before enforcement', async () => {
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({ legacyLedgerSeedingEnabled: 1 })]),
        result([]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    const output = await store.authorizeLegacy({
        verifierKeyId: 'synthetic-legacy-key',
        verifier: Buffer.alloc(32, 31),
        issuedAt: new Date('2042-06-01T11:00:00.000Z'),
        expiresAt: new Date('2042-06-01T15:00:00.000Z'),
    });

    assert.deepEqual(Object.keys(output).sort(), ['control', 'serverTime', 'unbound']);
    assert.equal(output.unbound, true);
    assert.equal(output.control.legacyCompatibilityEnforcementEnabled, false);
    assert.deepEqual(output.serverTime, SERVER_TIME);
    assert.equal(fake.state.commits, 1);
});

test('legacy authorization requires a binding after enforcement and central controls dominate', async () => {
    const enforced = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({ legacyCompatibilityEnforcementEnabled: 1 })]),
        result([]),
    ]);
    const enforcedStore = createAzureSqlSessionStore({
        sql: enforced.sql,
        connectionString: 'synthetic-only',
    });
    await assert.rejects(
        enforcedStore.authorizeLegacy({
            verifierKeyId: 'synthetic-legacy-key',
            verifier: Buffer.alloc(32, 32),
            issuedAt: new Date('2042-06-01T11:00:00.000Z'),
            expiresAt: new Date('2042-06-01T15:00:00.000Z'),
        }),
        (error) => error.errorClass === 'invalid-authority'
            && error.reason === 'legacy-binding-missing',
    );

    const controls = [
        {
            expectedClass: 'authority-unavailable',
            expectedReason: 'authority-incident',
            row: { incidentState: 'suspended' },
        },
        {
            expectedClass: 'invalid-authority',
            expectedReason: 'legacy-acceptance-disabled',
            row: {
                legacyIssuanceEnabled: 0,
                legacyStopIssuanceAt: new Date(SERVER_TIME.getTime() - 4 * 60 * 60 * 1000),
                legacyAcceptanceEnabled: 0,
                legacyAcceptanceDisabledAt: SERVER_TIME,
            },
        },
        {
            expectedClass: 'invalid-authority',
            expectedReason: 'legacy-acceptance-disabled',
            row: { hardSunsetAt: SERVER_TIME },
        },
    ];
    for (const entry of controls) {
        const fake = createFakeSql([
            result([{ serverTime: SERVER_TIME }]),
            result([controlRow(entry.row)]),
        ]);
        const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });
        await assert.rejects(
            store.authorizeLegacy({
                verifierKeyId: 'synthetic-legacy-key',
                verifier: Buffer.alloc(32, 33),
                issuedAt: new Date('2042-06-01T11:00:00.000Z'),
                expiresAt: new Date('2042-06-01T15:00:00.000Z'),
            }),
            (error) => error.errorClass === entry.expectedClass
                && error.reason === entry.expectedReason,
        );
        assert.equal(fake.state.queries.length, 3);
    }
});

test('driver failures roll back and become privacy-safe unavailable errors', async () => {
    const driverMessage = 'credential=should-never-escape; verifier=private';
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
        new Error(driverMessage),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    await assert.rejects(
        store.transitionControl({ expectedVersion: 3, changes: { targetRoutesEnabled: true } }),
        (error) => {
            assert.equal(error.errorClass, 'authority-unavailable');
            assert.equal(error.reason, 'session-store-unavailable');
            assert.doesNotMatch(error.message, /credential|verifier|private/);
            assert.equal(error.cause, undefined);
            return true;
        },
    );
    assert.equal(fake.state.rollbacks, 1);
});

test('commit failures are reported as unknown transaction outcomes without rollback', async () => {
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
        result([], [1]),
        result([controlRow({ controlVersion: 4 })]),
    ], { commitError: new Error('ambiguous transport after commit') });
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    await assert.rejects(
        store.transitionControl({ expectedVersion: 3, changes: { targetRoutesEnabled: true } }),
        (error) => {
            assert.equal(error.errorClass, 'authority-unavailable');
            assert.equal(error.reason, 'transaction-outcome-unknown');
            assert.doesNotMatch(error.message, /transport|commit/);
            return true;
        },
    );
    assert.equal(fake.state.commitAttempts, 1);
    assert.equal(fake.state.rollbacks, 0);
});

test('rollback failures are conservatively reported as unknown outcomes', async () => {
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
        new Error('synthetic query failure'),
    ], { rollbackError: new Error('synthetic rollback failure') });
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    await assert.rejects(
        store.transitionControl({ expectedVersion: 3, changes: { targetRoutesEnabled: true } }),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'transaction-outcome-unknown',
    );
    assert.equal(fake.state.rollbackAttempts, 1);
});

test('eligibility refresh uses transaction time and atomically revokes when entitlement elapsed', async () => {
    const subjectId = '00000000-0000-4000-8000-000000000042';
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
        result([subjectRow({ subjectId })]),
        result([], [1]),
        result([], [2]),
        result([], [1]),
        result([subjectRow({
            subjectId,
            eligibilityState: 'ineligible',
            entitlementExpiresAt: new Date('2042-06-01T11:59:59.000Z'),
            eligibilityObservedAt: SERVER_TIME,
            eligibilityRevalidateAt: SERVER_TIME,
            sessionEpoch: 8,
        })]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    const output = await store.updateEligibility({
        subjectId,
        observationStartedAt: SERVER_TIME,
        expectedCredentialVersion: 2,
        expectedCredentialFingerprintKeyId: 'synthetic-fingerprint-key',
        expectedCredentialFingerprint: Buffer.alloc(32, 2),
        expectedControlVersion: 3,
        rowHint: 4,
        eligibilityState: 'eligible',
        entitlementExpiresAt: new Date('2042-06-01T11:59:59.000Z'),
        eligibilityObservedAt: new Date('2042-06-01T11:55:00.000Z'),
        eligibilityRevalidateAt: new Date('2042-06-01T12:00:00.000Z'),
    });

    assert.equal(output.eligible, false);
    assert.equal(output.subject.sessionEpoch, 8);
    assert.equal(fake.state.commits, 1);
    assert.equal(fake.state.rollbacks, 0);
    const subjectMutation = fake.state.queries.find(({ statement }) => statement.includes('update-eligibility:expired'));
    const sessionMutation = fake.state.queries.find(({ statement }) => statement.includes('revoke-subject-sessions'));
    assert.deepEqual(subjectMutation.parameters.serverTime, SERVER_TIME);
    assert.match(subjectMutation.statement, /subject_session_epoch = subject_session_epoch \+ 1/);
    assert.match(sessionMutation.statement, /phase IN \('credential-verified', 'registration-pending', 'face-pending', 'authenticated'\)/);
    assert.doesNotMatch(sessionMutation.statement, /epoch_snapshot/);
});

test('subject observation mutations use credential and observation compare-and-replace guards', async () => {
    const common = {
        subjectId: subjectRow().subjectId,
        observationStartedAt: new Date('2042-06-01T11:59:00.000Z'),
        expectedCredentialVersion: 2,
        expectedCredentialFingerprintKeyId: 'synthetic-fingerprint-key',
        expectedCredentialFingerprint: Buffer.alloc(32, 2),
        expectedControlVersion: 3,
        eligibilityRevalidateAt: new Date('2042-06-01T12:04:00.000Z'),
    };
    const cases = [
        {
            marker: 'update-eligibility */',
            invoke: (store) => store.updateEligibility({
                ...common,
                rowHint: 7,
                entitlementExpiresAt: new Date('2042-06-02T03:00:00.000Z'),
            }),
        },
        {
            marker: 'revoke-for-ineligibility:subject',
            invoke: (store) => store.revokeForIneligibility({
                ...common,
                eligibilityState: 'ineligible',
                entitlementExpiresAt: SERVER_TIME,
                reason: 'account-inactive',
            }),
        },
        {
            marker: 'revoke-for-credential-change:subject',
            invoke: (store) => store.revokeForCredentialChange({
                ...common,
                credentialFingerprintKeyId: 'synthetic-fingerprint-key-v2',
                credentialFingerprint: Buffer.alloc(32, 9),
            }),
        },
    ];

    for (const entry of cases) {
        const fake = createFakeSql([
            result([{ serverTime: SERVER_TIME }]),
            result([controlRow()]),
            result([], [0]),
        ]);
        const store = createAzureSqlSessionStore({
            sql: fake.sql,
            connectionString: 'synthetic-only',
        });
        await assert.rejects(
            entry.invoke(store),
            (error) => error.errorClass === 'authority-conflict'
                && error.reason === 'subject-observation-compare-and-replace',
        );
        assert.equal(fake.state.commits, 0);
        assert.equal(fake.state.rollbacks, 1);
        const mutation = fake.state.queries.find(({ statement }) => statement.includes(entry.marker));
        assert.ok(mutation);
        assert.match(mutation.statement, /eligibility_observed_at <= @observationStartedAt/);
        assert.match(mutation.statement, /credential_version = @expectedCredentialVersion/);
        assert.match(
            mutation.statement,
            /credential_fingerprint_key_id = @expectedCredentialFingerprintKeyId/,
        );
        assert.match(mutation.statement, /credential_fingerprint = @expectedCredentialFingerprint/);
        assert.deepEqual(mutation.parameters.observationStartedAt, common.observationStartedAt);
        assert.equal(mutation.parameters.expectedCredentialVersion, 2);
        assert.deepEqual(
            mutation.parameters.expectedCredentialFingerprint,
            common.expectedCredentialFingerprint,
        );
    }
});

test('memory store rejects stale observations and permits same-flow credential CAS continuation', async () => {
    const backing = createTestSessionAuthorityBacking();
    const store = createTestSessionAuthorityStore({
        testOnly: true,
        backing,
        now: () => SERVER_TIME,
    });
    const subjectId = '00000000-0000-4000-8000-000000000055';
    const originalFingerprint = Buffer.alloc(32, 3);
    const replacementFingerprint = Buffer.alloc(32, 4);
    const initialObservation = new Date('2042-06-01T11:00:00.000Z');
    const observationStartedAt = new Date('2042-06-01T11:30:00.000Z');
    await store.createOrLoadSubject({
        subjectId,
        loginLookupKeyId: 'synthetic-lookup-key',
        loginLookupToken: Buffer.alloc(32, 5),
        encryptedAccountMapping: Buffer.from('synthetic-ciphertext'),
        accountMappingKeyId: 'synthetic-mapping-key',
        rowHint: 4,
        credentialFingerprintKeyId: 'synthetic-fingerprint-key',
        credentialFingerprint: originalFingerprint,
        eligibilityState: 'eligible',
        entitlementExpiresAt: new Date('2042-06-02T03:00:00.000Z'),
        eligibilityObservedAt: initialObservation,
        eligibilityRevalidateAt: new Date('2042-06-01T11:05:00.000Z'),
        expectedControlVersion: 1,
    });

    const originalExpectation = {
        subjectId,
        expectedCredentialVersion: 1,
        expectedCredentialFingerprintKeyId: 'synthetic-fingerprint-key',
        expectedCredentialFingerprint: originalFingerprint,
        expectedControlVersion: 1,
    };
    await assert.rejects(
        store.updateEligibility({
            ...originalExpectation,
            observationStartedAt: new Date('2042-06-01T10:59:59.999Z'),
            rowHint: 8,
            entitlementExpiresAt: new Date('2042-06-02T03:00:00.000Z'),
            eligibilityRevalidateAt: new Date('2042-06-01T11:04:59.999Z'),
        }),
        (error) => error.errorClass === 'authority-conflict'
            && error.reason === 'subject-observation-compare-and-replace',
    );
    await assert.rejects(
        store.revokeForIneligibility({
            ...originalExpectation,
            observationStartedAt,
            expectedCredentialFingerprint: Buffer.alloc(32, 99),
            eligibilityState: 'ineligible',
            entitlementExpiresAt: SERVER_TIME,
            reason: 'account-inactive',
        }),
        (error) => error.errorClass === 'authority-conflict'
            && error.reason === 'subject-observation-compare-and-replace',
    );

    const changed = await store.revokeForCredentialChange({
        ...originalExpectation,
        observationStartedAt,
        credentialFingerprintKeyId: 'synthetic-fingerprint-key-v2',
        credentialFingerprint: replacementFingerprint,
    });
    assert.equal(changed.subject.credentialVersion, 2);
    assert.deepEqual(changed.subject.eligibilityObservedAt, observationStartedAt);

    const refreshed = await store.updateEligibility({
        subjectId,
        observationStartedAt,
        expectedCredentialVersion: 2,
        expectedCredentialFingerprintKeyId: 'synthetic-fingerprint-key-v2',
        expectedCredentialFingerprint: replacementFingerprint,
        expectedControlVersion: 1,
        rowHint: 9,
        entitlementExpiresAt: new Date('2042-06-02T03:00:00.000Z'),
        eligibilityRevalidateAt: new Date('2042-06-01T12:04:00.000Z'),
    });
    assert.equal(refreshed.eligible, true);
    assert.deepEqual(refreshed.subject.eligibilityObservedAt, SERVER_TIME);

    await assert.rejects(
        store.issueSession({
            sessionId: '00000000-0000-4000-8000-000000000156',
            subjectId,
            expectedCredentialVersion: 1,
            expectedCredentialFingerprintKeyId: 'synthetic-fingerprint-key',
            expectedCredentialFingerprint: originalFingerprint,
            verifierKeyId: 'synthetic-target-key',
            verifier: Buffer.alloc(32, 6),
            phase: 'authenticated',
            faceRequired: false,
            registrationRequired: false,
        }),
        (error) => error.errorClass === 'authority-conflict'
            && error.reason === 'subject-credential-compare-and-replace',
    );
    assert.equal(backing.state.sessions.size, 0);
    assert.equal(backing.state.subjects.get(subjectId).legacyAuthorityDisabledAt, null);
});

test('internal subject revocation increments the epoch and revokes every active phase atomically', async () => {
    const subjectId = '00000000-0000-4000-8000-000000000043';
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({ legacyCompatibilityEnforcementEnabled: 1 })]),
        result([subjectRow({ subjectId, sessionEpoch: 8 })]),
        result([], [1]),
        result([], [3]),
        result([], [1]),
        result([subjectRow({ subjectId, sessionEpoch: 9 })]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    const output = await store.revokeSubject({ subjectId, reason: 'administrator-revocation' });

    assert.equal(output.subject.sessionEpoch, 9);
    assert.equal(fake.state.commits, 1);
    const sessionMutation = fake.state.queries.find(({ statement }) => statement.includes('revoke-subject-sessions'));
    assert.equal(sessionMutation.parameters.reason, 'administrator-revocation');
    assert.doesNotMatch(sessionMutation.statement, /administrator-revocation/);
    assert.doesNotMatch(sessionMutation.statement, /epoch_snapshot/);

    const invalidFake = createFakeSql();
    const invalidStore = createAzureSqlSessionStore({
        sql: invalidFake.sql,
        connectionString: 'synthetic-only',
    });
    await assert.rejects(
        invalidStore.revokeSubject({ subjectId, reason: 'private value' }),
        /privacy-safe machine value/,
    );
    assert.equal(invalidFake.state.connects, 0);
});

test('authority incident makes target status and logout unavailable', async () => {
    for (const operation of ['readSession', 'logout']) {
        const fake = createFakeSql([
            result([{ serverTime: SERVER_TIME }]),
            result([controlRow({ incidentState: 'suspended' })]),
        ]);
        const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });
        await assert.rejects(
            store[operation]({ verifierKeyId: 'synthetic-target-key', verifier: Buffer.alloc(32, 4) }),
            (error) => error.errorClass === 'authority-unavailable'
                && error.reason === 'authority-incident',
        );
        assert.equal(fake.state.rollbacks, 1);
        assert.equal(fake.state.commits, 0);
    }
});

test('logout atomically rejects a disabled target-route gate without mutating a session', async () => {
    const fake = createFakeSql([
        result([controlRow({ targetRoutesEnabled: 0 })]),
        result([sessionRow()]),
        result([subjectRow()]),
        result([{ serverTime: SERVER_TIME }]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    await assert.rejects(
        store.logout({
            verifierKeyId: 'synthetic-target-key',
            verifier: Buffer.alloc(32, 4),
        }),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'target-routes-disabled',
    );

    assert.equal(fake.state.commits, 0);
    assert.equal(fake.state.rollbacks, 1);
    assert.equal(fake.state.queries.some(({ statement }) => (
        statement.includes('revoke-session')
        || statement.includes('mark-expired')
        || statement.includes('enforce-stored-ineligibility')
    )), false);
});

test('verifier-key incident evidence cannot advance without suspension metadata', async () => {
    for (const field of ['targetVerifierKeyIncidentAt', 'legacyVerifierKeyIncidentAt']) {
        const fake = createFakeSql([
            result([{ serverTime: SERVER_TIME }]),
            result([controlRow()]),
        ]);
        const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });
        await assert.rejects(
            store.transitionControl({
                expectedVersion: 3,
                changes: { [field]: SERVER_TIME },
            }),
            (error) => error.errorClass === 'forbidden-authority'
                && error.reason === 'verifier-key-incident-requires-suspension',
        );
        assert.equal(fake.state.rollbacks, 1);
    }
});

test('target verifier incidents suspend authority without prematurely advancing configured generation', async () => {
    const changes = {
        incidentCode: 'target-verifier-compromise',
        incidentRecordedAt: SERVER_TIME,
        incidentState: 'suspended',
        targetVerifierKeyIncidentAt: SERVER_TIME,
    };
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
        result([], [1]),
        result([controlRow({ controlVersion: 4, ...changes })]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    const output = await store.transitionControl({ expectedVersion: 3, changes });

    assert.equal(output.control.incidentState, 'suspended');
    assert.equal(output.control.authorityGeneration, 3);
    assert.equal(output.control.globalSessionEpoch, 2);
    assert.deepEqual(output.control.targetVerifierKeyIncidentAt, SERVER_TIME);
    assert.equal(fake.state.commits, 1);
});

test('legacy verifier incidents atomically suspend and retire issuance and acceptance', async () => {
    const incomplete = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
    ]);
    const incompleteStore = createAzureSqlSessionStore({
        sql: incomplete.sql,
        connectionString: 'synthetic-only',
    });
    await assert.rejects(
        incompleteStore.transitionControl({
            expectedVersion: 3,
            changes: {
                incidentCode: 'legacy-verifier-compromise',
                incidentRecordedAt: SERVER_TIME,
                incidentState: 'suspended',
                legacyVerifierKeyIncidentAt: SERVER_TIME,
            },
        }),
        (error) => error.errorClass === 'forbidden-authority'
            && error.reason === 'legacy-key-incident-requires-retirement',
    );

    const changes = {
        incidentCode: 'legacy-verifier-compromise',
        incidentRecordedAt: SERVER_TIME,
        incidentState: 'suspended',
        legacyAcceptanceDisabledAt: SERVER_TIME,
        legacyAcceptanceEnabled: false,
        legacyIssuanceEnabled: false,
        legacyStopIssuanceAt: SERVER_TIME,
        legacyVerifierKeyIncidentAt: SERVER_TIME,
    };
    const qualified = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
        result([], [1]),
        result([controlRow({ controlVersion: 4, ...changes })]),
    ]);
    const qualifiedStore = createAzureSqlSessionStore({
        sql: qualified.sql,
        connectionString: 'synthetic-only',
    });

    const output = await qualifiedStore.transitionControl({ expectedVersion: 3, changes });

    assert.equal(output.control.legacyIssuanceEnabled, false);
    assert.equal(output.control.legacyAcceptanceEnabled, false);
    assert.deepEqual(output.control.legacyStopIssuanceAt, SERVER_TIME);
    assert.deepEqual(output.control.legacyAcceptanceDisabledAt, SERVER_TIME);
    assert.equal(qualified.state.commits, 1);
});

test('incident resume requires a separately fenced recovering generation', async () => {
    const resumedControl = controlRow({
        authorityGeneration: 4,
        controlVersion: 4,
        globalSessionEpoch: 3,
        incidentState: 'normal',
        legacyAcceptanceDisabledAt: SERVER_TIME,
        legacyAcceptanceEnabled: 0,
        legacyIssuanceEnabled: 0,
        legacyLedgerSeedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
        legacyStopIssuanceAt: new Date('2042-06-01T08:00:00.000Z'),
        seedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
    });
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({ incidentState: 'suspended' })]),
    ]);
    const rejectedStore = createAzureSqlSessionStore({
        sql: fake.sql,
        connectionString: 'synthetic-only',
    });
    await assert.rejects(
        rejectedStore.transitionControl({ expectedVersion: 3, changes: { incidentState: 'normal' } }),
        (error) => error.errorClass === 'forbidden-authority'
            && error.reason === 'incident-resume-requires-fenced-recovery',
    );

    const qualified = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({
            authorityGeneration: 4,
            globalSessionEpoch: 3,
            incidentState: 'recovering',
            legacyAcceptanceDisabledAt: SERVER_TIME,
            legacyAcceptanceEnabled: 0,
            legacyIssuanceEnabled: 0,
            legacyLedgerSeedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
            legacyStopIssuanceAt: new Date('2042-06-01T08:00:00.000Z'),
            seedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
        })]),
        result([{ liveFaceAuthorityExists: 0 }]),
        result([], [1]),
        result([resumedControl]),
        result([{ serverTime: SERVER_TIME }]),
        result([resumedControl]),
        result([sessionRow()]),
        result([subjectRow()]),
        result([{ serverTime: SERVER_TIME }]),
        result([resumedControl]),
    ]);
    const qualifiedStore = createAzureSqlSessionStore({
        sql: qualified.sql,
        connectionString: 'synthetic-only',
        expectedAuthorityGeneration: 4,
    });
    await qualifiedStore.transitionControl({
        expectedVersion: 3,
        changes: {
            incidentState: 'normal',
        },
    });
    await assert.rejects(
        qualifiedStore.readSession({
            verifierKeyId: 'synthetic-target-key',
            verifier: Buffer.alloc(32, 4),
        }),
        (error) => error.errorClass === 'invalid-authority' && error.reason === 'epoch-mismatch',
    );
    await assert.rejects(
        qualifiedStore.authorizeLegacy({
            verifierKeyId: 'synthetic-legacy-key',
            verifier: Buffer.alloc(32, 5),
            issuedAt: new Date('2042-06-01T10:00:00.000Z'),
            expiresAt: new Date('2042-06-01T14:00:00.000Z'),
        }),
        (error) => error.errorClass === 'invalid-authority'
            && error.reason === 'legacy-acceptance-disabled',
    );
    assert.equal(qualified.state.commits, 1);
    assert.equal(qualified.state.rollbacks, 2);
});

test('test-only store preserves incident retirement across restored target and legacy records', async () => {
    const backing = createTestSessionAuthorityBacking();
    const subjectId = '00000000-0000-4000-8000-000000000042';
    const sessionId = '00000000-0000-4000-8000-000000000142';
    const targetVerifier = Buffer.alloc(32, 4);
    const legacyVerifier = Buffer.alloc(32, 5);
    const subject = subjectRow({ subjectId });
    const session = sessionRow({ subjectId, sessionId });
    const binding = legacyRow({ subjectId });
    backing.state.subjects.set(subjectId, subject);
    backing.state.sessions.set(sessionId, session);
    backing.state.sessionsByVerifier.set(`synthetic-target-key:${targetVerifier.toString('hex')}`, session);
    backing.state.legacyBindings.set(`synthetic-legacy-key:${legacyVerifier.toString('hex')}`, binding);
    Object.assign(backing.state.control, {
        authorityGeneration: 4,
        globalSessionEpoch: 3,
        loginLookupKeyInitialized: 1,
        loginLookupKeyMatches: 1,
        accountMappingKeyMatches: 1,
        authorityKeysetInitialized: 1,
        authorityKeysetAggregateMatches: 1,
        targetVerifierKeyMatches: 1,
        legacyCompatibilityKeyMatches: 1,
        credentialFingerprintKeyMatches: 1,
        faceChallengeKeyMatches: 1,
        legacySigningKeyInitialized: 1,
        legacySigningKeyMatches: 1,
        incidentState: 'recovering',
        legacyAcceptanceDisabledAt: SERVER_TIME,
        legacyAcceptanceEnabled: false,
        legacyIssuanceEnabled: false,
        legacyLedgerSeedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
        legacyStopIssuanceAt: new Date('2042-06-01T08:00:00.000Z'),
        seedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
    });
    const store = createTestSessionAuthorityStore({
        testOnly: true,
        backing,
        expectedAuthorityGeneration: 4,
        now: () => SERVER_TIME,
    });

    await store.transitionControl({
        expectedVersion: 1,
        changes: {
            incidentState: 'normal',
        },
    });
    await assert.rejects(
        store.readSession({ verifierKeyId: 'synthetic-target-key', verifier: targetVerifier }),
        (error) => error.errorClass === 'invalid-authority' && error.reason === 'epoch-mismatch',
    );
    await assert.rejects(
        store.authorizeLegacy({
            verifierKeyId: 'synthetic-legacy-key',
            verifier: legacyVerifier,
            issuedAt: binding.issuedAt,
            expiresAt: binding.expiresAt,
        }),
        (error) => error.errorClass === 'invalid-authority'
            && error.reason === 'legacy-acceptance-disabled',
    );
});

test('Face reconciliation changes registration evidence only for registration ambiguity', async () => {
    for (const registrationReconciliationRequired of [true, false]) {
        const flowId = '00000000-0000-4000-8000-000000000151';
        const fake = createFakeSql([
            result([{ serverTime: SERVER_TIME }]),
            result([controlRow()]),
            result([flowRow({ flowId })]),
            result([], [1]),
        ]);
        const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });
        await store.markFaceFlowReconciliation({
            flowId,
            registrationReconciliationRequired,
        });
        const mutation = fake.state.queries.find(({ statement }) => (
            statement.includes('mark-face-flow-reconciliation')
        ));
        assert.match(mutation.statement, /WHEN @registrationReconciliationRequired = 1/);
        assert.match(mutation.statement, /ELSE registration_state/);
        assert.match(mutation.statement, /WHERE flow_id = @flowId/);
        assert.equal(mutation.parameters.flowId, flowId);
        assert.equal(
            mutation.parameters.registrationReconciliationRequired,
            registrationReconciliationRequired,
        );
    }
});

test('Face reconciliation requires one immutable flow match', async () => {
    const flowId = '00000000-0000-4000-8000-000000000152';
    const fake = createFakeSql([
        result([controlRow()]),
        result([flowRow({ flowId })]),
        result([{ serverTime: SERVER_TIME }]),
        result([], [0]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    await assert.rejects(
        store.markFaceFlowReconciliation({ flowId }),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'face-flow-reconciliation-unavailable',
    );
    assert.equal(fake.state.commits, 0);
    assert.equal(fake.state.rollbacks, 1);
    const mutation = fake.state.queries.find(({ statement }) => (
        statement.includes('mark-face-flow-reconciliation')
    ));
    assert.match(mutation.statement, /WHERE flow_id = @flowId/);
});

test('active session phase and lifetime mismatches fail closed before expiry mutation', async () => {
    const cases = [
        sessionRow({
            phase: 'authenticated',
            originalIssuedAt: new Date('2042-06-01T11:00:00.000Z'),
            phaseStartedAt: new Date('2042-06-01T11:00:00.000Z'),
            expiresAt: new Date('2042-06-01T11:20:00.000Z'),
        }),
        sessionRow({
            phase: 'credential-verified',
            originalIssuedAt: new Date('2042-06-01T11:00:00.000Z'),
            phaseStartedAt: new Date('2042-06-01T11:00:00.000Z'),
            expiresAt: new Date('2042-06-01T15:00:00.000Z'),
        }),
    ];

    for (const session of cases) {
        const fake = createFakeSql([
            result([controlRow()]),
            result([session]),
            result([subjectRow({ subjectId: session.subjectId })]),
            result([{ serverTime: SERVER_TIME }]),
        ]);
        const store = createAzureSqlSessionStore({
            sql: fake.sql,
            connectionString: 'synthetic-only',
        });
        await assert.rejects(
            store.readSession({
                verifierKeyId: session.verifierKeyId,
                verifier: session.verifier,
            }),
            (error) => error.errorClass === 'authority-unavailable'
                && error.reason === 'session-store-integrity',
        );
        assert.equal(fake.state.commits, 0);
        assert.equal(fake.state.queries.some(({ statement }) => (
            statement.includes('mark-expired')
        )), false);
    }
});

test('stored entitlement expiry revokes the subject and returns 403 after commit', async () => {
    const subjectId = '00000000-0000-4000-8000-000000000044';
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
        result([sessionRow({ subjectId })]),
        result([subjectRow({
            subjectId,
            entitlementExpiresAt: new Date('2042-06-01T11:59:59.000Z'),
        })]),
        result([], [1]),
        result([], [2]),
        result([], [1]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    await assert.rejects(
        store.readSession({ verifierKeyId: 'synthetic-target-key', verifier: Buffer.alloc(32, 4) }),
        (error) => error.errorClass === 'forbidden-authority' && error.reason === 'ineligible',
    );
    assert.equal(fake.state.commits, 1);
    assert.equal(fake.state.rollbacks, 0);
    assert.ok(fake.state.queries.some(({ statement }) => statement.includes('enforce-stored-ineligibility')));
    const revoke = fake.state.queries.find(({ statement }) => statement.includes('revoke-subject-sessions'));
    assert.doesNotMatch(revoke.statement, /epoch_snapshot/);
});

test('memory authority gives an exact entitlement/session expiry tie 403 precedence', async () => {
    const backing = createTestSessionAuthorityBacking();
    const subjectId = '00000000-0000-4000-8000-000000000054';
    const sessionId = '00000000-0000-4000-8000-000000000154';
    const verifier = Buffer.alloc(32, 54);
    const subject = subjectRow({
        subjectId,
        entitlementExpiresAt: SERVER_TIME,
        eligibilityObservedAt: new Date('2042-06-01T11:55:00.000Z'),
        eligibilityRevalidateAt: SERVER_TIME,
    });
    const session = sessionRow({
        subjectId,
        sessionId,
        verifier,
        originalIssuedAt: new Date('2042-06-01T08:00:00.000Z'),
        phaseStartedAt: new Date('2042-06-01T08:00:00.000Z'),
        createdAt: new Date('2042-06-01T08:00:00.000Z'),
        expiresAt: SERVER_TIME,
    });
    backing.state.subjects.set(subjectId, subject);
    backing.state.sessions.set(sessionId, session);
    backing.state.sessionsByVerifier.set(`synthetic-target-key:${verifier.toString('hex')}`, session);
    const store = createTestSessionAuthorityStore({
        testOnly: true,
        backing,
        now: () => SERVER_TIME,
    });

    await assert.rejects(
        store.readSession({ verifierKeyId: 'synthetic-target-key', verifier }),
        (error) => error.errorClass === 'forbidden-authority' && error.reason === 'ineligible',
    );
    assert.equal(subject.eligibilityState, 'ineligible');
    assert.equal(subject.sessionEpoch, 8);
    assert.equal(session.phase, 'revoked');
    assert.equal(session.revocationReason, 'entitlement-expired');
});

test('logout commits stored-entitlement revocation but remains effect-idempotent', async () => {
    const subjectId = '00000000-0000-4000-8000-000000000045';
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({ targetRoutesEnabled: 1 })]),
        result([sessionRow({ subjectId })]),
        result([subjectRow({
            subjectId,
            entitlementExpiresAt: new Date('2042-06-01T11:59:59.000Z'),
        })]),
        result([], [1]),
        result([], [2]),
        result([], [1]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    const output = await store.logout({
        verifierKeyId: 'synthetic-target-key',
        verifier: Buffer.alloc(32, 4),
    });

    assert.deepEqual(output, { revoked: false, serverTime: SERVER_TIME });
    assert.equal(fake.state.commits, 1);
    assert.equal(fake.state.rollbacks, 0);
});

test('legacy authorization commits stored-entitlement revocation before returning 403', async () => {
    const subjectId = '00000000-0000-4000-8000-000000000046';
    const issuedAt = new Date('2042-06-01T10:00:00.000Z');
    const expiresAt = new Date('2042-06-01T14:00:00.000Z');
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({ legacyCompatibilityEnforcementEnabled: 1 })]),
        result([legacyRow({ subjectId, issuedAt, expiresAt })]),
        result([subjectRow({
            subjectId,
            entitlementExpiresAt: new Date('2042-06-01T11:59:59.000Z'),
        })]),
        result([], [1]),
        result([], [2]),
        result([], [1]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    await assert.rejects(
        store.authorizeLegacy({
            verifierKeyId: 'synthetic-legacy-key',
            verifier: Buffer.alloc(32, 5),
            issuedAt,
            expiresAt,
        }),
        (error) => error.errorClass === 'forbidden-authority' && error.reason === 'ineligible',
    );
    assert.equal(fake.state.commits, 1);
    assert.ok(fake.state.queries.some(({ statement }) => statement.includes('revoke-subject-sessions')));
});

test('enforced legacy authorization gives earlier or tied entitlement expiry precedence', async () => {
    const cases = [
        {
            bindingExpiresAt: new Date('2042-06-01T11:59:00.000Z'),
            entitlementExpiresAt: new Date('2042-06-01T11:58:00.000Z'),
            expectedClass: 'forbidden-authority',
            expectedReason: 'ineligible',
        },
        {
            bindingExpiresAt: new Date('2042-06-01T11:59:00.000Z'),
            entitlementExpiresAt: new Date('2042-06-01T11:59:00.000Z'),
            expectedClass: 'forbidden-authority',
            expectedReason: 'ineligible',
        },
        {
            bindingExpiresAt: new Date('2042-06-01T11:59:00.000Z'),
            entitlementExpiresAt: SERVER_TIME,
            expectedClass: 'invalid-authority',
            expectedReason: 'legacy-binding-terminal',
        },
    ];

    for (const [index, entry] of cases.entries()) {
        const subjectId = `00000000-0000-4000-8000-00000000006${index}`;
        const issuedAt = new Date(entry.bindingExpiresAt.getTime() - 4 * 60 * 60 * 1000);
        const steps = [
            result([{ serverTime: SERVER_TIME }]),
            result([controlRow({ legacyCompatibilityEnforcementEnabled: 1 })]),
            result([legacyRow({ subjectId, issuedAt, expiresAt: entry.bindingExpiresAt })]),
            result([subjectRow({ subjectId, entitlementExpiresAt: entry.entitlementExpiresAt })]),
        ];
        if (entry.expectedClass === 'forbidden-authority') {
            steps.push(result([], [1]), result([], [2]), result([], [1]));
        }
        const fake = createFakeSql(steps);
        const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

        await assert.rejects(
            store.authorizeLegacy({
                verifierKeyId: 'synthetic-legacy-key',
                verifier: Buffer.alloc(32, 5),
                issuedAt,
                expiresAt: entry.bindingExpiresAt,
            }),
            (error) => error.errorClass === entry.expectedClass
                && error.reason === entry.expectedReason,
        );
        assert.equal(fake.state.commits, entry.expectedClass === 'forbidden-authority' ? 1 : 0);
        assert.equal(fake.state.rollbacks, entry.expectedClass === 'invalid-authority' ? 1 : 0);
    }
});

test('credential revocation leaves durable legacy evidence that eligibility reactivation cannot undo', async () => {
    const subjectId = '00000000-0000-4000-8000-000000000047';
    const issuedAt = new Date('2042-06-01T10:00:00.000Z');
    const expiresAt = new Date('2042-06-01T14:00:00.000Z');
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
        result([], [1]),
        result([], [2]),
        result([], [1]),
        result([subjectRow({ subjectId, credentialVersion: 3, sessionEpoch: 8 })]),
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
        result([], [1]),
        result([subjectRow({
            subjectId,
            credentialVersion: 3,
            credentialFingerprint: Buffer.alloc(32, 8),
            credentialFingerprintKeyId: 'synthetic-fingerprint-key-v2',
            sessionEpoch: 8,
        })]),
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({ legacyCompatibilityEnforcementEnabled: 1 })]),
        result([legacyRow({
            subjectId,
            issuedAt,
            expiresAt,
            compatibilityState: 'revoked',
            revokedAt: SERVER_TIME,
            revocationReason: 'credential-reset',
        })]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    await store.revokeForCredentialChange({
        subjectId,
        observationStartedAt: SERVER_TIME,
        expectedCredentialVersion: 2,
        expectedCredentialFingerprintKeyId: 'synthetic-fingerprint-key',
        expectedCredentialFingerprint: Buffer.alloc(32, 2),
        expectedControlVersion: 3,
        credentialFingerprintKeyId: 'synthetic-fingerprint-key-v2',
        credentialFingerprint: Buffer.alloc(32, 8),
    });
    await store.updateEligibility({
        subjectId,
        observationStartedAt: SERVER_TIME,
        expectedCredentialVersion: 3,
        expectedCredentialFingerprintKeyId: 'synthetic-fingerprint-key-v2',
        expectedCredentialFingerprint: Buffer.alloc(32, 8),
        expectedControlVersion: 3,
        rowHint: 8,
        eligibilityState: 'eligible',
        entitlementExpiresAt: new Date('2042-06-02T03:00:00.000Z'),
        eligibilityObservedAt: SERVER_TIME,
        eligibilityRevalidateAt: new Date('2042-06-01T12:05:00.000Z'),
    });
    await assert.rejects(
        store.authorizeLegacy({
            verifierKeyId: 'synthetic-legacy-key',
            verifier: Buffer.alloc(32, 5),
            issuedAt,
            expiresAt,
        }),
        (error) => error.errorClass === 'invalid-authority'
            && error.reason === 'legacy-binding-terminal',
    );

    const legacyRevocation = fake.state.queries.find(({ statement }) => (
        statement.includes('revoke-subject-legacy-bindings')
    ));
    assert.equal(legacyRevocation.parameters.reason, 'credential-reset');
    assert.equal(fake.state.commits, 2);
    assert.equal(fake.state.rollbacks, 1);
});

test('bindLegacy never treats an identical terminal verifier as idempotent success', async () => {
    const subjectId = '00000000-0000-4000-8000-000000000048';
    const issuedAt = new Date('2042-06-01T10:00:00.000Z');
    const expiresAt = new Date('2042-06-01T14:00:00.000Z');
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({
            legacyCompatibilityEnforcementEnabled: 1,
            legacyLedgerSeedingEnabled: 1,
        })]),
        result([subjectRow({ subjectId })]),
        result([legacyRow({
            subjectId,
            issuedAt,
            expiresAt,
            compatibilityState: 'revoked',
            revokedAt: SERVER_TIME,
            revocationReason: 'administrator-revocation',
        })]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    await assert.rejects(
        store.bindLegacy({
            legacyCompatibilityId: '00000000-0000-4000-8000-000000000348',
            subjectId,
            ...subjectCredentialExpectation(),
            verifierKeyId: 'synthetic-legacy-key',
            verifier: Buffer.alloc(32, 5),
            issuedAt,
            expiresAt,
        }),
        (error) => error.errorClass === 'authority-conflict'
            && error.reason === 'legacy-binding-terminal',
    );
    assert.equal(fake.state.rollbacks, 1);
    assert.equal(fake.state.queries.some(({ statement }) => statement.includes('bind-legacy */')), false);
});

test('legacy-handle leak cutoff is irreversible and cannot revoke valid target sessions', async () => {
    const subjectId = '00000000-0000-4000-8000-000000000049';
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([subjectRow({ subjectId })]),
        result([], [1]),
        result([], [2]),
        result([subjectRow({ subjectId, legacyAuthorityDisabledAt: SERVER_TIME })]),
        result([controlRow({ legacyCompatibilityEnforcementEnabled: 1 })]),
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({
            legacyCompatibilityEnforcementEnabled: 1,
            legacyLedgerSeedingEnabled: 1,
        })]),
        result([subjectRow({ subjectId, legacyAuthorityDisabledAt: SERVER_TIME })]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });

    const disabled = await store.disableLegacyAuthority({
        subjectId,
        reason: 'legacy-handle-leak',
    });
    assert.deepEqual(disabled.subject.legacyAuthorityDisabledAt, SERVER_TIME);
    assert.equal(
        fake.state.queries.some(({ statement }) => statement.includes('revoke-subject-sessions')),
        false,
    );
    assert.ok(fake.state.queries.some(({ statement }) => (
        statement.includes('revoke-subject-legacy-bindings')
    )));

    await assert.rejects(
        store.bindLegacy({
            legacyCompatibilityId: '00000000-0000-4000-8000-000000000349',
            subjectId,
            ...subjectCredentialExpectation(),
            verifierKeyId: 'synthetic-legacy-key',
            verifier: Buffer.alloc(32, 9),
            issuedAt: SERVER_TIME,
            expiresAt: new Date('2042-06-01T16:00:00.000Z'),
        }),
        (error) => error.errorClass === 'authority-conflict'
            && error.reason === 'target-authority-established',
    );
    assert.equal(fake.state.commits, 1);
    assert.equal(fake.state.rollbacks, 1);
});

test('control transition stamps rollout evidence from SQL time and rejects one-shot aging', async () => {
    const oneShot = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
    ]);
    await assert.rejects(
        createAzureSqlSessionStore({
            sql: oneShot.sql,
            connectionString: 'synthetic-only',
        }).transitionControl({
            expectedVersion: 3,
            changes: {
                legacyLedgerSeedingEnabled: true,
                legacyLedgerSeedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
                seedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
                seedingQualifiedAt: new Date('2042-06-01T08:00:00.000Z'),
                legacyCompatibilityEnforcementEnabled: true,
                legacyCompatibilityEnforcedAt: new Date('2042-06-01T08:00:00.000Z'),
            },
        }),
        (error) => error.errorClass === 'forbidden-authority'
            && error.reason === 'legacy-ledger-not-qualified',
    );
    assert.equal(oneShot.state.queries.some(({ statement }) => statement.includes('transition-control')), false);

    const qualified = {
        legacyCompatibilityEnforcementEnabled: 1,
        legacyCompatibilityEnforcedAt: new Date('2042-06-01T08:00:00.000Z'),
        legacyLedgerSeedingEnabled: 1,
        legacyLedgerSeedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
        seedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
        seedingQualifiedAt: new Date('2042-06-01T08:00:00.000Z'),
    };
    const activatedRow = controlRow({
        ...qualified,
        controlVersion: 4,
        targetRoutesEnabled: 1,
        targetSessionIssuanceEnabled: 1,
        targetSessionIssuanceStartedAt: SERVER_TIME,
        subjectTargetAdoptionEnabled: 1,
        subjectTargetAdoptionStartedAt: SERVER_TIME,
        dualStackStartedAt: SERVER_TIME,
        hardSunsetAt: new Date('2042-06-08T12:00:00.000Z'),
    });
    const activation = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow(qualified)]),
        result([], [1]),
        result([activatedRow]),
    ]);
    await createAzureSqlSessionStore({
        sql: activation.sql,
        connectionString: 'synthetic-only',
    }).transitionControl({
        expectedVersion: 3,
        changes: {
            targetRoutesEnabled: true,
            targetSessionIssuanceEnabled: true,
            targetSessionIssuanceStartedAt: new Date('2042-05-01T00:00:00.000Z'),
            subjectTargetAdoptionEnabled: true,
            subjectTargetAdoptionStartedAt: new Date('2042-05-01T00:00:00.000Z'),
            dualStackStartedAt: new Date('2042-05-01T00:00:00.000Z'),
            hardSunsetAt: new Date('2042-05-08T00:00:00.000Z'),
        },
    });
    const mutation = activation.state.queries.find(({ statement }) => statement.includes('transition-control'));
    assert.deepEqual(mutation.parameters.change_targetSessionIssuanceStartedAt, SERVER_TIME);
    assert.deepEqual(mutation.parameters.change_subjectTargetAdoptionStartedAt, SERVER_TIME);
    assert.deepEqual(mutation.parameters.change_dualStackStartedAt, SERVER_TIME);
    assert.deepEqual(
        mutation.parameters.change_hardSunsetAt,
        new Date('2042-06-08T12:00:00.000Z'),
    );
});

test('target disable and re-enable preserve the first SQL-stamped activation evidence', async () => {
    let clock = SERVER_TIME.getTime();
    const backing = createTestSessionAuthorityBacking();
    Object.assign(backing.state.control, {
        legacyCompatibilityEnforcementEnabled: true,
        legacyCompatibilityEnforcedAt: new Date('2042-06-01T08:00:00.000Z'),
        legacyLedgerSeedingEnabled: true,
        legacyLedgerSeedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
        seedingStartedAt: new Date('2042-06-01T04:00:00.000Z'),
        seedingQualifiedAt: new Date('2042-06-01T08:00:00.000Z'),
    });
    const store = createTestSessionAuthorityStore({
        testOnly: true,
        backing,
        now: () => new Date(clock),
    });
    const first = await store.transitionControl({
        expectedVersion: 1,
        changes: {
            targetRoutesEnabled: true,
            targetSessionIssuanceEnabled: true,
            subjectTargetAdoptionEnabled: true,
            dualStackStartedAt: new Date('2000-01-01T00:00:00.000Z'),
        },
    });
    const issuanceStartedAt = first.control.targetSessionIssuanceStartedAt;
    const adoptionStartedAt = first.control.subjectTargetAdoptionStartedAt;
    clock += 1_000;
    const disabled = await store.transitionControl({
        expectedVersion: first.control.version,
        changes: {
            targetSessionIssuanceEnabled: false,
            subjectTargetAdoptionEnabled: false,
        },
    });
    clock += 1_000;
    const reenabled = await store.transitionControl({
        expectedVersion: disabled.control.version,
        changes: {
            targetSessionIssuanceEnabled: true,
            subjectTargetAdoptionEnabled: true,
        },
    });
    assert.deepEqual(reenabled.control.targetSessionIssuanceStartedAt, issuanceStartedAt);
    assert.deepEqual(reenabled.control.subjectTargetAdoptionStartedAt, adoptionStartedAt);
});

test('normal legacy retirement requires a prior SQL-stamped four-hour stop', async () => {
    const stoppedAt = new Date('2042-06-01T08:00:00.000Z');
    const current = controlRow({
        legacyIssuanceEnabled: 0,
        legacyStopIssuanceAt: stoppedAt,
    });
    const retired = controlRow({
        controlVersion: 4,
        legacyIssuanceEnabled: 0,
        legacyStopIssuanceAt: stoppedAt,
        legacyAcceptanceEnabled: 0,
        legacyAcceptanceDisabledAt: SERVER_TIME,
    });
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([current]),
        result([], [1]),
        result([retired]),
    ]);
    const output = await createAzureSqlSessionStore({
        sql: fake.sql,
        connectionString: 'synthetic-only',
    }).transitionControl({
        expectedVersion: 3,
        changes: {
            legacyAcceptanceEnabled: false,
            legacyAcceptanceDisabledAt: new Date('2042-01-01T00:00:00.000Z'),
        },
    });
    assert.deepEqual(output.control.legacyAcceptanceDisabledAt, SERVER_TIME);
    const mutation = fake.state.queries.find(({ statement }) => statement.includes('transition-control'));
    assert.deepEqual(mutation.parameters.change_legacyAcceptanceDisabledAt, SERVER_TIME);
});

test('subject login remap is transactional, parameterized, and retry-idempotent', async () => {
    const subjectId = subjectRow().subjectId;
    const oldToken = Buffer.alloc(32, 1);
    const newToken = Buffer.alloc(32, 21);
    const newMapping = Buffer.from('synthetic-new-ciphertext');
    const updatedSubject = subjectRow({
        subjectId,
        loginLookupKeyId: 'synthetic-lookup-key',
        loginLookupToken: newToken,
        encryptedAccountMapping: newMapping,
        accountMappingKeyId: 'synthetic-mapping-key-v2',
    });
    const fake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([subjectRow({ subjectId, loginLookupToken: oldToken })]),
        result([]),
        result([], [1]),
        result([updatedSubject]),
        result([controlRow()]),
    ]);
    const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });
    const input = {
        subjectId,
        expectedLoginLookupKeyId: 'synthetic-lookup-key',
        expectedLoginLookupToken: oldToken,
        loginLookupKeyId: 'synthetic-lookup-key',
        loginLookupToken: newToken,
        encryptedAccountMapping: newMapping,
        accountMappingKeyId: 'synthetic-mapping-key-v2',
    };
    const output = await store.remapSubjectLogin(input);
    assert.equal(output.idempotent, false);
    assert.equal(output.subject.sessionEpoch, 7);
    assert.equal(output.subject.credentialVersion, 2);
    const mutation = fake.state.queries.find(({ statement }) => (
        statement.includes('remap-subject-login:update')
    ));
    assert.match(mutation.statement, /login_lookup_key_id = @expectedLoginLookupKeyId/);
    assert.match(mutation.statement, /login_lookup_token = @expectedLoginLookupToken/);
    assert.deepEqual(mutation.parameters.loginLookupToken, newToken);
    assert.deepEqual(mutation.parameters.expectedLoginLookupToken, oldToken);

    const retryFake = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([updatedSubject]),
        result([controlRow()]),
    ]);
    const retry = await createAzureSqlSessionStore({
        sql: retryFake.sql,
        connectionString: 'synthetic-only',
    }).remapSubjectLogin({
        ...input,
        encryptedAccountMapping: Buffer.from('synthetic-randomized-retry-ciphertext'),
    });
    assert.equal(retry.idempotent, true);
    assert.deepEqual(retry.subject.encryptedAccountMapping, newMapping);
    assert.equal(retryFake.state.queries.some(({ statement }) => (
        statement.includes('remap-subject-login:update')
    )), false);
});

test('memory login remap removes the old tuple and concurrent old-CAS attempts have one winner', async () => {
    const backing = createTestSessionAuthorityBacking();
    const subjectId = subjectRow().subjectId;
    const oldToken = Buffer.alloc(32, 1);
    const subject = subjectRow({ subjectId, loginLookupToken: oldToken });
    backing.state.subjects.set(subjectId, subject);
    backing.state.subjectsByLookup.set(`synthetic-lookup-key:${oldToken.toString('hex')}`, subjectId);
    const storeA = createTestSessionAuthorityStore({ testOnly: true, backing, now: () => SERVER_TIME });
    const storeB = createTestSessionAuthorityStore({ testOnly: true, backing, now: () => SERVER_TIME });
    const base = {
        subjectId,
        expectedLoginLookupKeyId: 'synthetic-lookup-key',
        expectedLoginLookupToken: oldToken,
        accountMappingKeyId: 'synthetic-mapping-key-v2',
    };
    const attempts = await Promise.allSettled([
        storeA.remapSubjectLogin({
            ...base,
            loginLookupKeyId: 'synthetic-lookup-key',
            loginLookupToken: Buffer.alloc(32, 22),
            encryptedAccountMapping: Buffer.from('synthetic-ciphertext-a'),
        }),
        storeB.remapSubjectLogin({
            ...base,
            loginLookupKeyId: 'synthetic-lookup-key',
            loginLookupToken: Buffer.alloc(32, 23),
            encryptedAccountMapping: Buffer.from('synthetic-ciphertext-b'),
        }),
    ]);
    assert.equal(attempts.filter(({ status }) => status === 'fulfilled').length, 1);
    const loser = attempts.find(({ status }) => status === 'rejected');
    assert.equal(loser.reason.errorClass, 'authority-unavailable');
    assert.equal(loser.reason.reason, 'subject-mapping-conflict');
    assert.equal(backing.state.subjectsByLookup.has(
        `synthetic-lookup-key:${oldToken.toString('hex')}`,
    ), false);

    const winner = attempts.find(({ status }) => status === 'fulfilled').value;
    const retry = await storeA.remapSubjectLogin({
        ...base,
        loginLookupKeyId: winner.subject.loginLookupKeyId,
        loginLookupToken: winner.subject.loginLookupToken,
        encryptedAccountMapping: Buffer.from('synthetic-fresh-randomized-ciphertext'),
    });
    assert.equal(retry.idempotent, true);
    assert.deepEqual(
        retry.subject.encryptedAccountMapping,
        winner.subject.encryptedAccountMapping,
    );

    const otherSubjectId = '00000000-0000-4000-8000-000000000099';
    const occupiedToken = Buffer.alloc(32, 24);
    backing.state.subjects.set(otherSubjectId, subjectRow({
        subjectId: otherSubjectId,
        loginLookupKeyId: 'synthetic-lookup-key',
        loginLookupToken: occupiedToken,
    }));
    backing.state.subjectsByLookup.set(
        `synthetic-lookup-key:${occupiedToken.toString('hex')}`,
        otherSubjectId,
    );
    await assert.rejects(
        storeA.remapSubjectLogin({
            subjectId,
            expectedLoginLookupKeyId: winner.subject.loginLookupKeyId,
            expectedLoginLookupToken: winner.subject.loginLookupToken,
            loginLookupKeyId: 'synthetic-lookup-key',
            loginLookupToken: occupiedToken,
            encryptedAccountMapping: Buffer.from('synthetic-conflicting-ciphertext'),
            accountMappingKeyId: 'synthetic-mapping-key-v3',
        }),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'subject-mapping-conflict',
    );
});

test('login subject mutations reject a same-generation incident after their control observation', async () => {
    const common = {
        subjectId: subjectRow().subjectId,
        observationStartedAt: SERVER_TIME,
        expectedCredentialVersion: 2,
        expectedCredentialFingerprintKeyId: 'synthetic-fingerprint-key',
        expectedCredentialFingerprint: Buffer.alloc(32, 2),
        expectedControlVersion: 3,
    };
    const cases = [
        {
            marker: 'create-or-load-subject:lookup',
            invoke: (store) => store.createOrLoadSubject({
                ...subjectRow(),
                expectedControlVersion: 3,
            }),
        },
        {
            marker: 'read-subject-by-lookup',
            invoke: (store) => store.readSubjectByLookup({
                loginLookupKeyId: 'synthetic-lookup-key',
                loginLookupToken: Buffer.alloc(32, 1),
                expectedControlVersion: 3,
            }),
        },
        {
            marker: 'update-eligibility */',
            invoke: (store) => store.updateEligibility({
                ...common,
                rowHint: 9,
                entitlementExpiresAt: new Date('2042-06-02T03:00:00.000Z'),
            }),
        },
        {
            marker: 'revoke-for-ineligibility:subject',
            invoke: (store) => store.revokeForIneligibility({
                ...common,
                eligibilityState: 'ineligible',
                entitlementExpiresAt: SERVER_TIME,
                reason: 'account-inactive',
            }),
        },
        {
            marker: 'revoke-for-credential-change:subject',
            invoke: (store) => store.revokeForCredentialChange({
                ...common,
                credentialFingerprintKeyId: 'synthetic-fingerprint-key-v2',
                credentialFingerprint: Buffer.alloc(32, 8),
            }),
        },
    ];
    for (const entry of cases) {
        const fake = createFakeSql([
            result([{ serverTime: SERVER_TIME }]),
            result([controlRow({ controlVersion: 4, incidentState: 'suspended' })]),
        ]);
        const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });
        await assert.rejects(
            entry.invoke(store),
            (error) => error.errorClass === 'authority-unavailable'
                && error.reason === 'authority-incident',
        );
        assert.equal(fake.state.queries.some(({ statement }) => statement.includes(entry.marker)), false);
    }
});

test('legacy store operations reject future, expired, and future-bound SQL-time evidence', async () => {
    const candidates = [
        {
            issuedAt: new Date('2042-06-01T12:00:00.001Z'),
            expiresAt: new Date('2042-06-01T16:00:00.001Z'),
        },
        {
            issuedAt: new Date('2042-06-01T08:00:00.000Z'),
            expiresAt: SERVER_TIME,
        },
    ];
    for (const metadata of candidates) {
        for (const mode of ['admit', 'bind']) {
            const fake = createFakeSql([
                result([{ serverTime: SERVER_TIME }]),
                result([controlRow(mode === 'bind' ? {
                    legacyCompatibilityEnforcementEnabled: 1,
                    legacyLedgerSeedingEnabled: 1,
                } : {})]),
            ]);
            const store = createAzureSqlSessionStore({ sql: fake.sql, connectionString: 'synthetic-only' });
            const promise = mode === 'bind'
                ? store.bindLegacy({
                    legacyCompatibilityId: '00000000-0000-4000-8000-000000000390',
                    subjectId: subjectRow().subjectId,
                    ...subjectCredentialExpectation(),
                    verifierKeyId: 'synthetic-legacy-key',
                    verifier: Buffer.alloc(32, 30),
                    ...metadata,
                })
                : store.admitUnboundLegacyIssuance({
                    loginLookupKeyId: 'synthetic-lookup-key',
                    loginLookupToken: Buffer.alloc(32, 31),
                    ...metadata,
                });
            await assert.rejects(
                promise,
                (error) => error.errorClass === 'authority-unavailable'
                    && error.reason === 'legacy-issuance-time-invalid',
            );
        }
    }

    const futureIssuedAt = new Date('2042-06-01T12:00:00.001Z');
    const futureExpiresAt = new Date('2042-06-01T16:00:00.001Z');
    const futureBinding = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow()]),
        result([legacyRow({ issuedAt: futureIssuedAt, expiresAt: futureExpiresAt })]),
    ]);
    await assert.rejects(
        createAzureSqlSessionStore({
            sql: futureBinding.sql,
            connectionString: 'synthetic-only',
        }).authorizeLegacy({
            verifierKeyId: 'synthetic-legacy-key',
            verifier: Buffer.alloc(32, 5),
            issuedAt: futureIssuedAt,
            expiresAt: futureExpiresAt,
        }),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'legacy-binding-integrity',
    );
});

test('legacy bind fences a credential reset and positive mutations require fresh eligibility', async () => {
    const staleBind = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([controlRow({
            legacyCompatibilityEnforcementEnabled: 1,
            legacyLedgerSeedingEnabled: 1,
        })]),
        result([subjectRow({
            credentialVersion: 3,
            credentialFingerprintKeyId: 'synthetic-fingerprint-key-v2',
            credentialFingerprint: Buffer.alloc(32, 8),
        })]),
    ]);
    await assert.rejects(
        createAzureSqlSessionStore({
            sql: staleBind.sql,
            connectionString: 'synthetic-only',
        }).bindLegacy({
            legacyCompatibilityId: '00000000-0000-4000-8000-000000000391',
            subjectId: subjectRow().subjectId,
            ...subjectCredentialExpectation(),
            verifierKeyId: 'synthetic-legacy-key',
            verifier: Buffer.alloc(32, 32),
            issuedAt: new Date('2042-06-01T11:00:00.000Z'),
            expiresAt: new Date('2042-06-01T15:00:00.000Z'),
        }),
        (error) => error.errorClass === 'authority-conflict'
            && error.reason === 'subject-credential-compare-and-replace',
    );
    assert.equal(staleBind.state.queries.some(({ statement }) => statement.includes('bind-legacy */')), false);

    const dueSubject = subjectRow({ eligibilityRevalidateAt: SERVER_TIME });
    const targetControl = controlRow({
        legacyCompatibilityEnforcementEnabled: 1,
        targetRoutesEnabled: 1,
        targetSessionIssuanceEnabled: 1,
        subjectTargetAdoptionEnabled: 1,
        dualStackStartedAt: new Date('2042-05-25T12:00:00.000Z'),
        hardSunsetAt: SERVER_TIME,
    });
    const issue = createFakeSql([
        result([{ serverTime: SERVER_TIME }]),
        result([dueSubject]),
        result([targetControl]),
    ]);
    await assert.rejects(
        createAzureSqlSessionStore({
            sql: issue.sql,
            connectionString: 'synthetic-only',
        }).issueSession({
            sessionId: '00000000-0000-4000-8000-000000000392',
            subjectId: dueSubject.subjectId,
            ...subjectCredentialExpectation(),
            verifierKeyId: 'synthetic-target-key',
            verifier: Buffer.alloc(32, 33),
            phase: 'authenticated',
            faceRequired: false,
            registrationRequired: false,
        }),
        (error) => error.errorClass === 'authority-unavailable'
            && error.reason === 'eligibility-revalidation-required',
    );
    assert.equal(issue.state.queries.some(({ statement }) => statement.includes('insert-session')), false);
});

test('due eligibility preserves provisional predecessors across every positive flow boundary', async () => {
    const operations = [
        {
            phase: 'credential-verified',
            invoke: (store, input) => store.rotateSession({
                ...input,
                allowedPhases: ['credential-verified'],
                sessionId: '00000000-0000-4000-8000-000000000401',
                verifierKeyId: 'synthetic-target-key-v2',
                verifier: Buffer.alloc(32, 41),
                phase: 'registration-pending',
            }),
        },
        {
            phase: 'credential-verified',
            invoke: (store, input) => store.reserveFaceFlow({
                ...input,
                allowedPhases: ['credential-verified'],
                flowId: '00000000-0000-4000-8000-000000000402',
                registrationState: 'registered',
            }),
        },
        {
            phase: 'credential-verified',
            flow: { challengeState: 'creating', registrationState: 'registered' },
            invoke: (store, input) => store.bindFaceChallengeAndRotate({
                ...input,
                allowedPhases: ['credential-verified'],
                sessionId: '00000000-0000-4000-8000-000000000403',
                verifierKeyId: 'synthetic-target-key-v2',
                verifier: Buffer.alloc(32, 43),
                challengeKeyId: 'synthetic-face-key',
                encryptedChallenge: Buffer.from('synthetic-challenge'),
            }),
        },
        {
            phase: 'face-pending',
            flow: { challengeState: 'active', registrationState: 'registered' },
            invoke: (store, input) => store.readFaceFlow(input),
        },
        {
            phase: 'face-pending',
            flow: { challengeState: 'active', registrationState: 'registered' },
            invoke: (store, input) => store.completeFaceSuccessAndRotate({
                ...input,
                sessionId: '00000000-0000-4000-8000-000000000404',
                verifierKeyId: 'synthetic-target-key-v2',
                verifier: Buffer.alloc(32, 44),
            }),
        },
    ];
    for (const [index, entry] of operations.entries()) {
        const backing = createTestSessionAuthorityBacking();
        const subjectId = `00000000-0000-4000-8000-${String(500 + index).padStart(12, '0')}`;
        const sessionId = `00000000-0000-4000-8000-${String(600 + index).padStart(12, '0')}`;
        const verifier = Buffer.alloc(32, 50 + index);
        const flowId = `00000000-0000-4000-8000-${String(700 + index).padStart(12, '0')}`;
        const subject = subjectRow({ subjectId, eligibilityRevalidateAt: SERVER_TIME });
        const session = sessionRow({
            subjectId,
            sessionId,
            verifier,
            phase: entry.phase,
            faceRequired: entry.phase !== 'authenticated',
            registrationRequired: entry.phase === 'registration-pending',
            version: 1,
            subjectEpochSnapshot: subject.sessionEpoch,
            credentialVersionSnapshot: subject.credentialVersion,
            globalEpochSnapshot: 1,
            authorityGenerationSnapshot: 1,
            originalIssuedAt: entry.phase === 'authenticated'
                ? new Date('2042-06-01T11:00:00.000Z')
                : new Date('2042-06-01T11:50:00.000Z'),
            phaseStartedAt: entry.phase === 'authenticated'
                ? new Date('2042-06-01T11:00:00.000Z')
                : new Date('2042-06-01T11:50:00.000Z'),
            createdAt: entry.phase === 'authenticated'
                ? new Date('2042-06-01T11:00:00.000Z')
                : new Date('2042-06-01T11:50:00.000Z'),
            expiresAt: entry.phase === 'authenticated'
                ? new Date('2042-06-01T15:00:00.000Z')
                : new Date('2042-06-01T12:10:00.000Z'),
        });
        backing.state.subjects.set(subjectId, subject);
        backing.state.sessions.set(sessionId, session);
        backing.state.sessionsByVerifier.set(`synthetic-target-key:${verifier.toString('hex')}`, session);
        if (entry.flow) {
            backing.state.flows.set(sessionId, {
                flowId,
                subjectId,
                currentSessionId: sessionId,
                challengeSessionId: entry.flow.challengeState === 'active' ? sessionId : null,
                challengeKeyId: entry.flow.challengeState === 'active'
                    ? 'synthetic-face-key'
                    : null,
                encryptedChallenge: entry.flow.challengeState === 'active'
                    ? Buffer.from('synthetic-challenge')
                    : null,
                challengeCreatedAt: entry.flow.challengeState === 'active'
                    ? session.phaseStartedAt
                    : null,
                createdAt: session.phaseStartedAt,
                consumedAt: null,
                ...entry.flow,
            });
        }
        const store = createTestSessionAuthorityStore({ testOnly: true, backing, now: () => SERVER_TIME });
        await assert.rejects(
            entry.invoke(store, {
                expectedSessionId: sessionId,
                expectedVersion: 1,
                flowId,
            }),
            (error) => error.errorClass === 'authority-unavailable'
                && error.reason === 'eligibility-revalidation-required',
        );
        assert.equal(backing.state.sessions.get(sessionId).phase, entry.phase);
        assert.equal(backing.state.sessions.size, 1);
    }
});

test('source contains no process-memory fallback or credential output and parameterizes mutations', () => {
    const source = fs.readFileSync(
        path.join(__dirname, '..', 'integrations', 'azure-sql-session-store.js'),
        'utf8',
    );
    assert.doesNotMatch(source, /console\.|process-memory|IndexVerificado|connectionString\s*\}/);
    assert.match(source, /request\.input\(name, value\)/);
    assert.match(source, /WITH \(UPDLOCK, HOLDLOCK\)/);
    assert.match(source, /SERIALIZABLE/);
    assert.match(source, /SYSUTCDATETIME\(\)/);
    assert.match(source, /DATALENGTH\(c\.legacy_signing_key_id\) = DATALENGTH\(@legacySigningKeyId\)/);
    assert.match(source, /registration_state = 'registered'/);
    assert.match(source, /\['enrollment-accepted', 'registered'\]/);
    assert.doesNotMatch(source, /identifier_verifier\s*=\s*'\$\{/);
});

function createFakeSql(steps = [], options = {}) {
    const queue = [...steps];
    const commitOutcomes = Array.isArray(options.commitOutcomes)
        ? [...options.commitOutcomes]
        : null;
    const state = {
        begins: [],
        closes: 0,
        commitAttempts: 0,
        commits: 0,
        connects: 0,
        constructorConfigurations: [],
        poolConstructions: 0,
        poolErrorListener: null,
        queries: [],
        rollbackAttempts: 0,
        rollbacks: 0,
    };

    class ConnectionPool {
        constructor(configuration) {
            state.poolConstructions += 1;
            this.configuration = configuration;
            state.constructorConfigurations.push(configuration);
            state.pool = this;
        }

        async connect() {
            state.connects += 1;
            if (options.connectError) throw options.connectError;
            return this;
        }

        on(event, listener) {
            assert.equal(event, 'error');
            state.poolErrorListener = listener;
            return this;
        }

        async close() {
            state.closes += 1;
        }
    }

    ConnectionPool.parseConnectionString = () => ({
        database: 'inert',
        options: { encrypt: false, trustServerCertificate: true },
        pool: { max: 10, min: 0 },
        server: 'synthetic.invalid',
    });

    class Transaction {
        constructor(pool) {
            this.pool = pool;
        }

        async begin(isolation) {
            state.begins.push(isolation);
            if (options.beginError) throw options.beginError;
        }

        async commit() {
            state.commitAttempts += 1;
            if (commitOutcomes && commitOutcomes.length > 0) {
                const outcome = commitOutcomes.shift();
                if (outcome instanceof Error) throw outcome;
            }
            if (options.commitError) throw options.commitError;
            state.commits += 1;
        }

        async rollback() {
            state.rollbackAttempts += 1;
            if (options.rollbackError) throw options.rollbackError;
            state.rollbacks += 1;
        }
    }

    class Request {
        constructor(owner) {
            this.owner = owner;
            this.parameters = {};
            this.parameterTypes = {};
        }

        input(name, type, value) {
            if (arguments.length === 2) {
                value = type;
                type = undefined;
            }
            this.parameters[name] = value;
            this.parameterTypes[name] = type;
            return this;
        }

        async query(statement) {
            state.queries.push({
                statement,
                parameters: { ...this.parameters },
                parameterTypes: { ...this.parameterTypes },
                owner: this.owner,
            });
            if (typeof options.queryHook === 'function') {
                const hooked = await options.queryHook({
                    parameters: { ...this.parameters },
                    state,
                    statement,
                });
                if (hooked !== undefined) return hooked;
            }
            const matchingIndex = fakeSelectStepIndex(queue, statement);
            const synthetic = matchingIndex < 0
                ? fakeSyntheticSelect(statement, this.parameters)
                : null;
            if (queue.length === 0 && !synthetic) throw new Error('Unexpected fake SQL query');
            const next = synthetic || (
                matchingIndex < 0 ? queue.shift() : queue.splice(matchingIndex, 1)[0]
            );
            if (next instanceof Error) throw next;
            return next;
        }
    }

    return {
        sql: {
            BigInt: Object.freeze({ name: 'BigInt' }),
            Binary: (length) => Object.freeze({ length, name: 'Binary' }),
            Bit: Object.freeze({ name: 'Bit' }),
            ConnectionPool,
            DateTime2: (scale) => Object.freeze({ name: 'DateTime2', scale }),
            Int: Object.freeze({ name: 'Int' }),
            ISOLATION_LEVEL: { SERIALIZABLE: 'SERIALIZABLE' },
            MAX: 'MAX',
            Request,
            Transaction,
            UniqueIdentifier: Object.freeze({ name: 'UniqueIdentifier' }),
            VarBinary: (length) => Object.freeze({ length, name: 'VarBinary' }),
            VarChar: (length) => Object.freeze({ length, name: 'VarChar' }),
        },
        state,
    };
}

function fakeSelectStepIndex(queue, statement) {
    const marker = [
        ['session-authority:server-time', (row) => (
            row && Object.keys(row).length === 1 && row.serverTime instanceof Date
        )],
        ['session-authority:select-control', (row) => row && row.controlId === 1],
        ['session-authority:select-subject', (row) => row && typeof row.credentialVersion === 'number'],
        ['session-authority:select-session', (row) => row && typeof row.phase === 'string', true],
        ['session-authority:select-authority-by-verifier', (row) => (
            row && typeof row.phase === 'string'
        ), true],
        ['session-authority:select-face-flow', (row) => row && typeof row.challengeState === 'string', true],
        ['session-authority:select-legacy-binding', (row) => row && typeof row.compatibilityState === 'string', true],
        ['session-authority:admit-unbound-legacy-issuance', (row) => (
            row && Object.hasOwn(row, 'legacyAuthorityDisabledAt')
        ), true],
        ['session-authority:remap-subject-login:target', (row) => (
            row && typeof row.credentialVersion === 'number'
        ), true],
    ].find(([name]) => statement.includes(name));
    if (!marker) return -1;
    const [, matches, allowEmpty] = marker;
    return queue.findIndex((step) => (
        !(step instanceof Error)
        && step
        && Array.isArray(step.recordset)
        && (
            (step.recordset.length > 0 && matches(step.recordset[0]))
            || (
                allowEmpty
                && step.recordset.length === 0
                && step.rowsAffected.length === 0
            )
        )
    ));
}

function fakeSyntheticSelect(statement, parameters) {
    if (statement.includes('session-authority:select-control')) {
        return result([controlRow()]);
    }
    if (statement.includes('session-authority:select-subject')) {
        return result([subjectRow({ subjectId: parameters.subjectId })]);
    }
    if (statement.includes('session-authority:select-face-flow')) return result([]);
    if (statement.includes('session-authority:select-session')) return result([]);
    if (statement.includes('session-authority:select-authority-by-verifier')) return result([]);
    if (statement.includes('session-authority:select-legacy-binding')) return result([]);
    if (statement.includes('session-authority:admit-unbound-legacy-issuance')) return result([]);
    if (statement.includes('session-authority:remap-subject-login:target')) return result([]);
    return null;
}

function result(recordset, rowsAffected = []) {
    return { recordset, recordsets: [recordset], rowsAffected };
}

function controlRow(overrides = {}) {
    const row = {
        controlId: 1,
        controlVersion: 3,
        authorityGeneration: 3,
        globalSessionEpoch: 2,
        loginLookupKeyInitialized: 1,
        loginLookupKeyMatches: 1,
        accountMappingKeyMatches: 1,
        keysetLoginLookupKeyMatches: 1,
        keysetAccountMappingKeyMatches: 1,
        authorityKeysetInitialized: 1,
        authorityKeysetAggregateMatches: 1,
        targetVerifierKeyMatches: 1,
        legacyCompatibilityKeyMatches: 1,
        credentialFingerprintKeyMatches: 1,
        faceChallengeKeyMatches: 1,
        legacySigningKeyInitialized: 1,
        legacySigningKeyMatches: 1,
        targetRoutesEnabled: 0,
        targetSessionIssuanceEnabled: 0,
        targetSessionIssuanceStartedAt: null,
        legacyLedgerSeedingEnabled: 0,
        legacyLedgerSeedingStartedAt: null,
        seedingStartedAt: null,
        seedingContinuityVersion: 1,
        seedingHeartbeatOwnerId: null,
        seedingHeartbeatAt: null,
        seedingLeaseExpiresAt: null,
        seedingQualifiedAt: null,
        legacyCompatibilityEnforcementEnabled: 0,
        legacyCompatibilityEnforcedAt: null,
        subjectTargetAdoptionEnabled: 0,
        subjectTargetAdoptionStartedAt: null,
        dualStackStartedAt: null,
        legacyIssuanceEnabled: 1,
        legacyStopIssuanceAt: null,
        legacyAcceptanceEnabled: 1,
        legacyAcceptanceDisabledAt: null,
        hardSunsetAt: null,
        incidentState: 'normal',
        incidentRecordedAt: null,
        incidentCode: null,
        targetVerifierKeyIncidentAt: null,
        legacyVerifierKeyIncidentAt: null,
        createdAt: new Date('2042-01-01T00:00:00.000Z'),
        updatedAt: SERVER_TIME,
        serverTime: SERVER_TIME,
        ...overrides,
    };
    if (row.incidentState !== 'normal') {
        if (!Object.hasOwn(overrides, 'incidentRecordedAt')) row.incidentRecordedAt = SERVER_TIME;
        if (!Object.hasOwn(overrides, 'incidentCode')) row.incidentCode = 'synthetic-incident';
    }
    if (row.legacyCompatibilityEnforcementEnabled) {
        const horizonStart = new Date(SERVER_TIME.getTime() - 4 * 60 * 60 * 1000);
        if (!Object.hasOwn(overrides, 'legacyLedgerSeedingEnabled')) {
            row.legacyLedgerSeedingEnabled = 1;
        }
        if (!Object.hasOwn(overrides, 'legacyLedgerSeedingStartedAt')) {
            row.legacyLedgerSeedingStartedAt = horizonStart;
        }
        if (!Object.hasOwn(overrides, 'seedingStartedAt')) row.seedingStartedAt = horizonStart;
        if (!Object.hasOwn(overrides, 'seedingQualifiedAt')) row.seedingQualifiedAt = SERVER_TIME;
        if (!Object.hasOwn(overrides, 'legacyCompatibilityEnforcedAt')) {
            row.legacyCompatibilityEnforcedAt = SERVER_TIME;
        }
    } else if (row.legacyLedgerSeedingEnabled) {
        if (!Object.hasOwn(overrides, 'legacyLedgerSeedingStartedAt')) {
            row.legacyLedgerSeedingStartedAt = row.seedingStartedAt || SERVER_TIME;
        }
        if (!row.seedingStartedAt) row.seedingStartedAt = row.legacyLedgerSeedingStartedAt;
    }
    if (
        row.targetSessionIssuanceEnabled
        || row.subjectTargetAdoptionEnabled
        || row.targetSessionIssuanceStartedAt
        || row.subjectTargetAdoptionStartedAt
        || row.dualStackStartedAt
        || row.hardSunsetAt
    ) {
        if (!Object.hasOwn(overrides, 'legacyCompatibilityEnforcementEnabled')) {
            row.legacyCompatibilityEnforcementEnabled = 1;
            row.legacyLedgerSeedingEnabled = 1;
            const horizonStart = new Date(SERVER_TIME.getTime() - 4 * 60 * 60 * 1000);
            if (!Object.hasOwn(overrides, 'legacyLedgerSeedingStartedAt')) {
                row.legacyLedgerSeedingStartedAt = horizonStart;
            }
            if (!Object.hasOwn(overrides, 'seedingStartedAt')) {
                row.seedingStartedAt = horizonStart;
            }
            if (!Object.hasOwn(overrides, 'seedingQualifiedAt')) row.seedingQualifiedAt = SERVER_TIME;
            if (!Object.hasOwn(overrides, 'legacyCompatibilityEnforcedAt')) {
                row.legacyCompatibilityEnforcedAt = SERVER_TIME;
            }
        }
        const startedAt = row.dualStackStartedAt
            || row.targetSessionIssuanceStartedAt
            || row.subjectTargetAdoptionStartedAt
            || (row.hardSunsetAt
                ? new Date(row.hardSunsetAt.getTime() - 7 * 24 * 60 * 60 * 1000)
                : SERVER_TIME);
        if (!Object.hasOwn(overrides, 'targetSessionIssuanceStartedAt')) {
            row.targetSessionIssuanceStartedAt = startedAt;
        }
        if (!Object.hasOwn(overrides, 'subjectTargetAdoptionStartedAt')) {
            row.subjectTargetAdoptionStartedAt = startedAt;
        }
        if (!Object.hasOwn(overrides, 'dualStackStartedAt')) row.dualStackStartedAt = startedAt;
        if (!Object.hasOwn(overrides, 'hardSunsetAt')) {
            row.hardSunsetAt = new Date(startedAt.getTime() + 7 * 24 * 60 * 60 * 1000);
        }
    }
    if (!row.legacyIssuanceEnabled && row.legacyStopIssuanceAt === null) {
        row.legacyStopIssuanceAt = row.hardSunsetAt
            ? new Date(row.hardSunsetAt.getTime() - 4 * 60 * 60 * 1000)
            : SERVER_TIME;
    }
    if (!row.legacyAcceptanceEnabled && row.legacyAcceptanceDisabledAt === null) {
        row.legacyAcceptanceDisabledAt = row.hardSunsetAt || SERVER_TIME;
    }
    if (
        row.legacyLedgerSeedingEnabled
        && !row.legacyCompatibilityEnforcementEnabled
        && !Object.hasOwn(overrides, 'seedingHeartbeatOwnerId')
    ) {
        row.seedingStartedAt ||= SERVER_TIME;
        row.seedingHeartbeatOwnerId = '00000000-0000-4000-8000-000000000091';
        row.seedingHeartbeatAt = SERVER_TIME;
        row.seedingLeaseExpiresAt = new Date(SERVER_TIME.getTime() + 2 * 60 * 1000);
    }
    return row;
}

function subjectRow(overrides = {}) {
    return {
        subjectId: '00000000-0000-4000-8000-000000000042',
        loginLookupToken: Buffer.alloc(32, 1),
        loginLookupKeyId: 'synthetic-lookup-key',
        encryptedAccountMapping: Buffer.from('synthetic-ciphertext'),
        accountMappingKeyId: 'synthetic-mapping-key',
        rowHint: 4,
        credentialVersion: 2,
        credentialFingerprint: Buffer.alloc(32, 2),
        credentialFingerprintKeyId: 'synthetic-fingerprint-key',
        sessionEpoch: 7,
        legacyAuthorityDisabledAt: null,
        eligibilityState: 'eligible',
        entitlementExpiresAt: new Date('2042-06-02T03:00:00.000Z'),
        eligibilityObservedAt: SERVER_TIME,
        eligibilityRevalidateAt: new Date('2042-06-01T12:05:00.000Z'),
        createdAt: new Date('2042-01-01T00:00:00.000Z'),
        ...overrides,
    };
}

function subjectCredentialExpectation(overrides = {}) {
    return {
        expectedCredentialVersion: 2,
        expectedCredentialFingerprint: Buffer.alloc(32, 2),
        expectedCredentialFingerprintKeyId: 'synthetic-fingerprint-key',
        ...overrides,
    };
}

function sessionRow(overrides = {}) {
    return {
        sessionId: '00000000-0000-4000-8000-000000000144',
        verifier: Buffer.alloc(32, 4),
        verifierKeyId: 'synthetic-target-key',
        subjectId: '00000000-0000-4000-8000-000000000044',
        phase: 'authenticated',
        originalIssuedAt: new Date('2042-06-01T11:00:00.000Z'),
        phaseStartedAt: new Date('2042-06-01T11:00:00.000Z'),
        expiresAt: new Date('2042-06-01T15:00:00.000Z'),
        faceRequired: false,
        registrationRequired: false,
        subjectEpochSnapshot: 7,
        credentialVersionSnapshot: 2,
        globalEpochSnapshot: 2,
        authorityGenerationSnapshot: 3,
        revokedAt: null,
        revocationReason: null,
        replacementSessionId: null,
        createdAt: new Date('2042-06-01T11:00:00.000Z'),
        ...overrides,
    };
}

function legacyRow(overrides = {}) {
    return {
        legacyCompatibilityId: '00000000-0000-4000-8000-000000000246',
        verifier: Buffer.alloc(32, 5),
        verifierKeyId: 'synthetic-legacy-key',
        subjectId: '00000000-0000-4000-8000-000000000046',
        issuedAt: new Date('2042-06-01T10:00:00.000Z'),
        expiresAt: new Date('2042-06-01T14:00:00.000Z'),
        compatibilityState: 'active',
        revokedAt: null,
        revocationReason: null,
        incidentAt: null,
        incidentCode: null,
        createdAt: new Date('2042-06-01T10:00:00.000Z'),
        ...overrides,
    };
}

function flowRow(overrides = {}) {
    return {
        flowId: '00000000-0000-4000-8000-000000000151',
        subjectId: '00000000-0000-4000-8000-000000000044',
        currentSessionId: '00000000-0000-4000-8000-000000000144',
        challengeSessionId: '00000000-0000-4000-8000-000000000144',
        registrationState: 'registered',
        challengeState: 'active',
        encryptedChallenge: Buffer.from('synthetic-challenge'),
        challengeKeyId: 'synthetic-face-key',
        challengeCreatedAt: new Date('2042-06-01T11:30:00.000Z'),
        consumedAt: null,
        createdAt: new Date('2042-06-01T11:20:00.000Z'),
        updatedAt: new Date('2042-06-01T11:30:00.000Z'),
        ...overrides,
    };
}
