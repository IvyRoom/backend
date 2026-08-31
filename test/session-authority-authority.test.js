'use strict';

process.env.NODE_ENV = 'test';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
    createHmac,
    randomBytes,
    randomUUID,
} = require('node:crypto');
const {
    createSessionAuthority,
} = require('../domains/session-authority/authority');
const {
    AUTHENTICATED_LIFETIME_MS,
    ELIGIBILITY_REVALIDATION_MS,
    LEGACY_LIFETIME_MS,
    LEGACY_SEEDING_HEARTBEAT_INTERVAL_MS,
    LEGACY_SUNSET_MAXIMUM_MS,
    NEXT_OPERATION_ROLES,
    PROVISIONAL_LIFETIME_MS,
    SESSION_PHASES,
} = require('../domains/session-authority/constants');
const {
    createLoginLookup,
} = require('../domains/session-authority/cryptography');
const {
    ERROR_CLASSES,
    authorityUnavailable,
    invalidAuthority,
} = require('../domains/session-authority/errors');
const {
    createTestSessionAuthorityBacking,
    createTestSessionAuthorityStore,
} = require('./session-authority-store-support');

const DAY_MS = 24 * 60 * 60 * 1000;
const DEFAULT_NOW = Date.parse('2042-06-01T12:00:00.000Z');
const DEFAULT_LEGACY_INSPECTOR_SKEW_MS = -37 * 60 * 1000;

function excelSerial(year, month, day) {
    const instant = Date.UTC(year, month - 1, day);
    const daysAfterEpoch = Math.floor((instant - Date.UTC(1899, 11, 31)) / DAY_MS);
    return daysAfterEpoch + (instant >= Date.UTC(1900, 2, 1) ? 1 : 0);
}

function syntheticPrivateValue(prefix) {
    return `${prefix}-${randomBytes(18).toString('base64url')}`;
}

function safeBuffersEqual(left, right) {
    return Buffer.isBuffer(left) && Buffer.isBuffer(right) && left.equals(right);
}

function createAccount({
    accessDateSerial = excelSerial(2042, 6, 30),
    accountStatus = 'Ativo',
    facePolicy = 'Ativo',
    login = syntheticPrivateValue('login'),
    password = syntheticPrivateValue('credential'),
    photoRegistrationStatus = 'Sim',
} = {}) {
    return {
        login,
        password,
        row: [
            syntheticPrivateValue('account-label'),
            syntheticPrivateValue('unused'),
            login,
            password,
            facePolicy,
            photoRegistrationStatus,
            accessDateSerial,
            accountStatus,
        ],
    };
}

function keyDescriptor(purpose) {
    return {
        keyId: `${purpose}-${randomBytes(8).toString('hex')}`,
        key: randomBytes(32),
    };
}

function createKeys() {
    return {
        targetVerifier: keyDescriptor('target'),
        legacyCompatibility: keyDescriptor('legacy'),
        credentialFingerprint: keyDescriptor('credential'),
        loginLookup: keyDescriptor('lookup'),
        accountMappingEncryption: keyDescriptor('mapping'),
        faceChallengeEncryption: keyDescriptor('face'),
    };
}

function createRuntimeControls(overrides = {}) {
    return {
        durableStoreRequired: true,
        targetRoutesEnabled: true,
        targetSessionIssuanceEnabled: true,
        legacyLedgerSeedingEnabled: true,
        legacyCompatibilityEnforcementEnabled: true,
        subjectTargetAdoptionEnabled: true,
        protectedRoutesEnabled: true,
        ...overrides,
    };
}

function createAccountSource(accountState) {
    return {
        async readRows() {
            accountState.readCalls += 1;
            if (accountState.readUnavailable) throw new Error('Synthetic account source unavailable');
            const rows = accountState.rows.map((row) => [...row]);
            if (typeof accountState.readRowsHook === 'function') {
                await accountState.readRowsHook({
                    callNumber: accountState.readCalls,
                    rows,
                });
            }
            return rows;
        },
        async downloadReferencePhoto(rowIndex) {
            accountState.downloadCalls += 1;
            accountState.lastDownloadRow = rowIndex;
            if (accountState.downloadUnavailable) throw new Error('Synthetic photo unavailable');
            return randomBytes(48);
        },
        async uploadReferencePhoto(rowIndex, referenceImage) {
            accountState.uploadCalls += 1;
            accountState.lastUploadRow = rowIndex;
            accountState.uploadWasBuffer = Buffer.isBuffer(referenceImage);
            if (accountState.uploadUnavailable) throw new Error('Synthetic upload unavailable');
        },
        async markPhotoRegistered(rowIndex) {
            accountState.markCalls += 1;
            if (accountState.markUnavailable) throw new Error('Synthetic registration update unavailable');
            if (accountState.rows[rowIndex]) accountState.rows[rowIndex][5] = 'Sim';
        },
    };
}

function createFaceSource(faceState) {
    const privateChallenges = new Set();

    return {
        async createLivenessSession(referenceImage, correlationId) {
            faceState.createCalls += 1;
            faceState.createReceivedBuffer = Buffer.isBuffer(referenceImage);
            faceState.createReceivedCorrelation = typeof correlationId === 'string'
                && correlationId.length > 0;
            if (faceState.createUnavailable) throw new Error('Synthetic Face create unavailable');
            if (faceState.invalidCreateResponse) return {};

            const privateChallengeId = randomBytes(32).toString('base64url');
            privateChallenges.add(privateChallengeId);
            return {
                authToken: randomBytes(24).toString('base64url'),
                privateChallengeId,
            };
        },
        async readLivenessSessionResult(privateChallengeId) {
            faceState.readCalls += 1;
            if (privateChallenges.has(privateChallengeId)) {
                faceState.boundReadCalls += 1;
            } else {
                faceState.unboundReadCalls += 1;
            }
            if (faceState.readUnavailable) throw new Error('Synthetic Face result unavailable');
            return typeof faceState.result === 'function'
                ? faceState.result()
                : faceState.result;
        },
    };
}

function createLegacyHandleAuthority(readNow) {
    const signingKey = randomBytes(32);
    const metadataByHandle = new Map();
    const createNowValues = [];
    const inspectNowValues = [];

    function createHandle(
        rowIndex,
        nowMs = readNow().getTime() + DEFAULT_LEGACY_INSPECTOR_SKEW_MS,
    ) {
        createNowValues.push(nowMs);
        const issuedAtMs = Math.floor(nowMs / 1000) * 1000;
        const handle = createHmac('sha256', signingKey)
            .update(`${rowIndex}:${issuedAtMs}`)
            .digest('base64url');
        const issuedAt = new Date(issuedAtMs);
        metadataByHandle.set(handle, {
            rowIndex,
            issuedAt,
            expiresAt: new Date(issuedAtMs + LEGACY_LIFETIME_MS),
        });
        return handle;
    }

    return {
        createNowValues,
        createHandle,
        defaultNow: () => readNow().getTime() + DEFAULT_LEGACY_INSPECTOR_SKEW_MS,
        inspectNowValues,
        inspectHandle(
            rawHandle,
            nowMs = readNow().getTime() + DEFAULT_LEGACY_INSPECTOR_SKEW_MS,
        ) {
            inspectNowValues.push(nowMs);
            const metadata = metadataByHandle.get(rawHandle);
            if (!metadata) throw invalidAuthority('invalid-legacy-handle');
            return {
                rowIndex: metadata.rowIndex,
                issuedAt: new Date(metadata.issuedAt.getTime()),
                expiresAt: new Date(metadata.expiresAt.getTime()),
            };
        },
    };
}

async function createHarness({
    accounts = [createAccount()],
    controlChanges = {},
    qualify = true,
    runtimeControls = {},
    startTime = DEFAULT_NOW,
} = {}) {
    let clockMs = qualify ? startTime - LEGACY_LIFETIME_MS : startTime;
    const keys = createKeys();
    const backing = createTestSessionAuthorityBacking();
    const loginLookupKeyCommitment = createHmac('sha256', keys.loginLookup.key)
        .update('session-authority/login-lookup-key-binding/v1', 'utf8')
        .digest();
    backing.state.control.loginLookupKeyId = keys.loginLookup.keyId;
    backing.state.control.loginLookupKeyCommitment = Buffer.from(loginLookupKeyCommitment);
    const failureA = {};
    const failureB = {};
    const readNow = () => new Date(clockMs);
    const createStore = (expectedAuthorityGeneration = 1, failure = {}) => (
        createTestSessionAuthorityStore({
            testOnly: true,
            backing,
            now: readNow,
            failure,
            expectedAuthorityGeneration,
            loginLookupKeyId: keys.loginLookup.keyId,
            loginLookupKeyCommitment,
        })
    );
    const storeA = createStore(1, failureA);
    const storeB = createStore(1, failureB);
    const accountState = {
        rows: accounts.map(({ row }) => [...row]),
        readCalls: 0,
        readRowsHook: null,
        readUnavailable: false,
        downloadCalls: 0,
        downloadUnavailable: false,
        lastDownloadRow: null,
        uploadCalls: 0,
        uploadUnavailable: false,
        lastUploadRow: null,
        uploadWasBuffer: false,
        markCalls: 0,
        markUnavailable: false,
    };
    const faceState = {
        createCalls: 0,
        createReceivedBuffer: false,
        createReceivedCorrelation: false,
        createUnavailable: false,
        invalidCreateResponse: false,
        readCalls: 0,
        boundReadCalls: 0,
        unboundReadCalls: 0,
        readUnavailable: false,
        result: {},
    };
    const accountSource = createAccountSource(accountState);
    const faceSource = createFaceSource(faceState);
    const legacyHandleAuthority = createLegacyHandleAuthority(readNow);
    const selectedRuntimeControls = createRuntimeControls(runtimeControls);

    function createAuthority(store = storeA, authorityRuntimeControls = selectedRuntimeControls) {
        return createSessionAuthority({
            store,
            accountSource,
            faceSource,
            legacyHandleAuthority,
            keys,
            runtimeControls: authorityRuntimeControls,
            randomBytes,
            createSubjectId: randomUUID,
            createSessionId: randomUUID,
            createFlowId: randomUUID,
            createCorrelationId: randomUUID,
            formatLegacyAccessDate: (serial) => `serial-${serial}`,
        });
    }

    const authorityA = createAuthority(storeA);
    const authorityB = createAuthority(storeB);
    if (qualify) {
        const started = await storeA.transitionControl({
            expectedVersion: 1,
            changes: { legacyLedgerSeedingEnabled: true },
        });
        const continuityOwnerId = randomUUID();
        await storeA.heartbeatLegacySeedingContinuity({ ownerId: continuityOwnerId });
        for (
            let elapsed = 0;
            elapsed < LEGACY_LIFETIME_MS;
            elapsed += LEGACY_SEEDING_HEARTBEAT_INTERVAL_MS
        ) {
            clockMs += LEGACY_SEEDING_HEARTBEAT_INTERVAL_MS;
            await storeA.heartbeatLegacySeedingContinuity({ ownerId: continuityOwnerId });
        }
        const qualifiedAt = readNow();
        await storeA.transitionControl({
            expectedVersion: started.control.version,
            changes: {
                targetRoutesEnabled: true,
                targetSessionIssuanceEnabled: true,
                seedingQualifiedAt: qualifiedAt,
                legacyCompatibilityEnforcementEnabled: true,
                subjectTargetAdoptionEnabled: true,
                dualStackStartedAt: qualifiedAt,
                ...controlChanges,
            },
        });
    }

    return {
        accounts,
        accountSource,
        accountState,
        authorityA,
        authorityB,
        backing,
        createAuthority,
        createStore,
        faceSource,
        faceState,
        failureA,
        failureB,
        keys,
        legacyHandleAuthority,
        readNow,
        storeA,
        storeB,
        advance(milliseconds) {
            clockMs += milliseconds;
        },
        setTime(value) {
            const parsed = value instanceof Date ? value.getTime() : new Date(value).getTime();
            if (!Number.isFinite(parsed)) throw new TypeError('Synthetic clock must be valid');
            clockMs = parsed;
        },
    };
}

function rowFor(harness, account) {
    return harness.accountState.rows.find((row) => row[2] === account.login);
}

function readSubjectForExactLogin(harness, exactLogin, store = harness.storeA) {
    const lookup = createLoginLookup(exactLogin, harness.keys.loginLookup);
    return store.readSubjectByLookup({
        loginLookupKeyId: lookup.keyId,
        loginLookupToken: lookup.token,
    });
}

function generatedIdentifier() {
    return randomBytes(32).toString('base64url');
}

function deferred() {
    let resolve;
    const promise = new Promise((resolvePromise) => {
        resolve = resolvePromise;
    });
    return { promise, resolve };
}

async function advanceWithContinuity(harness, store, ownerId, durationMs) {
    for (let elapsed = 0; elapsed < durationMs;) {
        const step = Math.min(LEGACY_SEEDING_HEARTBEAT_INTERVAL_MS, durationMs - elapsed);
        harness.advance(step);
        elapsed += step;
        await store.heartbeatLegacySeedingContinuity({ ownerId });
    }
}

function phaseOf(result) {
    return result.body.authenticationPhase;
}

function isAuthorityError(error, errorClass, reason) {
    return Boolean(error && error.errorClass === errorClass && error.reason === reason);
}

async function expectAuthorityError(operation, errorClass, reason) {
    await assert.rejects(
        typeof operation === 'function' ? operation() : operation,
        (error) => isAuthorityError(error, errorClass, reason),
    );
}

async function expectAuthorityErrorClass(operation, errorClass) {
    await assert.rejects(
        typeof operation === 'function' ? operation() : operation,
        (error) => Boolean(error && error.errorClass === errorClass),
    );
}

async function targetLogin(harness, account = harness.accounts[0], {
    authority = harness.authorityA,
    presentedIdentifier,
    password = account.password,
} = {}) {
    return authority.loginTarget({
        login: account.login,
        password,
        presentedIdentifier,
    });
}

function activeSessions(backing) {
    return [...backing.state.sessions.values()].filter((session) => (
        [
            SESSION_PHASES.credentialVerified,
            SESSION_PHASES.registrationPending,
            SESSION_PHASES.facePending,
            SESSION_PHASES.authenticated,
        ].includes(session.phase)
        && session.revokedAt === null
    ));
}

test('fresh credentials capture exact Face policy and issue only the selected phase', async (t) => {
    await t.test('exact Ativo starts a 20-minute registered-photo provisional session', async () => {
        const account = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Sim' });
        const harness = await createHarness({ accounts: [account] });
        const login = await targetLogin(harness, account);

        assert.equal(phaseOf(login), SESSION_PHASES.credentialVerified);
        assert.equal(
            Date.parse(login.body.expiresAt) - harness.readNow().getTime(),
            PROVISIONAL_LIFETIME_MS,
        );
        assert.deepEqual(login.body.allowedNextOperations, [NEXT_OPERATION_ROLES.faceChallenge]);
        await expectAuthorityError(
            harness.authorityA.authorizeProtected(login.issuance.identifier),
            ERROR_CLASSES.forbidden,
            'wrong-phase',
        );
        await expectAuthorityError(
            harness.authorityA.registrationEnrollment(login.issuance.identifier),
            ERROR_CLASSES.forbidden,
            'registration-not-required',
        );
    });

    await t.test('exact Inativo starts a four-hour authenticated session even without a photo', async () => {
        const account = createAccount({ facePolicy: 'Inativo', photoRegistrationStatus: 'Não' });
        const harness = await createHarness({ accounts: [account] });
        const login = await targetLogin(harness, account);

        assert.equal(phaseOf(login), SESSION_PHASES.authenticated);
        assert.equal(
            Date.parse(login.body.expiresAt) - harness.readNow().getTime(),
            AUTHENTICATED_LIFETIME_MS,
        );
        assert.deepEqual(login.body.allowedNextOperations, [
            NEXT_OPERATION_ROLES.protectedLearning,
            NEXT_OPERATION_ROLES.revokeAll,
        ]);
        const authorized = await harness.authorityA.authorizeProtected(login.issuance.identifier);
        assert.equal(authorized.platformRowIndex, 0);
        await expectAuthorityError(
            harness.authorityA.createExistingPhotoChallenge(login.issuance.identifier),
            ERROR_CLASSES.forbidden,
            'wrong-phase',
        );
    });

    await t.test('missing, blank, unreadable, and non-exact policy values fail closed', async () => {
        for (const facePolicy of [undefined, null, '', 'ativo', 'Inativo ', true]) {
            const account = createAccount();
            account.row[4] = facePolicy;
            const harness = await createHarness({ accounts: [account] });
            await expectAuthorityError(
                targetLogin(harness, account),
                ERROR_CLASSES.unavailable,
                'invalid-face-policy',
            );
            assert.equal(harness.backing.state.sessions.size, 0);
        }
    });

    await t.test('a later workbook policy edit neither promotes nor downgrades existing authority', async () => {
        const faceAccount = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Sim' });
        const directAccount = createAccount({ facePolicy: 'Inativo', photoRegistrationStatus: 'Não' });
        const harness = await createHarness({ accounts: [faceAccount, directAccount] });
        const provisional = await targetLogin(harness, faceAccount);
        const authenticated = await targetLogin(harness, directAccount);

        rowFor(harness, faceAccount)[4] = 'Inativo';
        rowFor(harness, directAccount)[4] = 'Ativo';

        const challenge = await harness.authorityA.createExistingPhotoChallenge(
            provisional.issuance.identifier,
        );
        assert.equal(phaseOf(await harness.authorityA.current(challenge.issuance.identifier)), SESSION_PHASES.facePending);
        assert.equal(
            (await harness.authorityA.authorizeProtected(authenticated.issuance.identifier)).platformRowIndex,
            1,
        );
        assert.equal(phaseOf(await targetLogin(harness, faceAccount)), SESSION_PHASES.authenticated);
        assert.equal(phaseOf(await targetLogin(harness, directAccount)), SESSION_PHASES.credentialVerified);
        assert.equal(
            phaseOf(await harness.authorityB.current(authenticated.issuance.identifier)),
            SESSION_PHASES.authenticated,
        );
    });
});

test('registration enrollment and registration-bound Face creation rotate exactly once', async () => {
    const account = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Não' });
    const harness = await createHarness({ accounts: [account] });
    const login = await targetLogin(harness, account);
    const loginIdentifier = login.issuance.identifier;

    assert.deepEqual(login.body.allowedNextOperations, [NEXT_OPERATION_ROLES.registrationEnrollment]);
    await expectAuthorityError(
        harness.authorityA.createExistingPhotoChallenge(loginIdentifier),
        ERROR_CLASSES.forbidden,
        'face-challenge-not-allowed',
    );

    const enrollment = await harness.authorityA.registrationEnrollment(loginIdentifier);
    assert.equal(enrollment.status, 204);
    assert.equal(Boolean(enrollment.issuance), true);
    assert.equal(enrollment.issuance.identifier === loginIdentifier, false);
    assert.equal(enrollment.issuance.expiresAt.getTime(), login.issuance.expiresAt.getTime());
    await expectAuthorityError(
        harness.authorityA.current(loginIdentifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.rotatedOut,
    );

    const enrolledIdentifier = enrollment.issuance.identifier;
    const enrolled = await harness.authorityB.current(enrolledIdentifier);
    assert.equal(phaseOf(enrolled), SESSION_PHASES.registrationPending);
    assert.deepEqual(enrolled.body.allowedNextOperations, [NEXT_OPERATION_ROLES.registrationChallenge]);
    const repeatedEnrollment = await harness.authorityA.registrationEnrollment(enrolledIdentifier);
    assert.equal(repeatedEnrollment.status, 204);
    assert.equal(Object.hasOwn(repeatedEnrollment, 'issuance'), false);
    await expectAuthorityError(
        harness.authorityA.createRegistrationChallenge(enrolledIdentifier, Buffer.alloc(0)),
        ERROR_CLASSES.forbidden,
        'reference-photo-required',
    );
    assert.equal(harness.accountState.uploadCalls, 0);

    const challenge = await harness.authorityB.createRegistrationChallenge(
        enrolledIdentifier,
        randomBytes(64),
    );
    assert.equal(challenge.status, 200);
    assert.deepEqual(Object.keys(challenge.body), ['Azure_Face_API_LivenessSession_authToken']);
    assert.equal(challenge.issuance.expiresAt.getTime(), login.issuance.expiresAt.getTime());
    assert.equal(harness.accountState.uploadCalls, 1);
    assert.equal(harness.accountState.markCalls, 1);
    assert.equal(harness.faceState.createCalls, 1);
    assert.equal(harness.accountState.uploadWasBuffer, true);
    assert.equal(harness.faceState.createReceivedBuffer, true);
    assert.equal(harness.faceState.createReceivedCorrelation, true);
    assert.equal([...harness.backing.state.flows.values()][0].registrationState, 'registered');
    await expectAuthorityError(
        harness.authorityA.current(enrolledIdentifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.rotatedOut,
    );

    const facePendingIdentifier = challenge.issuance.identifier;
    assert.equal(phaseOf(await harness.authorityA.current(facePendingIdentifier)), SESSION_PHASES.facePending);
    await expectAuthorityError(
        harness.authorityA.createRegistrationChallenge(facePendingIdentifier, randomBytes(32)),
        ERROR_CLASSES.conflict,
        'face-challenge-active',
    );
    assert.equal(harness.accountState.uploadCalls, 1);
    assert.equal(harness.faceState.createCalls, 1);
    await expectAuthorityError(
        harness.authorityA.authorizeProtected(facePendingIdentifier),
        ERROR_CLASSES.forbidden,
        'wrong-phase',
    );
});

test('existing-photo Face completion is backend-bound and preserves state by result class', async (t) => {
    async function facePendingHarness() {
        const account = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Sim' });
        const harness = await createHarness({ accounts: [account] });
        const login = await targetLogin(harness, account);
        const challenge = await harness.authorityA.createExistingPhotoChallenge(
            login.issuance.identifier,
        );
        return { account, challenge, harness, login };
    }

    await t.test('challenge binds one private provider reference and rejects an active repeat as conflict', async () => {
        const { challenge, harness, login } = await facePendingHarness();

        assert.equal(harness.accountState.downloadCalls, 1);
        assert.equal(harness.accountState.lastDownloadRow, 0);
        assert.equal(harness.faceState.createCalls, 1);
        assert.deepEqual(Object.keys(challenge.body), ['Azure_Face_API_LivenessSession_authToken']);
        assert.equal(challenge.issuance.expiresAt.getTime(), login.issuance.expiresAt.getTime());
        await expectAuthorityError(
            harness.authorityB.createExistingPhotoChallenge(challenge.issuance.identifier),
            ERROR_CLASSES.conflict,
            'face-challenge-active',
        );
        assert.equal(harness.accountState.downloadCalls, 1);
        assert.equal(harness.faceState.createCalls, 1);
    });

    await t.test('pending and provider-unavailable results preserve face-pending authority', async () => {
        const { challenge, harness } = await facePendingHarness();
        const identifier = challenge.issuance.identifier;
        harness.faceState.result = { providerState: 'pending' };

        await expectAuthorityError(
            harness.authorityA.completeFace(identifier, {
                livenessDecision: 'realface',
                matchDecision: true,
            }),
            ERROR_CLASSES.conflict,
            'face-result-pending',
        );
        assert.equal(phaseOf(await harness.authorityB.current(identifier)), SESSION_PHASES.facePending);
        assert.equal(harness.faceState.boundReadCalls, 1);
        assert.equal(harness.faceState.unboundReadCalls, 0);

        harness.faceState.readUnavailable = true;
        await expectAuthorityError(
            harness.authorityB.completeFace(identifier),
            ERROR_CLASSES.unavailable,
            'face-provider-unavailable',
        );
        harness.faceState.readUnavailable = false;
        assert.equal(phaseOf(await harness.authorityA.current(identifier)), SESSION_PHASES.facePending);

        harness.faceState.result = { livenessDecision: 'unexpected', matchDecision: false };
        await expectAuthorityError(
            harness.authorityB.completeFace(identifier),
            ERROR_CLASSES.unavailable,
            'face-provider-response-invalid',
        );
        assert.equal(phaseOf(await harness.authorityA.current(identifier)), SESSION_PHASES.facePending);

        harness.faceState.result = () => {
            throw new TypeError('Synthetic provider failed status');
        };
        await expectAuthorityError(
            harness.authorityB.completeFace(identifier),
            ERROR_CLASSES.unavailable,
            'face-provider-unavailable',
        );
        assert.equal(phaseOf(await harness.authorityA.current(identifier)), SESSION_PHASES.facePending);
    });

    await t.test('an invalid provider create response remains provisional and blocks blind retry', async () => {
        const account = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Sim' });
        const harness = await createHarness({ accounts: [account] });
        const login = await targetLogin(harness, account);
        harness.faceState.invalidCreateResponse = true;

        await expectAuthorityError(
            harness.authorityA.createExistingPhotoChallenge(login.issuance.identifier),
            ERROR_CLASSES.unavailable,
            'face-provider-response-invalid',
        );
        assert.equal(
            phaseOf(await harness.authorityB.current(login.issuance.identifier)),
            SESSION_PHASES.credentialVerified,
        );
        assert.equal([...harness.backing.state.flows.values()][0].registrationState, 'registered');
        assert.equal(
            [...harness.backing.state.flows.values()][0].challengeState,
            'reconciliation-required',
        );
        harness.faceState.invalidCreateResponse = false;
        await expectAuthorityError(
            harness.authorityB.createExistingPhotoChallenge(login.issuance.identifier),
            ERROR_CLASSES.conflict,
            'face-challenge-active',
        );
        assert.equal(harness.faceState.createCalls, 1);
    });

    await t.test('a definitive failed factor revokes only the active verifier', async () => {
        const { challenge, harness } = await facePendingHarness();
        harness.faceState.result = { livenessDecision: 'spoofface', matchDecision: false };

        await expectAuthorityError(
            harness.authorityA.completeFace(challenge.issuance.identifier),
            ERROR_CLASSES.forbidden,
            'face-factor-failed',
        );
        await expectAuthorityError(
            harness.authorityB.current(challenge.issuance.identifier),
            ERROR_CLASSES.invalid,
            SESSION_PHASES.revoked,
        );
        assert.equal(activeSessions(harness.backing).length, 0);
    });

    await t.test('a passing backend result consumes once, rotates, and starts a fresh four-hour clock', async () => {
        const { challenge, harness, login } = await facePendingHarness();
        harness.advance(7 * 60 * 1000);
        harness.faceState.result = { livenessDecision: 'realface', matchDecision: true };

        const completed = await harness.authorityA.completeFace(challenge.issuance.identifier);
        assert.equal(phaseOf(completed), SESSION_PHASES.authenticated);
        assert.equal(completed.issuance.identifier === challenge.issuance.identifier, false);
        assert.equal(
            completed.issuance.expiresAt.getTime() - harness.readNow().getTime(),
            AUTHENTICATED_LIFETIME_MS,
        );
        assert.equal(completed.issuance.expiresAt.getTime() === login.issuance.expiresAt.getTime(), false);
        await expectAuthorityError(
            harness.authorityB.current(challenge.issuance.identifier),
            ERROR_CLASSES.invalid,
            SESSION_PHASES.rotatedOut,
        );

        const repeated = await harness.authorityB.completeFace(completed.issuance.identifier);
        assert.equal(phaseOf(repeated), SESSION_PHASES.authenticated);
        assert.equal(Object.hasOwn(repeated, 'issuance'), false);
        assert.equal(harness.faceState.readCalls, 1);
        assert.equal(
            (await harness.authorityA.authorizeProtected(completed.issuance.identifier)).platformRowIndex,
            0,
        );
    });
});

test('direct authenticated authority cannot masquerade as completed Face authority', async () => {
    const account = createAccount({ facePolicy: 'Inativo', photoRegistrationStatus: 'Não' });
    const harness = await createHarness({ accounts: [account] });
    const login = await targetLogin(harness, account);

    await expectAuthorityError(
        harness.authorityA.completeFace(login.issuance.identifier),
        ERROR_CLASSES.forbidden,
        'face-completion-not-applicable',
    );
    assert.equal(harness.faceState.readCalls, 0);
});

test('Face provider references are bound to their exact flow, subject, and session', async () => {
    const firstAccount = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Sim' });
    const secondAccount = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Sim' });
    const harness = await createHarness({ accounts: [firstAccount, secondAccount] });
    const firstLogin = await targetLogin(harness, firstAccount);
    const secondLogin = await targetLogin(harness, secondAccount);
    const firstChallenge = await harness.authorityA.createExistingPhotoChallenge(
        firstLogin.issuance.identifier,
    );
    await harness.authorityB.createExistingPhotoChallenge(secondLogin.issuance.identifier);
    const flows = [...harness.backing.state.flows.values()];
    assert.equal(flows.length, 2);
    const firstCiphertext = Buffer.from(flows[0].encryptedChallenge);
    flows[0].encryptedChallenge = Buffer.from(flows[1].encryptedChallenge);
    flows[1].encryptedChallenge = firstCiphertext;
    const activeBefore = activeSessions(harness.backing).map(({ sessionId, phase }) => ({
        sessionId,
        phase,
    }));

    await expectAuthorityError(
        harness.authorityA.completeFace(firstChallenge.issuance.identifier),
        ERROR_CLASSES.unavailable,
        'face-challenge-integrity',
    );
    assert.equal(harness.faceState.readCalls, 0);
    assert.deepEqual(
        activeSessions(harness.backing).map(({ sessionId, phase }) => ({ sessionId, phase })),
        activeBefore,
    );
});

test('phase permissions and independently owned subjects fail closed', async () => {
    const provisionalAccount = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Sim' });
    const authenticatedAccount = createAccount({ facePolicy: 'Inativo' });
    const otherAccount = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({
        accounts: [provisionalAccount, authenticatedAccount, otherAccount],
    });
    const provisional = await targetLogin(harness, provisionalAccount);
    const authenticated = await targetLogin(harness, authenticatedAccount);
    const other = await targetLogin(harness, otherAccount);
    const otherAuthority = await harness.authorityA.authorizeProtected(other.issuance.identifier);

    assert.equal(phaseOf(await harness.authorityA.current(provisional.issuance.identifier)), SESSION_PHASES.credentialVerified);
    assert.equal(phaseOf(await harness.authorityA.current(authenticated.issuance.identifier)), SESSION_PHASES.authenticated);
    await expectAuthorityError(
        harness.authorityA.revokeAll(provisional.issuance.identifier),
        ERROR_CLASSES.forbidden,
        'wrong-phase',
    );
    await expectAuthorityError(
        harness.authorityA.completeFace(provisional.issuance.identifier),
        ERROR_CLASSES.forbidden,
        'wrong-phase',
    );
    await expectAuthorityError(
        harness.authorityA.registrationEnrollment(authenticated.issuance.identifier),
        ERROR_CLASSES.forbidden,
        'wrong-phase',
    );
    await expectAuthorityError(
        harness.authorityA.createRegistrationChallenge(
            authenticated.issuance.identifier,
            randomBytes(24),
        ),
        ERROR_CLASSES.forbidden,
        'wrong-phase',
    );
    await expectAuthorityError(
        harness.authorityA.authorizeProtected(
            authenticated.issuance.identifier,
            otherAuthority.subjectId,
        ),
        ERROR_CLASSES.forbidden,
        'wrong-subject',
    );
    assert.equal(
        (await harness.authorityA.authorizeProtected(authenticated.issuance.identifier)).platformRowIndex,
        1,
    );
});

test('same-predecessor rotations serialize across store instances', async () => {
    const account = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Não' });
    const harness = await createHarness({ accounts: [account] });
    const login = await targetLogin(harness, account);
    const outcomes = await Promise.allSettled([
        harness.authorityA.registrationEnrollment(login.issuance.identifier),
        harness.authorityB.registrationEnrollment(login.issuance.identifier),
    ]);
    const winners = outcomes.filter(({ status }) => status === 'fulfilled');
    const losers = outcomes.filter(({ status }) => status === 'rejected');

    assert.equal(winners.length, 1);
    assert.equal(losers.length, 1);
    assert.equal(
        isAuthorityError(losers[0].reason, ERROR_CLASSES.conflict, 'session-compare-and-replace'),
        true,
    );
    assert.equal(Boolean(winners[0].value.issuance), true);
    assert.equal(
        phaseOf(await harness.authorityB.current(winners[0].value.issuance.identifier)),
        SESSION_PHASES.registrationPending,
    );
    await expectAuthorityError(
        harness.authorityA.current(login.issuance.identifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.rotatedOut,
    );
    assert.equal(activeSessions(harness.backing).length, 1);
});

test('same active login predecessor has one replacement winner across instances', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const first = await targetLogin(harness, account);
    const outcomes = await Promise.allSettled([
        targetLogin(harness, account, {
            authority: harness.authorityA,
            presentedIdentifier: first.issuance.identifier,
        }),
        targetLogin(harness, account, {
            authority: harness.authorityB,
            presentedIdentifier: first.issuance.identifier,
        }),
    ]);
    const winners = outcomes.filter(({ status }) => status === 'fulfilled');
    const losers = outcomes.filter(({ status }) => status === 'rejected');

    assert.equal(winners.length, 1);
    assert.equal(losers.length, 1);
    assert.equal(
        isAuthorityError(losers[0].reason, ERROR_CLASSES.conflict, 'session-compare-and-replace'),
        true,
    );
    assert.equal(
        phaseOf(await harness.authorityA.current(winners[0].value.issuance.identifier)),
        SESSION_PHASES.authenticated,
    );
    await expectAuthorityError(
        harness.authorityB.current(first.issuance.identifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.rotatedOut,
    );
});

test('absolute clocks use store UTC, never extend, and promote only from a fresh clock', async (t) => {
    await t.test('provisional status and rotations preserve the original 20-minute deadline', async () => {
        const account = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Não' });
        const harness = await createHarness({ accounts: [account] });
        const login = await targetLogin(harness, account);
        const deadline = login.issuance.expiresAt.getTime();

        harness.advance(ELIGIBILITY_REVALIDATION_MS + 17_000);
        const status = await harness.authorityB.current(login.issuance.identifier);
        assert.equal(status.body.serverTime, harness.readNow().toISOString());
        assert.equal(Date.parse(status.body.expiresAt), deadline);

        const enrollment = await harness.authorityA.registrationEnrollment(login.issuance.identifier);
        assert.equal(enrollment.issuance.expiresAt.getTime(), deadline);
        harness.advance(PROVISIONAL_LIFETIME_MS - (ELIGIBILITY_REVALIDATION_MS + 17_000) - 1);
        assert.equal(
            phaseOf(await harness.authorityB.current(enrollment.issuance.identifier)),
            SESSION_PHASES.registrationPending,
        );
        harness.advance(1);
        await expectAuthorityError(
            harness.authorityA.current(enrollment.issuance.identifier),
            ERROR_CLASSES.invalid,
            'expired',
        );
    });

    await t.test('authenticated status has no idle extension and expires at exactly four hours', async () => {
        const account = createAccount({ facePolicy: 'Inativo' });
        const harness = await createHarness({ accounts: [account] });
        const login = await targetLogin(harness, account);
        const deadline = login.issuance.expiresAt.getTime();

        harness.advance(2 * 60 * 60 * 1000);
        assert.equal(
            Date.parse((await harness.authorityA.current(login.issuance.identifier)).body.expiresAt),
            deadline,
        );
        harness.advance((2 * 60 * 60 * 1000) - 1);
        assert.equal(
            Date.parse((await harness.authorityB.current(login.issuance.identifier)).body.expiresAt),
            deadline,
        );
        harness.advance(1);
        await expectAuthorityError(
            harness.authorityA.current(login.issuance.identifier),
            ERROR_CLASSES.invalid,
            'expired',
        );
    });
});

test('current logout is effect-idempotent and isolated from other devices', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const firstDevice = await targetLogin(harness, account);
    const secondDevice = await targetLogin(harness, account, { authority: harness.authorityB });

    const outcomes = await Promise.all([
        harness.authorityA.logout(firstDevice.issuance.identifier),
        harness.authorityB.logout(firstDevice.issuance.identifier),
    ]);
    assert.equal(outcomes.every(({ status }) => status === 204), true);
    assert.equal((await harness.authorityA.logout(undefined)).status, 204);
    assert.equal((await harness.authorityA.logout(syntheticPrivateValue('malformed'))).status, 204);
    assert.equal((await harness.authorityA.logout(generatedIdentifier())).status, 204);
    await expectAuthorityError(
        harness.authorityA.current(firstDevice.issuance.identifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.revoked,
    );
    assert.equal(
        phaseOf(await harness.authorityB.current(secondDevice.issuance.identifier)),
        SESSION_PHASES.authenticated,
    );

    const replacement = await targetLogin(harness, account, {
        presentedIdentifier: secondDevice.issuance.identifier,
    });
    await expectAuthorityError(
        harness.authorityB.current(secondDevice.issuance.identifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.rotatedOut,
    );
    assert.equal(phaseOf(await harness.authorityA.current(replacement.issuance.identifier)), SESSION_PHASES.authenticated);
});

test('current logout fails closed when the central target-route gate is disabled or unavailable', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const login = await targetLogin(harness, account);
    const session = activeSessions(harness.backing)[0];
    const originalVersion = session.version;

    const currentControl = await harness.storeA.readControl();
    await harness.storeA.transitionControl({
        expectedVersion: currentControl.control.version,
        changes: {
            targetRoutesEnabled: false,
            targetSessionIssuanceEnabled: false,
            subjectTargetAdoptionEnabled: false,
        },
    });

    for (const identifier of [login.issuance.identifier, generatedIdentifier()]) {
        await expectAuthorityError(
            harness.authorityA.logout(identifier),
            ERROR_CLASSES.unavailable,
            'target-routes-disabled',
        );
    }
    assert.equal((await harness.authorityA.logout(undefined)).status, 204);
    assert.equal((await harness.authorityA.logout(syntheticPrivateValue('malformed'))).status, 204);
    assert.equal(session.phase, SESSION_PHASES.authenticated);
    assert.equal(session.version, originalVersion);
    assert.equal(session.revokedAt, null);

    harness.failureA.unavailable = true;
    await expectAuthorityError(
        harness.authorityA.logout(login.issuance.identifier),
        ERROR_CLASSES.unavailable,
        'session-store-unavailable',
    );
    assert.equal((await harness.authorityA.logout(undefined)).status, 204);
    assert.equal((await harness.authorityA.logout(syntheticPrivateValue('malformed'))).status, 204);
    assert.equal(session.phase, SESSION_PHASES.authenticated);
    assert.equal(session.version, originalVersion);
    assert.equal(session.revokedAt, null);
});

test('revoke-all invalidates every prior same-subject device but no other subject', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const otherAccount = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account, otherAccount] });
    const firstDevice = await targetLogin(harness, account);
    const secondDevice = await targetLogin(harness, account, { authority: harness.authorityB });
    const otherSubject = await targetLogin(harness, otherAccount);

    assert.equal((await harness.authorityA.revokeAll(firstDevice.issuance.identifier)).status, 204);
    for (const identifier of [firstDevice.issuance.identifier, secondDevice.issuance.identifier]) {
        await expectAuthorityError(
            harness.authorityB.current(identifier),
            ERROR_CLASSES.invalid,
            SESSION_PHASES.revoked,
        );
    }
    assert.equal(phaseOf(await harness.authorityA.current(otherSubject.issuance.identifier)), SESSION_PHASES.authenticated);
    assert.equal(
        [...harness.backing.state.sessions.values()].filter((session) => (
            session.revocationReason === 'user-revoke-all'
        )).length,
        2,
    );

    const afterEpoch = await targetLogin(harness, account);
    assert.equal(phaseOf(await harness.authorityB.current(afterEpoch.issuance.identifier)), SESSION_PHASES.authenticated);
});

test('Face success race has one promotion winner and one stale conflict', async () => {
    const account = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Sim' });
    const harness = await createHarness({ accounts: [account] });
    const login = await targetLogin(harness, account);
    const challenge = await harness.authorityA.createExistingPhotoChallenge(login.issuance.identifier);
    harness.faceState.result = { livenessDecision: 'realface', matchDecision: true };

    const outcomes = await Promise.allSettled([
        harness.authorityA.completeFace(challenge.issuance.identifier),
        harness.authorityB.completeFace(challenge.issuance.identifier),
    ]);
    const winners = outcomes.filter(({ status }) => status === 'fulfilled');
    const losers = outcomes.filter(({ status }) => status === 'rejected');
    assert.equal(winners.length, 1);
    assert.equal(losers.length, 1);
    assert.equal(
        isAuthorityError(losers[0].reason, ERROR_CLASSES.conflict, 'session-compare-and-replace'),
        true,
    );
    assert.equal(phaseOf(winners[0].value), SESSION_PHASES.authenticated);
    assert.equal(
        phaseOf(await harness.authorityB.current(winners[0].value.issuance.identifier)),
        SESSION_PHASES.authenticated,
    );
    await expectAuthorityError(
        harness.authorityA.current(challenge.issuance.identifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.rotatedOut,
    );
});

test('ambiguous registration work remains provisional and reconciliation-blocked', async () => {
    const account = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Não' });
    const harness = await createHarness({ accounts: [account] });
    const login = await targetLogin(harness, account);
    const enrollment = await harness.authorityA.registrationEnrollment(login.issuance.identifier);
    harness.accountState.uploadUnavailable = true;

    await expectAuthorityError(
        harness.authorityA.createRegistrationChallenge(
            enrollment.issuance.identifier,
            randomBytes(40),
        ),
        ERROR_CLASSES.unavailable,
        'registration-reconciliation-required',
    );
    assert.equal(
        phaseOf(await harness.authorityB.current(enrollment.issuance.identifier)),
        SESSION_PHASES.registrationPending,
    );
    assert.equal(
        [...harness.backing.state.flows.values()][0].registrationState,
        'reconciliation-required',
    );
    harness.accountState.uploadUnavailable = false;
    await expectAuthorityError(
        harness.authorityB.createRegistrationChallenge(
            enrollment.issuance.identifier,
            randomBytes(40),
        ),
        ERROR_CLASSES.conflict,
        'face-challenge-active',
    );
    assert.equal(harness.faceState.createCalls, 0);
});

test('store and eligibility outages fail closed without destroying retained authority', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const login = await targetLogin(harness, account);

    await expectAuthorityError(
        targetLogin(harness, account, { password: syntheticPrivateValue('wrong-credential') }),
        ERROR_CLASSES.invalid,
        'invalid-credentials',
    );
    assert.equal(phaseOf(await harness.authorityA.current(login.issuance.identifier)), SESSION_PHASES.authenticated);

    harness.failureA.unavailable = true;
    await expectAuthorityError(
        harness.authorityA.current(login.issuance.identifier),
        ERROR_CLASSES.unavailable,
        'session-store-unavailable',
    );
    await expectAuthorityError(
        harness.authorityA.logout(generatedIdentifier()),
        ERROR_CLASSES.unavailable,
        'session-store-unavailable',
    );
    assert.equal((await harness.authorityA.logout(undefined)).status, 204);
    harness.failureA.unavailable = false;
    assert.equal(phaseOf(await harness.authorityB.current(login.issuance.identifier)), SESSION_PHASES.authenticated);

    harness.advance(ELIGIBILITY_REVALIDATION_MS);
    harness.accountState.readUnavailable = true;
    await expectAuthorityError(
        harness.authorityB.current(login.issuance.identifier),
        ERROR_CLASSES.unavailable,
        'eligibility-source-unavailable',
    );
    harness.accountState.readUnavailable = false;
    const recovered = await harness.authorityA.current(login.issuance.identifier);
    assert.equal(phaseOf(recovered), SESSION_PHASES.authenticated);
    assert.equal(Date.parse(recovered.body.expiresAt), login.issuance.expiresAt.getTime());
});

test('an unknown rotation outcome never returns a replacement identifier', async () => {
    const account = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Não' });
    const harness = await createHarness({ accounts: [account] });
    const login = await targetLogin(harness, account);
    const uncertainStore = {
        ...harness.storeB,
        async rotateSession(input) {
            await harness.storeB.rotateSession(input);
            throw authorityUnavailable('transaction-outcome-unknown');
        },
    };
    const uncertainAuthority = harness.createAuthority(uncertainStore);

    await expectAuthorityError(
        uncertainAuthority.registrationEnrollment(login.issuance.identifier),
        ERROR_CLASSES.unavailable,
        'transaction-outcome-unknown',
    );
    await expectAuthorityError(
        harness.authorityA.current(login.issuance.identifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.rotatedOut,
    );
    assert.equal(activeSessions(harness.backing).length, 1);
    assert.equal(activeSessions(harness.backing)[0].phase, SESSION_PHASES.registrationPending);
});

test('an unknown committed Face reservation is quarantined before provider work', async () => {
    const account = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Sim' });
    const harness = await createHarness({ accounts: [account] });
    const login = await targetLogin(harness, account);
    const uncertainStore = {
        ...harness.storeB,
        async reserveFaceFlow(input) {
            await harness.storeB.reserveFaceFlow(input);
            throw authorityUnavailable('transaction-outcome-unknown');
        },
    };
    const uncertainAuthority = harness.createAuthority(uncertainStore);

    await expectAuthorityError(
        uncertainAuthority.createExistingPhotoChallenge(login.issuance.identifier),
        ERROR_CLASSES.unavailable,
        'transaction-outcome-unknown',
    );
    assert.equal(harness.faceState.createCalls, 0);
    const [flow] = [...harness.backing.state.flows.values()];
    assert.ok(flow);
    assert.equal(flow.challengeState, 'reconciliation-required');
    assert.equal(activeSessions(harness.backing).length, 1);
    assert.equal(activeSessions(harness.backing)[0].phase, SESSION_PHASES.credentialVerified);
    await expectAuthorityError(
        harness.authorityA.createExistingPhotoChallenge(login.issuance.identifier),
        ERROR_CLASSES.conflict,
        'face-challenge-active',
    );
    assert.equal(harness.faceState.createCalls, 0);
});

test('an unknown rolled-back Face reservation fabricates no flow and permits a fresh attempt', async () => {
    const account = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Sim' });
    const harness = await createHarness({ accounts: [account] });
    const login = await targetLogin(harness, account);
    const uncertainStore = {
        ...harness.storeB,
        async reserveFaceFlow() {
            throw authorityUnavailable('transaction-outcome-unknown');
        },
    };
    const uncertainAuthority = harness.createAuthority(uncertainStore);

    await expectAuthorityError(
        uncertainAuthority.createExistingPhotoChallenge(login.issuance.identifier),
        ERROR_CLASSES.unavailable,
        'face-flow-reconciliation-unavailable',
    );
    assert.equal(harness.faceState.createCalls, 0);
    assert.equal(harness.backing.state.flows.size, 0);
    assert.equal(activeSessions(harness.backing).length, 1);
    assert.equal(activeSessions(harness.backing)[0].phase, SESSION_PHASES.credentialVerified);

    const retry = await harness.authorityA.createExistingPhotoChallenge(
        login.issuance.identifier,
    );
    assert.equal(retry.status, 200);
    assert.equal(harness.faceState.createCalls, 1);
    assert.equal(phaseOf(await harness.authorityB.current(retry.issuance.identifier)), SESSION_PHASES.facePending);
});

test('an unknown Face-bind response reconciles the committed flow by immutable identity', async () => {
    const account = createAccount({ facePolicy: 'Ativo', photoRegistrationStatus: 'Sim' });
    const harness = await createHarness({ accounts: [account] });
    const login = await targetLogin(harness, account);
    const uncertainStore = {
        ...harness.storeB,
        async bindFaceChallengeAndRotate(input) {
            await harness.storeB.bindFaceChallengeAndRotate(input);
            throw authorityUnavailable('transaction-outcome-unknown');
        },
    };
    const uncertainAuthority = harness.createAuthority(uncertainStore);

    await expectAuthorityError(
        uncertainAuthority.createExistingPhotoChallenge(login.issuance.identifier),
        ERROR_CLASSES.unavailable,
        'transaction-outcome-unknown',
    );

    const [flow] = [...harness.backing.state.flows.values()];
    assert.ok(flow);
    assert.equal(flow.challengeState, 'reconciliation-required');
    assert.equal(flow.registrationState, 'registered');
    assert.equal(
        activeSessions(harness.backing).some((session) => (
            session.sessionId === flow.currentSessionId
            && session.phase === SESSION_PHASES.facePending
        )),
        true,
    );
    await expectAuthorityError(
        harness.authorityA.current(login.issuance.identifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.rotatedOut,
    );
});

test('eligibility loss revokes every same-subject session at the revalidation bound', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const first = await targetLogin(harness, account);
    const second = await targetLogin(harness, account, { authority: harness.authorityB });
    harness.advance(ELIGIBILITY_REVALIDATION_MS);
    rowFor(harness, account)[7] = 'Inativo';

    await expectAuthorityError(
        harness.authorityA.current(first.issuance.identifier),
        ERROR_CLASSES.forbidden,
        'ineligible',
    );
    await expectAuthorityError(
        harness.authorityB.current(second.issuance.identifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.revoked,
    );
    assert.equal(
        [...harness.backing.state.sessions.values()].filter((session) => (
            session.revocationReason === 'account-inactive'
        )).length,
        2,
    );
});

test('SQL-time eligibility loss between observation and commit fails the request', async () => {
    const account = createAccount({
        accessDateSerial: excelSerial(2042, 6, 1),
        facePolicy: 'Inativo',
    });
    const harness = await createHarness({
        accounts: [account],
        startTime: Date.parse('2042-06-02T02:54:00.000Z'),
    });
    const login = await targetLogin(harness, account);
    harness.advance(ELIGIBILITY_REVALIDATION_MS);
    const advancingStore = {
        ...harness.storeB,
        async updateEligibility(input) {
            harness.advance(2 * 60 * 1000);
            return harness.storeB.updateEligibility(input);
        },
    };
    const advancingAuthority = harness.createAuthority(advancingStore);

    await expectAuthorityError(
        advancingAuthority.current(login.issuance.identifier),
        ERROR_CLASSES.forbidden,
        'ineligible',
    );
    await expectAuthorityError(
        harness.authorityA.current(login.issuance.identifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.revoked,
    );
});

test('credential fingerprint changes revoke prior sessions while preserving stable subject identity', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const first = await targetLogin(harness, account);
    const second = await targetLogin(harness, account, { authority: harness.authorityB });
    const before = await harness.authorityA.authorizeProtected(first.issuance.identifier);
    const replacementCredential = syntheticPrivateValue('replacement-credential');
    rowFor(harness, account)[3] = replacementCredential;
    harness.advance(ELIGIBILITY_REVALIDATION_MS);

    await expectAuthorityError(
        harness.authorityA.current(first.issuance.identifier),
        ERROR_CLASSES.invalid,
        'credential-changed',
    );
    await expectAuthorityError(
        harness.authorityB.current(second.issuance.identifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.revoked,
    );
    const fresh = await targetLogin(harness, account, { password: replacementCredential });
    const after = await harness.authorityB.authorizeProtected(fresh.issuance.identifier);
    assert.equal(after.subjectId === before.subjectId, true);
    assert.equal(
        [...harness.backing.state.sessions.values()].filter((session) => (
            session.revocationReason === 'credential-reset'
        )).length,
        2,
    );
});

test('fresh credentials detect a changed fingerprint before issuing replacement authority', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const first = await targetLogin(harness, account);
    const second = await targetLogin(harness, account, { authority: harness.authorityB });
    const before = await harness.authorityA.authorizeProtected(first.issuance.identifier);
    const replacementCredential = syntheticPrivateValue('fresh-replacement-credential');
    rowFor(harness, account)[3] = replacementCredential;

    const replacement = await targetLogin(harness, account, { password: replacementCredential });
    const after = await harness.authorityB.authorizeProtected(replacement.issuance.identifier);
    assert.equal(after.subjectId === before.subjectId, true);
    for (const identifier of [first.issuance.identifier, second.issuance.identifier]) {
        await expectAuthorityError(
            harness.authorityA.current(identifier),
            ERROR_CLASSES.invalid,
            SESSION_PHASES.revoked,
        );
    }
    assert.equal(
        [...harness.backing.state.sessions.values()].filter((session) => (
            session.revocationReason === 'credential-reset'
        )).length,
        2,
    );
});

test('a stale credential observation cannot overwrite a newer credential reset', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const original = await targetLogin(harness, account);
    const staleReadEntered = deferred();
    const releaseStaleRead = deferred();
    const delayedCallNumber = harness.accountState.readCalls + 1;
    harness.accountState.readRowsHook = async ({ callNumber }) => {
        if (callNumber !== delayedCallNumber) return;
        staleReadEntered.resolve();
        await releaseStaleRead.promise;
    };

    const staleLogin = targetLogin(harness, account, { authority: harness.authorityB });
    await staleReadEntered.promise;
    const replacementCredential = syntheticPrivateValue('newer-credential');
    rowFor(harness, account)[3] = replacementCredential;
    harness.advance(1);
    const replacement = await targetLogin(harness, account, { password: replacementCredential });

    releaseStaleRead.resolve();
    await expectAuthorityErrorClass(staleLogin, ERROR_CLASSES.conflict);
    harness.accountState.readRowsHook = null;
    await expectAuthorityError(
        harness.authorityA.current(original.issuance.identifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.revoked,
    );
    assert.equal(
        phaseOf(await harness.authorityB.current(replacement.issuance.identifier)),
        SESSION_PHASES.authenticated,
    );
    assert.equal(activeSessions(harness.backing).length, 1);
});

test('a stale eligible observation cannot overwrite a newer ineligibility revocation', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const original = await targetLogin(harness, account);
    const staleReadEntered = deferred();
    const releaseStaleRead = deferred();
    const delayedCallNumber = harness.accountState.readCalls + 1;
    harness.accountState.readRowsHook = async ({ callNumber }) => {
        if (callNumber !== delayedCallNumber) return;
        staleReadEntered.resolve();
        await releaseStaleRead.promise;
    };

    const staleLogin = targetLogin(harness, account, { authority: harness.authorityB });
    await staleReadEntered.promise;
    rowFor(harness, account)[6] = excelSerial(2042, 5, 31);
    harness.advance(1);
    await expectAuthorityErrorClass(
        targetLogin(harness, account),
        ERROR_CLASSES.forbidden,
    );

    releaseStaleRead.resolve();
    await expectAuthorityErrorClass(staleLogin, ERROR_CLASSES.conflict);
    harness.accountState.readRowsHook = null;
    await expectAuthorityError(
        harness.authorityA.current(original.issuance.identifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.revoked,
    );
    assert.equal(activeSessions(harness.backing).length, 0);
});

test('a same-generation incident blocks a credential login paused after its control read', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const accountReadEntered = deferred();
    const releaseAccountRead = deferred();
    const delayedCallNumber = harness.accountState.readCalls + 1;
    harness.accountState.readRowsHook = async ({ callNumber }) => {
        if (callNumber !== delayedCallNumber) return;
        accountReadEntered.resolve();
        await releaseAccountRead.promise;
    };

    const pausedLogin = targetLogin(harness, account, { authority: harness.authorityB });
    await accountReadEntered.promise;
    const beforeIncident = await harness.storeA.readControl();
    const incident = await harness.storeA.transitionControl({
        expectedVersion: beforeIncident.control.version,
        changes: {
            incidentState: 'suspended',
            incidentRecordedAt: harness.readNow(),
            incidentCode: 'synthetic-login-interleaving',
        },
    });
    assert.equal(
        incident.control.authorityGeneration,
        beforeIncident.control.authorityGeneration,
    );

    releaseAccountRead.resolve();
    await expectAuthorityError(
        pausedLogin,
        ERROR_CLASSES.unavailable,
        'authority-incident',
    );
    harness.accountState.readRowsHook = null;
    assert.equal(harness.backing.state.subjects.size, 0);
    assert.equal(harness.backing.state.sessions.size, 0);
    assert.equal(harness.backing.state.legacyBindings.size, 0);
});

test('stable subject mapping follows exact account identity across workbook row movement', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const decoy = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const login = await targetLogin(harness, account);
    const before = await harness.authorityA.authorizeProtected(login.issuance.identifier);
    harness.accountState.rows = [[...decoy.row], [...account.row]];

    const after = await harness.authorityB.authorizeProtected(login.issuance.identifier);
    assert.equal(before.platformRowIndex, 0);
    assert.equal(after.platformRowIndex, 1);
    assert.equal(after.subjectId === before.subjectId, true);
    assert.equal(harness.backing.state.subjects.size, 1);
});

test('encrypted account mappings must reproduce their exact keyed lookup descriptor', async () => {
    const firstAccount = createAccount({ facePolicy: 'Inativo' });
    const secondAccount = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [firstAccount, secondAccount] });
    const firstLogin = await targetLogin(harness, firstAccount);
    await targetLogin(harness, secondAccount);
    const subjects = [...harness.backing.state.subjects.values()];
    assert.equal(subjects.length, 2);
    const firstCiphertext = Buffer.from(subjects[0].encryptedAccountMapping);
    subjects[0].encryptedAccountMapping = Buffer.from(subjects[1].encryptedAccountMapping);
    subjects[1].encryptedAccountMapping = firstCiphertext;
    const readsBefore = harness.accountState.readCalls;
    harness.advance(ELIGIBILITY_REVALIDATION_MS);

    await expectAuthorityError(
        harness.authorityA.current(firstLogin.issuance.identifier),
        ERROR_CLASSES.unavailable,
        'subject-mapping-integrity',
    );
    assert.equal(harness.accountState.readCalls, readsBefore);

    const firstSubject = subjects.find((subject) => (
        safeBuffersEqual(subject.loginLookupToken, createLoginLookup(
            firstAccount.login,
            harness.keys.loginLookup,
        ).token)
    ));
    const epochBefore = firstSubject.sessionEpoch;
    rowFor(harness, firstAccount)[7] = 'Inativo';
    await expectAuthorityError(
        targetLogin(harness, firstAccount),
        ERROR_CLASSES.unavailable,
        'subject-mapping-integrity',
    );
    assert.equal(firstSubject.sessionEpoch, epochBefore);
});

test('exact-login remap preserves subject, sessions, row movement, and the legacy cutoff', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const decoy = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const legacy = await harness.authorityA.loginLegacyWithSeeding({
        login: account.login,
        password: account.password,
    });
    const target = await targetLogin(harness, account);
    const before = await harness.authorityA.authorizeProtected(target.issuance.identifier);
    const oldExactLogin = account.login;
    const newExactLogin = syntheticPrivateValue('remapped-login');
    const remappedRow = [...rowFor(harness, account)];
    remappedRow[2] = newExactLogin;
    harness.accountState.rows = [[...decoy.row], remappedRow];

    const remapped = await harness.authorityA.remapSubjectLogin({
        subjectId: before.subjectId,
        expectedExactLogin: oldExactLogin,
        newExactLogin,
    });
    assert.deepEqual(Object.keys(remapped).sort(), ['idempotent', 'serverTime']);
    assert.equal(remapped.idempotent, false);
    assert.equal(remapped.serverTime instanceof Date, true);
    assert.equal(remapped.serverTime.getTime(), harness.readNow().getTime());

    const oldLookup = await readSubjectForExactLogin(harness, oldExactLogin);
    const newLookup = await readSubjectForExactLogin(harness, newExactLogin, harness.storeB);
    assert.equal(oldLookup.subject, null);
    assert.equal(Boolean(newLookup.subject), true);
    assert.equal(newLookup.subject.subjectId === before.subjectId, true);
    const existing = await harness.authorityB.authorizeProtected(target.issuance.identifier);
    assert.equal(existing.subjectId === before.subjectId, true);
    assert.equal(existing.platformRowIndex, 1);

    await expectAuthorityError(
        targetLogin(harness, account),
        ERROR_CLASSES.invalid,
        'invalid-credentials',
    );
    const remappedAccount = {
        ...account,
        login: newExactLogin,
        row: remappedRow,
    };
    const fresh = await targetLogin(harness, remappedAccount, { authority: harness.authorityB });
    const freshAuthority = await harness.authorityA.authorizeProtected(fresh.issuance.identifier);
    assert.equal(freshAuthority.subjectId === before.subjectId, true);

    await expectAuthorityErrorClass(
        harness.authorityA.authorizeLegacy(legacy.body.IndexVerificado),
        ERROR_CLASSES.invalid,
    );
    await expectAuthorityError(
        harness.authorityB.loginLegacyWithSeeding({
            login: newExactLogin,
            password: account.password,
        }),
        ERROR_CLASSES.conflict,
        'target-authority-established',
    );
    assert.equal(harness.backing.state.subjects.size, 1);
});

test('exact remap retry is idempotent despite a fresh encrypted candidate', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const target = await targetLogin(harness, account);
    const authority = await harness.authorityA.authorizeProtected(target.issuance.identifier);
    const nextLogin = syntheticPrivateValue('idempotent-remapped-login');
    const row = [...rowFor(harness, account)];
    row[2] = nextLogin;
    harness.accountState.rows = [row];

    const first = await harness.authorityA.remapSubjectLogin({
        subjectId: authority.subjectId,
        expectedExactLogin: account.login,
        newExactLogin: nextLogin,
    });
    const storedAfterFirst = Buffer.from(
        harness.backing.state.subjects.get(authority.subjectId).encryptedAccountMapping,
    );
    const retry = await harness.authorityB.remapSubjectLogin({
        subjectId: authority.subjectId,
        expectedExactLogin: account.login,
        newExactLogin: nextLogin,
    });
    const storedAfterRetry = harness.backing.state.subjects
        .get(authority.subjectId).encryptedAccountMapping;

    assert.equal(first.idempotent, false);
    assert.equal(retry.idempotent, true);
    assert.deepEqual(Object.keys(retry).sort(), ['idempotent', 'serverTime']);
    assert.equal(storedAfterFirst.equals(storedAfterRetry), true);
    assert.equal(
        (await harness.authorityA.authorizeProtected(target.issuance.identifier)).platformRowIndex,
        0,
    );
});

test('remap rejects a login already owned by another stable subject', async () => {
    const firstAccount = createAccount({ facePolicy: 'Inativo' });
    const secondAccount = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [firstAccount, secondAccount] });
    const first = await targetLogin(harness, firstAccount);
    const second = await targetLogin(harness, secondAccount, { authority: harness.authorityB });
    const firstAuthority = await harness.authorityA.authorizeProtected(first.issuance.identifier);
    const secondAuthority = await harness.authorityB.authorizeProtected(second.issuance.identifier);

    await expectAuthorityError(
        harness.authorityA.remapSubjectLogin({
            subjectId: firstAuthority.subjectId,
            expectedExactLogin: firstAccount.login,
            newExactLogin: secondAccount.login,
        }),
        ERROR_CLASSES.unavailable,
        'subject-mapping-conflict',
    );
    const firstLookup = await readSubjectForExactLogin(harness, firstAccount.login);
    const secondLookup = await readSubjectForExactLogin(harness, secondAccount.login, harness.storeB);
    assert.equal(firstLookup.subject.subjectId === firstAuthority.subjectId, true);
    assert.equal(secondLookup.subject.subjectId === secondAuthority.subjectId, true);
    assert.equal(
        (await harness.authorityA.authorizeProtected(first.issuance.identifier)).platformRowIndex,
        0,
    );
    assert.equal(
        (await harness.authorityB.authorizeProtected(second.issuance.identifier)).platformRowIndex,
        1,
    );
});

test('concurrent divergent exact-login remaps permit one durable winner', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const target = await targetLogin(harness, account);
    const before = await harness.authorityA.authorizeProtected(target.issuance.identifier);
    const candidates = [
        syntheticPrivateValue('concurrent-remap-a'),
        syntheticPrivateValue('concurrent-remap-b'),
    ];

    const outcomes = await Promise.allSettled(candidates.map((newExactLogin, index) => (
        (index === 0 ? harness.authorityA : harness.authorityB).remapSubjectLogin({
            subjectId: before.subjectId,
            expectedExactLogin: account.login,
            newExactLogin,
        })
    )));
    const winnerIndex = outcomes.findIndex(({ status }) => status === 'fulfilled');
    const loserIndex = outcomes.findIndex(({ status }) => status === 'rejected');
    assert.equal(winnerIndex >= 0, true);
    assert.equal(loserIndex >= 0, true);
    assert.equal(
        isAuthorityError(
            outcomes[loserIndex].reason,
            ERROR_CLASSES.unavailable,
            'subject-mapping-conflict',
        ),
        true,
    );
    assert.deepEqual(
        Object.keys(outcomes[winnerIndex].value).sort(),
        ['idempotent', 'serverTime'],
    );
    assert.equal(outcomes[winnerIndex].value.idempotent, false);

    const oldLookup = await readSubjectForExactLogin(harness, account.login);
    const winnerLookup = await readSubjectForExactLogin(harness, candidates[winnerIndex]);
    const loserLookup = await readSubjectForExactLogin(
        harness,
        candidates[loserIndex],
        harness.storeB,
    );
    assert.equal(oldLookup.subject, null);
    assert.equal(winnerLookup.subject.subjectId === before.subjectId, true);
    assert.equal(loserLookup.subject, null);

    const movedRow = [...account.row];
    movedRow[2] = candidates[winnerIndex];
    harness.accountState.rows = [movedRow];
    const existing = await harness.authorityB.authorizeProtected(target.issuance.identifier);
    assert.equal(existing.subjectId === before.subjectId, true);
});

test('missing and ambiguous exact account mappings fail closed', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const login = await targetLogin(harness, account);
    const originalRow = [...rowFor(harness, account)];
    harness.accountState.rows = [originalRow, [...originalRow]];

    await expectAuthorityErrorClass(
        harness.authorityA.authorizeProtected(login.issuance.identifier),
        ERROR_CLASSES.unavailable,
    );
    harness.advance(ELIGIBILITY_REVALIDATION_MS);
    await expectAuthorityErrorClass(
        harness.authorityB.current(login.issuance.identifier),
        ERROR_CLASSES.unavailable,
    );

    harness.accountState.rows = [];
    await expectAuthorityErrorClass(
        harness.authorityA.current(login.issuance.identifier),
        ERROR_CLASSES.unavailable,
    );
});

test('administrator and identifier-leak subject revocation end every active device', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const first = await targetLogin(harness, account);
    const second = await targetLogin(harness, account, { authority: harness.authorityB });
    const subject = await harness.authorityA.authorizeProtected(first.issuance.identifier);

    await harness.authorityA.revokeSubject(subject.subjectId, 'administrator');
    for (const identifier of [first.issuance.identifier, second.issuance.identifier]) {
        await expectAuthorityError(
            harness.authorityB.current(identifier),
            ERROR_CLASSES.invalid,
            SESSION_PHASES.revoked,
        );
    }
    assert.equal(
        [...harness.backing.state.sessions.values()].filter((session) => (
            session.revocationReason === 'administrator'
        )).length,
        2,
    );

    const recovered = await targetLogin(harness, account);
    await harness.authorityB.revokeSubject(subject.subjectId, 'identifier-leak');
    await expectAuthorityError(
        harness.authorityA.current(recovered.issuance.identifier),
        ERROR_CLASSES.invalid,
        SESSION_PHASES.revoked,
    );
    await assert.rejects(
        harness.authorityA.revokeSubject('', 'administrator'),
        /subject identifier is required/u,
    );
    await assert.rejects(
        harness.authorityA.revokeSubject(subject.subjectId, 'unsupported-reason'),
        /revocation reason is invalid/u,
    );
});

test('legacy ledger seeding is verifier-bound, deterministic, and stable across row movement', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const decoy = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    assert.equal(
        harness.legacyHandleAuthority.defaultNow() === harness.readNow().getTime(),
        false,
    );
    const first = await harness.authorityA.loginLegacyWithSeeding({
        login: account.login,
        password: account.password,
    });
    const repeated = await harness.authorityB.loginLegacyWithSeeding({
        login: account.login,
        password: account.password,
    });

    assert.deepEqual(
        harness.legacyHandleAuthority.createNowValues,
        [harness.readNow().getTime(), harness.readNow().getTime()],
    );
    assert.deepEqual(
        harness.legacyHandleAuthority.inspectNowValues,
        [harness.readNow().getTime(), harness.readNow().getTime()],
    );
    assert.equal(first.status, 200);
    assert.equal(typeof first.body.IndexVerificado, 'string');
    assert.equal(first.body.IndexVerificado.length > 0, true);
    assert.equal(first.body.IndexVerificado === repeated.body.IndexVerificado, true);
    assert.equal(harness.backing.state.legacyBindings.size, 1);
    assert.equal(harness.backing.state.subjects.size, 1);

    const before = await harness.authorityA.authorizeLegacy(first.body.IndexVerificado);
    assert.equal(
        harness.legacyHandleAuthority.inspectNowValues.at(-1),
        harness.readNow().getTime(),
    );
    harness.accountState.rows = [[...decoy.row], [...account.row]];
    const after = await harness.authorityB.authorizeLegacy(first.body.IndexVerificado);
    assert.equal(before.platformRowIndex, 0);
    assert.equal(after.platformRowIndex, 1);
    assert.equal(after.subjectId === before.subjectId, true);

    harness.advance(LEGACY_LIFETIME_MS);
    await expectAuthorityError(
        harness.authorityA.authorizeLegacy(first.body.IndexVerificado),
        ERROR_CLASSES.invalid,
        'legacy-binding-terminal',
    );
});

test('a legacy bind paused after credential resolution loses to a credential reset', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    await harness.authorityA.loginLegacyWithSeeding({
        login: account.login,
        password: account.password,
    });
    harness.advance(1_000);

    const bindEntered = deferred();
    const releaseBind = deferred();
    const delayedStore = {
        ...harness.storeB,
        async bindLegacy(input) {
            bindEntered.resolve();
            await releaseBind.promise;
            return harness.storeB.bindLegacy(input);
        },
    };
    const delayedAuthority = harness.createAuthority(delayedStore);
    const staleLogin = delayedAuthority.loginLegacyWithSeeding({
        login: account.login,
        password: account.password,
    });
    await bindEntered.promise;

    const replacementCredential = syntheticPrivateValue('replacement-legacy-credential');
    rowFor(harness, account)[3] = replacementCredential;
    harness.advance(1_000);
    const replacement = await harness.authorityA.loginLegacyWithSeeding({
        login: account.login,
        password: replacementCredential,
    });
    assert.equal(replacement.status, 200);
    const bindingsAfterReplacement = harness.backing.state.legacyBindings.size;

    releaseBind.resolve();
    await expectAuthorityError(
        staleLogin,
        ERROR_CLASSES.conflict,
        'subject-credential-compare-and-replace',
    );
    assert.equal(harness.backing.state.legacyBindings.size, bindingsAfterReplacement);
    const bindings = [...harness.backing.state.legacyBindings.values()];
    assert.equal(bindings.filter((binding) => binding.revokedAt === null).length, 1);
    assert.equal(
        bindings.filter((binding) => binding.revocationReason === 'credential-reset').length,
        1,
    );
});

test('legacy login preserves a distinct workbook-read failure for its compatibility envelope', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    harness.accountState.readUnavailable = true;

    await expectAuthorityError(
        harness.authorityA.loginLegacyWithSeeding({
            login: account.login,
            password: account.password,
        }),
        ERROR_CLASSES.unavailable,
        'legacy-platform-data-read-failed',
    );
    assert.equal(harness.backing.state.subjects.size, 0);
    assert.equal(harness.backing.state.legacyBindings.size, 0);
});

test('pre-enforcement legacy admission preserves bound seeding and unbound compatibility', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const seedingHarness = await createHarness({
        accounts: [account],
        qualify: false,
        runtimeControls: {
            targetRoutesEnabled: false,
            targetSessionIssuanceEnabled: false,
            legacyLedgerSeedingEnabled: true,
            legacyCompatibilityEnforcementEnabled: false,
            subjectTargetAdoptionEnabled: false,
            protectedRoutesEnabled: false,
        },
    });
    const seedingStartedAt = seedingHarness.readNow();
    await seedingHarness.storeA.transitionControl({
        expectedVersion: 1,
        changes: {
            legacyLedgerSeedingEnabled: true,
            legacyLedgerSeedingStartedAt: seedingStartedAt,
            seedingStartedAt,
        },
    });
    await seedingHarness.storeA.heartbeatLegacySeedingContinuity({ ownerId: randomUUID() });
    const seeded = await seedingHarness.authorityA.loginLegacyWithSeeding({
        login: account.login,
        password: account.password,
    });
    assert.equal(
        (await seedingHarness.authorityB.authorizeLegacy(seeded.body.IndexVerificado)).platformRowIndex,
        0,
    );

    const fallbackHarness = await createHarness({
        accounts: [account],
        qualify: false,
        runtimeControls: {
            targetRoutesEnabled: false,
            targetSessionIssuanceEnabled: false,
            legacyLedgerSeedingEnabled: false,
            legacyCompatibilityEnforcementEnabled: false,
            subjectTargetAdoptionEnabled: false,
            protectedRoutesEnabled: false,
        },
    });
    const preLedgerHandle = fallbackHarness.legacyHandleAuthority.createHandle(0);
    assert.equal(
        (await fallbackHarness.authorityA.authorizeLegacy(preLedgerHandle)).platformRowIndex,
        0,
    );

    const mismatchedHarness = await createHarness({
        accounts: [account],
        qualify: false,
        runtimeControls: {
            targetRoutesEnabled: false,
            targetSessionIssuanceEnabled: false,
            legacyLedgerSeedingEnabled: false,
            legacyCompatibilityEnforcementEnabled: true,
            subjectTargetAdoptionEnabled: false,
            protectedRoutesEnabled: false,
        },
    });
    await expectAuthorityError(
        mismatchedHarness.authorityA.authorizeLegacy(
            mismatchedHarness.legacyHandleAuthority.createHandle(0),
        ),
        ERROR_CLASSES.unavailable,
        'legacy-enforcement-gate-mismatch',
    );
});

test('pre-enforcement seeding preserves expired legacy login until enforcement', async () => {
    const account = createAccount({
        accessDateSerial: excelSerial(2042, 5, 31),
        accountStatus: 'Ativo',
        facePolicy: 'Inativo',
    });
    const harness = await createHarness({
        accounts: [account],
        qualify: false,
        runtimeControls: {
            targetRoutesEnabled: false,
            targetSessionIssuanceEnabled: false,
            legacyLedgerSeedingEnabled: true,
            legacyCompatibilityEnforcementEnabled: false,
            subjectTargetAdoptionEnabled: false,
            protectedRoutesEnabled: false,
        },
    });
    const started = await harness.storeA.transitionControl({
        expectedVersion: 1,
        changes: { legacyLedgerSeedingEnabled: true },
    });
    const ownerId = randomUUID();
    await harness.storeA.heartbeatLegacySeedingContinuity({ ownerId });

    const legacy = await harness.authorityA.loginLegacyWithSeeding({
        login: account.login,
        password: account.password,
    });
    assert.equal(legacy.status, 200);
    assert.equal(typeof legacy.body.IndexVerificado, 'string');
    assert.equal(harness.backing.state.legacyBindings.size, 1);
    assert.equal(
        (await harness.authorityA.authorizeLegacy(legacy.body.IndexVerificado)).platformRowIndex,
        0,
    );

    await advanceWithContinuity(harness, harness.storeA, ownerId, LEGACY_LIFETIME_MS);
    const qualifiedAt = harness.readNow();
    await harness.storeA.transitionControl({
        expectedVersion: started.control.version,
        changes: {
            seedingQualifiedAt: qualifiedAt,
            legacyCompatibilityEnforcementEnabled: true,
        },
    });
    const enforcedAuthority = harness.createAuthority(
        harness.storeB,
        createRuntimeControls({
            targetRoutesEnabled: false,
            targetSessionIssuanceEnabled: false,
            legacyLedgerSeedingEnabled: true,
            legacyCompatibilityEnforcementEnabled: true,
            subjectTargetAdoptionEnabled: false,
            protectedRoutesEnabled: false,
        }),
    );
    await expectAuthorityError(
        enforcedAuthority.authorizeLegacy(legacy.body.IndexVerificado),
        ERROR_CLASSES.forbidden,
        'ineligible',
    );
    const [subject] = [...harness.backing.state.subjects.values()];
    const [binding] = [...harness.backing.state.legacyBindings.values()];
    assert.equal(subject.eligibilityState, 'ineligible');
    assert.equal(binding.revocationReason, 'entitlement-expired');
});

test('legacy enforcement rejects missing and conflicting immutable bindings', async () => {
    const firstAccount = createAccount({ facePolicy: 'Inativo' });
    const secondAccount = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [firstAccount, secondAccount] });
    const first = await harness.authorityA.loginLegacyWithSeeding({
        login: firstAccount.login,
        password: firstAccount.password,
    });
    await harness.authorityB.loginLegacyWithSeeding({
        login: secondAccount.login,
        password: secondAccount.password,
    });
    await expectAuthorityError(
        harness.authorityA.authorizeLegacy(generatedIdentifier()),
        ERROR_CLASSES.invalid,
        'invalid-legacy-handle',
    );

    harness.advance(1_000);
    const unboundHandle = harness.legacyHandleAuthority.createHandle(0);
    await expectAuthorityError(
        harness.authorityA.authorizeLegacy(unboundHandle),
        ERROR_CLASSES.invalid,
        'legacy-binding-missing',
    );

    const bindings = [...harness.backing.state.legacyBindings.values()];
    assert.equal(bindings.length, 2);
    assert.equal(bindings[0].subjectId === bindings[1].subjectId, false);
    const conflictingSubject = harness.backing.state.subjects.get(bindings[1].subjectId);
    await expectAuthorityError(
        harness.storeA.bindLegacy({
            legacyCompatibilityId: randomUUID(),
            subjectId: bindings[1].subjectId,
            verifierKeyId: bindings[0].verifierKeyId,
            verifier: bindings[0].verifier,
            issuedAt: bindings[0].issuedAt,
            expiresAt: bindings[0].expiresAt,
            expectedCredentialVersion: conflictingSubject.credentialVersion,
            expectedCredentialFingerprintKeyId: conflictingSubject.credentialFingerprintKeyId,
            expectedCredentialFingerprint: conflictingSubject.credentialFingerprint,
        }),
        ERROR_CLASSES.unavailable,
        'legacy-binding-integrity',
    );
    await expectAuthorityError(
        harness.storeB.authorizeLegacy({
            verifierKeyId: bindings[0].verifierKeyId,
            verifier: bindings[0].verifier,
            issuedAt: new Date(bindings[0].issuedAt.getTime() + 1_000),
            expiresAt: new Date(bindings[0].expiresAt.getTime() + 1_000),
        }),
        ERROR_CLASSES.unavailable,
        'legacy-binding-integrity',
    );
    assert.equal(
        (await harness.authorityB.authorizeLegacy(first.body.IndexVerificado)).platformRowIndex,
        0,
    );
});

test('target adoption irreversibly cuts off legacy authorization and reissuance', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const legacy = await harness.authorityA.loginLegacyWithSeeding({
        login: account.login,
        password: account.password,
    });
    assert.equal((await harness.authorityA.authorizeLegacy(legacy.body.IndexVerificado)).platformRowIndex, 0);

    const target = await targetLogin(harness, account);
    assert.equal(phaseOf(await harness.authorityB.current(target.issuance.identifier)), SESSION_PHASES.authenticated);
    await expectAuthorityError(
        harness.authorityA.authorizeLegacy(legacy.body.IndexVerificado),
        ERROR_CLASSES.invalid,
        'target-authority-established',
    );
    await expectAuthorityError(
        harness.authorityB.loginLegacyWithSeeding({
            login: account.login,
            password: account.password,
        }),
        ERROR_CLASSES.conflict,
        'target-authority-established',
    );
    assert.equal(harness.backing.state.legacyBindings.size, 1);
});

test('global legacy stop controls are irreversible and never fall back to an unchecked handle', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const harness = await createHarness({ accounts: [account] });
    const legacy = await harness.authorityA.loginLegacyWithSeeding({
        login: account.login,
        password: account.password,
    });
    const currentControl = await harness.storeA.readControl();
    const issuanceStopped = await harness.storeA.transitionControl({
        expectedVersion: currentControl.control.version,
        changes: {
            legacyIssuanceEnabled: false,
            legacyStopIssuanceAt: harness.readNow(),
        },
    });

    await expectAuthorityError(
        harness.authorityA.loginLegacyWithSeeding({
            login: account.login,
            password: account.password,
        }),
        ERROR_CLASSES.conflict,
        'legacy-issuance-disabled',
    );
    assert.equal((await harness.authorityB.authorizeLegacy(legacy.body.IndexVerificado)).platformRowIndex, 0);

    harness.advance(LEGACY_LIFETIME_MS);
    const stopped = await harness.storeB.transitionControl({
        expectedVersion: issuanceStopped.control.version,
        changes: {
            legacyAcceptanceEnabled: false,
            legacyAcceptanceDisabledAt: harness.readNow(),
        },
    });
    await expectAuthorityError(
        harness.authorityB.authorizeLegacy(legacy.body.IndexVerificado),
        ERROR_CLASSES.invalid,
        'legacy-acceptance-disabled',
    );
    await expectAuthorityErrorClass(
        harness.storeB.transitionControl({
            expectedVersion: stopped.control.version,
            changes: {
                legacyIssuanceEnabled: true,
                legacyAcceptanceEnabled: true,
            },
        }),
        ERROR_CLASSES.forbidden,
    );
});

test('runtime and database gates both deny dormant target authority', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const runtimeDisabled = await createHarness({
        accounts: [account],
        runtimeControls: { targetSessionIssuanceEnabled: false },
    });
    await expectAuthorityError(
        targetLogin(runtimeDisabled, account),
        ERROR_CLASSES.unavailable,
        'targetSessionIssuanceEnabled-disabled',
    );
    assert.equal(runtimeDisabled.backing.state.sessions.size, 0);

    const databaseDisabled = await createHarness({ accounts: [account], qualify: false });
    const control = await databaseDisabled.storeA.transitionControl({
        expectedVersion: 1,
        changes: { targetRoutesEnabled: true },
    });
    assert.equal(control.control.targetRoutesEnabled, true);
    await expectAuthorityError(
        targetLogin(databaseDisabled, account),
        ERROR_CLASSES.unavailable,
        'targetSessionIssuanceEnabled-not-qualified',
    );
    assert.equal(databaseDisabled.backing.state.sessions.size, 0);
});

test('legacy enforcement requires one uninterrupted four-hour seeding horizon', async () => {
    const harness = await createHarness({ qualify: false });
    const firstOwnerId = randomUUID();
    const initialStart = harness.readNow();
    const started = await harness.storeA.transitionControl({
        expectedVersion: 1,
        changes: {
            legacyLedgerSeedingEnabled: true,
            legacyLedgerSeedingStartedAt: initialStart,
            seedingStartedAt: initialStart,
        },
    });
    await harness.storeA.heartbeatLegacySeedingContinuity({ ownerId: firstOwnerId });
    await advanceWithContinuity(harness, harness.storeA, firstOwnerId, 3 * 60 * 60 * 1000);
    await expectAuthorityErrorClass(
        harness.storeB.transitionControl({
            expectedVersion: started.control.version,
            changes: {
                seedingQualifiedAt: harness.readNow(),
                legacyCompatibilityEnforcementEnabled: true,
                legacyCompatibilityEnforcedAt: harness.readNow(),
            },
        }),
        ERROR_CLASSES.forbidden,
    );

    const stopped = await harness.storeA.transitionControl({
        expectedVersion: started.control.version,
        changes: { legacyLedgerSeedingEnabled: false },
    });
    harness.advance(60 * 1000);
    const restartedAt = harness.readNow();
    const restarted = await harness.storeB.transitionControl({
        expectedVersion: stopped.control.version,
        changes: {
            legacyLedgerSeedingEnabled: true,
            seedingStartedAt: restartedAt,
        },
    });
    const secondOwnerId = randomUUID();
    await harness.storeB.heartbeatLegacySeedingContinuity({ ownerId: secondOwnerId });
    await advanceWithContinuity(
        harness,
        harness.storeB,
        secondOwnerId,
        LEGACY_LIFETIME_MS - 1,
    );
    await expectAuthorityErrorClass(
        harness.storeA.transitionControl({
            expectedVersion: restarted.control.version,
            changes: {
                seedingQualifiedAt: harness.readNow(),
                legacyCompatibilityEnforcementEnabled: true,
                legacyCompatibilityEnforcedAt: harness.readNow(),
            },
        }),
        ERROR_CLASSES.forbidden,
    );

    await advanceWithContinuity(harness, harness.storeB, secondOwnerId, 1);
    const enforced = await harness.storeB.transitionControl({
        expectedVersion: restarted.control.version,
        changes: {
            seedingQualifiedAt: harness.readNow(),
            legacyCompatibilityEnforcementEnabled: true,
            legacyCompatibilityEnforcedAt: harness.readNow(),
        },
    });
    assert.equal(enforced.control.legacyCompatibilityEnforcementEnabled, true);
    assert.equal(
        enforced.control.seedingQualifiedAt.getTime() - enforced.control.seedingStartedAt.getTime(),
        LEGACY_LIFETIME_MS,
    );
    await expectAuthorityErrorClass(
        harness.storeA.transitionControl({
            expectedVersion: enforced.control.version,
            changes: { legacyCompatibilityEnforcementEnabled: false },
        }),
        ERROR_CLASSES.forbidden,
    );
});

test('global and authority epochs invalidate prior records and incident state blocks issuance', async () => {
    const account = createAccount({ facePolicy: 'Inativo' });
    const incidentLoginAccount = createAccount({
        login: syntheticPrivateValue('incident-login'),
        password: syntheticPrivateValue('incident-password'),
        facePolicy: 'Inativo',
    });
    const harness = await createHarness({ accounts: [account, incidentLoginAccount] });
    const original = await targetLogin(harness, account);
    let control = await harness.storeA.readControl();
    control = await harness.storeA.transitionControl({
        expectedVersion: control.control.version,
        changes: { globalSessionEpoch: control.control.globalSessionEpoch + 1 },
    });
    await expectAuthorityError(
        harness.authorityB.current(original.issuance.identifier),
        ERROR_CLASSES.invalid,
        'epoch-mismatch',
    );

    const beforeIncident = await targetLogin(harness, account);
    const generationTwoStore = harness.createStore(2);
    const generationTwoAuthority = harness.createAuthority(generationTwoStore);
    await expectAuthorityError(
        generationTwoAuthority.current(beforeIncident.issuance.identifier),
        ERROR_CLASSES.unavailable,
        'authority-generation-mismatch',
    );
    control = await harness.storeA.transitionControl({
        expectedVersion: control.control.version,
        changes: {
            incidentState: 'suspended',
            incidentRecordedAt: harness.readNow(),
            incidentCode: 'synthetic-recovery',
        },
    });
    await expectAuthorityError(
        harness.authorityA.current(beforeIncident.issuance.identifier),
        ERROR_CLASSES.unavailable,
        'authority-incident',
    );
    await expectAuthorityError(
        targetLogin(harness, incidentLoginAccount),
        ERROR_CLASSES.unavailable,
        'authority-incident',
    );
    assert.equal(harness.backing.state.subjects.size, 1);

    control = await harness.storeA.transitionControl({
        expectedVersion: control.control.version,
        changes: {
            legacyIssuanceEnabled: false,
            legacyStopIssuanceAt: harness.readNow(),
        },
    });
    harness.advance(LEGACY_LIFETIME_MS);
    Object.assign(harness.backing.state.control, {
        authorityGeneration: 2,
        globalSessionEpoch: control.control.globalSessionEpoch + 1,
        incidentState: 'recovering',
    });
    control = await generationTwoStore.transitionControl({
        expectedVersion: control.control.version,
        changes: {
            incidentState: 'normal',
            legacyAcceptanceEnabled: false,
            legacyAcceptanceDisabledAt: harness.readNow(),
        },
    });
    await expectAuthorityError(
        targetLogin(harness, account),
        ERROR_CLASSES.unavailable,
        'authority-generation-mismatch',
    );
    const recovered = await targetLogin(harness, account, {
        authority: generationTwoAuthority,
    });
    assert.equal(
        phaseOf(await generationTwoAuthority.current(recovered.issuance.identifier)),
        SESSION_PHASES.authenticated,
    );

    await expectAuthorityErrorClass(
        generationTwoStore.transitionControl({
            expectedVersion: control.control.version,
            changes: { globalSessionEpoch: control.control.globalSessionEpoch - 1 },
        }),
        ERROR_CLASSES.forbidden,
    );
    await expectAuthorityErrorClass(
        generationTwoStore.transitionControl({
            expectedVersion: control.control.version,
            changes: { authorityGeneration: control.control.authorityGeneration - 1 },
        }),
        ERROR_CLASSES.forbidden,
    );
});

test('control compare-and-replace permits one multi-instance winner', async () => {
    const harness = await createHarness();
    const current = await harness.storeA.readControl();
    const outcomes = await Promise.allSettled([
        harness.storeA.transitionControl({
            expectedVersion: current.control.version,
            changes: { globalSessionEpoch: current.control.globalSessionEpoch + 1 },
        }),
        harness.storeB.transitionControl({
            expectedVersion: current.control.version,
            changes: { globalSessionEpoch: current.control.globalSessionEpoch + 1 },
        }),
    ]);
    const winners = outcomes.filter(({ status }) => status === 'fulfilled');
    const losers = outcomes.filter(({ status }) => status === 'rejected');
    assert.equal(winners.length, 1);
    assert.equal(losers.length, 1);
    assert.equal(
        isAuthorityError(losers[0].reason, ERROR_CLASSES.conflict, 'control-compare-and-replace'),
        true,
    );
});

test('login, logout, and revoke-all races follow commit authority instead of response order', async (t) => {
    await t.test('an issuance committed before revoke-all is inert even when its response arrives later', async () => {
        const account = createAccount({ facePolicy: 'Inativo' });
        const harness = await createHarness({ accounts: [account] });
        const caller = await targetLogin(harness, account);
        const committed = deferred();
        const releaseResponse = deferred();
        const delayedResponseStore = {
            ...harness.storeB,
            async issueSession(input) {
                const result = await harness.storeB.issueSession(input);
                committed.resolve();
                await releaseResponse.promise;
                return result;
            },
        };
        const delayedAuthority = harness.createAuthority(delayedResponseStore);
        const delayedLogin = targetLogin(harness, account, { authority: delayedAuthority });

        await committed.promise;
        await harness.authorityA.revokeAll(caller.issuance.identifier);
        releaseResponse.resolve();
        const lateResponse = await delayedLogin;
        await expectAuthorityError(
            harness.authorityA.current(lateResponse.issuance.identifier),
            ERROR_CLASSES.invalid,
            SESSION_PHASES.revoked,
        );
    });

    await t.test('an issuance committing after revoke-all is a new active login', async () => {
        const account = createAccount({ facePolicy: 'Inativo' });
        const harness = await createHarness({ accounts: [account] });
        const caller = await targetLogin(harness, account);
        const enteredIssue = deferred();
        const releaseCommit = deferred();
        const delayedCommitStore = {
            ...harness.storeB,
            async issueSession(input) {
                enteredIssue.resolve();
                await releaseCommit.promise;
                return harness.storeB.issueSession(input);
            },
        };
        const delayedAuthority = harness.createAuthority(delayedCommitStore);
        const delayedLogin = targetLogin(harness, account, { authority: delayedAuthority });

        await enteredIssue.promise;
        await harness.authorityA.revokeAll(caller.issuance.identifier);
        releaseCommit.resolve();
        const afterEpoch = await delayedLogin;
        assert.equal(
            phaseOf(await harness.authorityA.current(afterEpoch.issuance.identifier)),
            SESSION_PHASES.authenticated,
        );
    });

    await t.test('current logout cannot revoke an independent login that commits afterward', async () => {
        const account = createAccount({ facePolicy: 'Inativo' });
        const harness = await createHarness({ accounts: [account] });
        const current = await targetLogin(harness, account);
        const enteredIssue = deferred();
        const releaseCommit = deferred();
        const delayedCommitStore = {
            ...harness.storeB,
            async issueSession(input) {
                enteredIssue.resolve();
                await releaseCommit.promise;
                return harness.storeB.issueSession(input);
            },
        };
        const delayedAuthority = harness.createAuthority(delayedCommitStore);
        const delayedLogin = targetLogin(harness, account, { authority: delayedAuthority });

        await enteredIssue.promise;
        await harness.authorityA.logout(current.issuance.identifier);
        releaseCommit.resolve();
        const later = await delayedLogin;
        await expectAuthorityError(
            harness.authorityB.current(current.issuance.identifier),
            ERROR_CLASSES.invalid,
            SESSION_PHASES.revoked,
        );
        assert.equal(
            phaseOf(await harness.authorityA.current(later.issuance.identifier)),
            SESSION_PHASES.authenticated,
        );
    });
});
