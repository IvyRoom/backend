'use strict';

const {
    AUTHENTICATED_LIFETIME_MS,
    LEGACY_LIFETIME_MS,
    LEGACY_SEEDING_LEASE_MS,
    LEGACY_SUNSET_MAXIMUM_MS,
    PROVISIONAL_LIFETIME_MS,
    SESSION_PHASES,
} = require('../domains/session-authority/constants');
const {
    authorityConflict,
    authorityUnavailable,
    forbiddenAuthority,
    invalidAuthority,
} = require('../domains/session-authority/errors');

function createTestSessionAuthorityBacking() {
    return {
        mutex: Promise.resolve(),
        state: {
            subjects: new Map(),
            subjectsByLookup: new Map(),
            sessions: new Map(),
            sessionsByVerifier: new Map(),
            flows: new Map(),
            legacyBindings: new Map(),
            control: {
                controlId: 1,
                version: 1,
                authorityGeneration: 1,
                globalSessionEpoch: 1,
                loginLookupKeyId: null,
                loginLookupKeyCommitment: null,
                accountMappingKeyBinding: null,
                authorityKeysetBinding: null,
                legacySigningKeyBinding: null,
                targetRoutesEnabled: false,
                targetSessionIssuanceEnabled: false,
                targetSessionIssuanceStartedAt: null,
                legacyLedgerSeedingEnabled: false,
                legacyCompatibilityEnforcementEnabled: false,
                subjectTargetAdoptionEnabled: false,
                subjectTargetAdoptionStartedAt: null,
                legacyIssuanceEnabled: true,
                legacyAcceptanceEnabled: true,
                legacyLedgerSeedingStartedAt: null,
                seedingStartedAt: null,
                seedingContinuityVersion: 1,
                seedingHeartbeatOwnerId: null,
                seedingHeartbeatAt: null,
                seedingLeaseExpiresAt: null,
                seedingQualifiedAt: null,
                legacyCompatibilityEnforcedAt: null,
                dualStackStartedAt: null,
                legacyStopIssuanceAt: null,
                legacyAcceptanceDisabledAt: null,
                hardSunsetAt: null,
                incidentState: 'normal',
                incidentRecordedAt: null,
                incidentCode: null,
                targetVerifierKeyIncidentAt: null,
                legacyVerifierKeyIncidentAt: null,
            },
        },
    };
}

function createTestSessionAuthorityStore({
    testOnly,
    backing = createTestSessionAuthorityBacking(),
    now = () => new Date('2042-06-01T12:00:00.000Z'),
    failure = {},
    expectedAuthorityGeneration = 1,
    loginLookupKeyId,
    loginLookupKeyCommitment,
    accountMappingKeyBinding,
    authorityKeysetBinding,
    legacySigningKeyBinding,
} = {}) {
    if (testOnly !== true) {
        throw new TypeError('The process-memory session store is test-only');
    }
    if (!Number.isSafeInteger(expectedAuthorityGeneration) || expectedAuthorityGeneration < 1) {
        throw new TypeError('Expected authority generation must be a positive safe integer');
    }
    const lookupFenceEnabled = loginLookupKeyId !== undefined
        || loginLookupKeyCommitment !== undefined;
    if (lookupFenceEnabled) {
        if (typeof loginLookupKeyId !== 'string' || loginLookupKeyId.length === 0) {
            throw new TypeError('Login lookup key ID must be non-empty');
        }
        if (!Buffer.isBuffer(loginLookupKeyCommitment) || loginLookupKeyCommitment.length !== 32) {
            throw new TypeError('Login lookup key commitment must be a 32-byte Buffer');
        }
    }
    const expectedLookupCommitment = lookupFenceEnabled
        ? Buffer.from(loginLookupKeyCommitment)
        : null;
    const keysetFenceEnabled = accountMappingKeyBinding !== undefined
        || authorityKeysetBinding !== undefined
        || legacySigningKeyBinding !== undefined;
    if (keysetFenceEnabled) {
        if (!lookupFenceEnabled) throw new TypeError('Keyset fencing requires lookup fencing');
        validateTestKeyBinding(accountMappingKeyBinding, 'Account mapping');
        validateTestKeyBinding(legacySigningKeyBinding, 'Legacy signing');
        validateTestAuthorityKeysetBinding(authorityKeysetBinding);
        if (
            authorityKeysetBinding.purposes.loginLookup.keyId !== loginLookupKeyId
            || authorityKeysetBinding.purposes.accountMappingEncryption.keyId
                !== accountMappingKeyBinding.keyId
        ) throw new TypeError('Immutable key IDs must match keyset descriptors');
    }
    const expectedAccountMappingKeyBinding = keysetFenceEnabled
        ? copyTestKeyBinding(accountMappingKeyBinding)
        : null;
    const expectedAuthorityKeysetBinding = keysetFenceEnabled
        ? copyTestAuthorityKeysetBinding(authorityKeysetBinding)
        : null;
    const expectedLegacySigningKeyBinding = keysetFenceEnabled
        ? copyTestKeyBinding(legacySigningKeyBinding)
        : null;
    let continuityCompromised = false;
    let continuityObservedActive = false;

    async function exclusive(operation, {
        fenceAuthority = true,
        fenceLookup = true,
        allowRotatableKeyMismatch = false,
    } = {}) {
        let release;
        const previous = backing.mutex;
        backing.mutex = new Promise((resolve) => { release = resolve; });
        await previous;
        try {
            if (failure.unavailable === true) throw authorityUnavailable('session-store-unavailable');
            if (
                fenceAuthority
                && backing.state.control.authorityGeneration !== expectedAuthorityGeneration
            ) throw authorityUnavailable('authority-generation-mismatch');
            if (lookupFenceEnabled && fenceLookup) {
                requireLoginLookupKeyFence(
                    backing.state.control,
                    loginLookupKeyId,
                    expectedLookupCommitment,
                );
            }
            const keyMismatches = keysetFenceEnabled && fenceLookup
                ? requireTestAuthorityKeysetFence(
                    backing.state.control,
                    expectedAccountMappingKeyBinding,
                    expectedAuthorityKeysetBinding,
                    expectedLegacySigningKeyBinding,
                    { allowRotatableKeyMismatch },
                )
                : null;
            const result = await operation(backing.state, readNow(), keyMismatches);
            if (failure.unknownTransaction === true) {
                throw authorityUnavailable('transaction-outcome-unknown');
            }
            return clone(result);
        } catch (error) {
            if (
                error
                && (
                    error.reason === 'session-store-unavailable'
                    || error.reason === 'transaction-outcome-unknown'
                )
            ) continuityCompromised = true;
            throw error;
        } finally {
            release();
        }
    }

    function readNow() {
        const value = now();
        const date = value instanceof Date ? new Date(value.getTime()) : new Date(value);
        if (!Number.isFinite(date.getTime())) throw new TypeError('Test store clock must return a valid instant');
        return date;
    }

    async function readControl() {
        return exclusive((state, serverTime) => ({ control: state.control, serverTime }));
    }

    function requireBoundLoginLookupKeyId(candidate) {
        if (lookupFenceEnabled && candidate !== loginLookupKeyId) {
            throw authorityUnavailable('login-lookup-key-mismatch');
        }
    }

    async function initializeLoginLookupKey({
        loginLookupKeyId: candidateKeyId,
        loginLookupKeyCommitment: candidateCommitment,
    }) {
        if (!lookupFenceEnabled) throw new TypeError('Test lookup-key fence is not configured');
        if (
            candidateKeyId !== loginLookupKeyId
            || !Buffer.isBuffer(candidateCommitment)
            || !candidateCommitment.equals(expectedLookupCommitment)
        ) throw authorityUnavailable('login-lookup-key-mismatch');
        return exclusive((state, serverTime) => {
            const control = state.control;
            if (control.loginLookupKeyId !== null || control.loginLookupKeyCommitment !== null) {
                requireLoginLookupKeyFence(control, loginLookupKeyId, expectedLookupCommitment);
                if (keysetFenceEnabled) {
                    requireTestAuthorityKeysetFence(
                        control,
                        expectedAccountMappingKeyBinding,
                        expectedAuthorityKeysetBinding,
                        expectedLegacySigningKeyBinding,
                    );
                }
                return { control, idempotent: true, serverTime };
            }
            requireDormantLookupInitialization(control);
            if (
                state.subjects.size !== 0
                || state.sessions.size !== 0
                || state.flows.size !== 0
                || state.legacyBindings.size !== 0
            ) throw authorityUnavailable('login-lookup-key-initialization-blocked');
            control.loginLookupKeyId = loginLookupKeyId;
            control.loginLookupKeyCommitment = Buffer.from(expectedLookupCommitment);
            if (keysetFenceEnabled) {
                control.accountMappingKeyBinding = copyTestKeyBinding(
                    expectedAccountMappingKeyBinding,
                );
                control.authorityKeysetBinding = copyTestAuthorityKeysetBinding(
                    expectedAuthorityKeysetBinding,
                );
                control.legacySigningKeyBinding = copyTestKeyBinding(
                    expectedLegacySigningKeyBinding,
                );
            }
            control.version += 1;
            return { control, idempotent: false, serverTime };
        }, { fenceLookup: false });
    }

    async function heartbeatLegacySeedingContinuity({ ownerId }) {
        validateContinuityOwnerId(ownerId);
        return exclusive((state, serverTime) => {
            const control = state.control;
            if (control.incidentState !== 'normal') {
                continuityObservedActive = false;
                continuityCompromised = true;
                return {
                    active: false,
                    owner: false,
                    reset: true,
                    control,
                    serverTime,
                };
            }
            if (!seedingHorizonActive(control)) {
                continuityObservedActive = false;
                continuityCompromised = false;
                return {
                    active: false,
                    owner: false,
                    reset: false,
                    control,
                    serverTime,
                };
            }
            const live = hasLiveSeedingContinuity(control, serverTime);
            const reset = !continuityObservedActive || continuityCompromised || !live;
            if (!reset && control.seedingHeartbeatOwnerId !== ownerId) {
                continuityObservedActive = true;
                continuityCompromised = false;
                return {
                    active: true,
                    owner: false,
                    reset: false,
                    control,
                    serverTime,
                };
            }
            control.seedingContinuityVersion += 1;
            if (reset) control.seedingStartedAt = serverTime;
            control.seedingHeartbeatOwnerId = ownerId;
            control.seedingHeartbeatAt = serverTime;
            control.seedingLeaseExpiresAt = new Date(
                serverTime.getTime() + LEGACY_SEEDING_LEASE_MS,
            );
            continuityObservedActive = true;
            continuityCompromised = false;
            return {
                active: true,
                owner: true,
                reset,
                control,
                serverTime,
            };
        });
    }

    async function transitionControl({ expectedVersion, changes }) {
        if (
            Object.hasOwn(changes, 'incidentCode')
            && changes.incidentCode !== null
            && (
                typeof changes.incidentCode !== 'string'
                || !/^[a-z0-9][a-z0-9-]{0,127}$/.test(changes.incidentCode)
            )
        ) throw new TypeError('Incident code must be a privacy-safe machine value');
        return exclusive((state, serverTime, keyMismatches) => {
            const control = state.control;
            let stampedChanges = stampControlTransition(control, changes, serverTime);
            const keyRecovery = prepareTestAuthorityKeyRecovery({
                state,
                current: control,
                changes: stampedChanges,
                serverTime,
                expectedAuthorityGeneration,
                keyMismatches,
                expectedAuthorityKeysetBinding,
                expectedLegacySigningKeyBinding,
            });
            stampedChanges = keyRecovery.changes;
            requireControlTransitionGeneration(
                control,
                stampedChanges,
                expectedAuthorityGeneration,
                keyRecovery.active,
            );
            if (control.version !== expectedVersion) throw authorityConflict('control-compare-and-replace');
            const next = { ...control, ...stampedChanges };
            if (keyRecovery.active) {
                next.authorityKeysetBinding = copyTestAuthorityKeysetBinding(
                    expectedAuthorityKeysetBinding,
                );
                if (keyRecovery.changedPurposes.includes('legacySigning')) {
                    next.legacySigningKeyBinding = copyTestKeyBinding(
                        expectedLegacySigningKeyBinding,
                    );
                }
            }
            requireIrreversibleControl(control, next, serverTime, stampedChanges, {
                keyRecoveryActive: keyRecovery.active,
            });
            if (
                control.incidentState === 'recovering'
                && next.incidentState === 'normal'
                && [...state.flows.values()].some((flow) => (
                    ['creating', 'active', 'reconciliation-required'].includes(flow.challengeState)
                    && ACTIVE_PHASES_FOR_EVIDENCE.has(
                        state.sessions.get(flow.currentSessionId)?.phase,
                    )
                ))
            ) throw forbiddenAuthority('incident-resume-face-reconciliation-required');
            keyRecovery.apply();
            next.version += 1;
            state.control = next;
            return { control: next, serverTime };
        }, { fenceAuthority: false, allowRotatableKeyMismatch: true });
    }

    async function createOrLoadSubject(input) {
        requireBoundLoginLookupKeyId(input.loginLookupKeyId);
        validateExpectedControlVersion(input.expectedControlVersion);
        return exclusive((state, serverTime) => {
            requireExpectedNormalControl(state.control, input.expectedControlVersion);
            const lookupKey = binaryKey(input.loginLookupKeyId, input.loginLookupToken);
            const existingId = state.subjectsByLookup.get(lookupKey);
            if (existingId !== undefined) {
                const existing = state.subjects.get(existingId);
                if (!existing) throw authorityUnavailable('subject-mapping-integrity');
                return { created: false, subject: existing, serverTime };
            }

            if (state.subjects.has(input.subjectId)) throw authorityUnavailable('subject-identifier-conflict');
            const subject = {
                subjectId: input.subjectId,
                loginLookupKeyId: input.loginLookupKeyId,
                loginLookupToken: copyBuffer(input.loginLookupToken),
                accountMappingKeyId: input.accountMappingKeyId,
                encryptedAccountMapping: copyBuffer(input.encryptedAccountMapping),
                rowHint: input.rowHint,
                credentialVersion: 1,
                credentialFingerprintKeyId: input.credentialFingerprintKeyId,
                credentialFingerprint: copyBuffer(input.credentialFingerprint),
                sessionEpoch: 1,
                legacyAuthorityDisabledAt: null,
                eligibilityState: input.eligibilityState,
                entitlementExpiresAt: copyDate(input.entitlementExpiresAt),
                eligibilityObservedAt: copyDate(input.eligibilityObservedAt),
                eligibilityRevalidateAt: copyDate(input.eligibilityRevalidateAt),
                createdAt: serverTime,
                updatedAt: serverTime,
            };
            state.subjects.set(subject.subjectId, subject);
            state.subjectsByLookup.set(lookupKey, subject.subjectId);
            return { created: true, subject, serverTime };
        });
    }

    async function readSubjectByLookup({ loginLookupKeyId, loginLookupToken, expectedControlVersion }) {
        requireBoundLoginLookupKeyId(loginLookupKeyId);
        if (expectedControlVersion !== undefined) validateExpectedControlVersion(expectedControlVersion);
        return exclusive((state, serverTime) => {
            if (expectedControlVersion !== undefined) {
                requireExpectedNormalControl(state.control, expectedControlVersion);
            }
            const subjectId = state.subjectsByLookup.get(binaryKey(loginLookupKeyId, loginLookupToken));
            return { subject: subjectId === undefined ? null : state.subjects.get(subjectId), serverTime };
        });
    }

    async function remapSubjectLogin(input) {
        requireBoundLoginLookupKeyId(input.expectedLoginLookupKeyId);
        requireBoundLoginLookupKeyId(input.loginLookupKeyId);
        validateSubjectLoginRemap(input);
        return exclusive((state, serverTime) => {
            const subject = requireSubject(state, input.subjectId);
            const currentKey = binaryKey(subject.loginLookupKeyId, subject.loginLookupToken);
            const targetKey = binaryKey(input.loginLookupKeyId, input.loginLookupToken);
            if (currentKey === targetKey) {
                return { subject, idempotent: true, serverTime };
            }
            if (
                subject.loginLookupKeyId !== input.expectedLoginLookupKeyId
                || !sameBuffer(subject.loginLookupToken, input.expectedLoginLookupToken)
            ) throw authorityUnavailable('subject-mapping-conflict');
            if (state.subjectsByLookup.has(targetKey)) {
                throw authorityUnavailable('subject-mapping-conflict');
            }

            state.subjectsByLookup.delete(currentKey);
            subject.loginLookupKeyId = input.loginLookupKeyId;
            subject.loginLookupToken = copyBuffer(input.loginLookupToken);
            subject.accountMappingKeyId = input.accountMappingKeyId;
            subject.encryptedAccountMapping = copyBuffer(input.encryptedAccountMapping);
            subject.updatedAt = serverTime;
            state.subjectsByLookup.set(targetKey, subject.subjectId);
            return { subject, idempotent: false, serverTime };
        });
    }

    async function updateEligibility(input) {
        validateSubjectObservationExpectation(input);
        validateExpectedControlVersion(input.expectedControlVersion);
        return exclusive((state, serverTime) => {
            requireExpectedNormalControl(state.control, input.expectedControlVersion);
            const subject = requireSubject(state, input.subjectId);
            requireSubjectObservationExpectation(subject, input);
            if (
                !(input.entitlementExpiresAt instanceof Date)
                || !Number.isFinite(input.entitlementExpiresAt.getTime())
                || !(input.eligibilityRevalidateAt instanceof Date)
                || !Number.isFinite(input.eligibilityRevalidateAt.getTime())
            ) throw authorityUnavailable('eligibility-data-integrity');
            if (input.entitlementExpiresAt <= serverTime) {
                subject.eligibilityState = 'ineligible';
                subject.entitlementExpiresAt = copyDate(input.entitlementExpiresAt);
                subject.eligibilityObservedAt = serverTime;
                subject.eligibilityRevalidateAt = serverTime;
                subject.sessionEpoch += 1;
                subject.updatedAt = serverTime;
                revokeSubjectSessions(state, subject.subjectId, serverTime, 'entitlement-expired');
                revokeLegacySubjectBindings(state, subject.subjectId, serverTime, 'entitlement-expired');
                return { eligible: false, subject, serverTime };
            }
            const eligibilityRevalidateAt = new Date(Math.min(
                input.eligibilityRevalidateAt.getTime(),
                input.entitlementExpiresAt.getTime(),
            ));
            if (eligibilityRevalidateAt <= serverTime) {
                throw authorityUnavailable('eligibility-revalidation-required');
            }
            subject.rowHint = input.rowHint;
            subject.eligibilityState = 'eligible';
            subject.entitlementExpiresAt = copyDate(input.entitlementExpiresAt);
            subject.eligibilityObservedAt = serverTime;
            subject.eligibilityRevalidateAt = eligibilityRevalidateAt;
            subject.updatedAt = serverTime;
            return { eligible: true, subject, serverTime };
        });
    }

    async function revokeForIneligibility(input) {
        validateSubjectObservationExpectation(input);
        validateExpectedControlVersion(input.expectedControlVersion);
        return exclusive((state, serverTime) => {
            requireExpectedNormalControl(state.control, input.expectedControlVersion);
            const subject = requireSubject(state, input.subjectId);
            requireSubjectObservationExpectation(subject, input);
            subject.eligibilityState = input.eligibilityState || 'ineligible';
            subject.entitlementExpiresAt = copyDate(input.entitlementExpiresAt);
            subject.eligibilityObservedAt = copyDate(input.observationStartedAt);
            subject.eligibilityRevalidateAt = copyDate(input.observationStartedAt);
            subject.sessionEpoch += 1;
            subject.updatedAt = serverTime;
            revokeSubjectSessions(state, subject.subjectId, serverTime, input.reason || 'ineligible');
            revokeLegacySubjectBindings(state, subject.subjectId, serverTime, input.reason || 'ineligible');
            return { subject, serverTime };
        });
    }

    async function revokeForCredentialChange(input) {
        validateSubjectObservationExpectation(input);
        validateExpectedControlVersion(input.expectedControlVersion);
        return exclusive((state, serverTime) => {
            requireExpectedNormalControl(state.control, input.expectedControlVersion);
            const subject = requireSubject(state, input.subjectId);
            requireSubjectObservationExpectation(subject, input);
            subject.credentialVersion += 1;
            subject.credentialFingerprintKeyId = input.credentialFingerprintKeyId;
            subject.credentialFingerprint = copyBuffer(input.credentialFingerprint);
            subject.eligibilityObservedAt = copyDate(input.observationStartedAt);
            subject.eligibilityRevalidateAt = copyDate(input.observationStartedAt);
            subject.sessionEpoch += 1;
            subject.updatedAt = serverTime;
            revokeSubjectSessions(state, subject.subjectId, serverTime, 'credential-reset');
            revokeLegacySubjectBindings(state, subject.subjectId, serverTime, 'credential-reset');
            return { subject, serverTime };
        });
    }

    async function inspectLoginPredecessor(input) {
        return exclusive((state, serverTime) => {
            const session = state.sessionsByVerifier.get(binaryKey(input.verifierKeyId, input.verifier));
            if (!session) return { kind: 'unusable', serverTime };
            const authority = evaluateSession(state, session, serverTime);
            if (authority.entitlementExpired) {
                expireSubjectEntitlement(state, authority.subject, serverTime);
                return { kind: 'unusable', serverTime };
            }
            if (authority.unavailable) throw authorityUnavailable(authority.reason);
            if (!authority.active) return { kind: 'unusable', serverTime };
            return {
                kind: 'active',
                expectedSessionId: session.sessionId,
                expectedVersion: session.version,
                serverTime,
            };
        });
    }

    async function issueSession(input) {
        validateSubjectCredentialExpectation(input);
        return exclusive((state, serverTime) => {
            const subject = requireSubject(state, input.subjectId);
            requireSubjectCredentialExpectation(subject, input);
            const control = state.control;
            requireEligibleSubject(subject, serverTime);
            requireFreshEligibility(subject, serverTime);
            requireTargetIssuanceControl(control, serverTime);

            if (input.predecessor) {
                const predecessor = state.sessions.get(input.predecessor.expectedSessionId);
                if (!predecessor || predecessor.version !== input.predecessor.expectedVersion) {
                    throw authorityConflict('session-compare-and-replace');
                }
                const predecessorAuthority = evaluateSession(state, predecessor, serverTime);
                if (predecessorAuthority.entitlementExpired) {
                    expireSubjectEntitlement(state, predecessorAuthority.subject, serverTime);
                    throw forbiddenAuthority('ineligible');
                }
                if (predecessorAuthority.unavailable) {
                    throw authorityUnavailable(predecessorAuthority.reason);
                }
                if (!predecessorAuthority.active) {
                    throw authorityConflict('session-compare-and-replace');
                }
                rotateOut(predecessor, input.sessionId, serverTime);
            }

            if (subject.legacyAuthorityDisabledAt === null) {
                subject.legacyAuthorityDisabledAt = serverTime;
                subject.updatedAt = serverTime;
            }

            const session = createSessionRecord({
                ...input,
                subject,
                control,
                serverTime,
                originalIssuedAt: serverTime,
                phaseStartedAt: serverTime,
                expiresAt: new Date(serverTime.getTime() + (
                    input.phase === SESSION_PHASES.authenticated
                        ? AUTHENTICATED_LIFETIME_MS
                        : PROVISIONAL_LIFETIME_MS
                )),
            });
            insertSession(state, session);
            return { session, subject, control, serverTime };
        });
    }

    async function readSession(input) {
        return exclusive((state, serverTime) => {
            const session = state.sessionsByVerifier.get(binaryKey(input.verifierKeyId, input.verifier));
            if (!session) throw invalidAuthority('unknown-session');
            const authority = evaluateSession(state, session, serverTime);
            if (authority.entitlementExpired) {
                expireSubjectEntitlement(state, authority.subject, serverTime);
                throw forbiddenAuthority('ineligible');
            }
            if (!authority.active) throwAuthorityState(authority);
            return {
                session,
                subject: authority.subject,
                control: state.control,
                serverTime,
            };
        });
    }

    async function rotateSession(input) {
        return exclusive((state, serverTime) => {
            const predecessor = requireExpectedActiveSession(state, input, serverTime);
            if (!input.allowedPhases.includes(predecessor.phase)) throw forbiddenAuthority('wrong-phase');

            const subject = requireSubject(state, predecessor.subjectId);
            requireFreshEligibility(subject, serverTime);
            const expiresAt = input.phase === SESSION_PHASES.authenticated
                ? new Date(serverTime.getTime() + AUTHENTICATED_LIFETIME_MS)
                : predecessor.expiresAt;
            const originalIssuedAt = input.phase === SESSION_PHASES.authenticated
                ? serverTime
                : predecessor.originalIssuedAt;
            const session = createSessionRecord({
                ...input,
                subject,
                control: state.control,
                serverTime,
                faceRequired: predecessor.faceRequired,
                registrationRequired: predecessor.registrationRequired,
                originalIssuedAt,
                phaseStartedAt: serverTime,
                expiresAt,
            });
            rotateOut(predecessor, session.sessionId, serverTime);
            insertSession(state, session);
            migrateFlow(state, predecessor.sessionId, session.sessionId);
            return { session, subject, control: state.control, serverTime };
        });
    }

    async function reserveFaceFlow(input) {
        return exclusive((state, serverTime) => {
            const session = requireExpectedActiveSession(state, input, serverTime);
            requireFreshEligibility(requireSubject(state, session.subjectId), serverTime);
            if (!input.allowedPhases.includes(session.phase)) throw forbiddenAuthority('wrong-phase');
            const existing = state.flows.get(session.sessionId);
            if (existing && ['creating', 'bound', 'reconciliation-required'].includes(existing.challengeState)) {
                throw authorityConflict('face-challenge-active');
            }
            state.flows.set(session.sessionId, {
                flowId: input.flowId,
                subjectId: session.subjectId,
                currentSessionId: session.sessionId,
                challengeSessionId: null,
                registrationState: input.registrationState,
                challengeState: 'creating',
                challengeKeyId: null,
                encryptedChallenge: null,
                challengeCreatedAt: null,
                createdAt: serverTime,
                consumedAt: null,
            });
            return { flowId: input.flowId, session, serverTime };
        });
    }

    async function markFaceFlowReconciliation(input) {
        return exclusive((state, serverTime) => {
            const matches = [...state.flows.values()].filter((flow) => flow.flowId === input.flowId);
            if (
                matches.length !== 1
                || !['creating', 'active', 'reconciliation-required'].includes(
                    matches[0].challengeState,
                )
            ) {
                throw authorityUnavailable('face-flow-reconciliation-unavailable');
            }
            const [flow] = matches;
            flow.challengeState = 'reconciliation-required';
            if (input.registrationReconciliationRequired) {
                flow.registrationState = 'reconciliation-required';
            }
            return { serverTime };
        });
    }

    async function bindFaceChallengeAndRotate(input) {
        return exclusive((state, serverTime) => {
            const predecessor = requireExpectedActiveSession(state, input, serverTime);
            const subject = requireSubject(state, predecessor.subjectId);
            requireFreshEligibility(subject, serverTime);
            const flow = state.flows.get(predecessor.sessionId);
            if (
                !flow
                || flow.flowId !== input.flowId
                || flow.challengeState !== 'creating'
            ) throw authorityConflict('face-flow-not-reserved');
            if (!input.allowedPhases.includes(predecessor.phase)) throw forbiddenAuthority('wrong-phase');
            flow.registrationState = 'registered';

            const session = createSessionRecord({
                ...input,
                phase: SESSION_PHASES.facePending,
                subject,
                control: state.control,
                serverTime,
                faceRequired: predecessor.faceRequired,
                registrationRequired: predecessor.registrationRequired,
                originalIssuedAt: predecessor.originalIssuedAt,
                phaseStartedAt: serverTime,
                expiresAt: predecessor.expiresAt,
            });
            flow.challengeState = 'active';
            flow.challengeSessionId = session.sessionId;
            flow.challengeKeyId = input.challengeKeyId;
            flow.encryptedChallenge = copyBuffer(input.encryptedChallenge);
            flow.challengeCreatedAt = serverTime;
            state.flows.delete(predecessor.sessionId);
            flow.currentSessionId = session.sessionId;
            state.flows.set(session.sessionId, flow);
            rotateOut(predecessor, session.sessionId, serverTime);
            insertSession(state, session);
            return { session, subject, control: state.control, serverTime };
        });
    }

    async function readFaceFlow(input) {
        return exclusive((state, serverTime) => {
            const session = requireExpectedActiveSession(state, input, serverTime);
            requireFreshEligibility(requireSubject(state, session.subjectId), serverTime);
            const flow = state.flows.get(session.sessionId);
            if (session.phase === SESSION_PHASES.authenticated) {
                if (
                    input.allowConsumed !== true
                    || !session.faceRequired
                    || !flow
                    || flow.challengeState !== 'consumed'
                    || !(flow.consumedAt instanceof Date)
                ) throw forbiddenAuthority('face-completion-not-applicable');
                return {
                    session,
                    flow,
                    subject: requireSubject(state, session.subjectId),
                    control: state.control,
                    serverTime,
                };
            }
            if (session.phase !== SESSION_PHASES.facePending) throw forbiddenAuthority('wrong-phase');
            if (!flow || flow.challengeState !== 'active' || !flow.encryptedChallenge) {
                throw authorityConflict('face-challenge-unavailable');
            }
            return {
                session,
                flow,
                subject: requireSubject(state, session.subjectId),
                control: state.control,
                serverTime,
            };
        });
    }

    async function completeFaceSuccessAndRotate(input) {
        return exclusive((state, serverTime) => {
            const predecessor = requireExpectedActiveSession(state, input, serverTime);
            const subject = requireSubject(state, predecessor.subjectId);
            requireFreshEligibility(subject, serverTime);
            if (predecessor.phase !== SESSION_PHASES.facePending) throw forbiddenAuthority('wrong-phase');
            const flow = state.flows.get(predecessor.sessionId);
            if (!flow || flow.challengeState !== 'active') throw authorityConflict('face-challenge-unavailable');

            const session = createSessionRecord({
                ...input,
                phase: SESSION_PHASES.authenticated,
                subject,
                control: state.control,
                serverTime,
                faceRequired: predecessor.faceRequired,
                registrationRequired: predecessor.registrationRequired,
                originalIssuedAt: serverTime,
                phaseStartedAt: serverTime,
                expiresAt: new Date(serverTime.getTime() + AUTHENTICATED_LIFETIME_MS),
            });
            flow.challengeState = 'consumed';
            flow.consumedAt = serverTime;
            rotateOut(predecessor, session.sessionId, serverTime);
            insertSession(state, session);
            state.flows.delete(predecessor.sessionId);
            state.flows.set(session.sessionId, flow);
            flow.currentSessionId = session.sessionId;
            return { session, subject, control: state.control, serverTime };
        });
    }

    async function completeFaceFailure(input) {
        return exclusive((state, serverTime) => {
            const session = requireExpectedActiveSession(state, input, serverTime);
            if (session.phase !== SESSION_PHASES.facePending) throw forbiddenAuthority('wrong-phase');
            session.phase = SESSION_PHASES.revoked;
            session.revokedAt = serverTime;
            session.revocationReason = 'face-factor-failed';
            session.version += 1;
            const flow = state.flows.get(session.sessionId);
            if (flow) {
                flow.challengeState = 'failed';
                flow.consumedAt = serverTime;
            }
            return { session, serverTime };
        });
    }

    async function logout(input) {
        return exclusive((state, serverTime) => {
            if (state.control.incidentState !== 'normal') {
                throw authorityUnavailable('authority-incident');
            }
            if (!state.control.targetRoutesEnabled) {
                throw authorityUnavailable('target-routes-disabled');
            }
            const session = state.sessionsByVerifier.get(binaryKey(input.verifierKeyId, input.verifier));
            if (!session) {
                return { revoked: false, serverTime };
            }
            const authority = evaluateSession(state, session, serverTime);
            if (authority.entitlementExpired) {
                expireSubjectEntitlement(state, authority.subject, serverTime);
                return { revoked: true, serverTime };
            }
            if (authority.unavailable) throw authorityUnavailable(authority.reason);
            if (!authority.active) return { revoked: false, serverTime };
            session.phase = SESSION_PHASES.revoked;
            session.revokedAt = serverTime;
            session.revocationReason = 'logout';
            session.version += 1;
            return { revoked: true, serverTime };
        });
    }

    async function revokeAll(input) {
        return exclusive((state, serverTime) => {
            const session = requireExpectedActiveSession(state, input, serverTime);
            if (session.phase !== SESSION_PHASES.authenticated) throw forbiddenAuthority('wrong-phase');
            const subject = requireSubject(state, session.subjectId);
            subject.sessionEpoch += 1;
            subject.updatedAt = serverTime;
            revokeSubjectSessions(state, subject.subjectId, serverTime, input.reason || 'revoke-all');
            revokeLegacySubjectBindings(state, subject.subjectId, serverTime, input.reason || 'revoke-all');
            return { subject, serverTime };
        });
    }

    async function revokeSubject(input) {
        return exclusive((state, serverTime) => {
            if (!state.control.legacyCompatibilityEnforcementEnabled) {
                throw authorityUnavailable('legacy-enforcement-required');
            }
            const subject = requireSubject(state, input.subjectId);
            subject.sessionEpoch += 1;
            subject.updatedAt = serverTime;
            revokeSubjectSessions(state, subject.subjectId, serverTime, input.reason);
            revokeLegacySubjectBindings(state, subject.subjectId, serverTime, input.reason);
            return { subject, serverTime };
        });
    }

    async function disableLegacyAuthority(input) {
        return exclusive((state, serverTime) => {
            if (!state.control.legacyCompatibilityEnforcementEnabled) {
                throw authorityUnavailable('legacy-enforcement-required');
            }
            const subject = requireSubject(state, input.subjectId);
            if (subject.legacyAuthorityDisabledAt === null) {
                subject.legacyAuthorityDisabledAt = serverTime;
                subject.updatedAt = serverTime;
            }
            revokeLegacySubjectBindings(
                state,
                subject.subjectId,
                serverTime,
                input.reason || 'legacy-handle-leak',
            );
            return { subject, serverTime };
        });
    }

    async function admitUnboundLegacyIssuance(input) {
        requireBoundLoginLookupKeyId(input.loginLookupKeyId);
        validateLegacyLifetime(input.issuedAt, input.expiresAt);
        return exclusive((state, serverTime) => {
            const control = state.control;
            requireUnboundLegacyIssuanceControl(control, serverTime);
            requireLegacyCandidateWithinSunset(control, input.expiresAt);
            if (input.issuedAt > serverTime || input.expiresAt <= serverTime) {
                throw authorityUnavailable('legacy-issuance-time-invalid');
            }
            const subjectId = state.subjectsByLookup.get(binaryKey(
                input.loginLookupKeyId,
                input.loginLookupToken,
            ));
            if (subjectId !== undefined) {
                const subject = state.subjects.get(subjectId);
                if (!subject) throw authorityUnavailable('subject-mapping-integrity');
                if (subject.legacyAuthorityDisabledAt !== null) {
                    throw authorityConflict('target-authority-established');
                }
            }
            return { admitted: true, control, serverTime };
        });
    }

    async function bindLegacy(input) {
        validateLegacyLifetime(input.issuedAt, input.expiresAt);
        validateSubjectCredentialExpectation(input);
        return exclusive((state, serverTime) => {
            const control = state.control;
            const subject = requireSubject(state, input.subjectId);
            requireLegacyIssuanceControl(control, serverTime);
            if (
                continuityCompromised
                || (seedingHorizonActive(control) && !continuityObservedActive)
            ) {
                throw authorityUnavailable('legacy-seeding-continuity-unavailable');
            }
            requireLiveSeedingContinuity(control, serverTime);
            requireLegacyCandidateWithinSunset(control, input.expiresAt);
            if (input.issuedAt > serverTime || input.expiresAt <= serverTime) {
                throw authorityUnavailable('legacy-issuance-time-invalid');
            }
            if (!control.legacyLedgerSeedingEnabled) {
                throw authorityUnavailable('legacy-ledger-seeding-disabled');
            }
            requireSubjectCredentialExpectation(subject, input);
            if (control.legacyCompatibilityEnforcementEnabled) {
                requireFreshEligibility(subject, serverTime);
            }
            if (subject.legacyAuthorityDisabledAt !== null) throw authorityConflict('target-authority-established');
            const key = binaryKey(input.verifierKeyId, input.verifier);
            const existing = state.legacyBindings.get(key);
            if (existing) {
                if (!hasValidLegacyBinding(existing)) {
                    throw authorityUnavailable('legacy-binding-integrity');
                }
                if (
                    existing.subjectId !== input.subjectId
                    || existing.issuedAt.getTime() !== input.issuedAt.getTime()
                    || existing.expiresAt.getTime() !== input.expiresAt.getTime()
                    || existing.verifierKeyId !== input.verifierKeyId
                ) {
                    throw authorityUnavailable('legacy-binding-integrity');
                }
                if (existing.revokedAt !== null || existing.incidentState !== 'normal') {
                    throw authorityUnavailable('legacy-binding-terminal');
                }
                return { binding: existing, idempotent: true, serverTime };
            }
            const binding = {
                legacyCompatibilityId: input.legacyCompatibilityId,
                verifierKeyId: input.verifierKeyId,
                verifier: copyBuffer(input.verifier),
                subjectId: input.subjectId,
                issuedAt: copyDate(input.issuedAt),
                expiresAt: copyDate(input.expiresAt),
                revokedAt: null,
                revocationReason: null,
                incidentState: 'normal',
                createdAt: serverTime,
            };
            state.legacyBindings.set(key, binding);
            return { binding, idempotent: false, serverTime };
        });
    }

    async function authorizeLegacy(input) {
        validateLegacyLifetime(input.issuedAt, input.expiresAt);
        return exclusive((state, serverTime) => {
            const control = state.control;
            if (control.incidentState !== 'normal') throw authorityUnavailable('authority-incident');
            if (!control.legacyAcceptanceEnabled || (control.hardSunsetAt && serverTime >= control.hardSunsetAt)) {
                throw invalidAuthority('legacy-acceptance-disabled');
            }
            const binding = state.legacyBindings.get(binaryKey(input.verifierKeyId, input.verifier));
            if (!binding) {
                if (control.legacyCompatibilityEnforcementEnabled) {
                    throw invalidAuthority('legacy-binding-missing');
                }
                if (input.issuedAt > serverTime || input.expiresAt <= serverTime) {
                    throw invalidAuthority('legacy-binding-terminal');
                }
                return { unbound: true, control, serverTime };
            }
            if (!hasValidLegacyBinding(binding)) {
                throw authorityUnavailable('legacy-binding-integrity');
            }
            if (
                binding.verifierKeyId !== input.verifierKeyId
                || binding.issuedAt.getTime() !== input.issuedAt.getTime()
                || binding.expiresAt.getTime() !== input.expiresAt.getTime()
            ) {
                throw authorityUnavailable('legacy-binding-integrity');
            }
            if (!(binding.issuedAt instanceof Date) || binding.issuedAt > serverTime) {
                throw authorityUnavailable('legacy-binding-integrity');
            }
            if (!(binding.expiresAt instanceof Date)) throw authorityUnavailable('legacy-binding-integrity');
            if (binding.incidentState === 'incident') throw authorityUnavailable('legacy-binding-integrity');
            if (!control.legacyCompatibilityEnforcementEnabled) {
                if (binding.expiresAt <= serverTime) throw invalidAuthority('legacy-binding-terminal');
                if (
                    binding.incidentState !== 'normal'
                    || binding.revokedAt !== null
                ) throw invalidAuthority('legacy-binding-terminal');
                return { unbound: true, control, serverTime };
            }
            if (
                binding.incidentState !== 'normal'
                || binding.revokedAt !== null
            ) throw invalidAuthority('legacy-binding-terminal');
            const subject = requireSubject(state, binding.subjectId);
            if (subject.legacyAuthorityDisabledAt !== null) throw invalidAuthority('target-authority-established');
            const bindingExpired = binding.expiresAt <= serverTime;
            const entitlementExpired = subject.entitlementExpiresAt <= serverTime;
            if (
                entitlementExpired
                && (!bindingExpired || subject.entitlementExpiresAt <= binding.expiresAt)
            ) {
                expireSubjectEntitlement(state, subject, serverTime);
                throw forbiddenAuthority('ineligible');
            }
            if (bindingExpired) throw invalidAuthority('legacy-binding-terminal');
            if (entitlementExpired) {
                expireSubjectEntitlement(state, subject, serverTime);
                throw forbiddenAuthority('ineligible');
            }
            requireEligibleSubject(subject, serverTime);
            return { binding, subject, control, serverTime };
        });
    }

    return {
        admitUnboundLegacyIssuance,
        authorizeLegacy,
        bindFaceChallengeAndRotate,
        bindLegacy,
        completeFaceFailure,
        completeFaceSuccessAndRotate,
        createOrLoadSubject,
        disableLegacyAuthority,
        heartbeatLegacySeedingContinuity,
        initializeLoginLookupKey,
        inspectLoginPredecessor,
        issueSession,
        logout,
        markFaceFlowReconciliation,
        readControl,
        readFaceFlow,
        readSession,
        readSubjectByLookup,
        remapSubjectLogin,
        reserveFaceFlow,
        revokeAll,
        revokeSubject,
        revokeForCredentialChange,
        revokeForIneligibility,
        rotateSession,
        transitionControl,
        updateEligibility,
    };
}

function createSessionRecord(input) {
    return {
        sessionId: input.sessionId,
        verifierKeyId: input.verifierKeyId,
        verifier: copyBuffer(input.verifier),
        subjectId: input.subject.subjectId,
        phase: input.phase,
        originalIssuedAt: copyDate(input.originalIssuedAt),
        phaseStartedAt: copyDate(input.phaseStartedAt),
        createdAt: copyDate(input.phaseStartedAt),
        expiresAt: copyDate(input.expiresAt),
        faceRequired: input.faceRequired,
        registrationRequired: input.registrationRequired,
        subjectEpochSnapshot: input.subject.sessionEpoch,
        credentialVersionSnapshot: input.subject.credentialVersion,
        globalEpochSnapshot: input.control.globalSessionEpoch,
        authorityGenerationSnapshot: input.control.authorityGeneration,
        revokedAt: null,
        revocationReason: null,
        replacementSessionId: null,
        version: 1,
    };
}

function insertSession(state, session) {
    const key = binaryKey(session.verifierKeyId, session.verifier);
    if (state.sessions.has(session.sessionId) || state.sessionsByVerifier.has(key)) {
        throw authorityUnavailable('session-verifier-conflict');
    }
    state.sessions.set(session.sessionId, session);
    state.sessionsByVerifier.set(key, session);
}

function requireExpectedActiveSession(state, input, serverTime) {
    const session = state.sessions.get(input.expectedSessionId);
    if (!session || session.version !== input.expectedVersion) {
        throw authorityConflict('session-compare-and-replace');
    }
    const authority = evaluateSession(state, session, serverTime);
    if (authority.entitlementExpired) {
        expireSubjectEntitlement(state, authority.subject, serverTime);
        throw forbiddenAuthority('ineligible');
    }
    if (!authority.active) throwAuthorityState(authority);
    return session;
}

function evaluateSession(state, session, serverTime) {
    const subject = state.subjects.get(session.subjectId);
    if (!subject) return { active: false, unavailable: true, reason: 'subject-mapping-integrity' };
    if (!hasValidStoredSessionRecord(session)) {
        return { active: false, unavailable: true, reason: 'session-store-integrity' };
    }
    if (![SESSION_PHASES.credentialVerified, SESSION_PHASES.registrationPending, SESSION_PHASES.facePending, SESSION_PHASES.authenticated].includes(session.phase)) {
        return { active: false, reason: session.phase };
    }
    if (session.revokedAt !== null) return { active: false, reason: 'revoked' };
    if (!hasValidActiveSessionLifetime(session)) {
        return { active: false, unavailable: true, reason: 'session-store-integrity' };
    }
    if (!hasValidActiveSessionEvidence(
        state,
        session,
        state.flows.get(session.sessionId),
        serverTime,
    )) {
        return { active: false, unavailable: true, reason: 'session-store-integrity' };
    }
    if (
        !(subject.eligibilityObservedAt instanceof Date)
        || !Number.isFinite(subject.eligibilityObservedAt.getTime())
        || !(subject.eligibilityRevalidateAt instanceof Date)
        || !Number.isFinite(subject.eligibilityRevalidateAt.getTime())
        || subject.eligibilityObservedAt > serverTime
        || subject.eligibilityRevalidateAt < subject.eligibilityObservedAt
        || subject.eligibilityRevalidateAt.getTime()
            - subject.eligibilityObservedAt.getTime() > 5 * 60 * 1000
    ) return { active: false, unavailable: true, reason: 'subject-eligibility-integrity' };
    const sessionExpired = session.expiresAt <= serverTime;
    const entitlementExpired = subject.entitlementExpiresAt <= serverTime;
    if (
        entitlementExpired
        && (!sessionExpired || subject.entitlementExpiresAt <= session.expiresAt)
    ) {
        return { active: false, entitlementExpired: true, reason: 'ineligible', subject };
    }
    if (sessionExpired) {
        session.phase = SESSION_PHASES.expired;
        session.version += 1;
        return { active: false, reason: 'expired' };
    }
    if (entitlementExpired) {
        return { active: false, entitlementExpired: true, reason: 'ineligible', subject };
    }
    if (state.control.incidentState !== 'normal') {
        return { active: false, unavailable: true, reason: 'authority-incident' };
    }
    if (
        session.subjectEpochSnapshot !== subject.sessionEpoch
        || session.credentialVersionSnapshot !== subject.credentialVersion
        || session.globalEpochSnapshot !== state.control.globalSessionEpoch
        || session.authorityGenerationSnapshot !== state.control.authorityGeneration
    ) {
        return { active: false, reason: 'epoch-mismatch' };
    }
    if (subject.eligibilityState !== 'eligible') return { active: false, reason: 'ineligible' };
    return { active: true, subject };
}

function hasValidActiveSessionLifetime(session) {
    if (
        !(session.originalIssuedAt instanceof Date)
        || !Number.isFinite(session.originalIssuedAt.getTime())
        || !(session.phaseStartedAt instanceof Date)
        || !Number.isFinite(session.phaseStartedAt.getTime())
        || !(session.expiresAt instanceof Date)
        || !Number.isFinite(session.expiresAt.getTime())
        || session.originalIssuedAt > session.phaseStartedAt
        || session.phaseStartedAt >= session.expiresAt
    ) return false;
    const expectedLifetime = session.phase === SESSION_PHASES.authenticated
        ? AUTHENTICATED_LIFETIME_MS
        : PROVISIONAL_LIFETIME_MS;
    return session.expiresAt.getTime() - session.originalIssuedAt.getTime() === expectedLifetime;
}

function hasValidStoredSessionRecord(session) {
    const phases = [
        SESSION_PHASES.credentialVerified,
        SESSION_PHASES.registrationPending,
        SESSION_PHASES.facePending,
        SESSION_PHASES.authenticated,
        SESSION_PHASES.expired,
        SESSION_PHASES.revoked,
        SESSION_PHASES.rotatedOut,
    ];
    if (
        !phases.includes(session.phase)
        || typeof session.faceRequired !== 'boolean'
        || typeof session.registrationRequired !== 'boolean'
        || !Buffer.isBuffer(session.verifier)
        || session.verifier.length !== 32
        || typeof session.verifierKeyId !== 'string'
        || session.verifierKeyId.length === 0
        || typeof session.sessionId !== 'string'
        || typeof session.subjectId !== 'string'
        || !(session.originalIssuedAt instanceof Date)
        || !Number.isFinite(session.originalIssuedAt.getTime())
        || !(session.phaseStartedAt instanceof Date)
        || !Number.isFinite(session.phaseStartedAt.getTime())
        || !(session.expiresAt instanceof Date)
        || !Number.isFinite(session.expiresAt.getTime())
        || session.originalIssuedAt > session.phaseStartedAt
        || session.phaseStartedAt >= session.expiresAt
        || ![
            session.subjectEpochSnapshot,
            session.credentialVersionSnapshot,
            session.globalEpochSnapshot,
            session.authorityGenerationSnapshot,
        ].every((value) => Number.isSafeInteger(value) && value >= 1)
    ) return false;
    const lifetime = session.expiresAt.getTime() - session.originalIssuedAt.getTime();
    if (
        ([SESSION_PHASES.credentialVerified, SESSION_PHASES.registrationPending,
            SESSION_PHASES.facePending].includes(session.phase)
            && lifetime !== PROVISIONAL_LIFETIME_MS)
        || (session.phase === SESSION_PHASES.authenticated
            && lifetime !== AUTHENTICATED_LIFETIME_MS)
        || ([SESSION_PHASES.expired, SESSION_PHASES.revoked, SESSION_PHASES.rotatedOut]
            .includes(session.phase)
            && ![PROVISIONAL_LIFETIME_MS, AUTHENTICATED_LIFETIME_MS].includes(lifetime))
    ) return false;
    if (ACTIVE_PHASES_FOR_EVIDENCE.has(session.phase) || session.phase === SESSION_PHASES.expired) {
        return session.revokedAt === null
            && session.revocationReason === null
            && session.replacementSessionId === null;
    }
    if (session.phase === SESSION_PHASES.revoked) {
        return session.revokedAt instanceof Date
            && Number.isFinite(session.revokedAt.getTime())
            && isPrivacySafeMachineToken(session.revocationReason)
            && session.replacementSessionId === null;
    }
    return session.revokedAt === null
        && session.revocationReason === null
        && typeof session.replacementSessionId === 'string'
        && session.replacementSessionId !== session.sessionId;
}

const ACTIVE_PHASES_FOR_EVIDENCE = new Set([
    SESSION_PHASES.credentialVerified,
    SESSION_PHASES.registrationPending,
    SESSION_PHASES.facePending,
    SESSION_PHASES.authenticated,
]);

function hasValidActiveSessionEvidence(state, session, flow, serverTime) {
    if (
        session.originalIssuedAt > serverTime
        || session.phaseStartedAt > serverTime
        || session.createdAt > serverTime
        || !sameInstant(session.createdAt, session.phaseStartedAt)
    ) return false;
    const provisional = [
        SESSION_PHASES.credentialVerified,
        SESSION_PHASES.registrationPending,
        SESSION_PHASES.facePending,
    ].includes(session.phase);
    if (provisional && !session.faceRequired) return false;
    if (session.phase === SESSION_PHASES.registrationPending && !session.registrationRequired) {
        return false;
    }
    if (flow !== undefined) {
        if (flow.subjectId !== session.subjectId || flow.currentSessionId !== session.sessionId) {
            return false;
        }
    }
    if (session.phase === SESSION_PHASES.facePending) {
        return flowHasResolvedReference(flow, serverTime, 'active', false)
            && flow.challengeSessionId === session.sessionId
            && hasValidChallengeSessionLineage(state, flow, session, false);
    }
    if (session.phase === SESSION_PHASES.authenticated) {
        if (!session.faceRequired) return flow === undefined;
        return flowHasResolvedReference(flow, serverTime, 'consumed', true)
            && hasValidChallengeSessionLineage(state, flow, session, true);
    }
    if (flow === undefined) return true;
    return ['creating', 'reconciliation-required'].includes(flow.challengeState)
        && flow.encryptedChallenge === null
        && flow.challengeKeyId === null
        && flow.challengeCreatedAt === null
        && flow.consumedAt === null;
}

function hasValidChallengeSessionLineage(state, flow, currentSession, consumed) {
    const challengeSession = state.sessions.get(flow?.challengeSessionId);
    if (
        !challengeSession
        || challengeSession.sessionId !== flow.challengeSessionId
        || challengeSession.subjectId !== flow.subjectId
        || challengeSession.subjectId !== currentSession.subjectId
    ) return false;
    if (!consumed) {
        return challengeSession.phase === SESSION_PHASES.facePending
            && challengeSession.sessionId === currentSession.sessionId
            && challengeSession.faceRequired === true
            && challengeSession.registrationRequired === currentSession.registrationRequired
            && challengeSession.expiresAt.getTime()
                - challengeSession.originalIssuedAt.getTime() === PROVISIONAL_LIFETIME_MS
            && sameInstant(flow.challengeCreatedAt, challengeSession.phaseStartedAt);
    }
    return challengeSession.phase === SESSION_PHASES.rotatedOut
        && challengeSession.faceRequired === true
        && challengeSession.registrationRequired === currentSession.registrationRequired
        && challengeSession.subjectEpochSnapshot === currentSession.subjectEpochSnapshot
        && challengeSession.credentialVersionSnapshot === currentSession.credentialVersionSnapshot
        && challengeSession.globalEpochSnapshot === currentSession.globalEpochSnapshot
        && challengeSession.authorityGenerationSnapshot
            === currentSession.authorityGenerationSnapshot
        && challengeSession.expiresAt.getTime()
            - challengeSession.originalIssuedAt.getTime() === PROVISIONAL_LIFETIME_MS
        && sameInstant(flow.challengeCreatedAt, challengeSession.phaseStartedAt)
        && sameInstant(currentSession.originalIssuedAt, flow.consumedAt)
        && sameInstant(currentSession.phaseStartedAt, flow.consumedAt)
        && sameInstant(currentSession.createdAt, flow.consumedAt)
        && challengeSession.replacementSessionId === currentSession.sessionId;
}

function flowHasResolvedReference(flow, serverTime, expectedState, resolved) {
    if (
        !flow
        || flow.challengeState !== expectedState
        || flow.registrationState !== 'registered'
        || !Buffer.isBuffer(flow.encryptedChallenge)
        || flow.encryptedChallenge.length === 0
        || typeof flow.challengeKeyId !== 'string'
        || flow.challengeKeyId.length === 0
        || !(flow.challengeCreatedAt instanceof Date)
        || !Number.isFinite(flow.challengeCreatedAt.getTime())
        || flow.challengeCreatedAt > serverTime
    ) return false;
    if (!resolved) return flow.consumedAt === null;
    return flow.consumedAt instanceof Date
        && Number.isFinite(flow.consumedAt.getTime())
        && flow.consumedAt >= flow.challengeCreatedAt
        && flow.consumedAt <= serverTime;
}

function throwAuthorityState(authority) {
    if (authority.unavailable) throw authorityUnavailable(authority.reason);
    throw invalidAuthority(authority.reason);
}

function expireSubjectEntitlement(state, subject, serverTime) {
    subject.eligibilityState = 'ineligible';
    subject.eligibilityObservedAt = serverTime;
    subject.eligibilityRevalidateAt = serverTime;
    subject.sessionEpoch += 1;
    subject.updatedAt = serverTime;
    revokeSubjectSessions(state, subject.subjectId, serverTime, 'entitlement-expired');
    revokeLegacySubjectBindings(state, subject.subjectId, serverTime, 'entitlement-expired');
}

function revokeLegacySubjectBindings(state, subjectId, serverTime, reason) {
    for (const binding of state.legacyBindings.values()) {
        if (
            binding.subjectId !== subjectId
            || binding.revokedAt !== null
            || binding.incidentState !== 'normal'
        ) continue;
        binding.revokedAt = serverTime;
        binding.revocationReason = reason;
        binding.incidentState = 'revoked';
    }
}

function hasValidLegacyBinding(binding) {
    if (
        !binding
        || !['normal', 'revoked', 'incident'].includes(binding.incidentState)
        || typeof binding.subjectId !== 'string'
        || !Buffer.isBuffer(binding.verifier)
        || binding.verifier.length !== 32
        || typeof binding.verifierKeyId !== 'string'
        || binding.verifierKeyId.length === 0
        || !(binding.issuedAt instanceof Date)
        || !Number.isFinite(binding.issuedAt.getTime())
        || !(binding.expiresAt instanceof Date)
        || !Number.isFinite(binding.expiresAt.getTime())
        || binding.expiresAt.getTime() - binding.issuedAt.getTime() !== LEGACY_LIFETIME_MS
    ) return false;
    const incidentAt = binding.incidentAt ?? null;
    const incidentCode = binding.incidentCode ?? null;
    if (binding.incidentState === 'normal') {
        return binding.revokedAt === null
            && binding.revocationReason === null
            && incidentAt === null
            && incidentCode === null;
    }
    if (binding.incidentState === 'revoked') {
        return binding.revokedAt instanceof Date
            && Number.isFinite(binding.revokedAt.getTime())
            && isPrivacySafeMachineToken(binding.revocationReason)
            && incidentAt === null
            && incidentCode === null;
    }
    return binding.revokedAt === null
        && binding.revocationReason === null
        && incidentAt instanceof Date
        && Number.isFinite(incidentAt.getTime())
        && isPrivacySafeMachineToken(incidentCode);
}

function isPrivacySafeMachineToken(value) {
    return typeof value === 'string' && /^[a-z0-9][a-z0-9-]{0,127}$/.test(value);
}

function requireSubject(state, subjectId) {
    const subject = state.subjects.get(subjectId);
    if (!subject) throw authorityUnavailable('subject-mapping-integrity');
    return subject;
}

function validateSubjectObservationExpectation(input) {
    if (!(input.observationStartedAt instanceof Date) || !Number.isFinite(input.observationStartedAt.getTime())) {
        throw new TypeError('observationStartedAt must be a valid Date');
    }
    validateSubjectCredentialExpectation(input);
}

function validateSubjectCredentialExpectation(input) {
    if (!Number.isSafeInteger(input.expectedCredentialVersion) || input.expectedCredentialVersion < 1) {
        throw new TypeError('Expected credential version is required');
    }
    if (
        typeof input.expectedCredentialFingerprintKeyId !== 'string'
        || input.expectedCredentialFingerprintKeyId.length === 0
    ) throw new TypeError('Expected credential fingerprint key ID is required');
    if (!Buffer.isBuffer(input.expectedCredentialFingerprint)) {
        throw new TypeError('Expected credential fingerprint must be a Buffer');
    }
}

function requireSubjectObservationExpectation(subject, input) {
    if (
        subject.eligibilityObservedAt > input.observationStartedAt
        || !subjectCredentialExpectationMatches(subject, input)
    ) {
        throw authorityConflict('subject-observation-compare-and-replace');
    }
}

function requireSubjectCredentialExpectation(subject, input) {
    if (!subjectCredentialExpectationMatches(subject, input)) {
        throw authorityConflict('subject-credential-compare-and-replace');
    }
}

function subjectCredentialExpectationMatches(subject, input) {
    return subject.credentialVersion === input.expectedCredentialVersion
        && subject.credentialFingerprintKeyId === input.expectedCredentialFingerprintKeyId
        && sameBuffer(subject.credentialFingerprint, input.expectedCredentialFingerprint);
}

function sameBuffer(left, right) {
    return Buffer.isBuffer(left) && Buffer.isBuffer(right) && left.equals(right);
}

function requireEligibleSubject(subject, serverTime) {
    if (subject.eligibilityState !== 'eligible' || subject.entitlementExpiresAt <= serverTime) {
        throw forbiddenAuthority('ineligible');
    }
}

function requireFreshEligibility(subject, serverTime) {
    if (
        !(subject.eligibilityRevalidateAt instanceof Date)
        || !Number.isFinite(subject.eligibilityRevalidateAt.getTime())
        || subject.eligibilityRevalidateAt <= serverTime
    ) throw authorityUnavailable('eligibility-revalidation-required');
}

function requireAuthorityOpen(control) {
    if (control.incidentState !== 'normal') throw authorityUnavailable('authority-incident');
}

function requireTargetIssuanceControl(control, serverTime) {
    requireAuthorityOpen(control);
    if (!control.targetSessionIssuanceEnabled || !control.subjectTargetAdoptionEnabled) {
        throw authorityUnavailable('target-session-issuance-disabled');
    }
    if (
        !control.legacyCompatibilityEnforcementEnabled
        || !(control.dualStackStartedAt instanceof Date)
        || !(control.hardSunsetAt instanceof Date)
        || control.hardSunsetAt.getTime() - control.dualStackStartedAt.getTime()
            !== LEGACY_SUNSET_MAXIMUM_MS
    ) throw authorityUnavailable('target-session-window-unqualified');
    if (serverTime < control.dualStackStartedAt) {
        throw authorityUnavailable('target-session-window-inactive');
    }
}

function requireLegacyIssuanceControl(control, serverTime) {
    requireAuthorityOpen(control);
    if (control.hardSunsetAt !== null && !(control.hardSunsetAt instanceof Date)) {
        throw authorityUnavailable('control-integrity');
    }
    if (
        !control.legacyIssuanceEnabled
        || !control.legacyAcceptanceEnabled
        || (
            control.hardSunsetAt !== null
            && serverTime >= new Date(control.hardSunsetAt.getTime() - LEGACY_LIFETIME_MS)
        )
    ) throw authorityConflict('legacy-issuance-disabled');
}

function requireLegacyCandidateWithinSunset(control, expiresAt) {
    if (control.hardSunsetAt !== null && expiresAt > control.hardSunsetAt) {
        throw authorityConflict('legacy-issuance-disabled');
    }
}

function requireUnboundLegacyIssuanceControl(control, serverTime) {
    requireLegacyIssuanceControl(control, serverTime);
    if (control.legacyLedgerSeedingEnabled) {
        throw authorityUnavailable('legacy-seeding-gate-mismatch');
    }
    if (
        control.legacyCompatibilityEnforcementEnabled
        || control.subjectTargetAdoptionEnabled
        || control.targetSessionIssuanceEnabled
    ) throw authorityUnavailable('legacy-unbound-issuance-disabled');
}

function validateLegacyLifetime(issuedAt, expiresAt) {
    if (
        !(issuedAt instanceof Date)
        || !Number.isFinite(issuedAt.getTime())
        || !(expiresAt instanceof Date)
        || !Number.isFinite(expiresAt.getTime())
        || expiresAt.getTime() - issuedAt.getTime() !== LEGACY_LIFETIME_MS
    ) throw new TypeError('Legacy issue metadata must describe one exact legacy lifetime');
}

function validateSubjectLoginRemap(input) {
    if (!input || typeof input.subjectId !== 'string' || input.subjectId.length === 0) {
        throw new TypeError('Subject ID is required');
    }
    for (const name of ['expectedLoginLookupKeyId', 'loginLookupKeyId', 'accountMappingKeyId']) {
        if (typeof input[name] !== 'string' || input[name].length === 0) {
            throw new TypeError('Subject mapping key ID is required');
        }
    }
    for (const name of [
        'expectedLoginLookupToken',
        'loginLookupToken',
        'encryptedAccountMapping',
    ]) {
        if (!Buffer.isBuffer(input[name]) || input[name].length === 0) {
            throw new TypeError('Subject mapping value must be a Buffer');
        }
    }
    if (input.expectedLoginLookupToken.length !== 32 || input.loginLookupToken.length !== 32) {
        throw new TypeError('Subject lookup tokens must be 32 bytes');
    }
}

function validateExpectedControlVersion(value) {
    if (!Number.isSafeInteger(value) || value < 1) {
        throw new TypeError('Expected control version must be a positive safe integer');
    }
}

function validateContinuityOwnerId(ownerId) {
    if (
        typeof ownerId !== 'string'
        || !/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/u.test(ownerId)
    ) throw new TypeError('Continuity owner ID must be a canonical random UUID');
}

function seedingHorizonActive(control) {
    return control.legacyLedgerSeedingEnabled === true
        && control.legacyCompatibilityEnforcementEnabled !== true
        && control.legacyIssuanceEnabled === true;
}

function hasLiveSeedingContinuity(control, serverTime) {
    return typeof control.seedingHeartbeatOwnerId === 'string'
        && control.seedingStartedAt instanceof Date
        && control.seedingHeartbeatAt instanceof Date
        && control.seedingLeaseExpiresAt instanceof Date
        && control.seedingHeartbeatAt >= control.seedingStartedAt
        && control.seedingHeartbeatAt <= serverTime
        && control.seedingLeaseExpiresAt > serverTime
        && control.seedingLeaseExpiresAt.getTime() - control.seedingHeartbeatAt.getTime()
            === LEGACY_SEEDING_LEASE_MS;
}

function requireLiveSeedingContinuity(control, serverTime) {
    if (seedingHorizonActive(control) && !hasLiveSeedingContinuity(control, serverTime)) {
        throw authorityUnavailable('legacy-seeding-continuity-unavailable');
    }
}

function requireExpectedNormalControl(control, expectedControlVersion) {
    if (control.incidentState !== 'normal') throw authorityUnavailable('authority-incident');
    if (control.version !== expectedControlVersion) {
        throw authorityUnavailable('authority-control-changed');
    }
}

function requireLoginLookupKeyFence(control, keyId, commitment) {
    if (control.loginLookupKeyId === null && control.loginLookupKeyCommitment === null) {
        throw authorityUnavailable('login-lookup-key-uninitialized');
    }
    if (
        control.loginLookupKeyId !== keyId
        || !Buffer.isBuffer(control.loginLookupKeyCommitment)
        || !control.loginLookupKeyCommitment.equals(commitment)
    ) throw authorityUnavailable('login-lookup-key-mismatch');
}

function validateTestKeyBinding(binding, label) {
    if (
        !binding
        || typeof binding.keyId !== 'string'
        || binding.keyId.length === 0
        || binding.keyId.length > 128
        || !Buffer.isBuffer(binding.commitment)
        || binding.commitment.length !== 32
    ) throw new TypeError(`${label} key binding must be complete`);
}

function validateTestAuthorityKeysetBinding(binding) {
    if (
        !binding
        || !Buffer.isBuffer(binding.commitment)
        || binding.commitment.length !== 32
        || !binding.purposes
    ) throw new TypeError('Authority keyset binding must be complete');
    for (const purpose of [
        'targetVerifier',
        'legacyCompatibility',
        'loginLookup',
        'credentialFingerprint',
        'accountMappingEncryption',
        'faceChallengeEncryption',
    ]) validateTestKeyBinding(binding.purposes[purpose], `Authority ${purpose}`);
}

function copyTestKeyBinding(binding) {
    return {
        keyId: binding.keyId,
        commitment: Buffer.from(binding.commitment),
    };
}

function copyTestAuthorityKeysetBinding(binding) {
    return {
        commitment: Buffer.from(binding.commitment),
        purposes: Object.fromEntries(Object.entries(binding.purposes).map(
            ([purpose, descriptor]) => [purpose, copyTestKeyBinding(descriptor)],
        )),
    };
}

function sameTestKeyBinding(left, right) {
    return Boolean(
        left
        && right
        && left.keyId === right.keyId
        && Buffer.isBuffer(left.commitment)
        && Buffer.isBuffer(right.commitment)
        && left.commitment.equals(right.commitment),
    );
}

function requireTestAuthorityKeysetFence(
    control,
    accountMappingBinding,
    authorityKeysetBinding,
    legacySigningBinding,
    { allowRotatableKeyMismatch = false } = {},
) {
    if (!control.accountMappingKeyBinding || !control.authorityKeysetBinding) {
        throw authorityUnavailable('authority-keyset-uninitialized');
    }
    if (!control.legacySigningKeyBinding) {
        throw authorityUnavailable('legacy-signing-key-uninitialized');
    }
    if (!sameTestKeyBinding(control.accountMappingKeyBinding, accountMappingBinding)) {
        throw authorityUnavailable('authority-keyset-mismatch');
    }
    const stored = control.authorityKeysetBinding;
    if (
        !sameTestKeyBinding(
            stored.purposes?.loginLookup,
            authorityKeysetBinding.purposes.loginLookup,
        )
        || !sameTestKeyBinding(
            stored.purposes?.accountMappingEncryption,
            authorityKeysetBinding.purposes.accountMappingEncryption,
        )
    ) throw authorityUnavailable('authority-keyset-mismatch');
    const mismatches = {
        aggregate: !Buffer.isBuffer(stored.commitment)
            || !stored.commitment.equals(authorityKeysetBinding.commitment),
        targetVerifier: !sameTestKeyBinding(
            stored.purposes?.targetVerifier,
            authorityKeysetBinding.purposes.targetVerifier,
        ),
        legacyCompatibility: !sameTestKeyBinding(
            stored.purposes?.legacyCompatibility,
            authorityKeysetBinding.purposes.legacyCompatibility,
        ),
        credentialFingerprint: !sameTestKeyBinding(
            stored.purposes?.credentialFingerprint,
            authorityKeysetBinding.purposes.credentialFingerprint,
        ),
        faceChallengeEncryption: !sameTestKeyBinding(
            stored.purposes?.faceChallengeEncryption,
            authorityKeysetBinding.purposes.faceChallengeEncryption,
        ),
        legacySigning: !sameTestKeyBinding(control.legacySigningKeyBinding, legacySigningBinding),
    };
    if (Object.values(mismatches).some(Boolean) && !allowRotatableKeyMismatch) {
        throw authorityUnavailable('authority-keyset-mismatch');
    }
    return mismatches;
}

function prepareTestAuthorityKeyRecovery({
    state,
    current,
    changes,
    serverTime,
    expectedAuthorityGeneration,
    keyMismatches,
    expectedAuthorityKeysetBinding,
    expectedLegacySigningKeyBinding,
}) {
    const mismatches = keyMismatches || {
        aggregate: false,
        targetVerifier: false,
        legacyCompatibility: false,
        credentialFingerprint: false,
        faceChallengeEncryption: false,
        legacySigning: false,
    };
    const leafPurposes = [
        'targetVerifier',
        'legacyCompatibility',
        'credentialFingerprint',
        'faceChallengeEncryption',
    ].filter((purpose) => mismatches[purpose]);
    const changedPurposes = [
        ...leafPurposes,
        ...(mismatches.legacySigning ? ['legacySigning'] : []),
    ];
    const anyMismatch = mismatches.aggregate
        || leafPurposes.length > 0
        || mismatches.legacySigning;
    if (!anyMismatch) {
        return { active: false, changedPurposes: [], changes, apply() {} };
    }
    if (
        (mismatches.aggregate || leafPurposes.length > 0)
        && (!mismatches.aggregate || leafPurposes.length === 0)
    ) throw authorityUnavailable('authority-keyset-mismatch');
    const next = { ...current, ...changes };
    if (
        current.authorityGeneration !== expectedAuthorityGeneration - 1
        || current.incidentState !== 'suspended'
        || changes.authorityGeneration !== expectedAuthorityGeneration
        || !Number.isSafeInteger(changes.globalSessionEpoch)
        || changes.globalSessionEpoch <= current.globalSessionEpoch
        || changes.incidentState !== 'recovering'
        || next.targetSessionIssuanceEnabled !== false
        || next.subjectTargetAdoptionEnabled !== false
    ) throw authorityUnavailable('authority-keyset-mismatch');
    const stamped = { ...changes };
    if (
        changedPurposes.includes('targetVerifier')
        && current.targetVerifierKeyIncidentAt === null
    ) {
        stamped.targetVerifierKeyIncidentAt = current.incidentRecordedAt;
    }
    const retireLegacyBindings = changedPurposes.includes('legacyCompatibility')
        || changedPurposes.includes('legacySigning')
        || legacyAuthorityWasInScope(current, next);
    if (retireLegacyBindings) {
        stamped.legacyIssuanceEnabled = false;
        stamped.legacyAcceptanceEnabled = false;
        if (current.legacyIssuanceEnabled) stamped.legacyStopIssuanceAt = serverTime;
        if (current.legacyAcceptanceEnabled) stamped.legacyAcceptanceDisabledAt = serverTime;
        if (current.legacyVerifierKeyIncidentAt === null) {
            stamped.legacyVerifierKeyIncidentAt = current.incidentRecordedAt;
        }
    }
    return {
        active: true,
        changedPurposes,
        changes: stamped,
        apply() {
            for (const flow of state.flows.values()) {
                if (!['creating', 'active', 'reconciliation-required'].includes(
                    flow.challengeState,
                )) continue;
                flow.challengeState = 'reconciliation-required';
                flow.updatedAt = serverTime;
                const session = state.sessions.get(flow.currentSessionId);
                if (session && ACTIVE_PHASES_FOR_EVIDENCE.has(session.phase)) {
                    session.phase = SESSION_PHASES.revoked;
                    session.revokedAt = serverTime;
                    session.revocationReason = 'key-recovery';
                    session.replacementSessionId = null;
                    session.version = 2;
                }
            }
            if (retireLegacyBindings) {
                for (const binding of state.legacyBindings.values()) {
                    if (binding.compatibilityState !== 'active') continue;
                    binding.compatibilityState = 'incident';
                    binding.incidentAt = serverTime;
                    binding.incidentCode = 'legacy-key-recovery';
                    binding.revokedAt = null;
                    binding.revocationReason = null;
                }
            }
        },
    };
}

function requireDormantLookupInitialization(control) {
    const disabled = [
        'targetRoutesEnabled',
        'targetSessionIssuanceEnabled',
        'legacyLedgerSeedingEnabled',
        'legacyCompatibilityEnforcementEnabled',
        'subjectTargetAdoptionEnabled',
    ].every((field) => control[field] === false);
    const absent = [
        'targetSessionIssuanceStartedAt',
        'legacyLedgerSeedingStartedAt',
        'seedingStartedAt',
        'seedingQualifiedAt',
        'seedingHeartbeatOwnerId',
        'seedingHeartbeatAt',
        'seedingLeaseExpiresAt',
        'legacyCompatibilityEnforcedAt',
        'subjectTargetAdoptionStartedAt',
        'dualStackStartedAt',
        'legacyStopIssuanceAt',
        'legacyAcceptanceDisabledAt',
        'hardSunsetAt',
        'incidentRecordedAt',
        'incidentCode',
        'targetVerifierKeyIncidentAt',
        'legacyVerifierKeyIncidentAt',
    ].every((field) => control[field] === null);
    if (
        !disabled
        || !absent
        || control.legacyIssuanceEnabled !== true
        || control.legacyAcceptanceEnabled !== true
        || control.incidentState !== 'normal'
    ) throw authorityUnavailable('login-lookup-key-initialization-blocked');
}

function rotateOut(session, replacementSessionId, serverTime) {
    session.phase = SESSION_PHASES.rotatedOut;
    session.replacementSessionId = replacementSessionId;
    session.revokedAt = null;
    session.revocationReason = null;
    session.version += 1;
}

function migrateFlow(state, predecessorId, replacementId) {
    const flow = state.flows.get(predecessorId);
    if (!flow) return;
    state.flows.delete(predecessorId);
    flow.currentSessionId = replacementId;
    state.flows.set(replacementId, flow);
}

function revokeSubjectSessions(state, subjectId, serverTime, reason) {
    for (const session of state.sessions.values()) {
        if (session.subjectId !== subjectId) continue;
        if (![
            SESSION_PHASES.credentialVerified,
            SESSION_PHASES.registrationPending,
            SESSION_PHASES.facePending,
            SESSION_PHASES.authenticated,
        ].includes(session.phase) || session.revokedAt !== null) continue;
        session.phase = SESSION_PHASES.revoked;
        session.revokedAt = serverTime;
        session.revocationReason = reason;
        session.version += 1;
    }
}

function stampControlTransition(current, changes, serverTime) {
    const stamped = { ...changes };
    const nextIncidentState = Object.hasOwn(changes, 'incidentState')
        ? changes.incidentState
        : current.incidentState;
    if (
        current.incidentState === 'normal'
        && nextIncidentState !== 'normal'
        && seedingHorizonActive(current)
    ) {
        stamped.seedingStartedAt = serverTime;
        stamped.seedingContinuityVersion = current.seedingContinuityVersion + 1;
        stamped.seedingHeartbeatOwnerId = null;
        stamped.seedingHeartbeatAt = null;
        stamped.seedingLeaseExpiresAt = null;
    }
    const nextTargetIssuanceEnabled = Object.hasOwn(changes, 'targetSessionIssuanceEnabled')
        ? changes.targetSessionIssuanceEnabled
        : current.targetSessionIssuanceEnabled;
    const nextSubjectAdoptionEnabled = Object.hasOwn(changes, 'subjectTargetAdoptionEnabled')
        ? changes.subjectTargetAdoptionEnabled
        : current.subjectTargetAdoptionEnabled;
    const firstTargetActivation = current.targetSessionIssuanceStartedAt === null
        && current.subjectTargetAdoptionStartedAt === null
        && current.dualStackStartedAt === null
        && nextTargetIssuanceEnabled === true
        && nextSubjectAdoptionEnabled === true;
    if (firstTargetActivation) {
        stamped.targetSessionIssuanceStartedAt = serverTime;
        stamped.subjectTargetAdoptionStartedAt = serverTime;
        stamped.dualStackStartedAt = serverTime;
        stamped.hardSunsetAt = new Date(serverTime.getTime() + LEGACY_SUNSET_MAXIMUM_MS);
    }
    const nextSeedingEnabled = Object.hasOwn(changes, 'legacyLedgerSeedingEnabled')
        ? changes.legacyLedgerSeedingEnabled
        : current.legacyLedgerSeedingEnabled;
    if (!current.legacyLedgerSeedingEnabled && nextSeedingEnabled) {
        if (current.legacyLedgerSeedingStartedAt === null) {
            stamped.legacyLedgerSeedingStartedAt = serverTime;
        }
        stamped.seedingStartedAt = serverTime;
        stamped.seedingHeartbeatOwnerId = null;
        stamped.seedingHeartbeatAt = null;
        stamped.seedingLeaseExpiresAt = null;
    } else if (
        Object.hasOwn(changes, 'seedingStartedAt')
        && changes.seedingStartedAt !== null
        && !sameInstant(changes.seedingStartedAt, current.seedingStartedAt)
    ) {
        stamped.seedingStartedAt = serverTime;
        stamped.seedingHeartbeatOwnerId = null;
        stamped.seedingHeartbeatAt = null;
        stamped.seedingLeaseExpiresAt = null;
    }

    const nextEnforcementEnabled = Object.hasOwn(changes, 'legacyCompatibilityEnforcementEnabled')
        ? changes.legacyCompatibilityEnforcementEnabled
        : current.legacyCompatibilityEnforcementEnabled;
    if (
        Object.hasOwn(changes, 'seedingQualifiedAt')
        && changes.seedingQualifiedAt !== null
        && !sameInstant(changes.seedingQualifiedAt, current.seedingQualifiedAt)
    ) stamped.seedingQualifiedAt = serverTime;
    if (!current.legacyCompatibilityEnforcementEnabled && nextEnforcementEnabled) {
        if (!(current.seedingQualifiedAt instanceof Date)) stamped.seedingQualifiedAt = serverTime;
        stamped.legacyCompatibilityEnforcedAt = serverTime;
    }

    const nextLegacyIssuanceEnabled = Object.hasOwn(changes, 'legacyIssuanceEnabled')
        ? changes.legacyIssuanceEnabled
        : current.legacyIssuanceEnabled;
    if (current.legacyIssuanceEnabled && !nextLegacyIssuanceEnabled) {
        stamped.legacyStopIssuanceAt = serverTime;
    }
    const nextLegacyAcceptanceEnabled = Object.hasOwn(changes, 'legacyAcceptanceEnabled')
        ? changes.legacyAcceptanceEnabled
        : current.legacyAcceptanceEnabled;
    if (current.legacyAcceptanceEnabled && !nextLegacyAcceptanceEnabled) {
        stamped.legacyAcceptanceDisabledAt = serverTime;
    }
    return stamped;
}

function requireControlTransitionGeneration(
    current,
    changes,
    expectedAuthorityGeneration,
    keyRecoveryActive = false,
) {
    const proposedGeneration = Object.hasOwn(changes, 'authorityGeneration')
        ? changes.authorityGeneration
        : current.authorityGeneration;
    if (current.authorityGeneration === expectedAuthorityGeneration) {
        if (proposedGeneration > expectedAuthorityGeneration) {
            throw authorityUnavailable('authority-generation-mismatch');
        }
        return;
    }
    if (
        current.authorityGeneration === expectedAuthorityGeneration - 1
        && keyRecoveryActive
        && current.incidentState !== 'normal'
        && proposedGeneration === expectedAuthorityGeneration
        && changes.globalSessionEpoch > current.globalSessionEpoch
        && ({ ...current, ...changes }).incidentState === 'recovering'
    ) return;
    throw authorityUnavailable('authority-generation-mismatch');
}

function requireIrreversibleControl(
    current,
    next,
    serverTime,
    changes,
    { keyRecoveryActive = false } = {},
) {
    if (
        next.incidentState === 'recovering'
        && current.incidentState !== 'recovering'
        && !keyRecoveryActive
    ) throw forbiddenAuthority('recovering-requires-key-recovery');
    if (
        next.authorityGeneration < current.authorityGeneration
        || next.globalSessionEpoch < current.globalSessionEpoch
    ) {
        throw forbiddenAuthority('irreversible-authority-epoch');
    }
    const targetKeyIncidentAdvanced = verifierIncidentAdvanced(
        current,
        next,
        'targetVerifierKeyIncidentAt',
        serverTime,
    );
    const legacyKeyIncidentAdvanced = verifierIncidentAdvanced(
        current,
        next,
        'legacyVerifierKeyIncidentAt',
        serverTime,
    );
    if (targetKeyIncidentAdvanced || legacyKeyIncidentAdvanced) {
        if (
            next.incidentState === 'normal'
            || !(next.incidentRecordedAt instanceof Date)
            || typeof next.incidentCode !== 'string'
            || next.incidentCode.length === 0
        ) throw forbiddenAuthority('verifier-key-incident-requires-suspension');
        const latestIncident = targetKeyIncidentAdvanced && legacyKeyIncidentAdvanced
            ? new Date(Math.max(
                next.targetVerifierKeyIncidentAt.getTime(),
                next.legacyVerifierKeyIncidentAt.getTime(),
            ))
            : next[targetKeyIncidentAdvanced
                ? 'targetVerifierKeyIncidentAt'
                : 'legacyVerifierKeyIncidentAt'];
        if (next.incidentRecordedAt < latestIncident || next.incidentRecordedAt > serverTime) {
            throw forbiddenAuthority('verifier-key-incident-time-invalid');
        }
    }
    if (legacyKeyIncidentAdvanced) {
        if (
            next.legacyIssuanceEnabled
            || next.legacyAcceptanceEnabled
            || !(next.legacyStopIssuanceAt instanceof Date)
            || !(next.legacyAcceptanceDisabledAt instanceof Date)
            || next.legacyStopIssuanceAt > next.legacyAcceptanceDisabledAt
            || next.legacyStopIssuanceAt > serverTime
            || next.legacyAcceptanceDisabledAt > serverTime
        ) throw forbiddenAuthority('legacy-key-incident-requires-retirement');
    }
    validateTargetControlTransition(current, next, serverTime);
    validateSeedingControlTransition(current, next, changes, serverTime);
    validateLegacyRetirementControlTransition(
        current,
        next,
        serverTime,
        legacyKeyIncidentAdvanced,
    );
    for (const name of ['legacyCompatibilityEnforcementEnabled']) {
        if (current[name] === true && next[name] === false) {
            throw forbiddenAuthority(`irreversible-${name}`);
        }
    }
    for (const name of ['legacyIssuanceEnabled', 'legacyAcceptanceEnabled']) {
        if (current[name] === false && next[name] === true) {
            throw forbiddenAuthority(`irreversible-${name}`);
        }
    }
    for (const name of [
        'legacyLedgerSeedingStartedAt',
        'legacyCompatibilityEnforcedAt',
        'dualStackStartedAt',
        'legacyStopIssuanceAt',
        'legacyAcceptanceDisabledAt',
        'hardSunsetAt',
    ]) {
        if (current[name] !== null && (next[name] === null || next[name].getTime() !== current[name].getTime())) {
            throw forbiddenAuthority(`irreversible-${name}`);
        }
    }
    for (const name of ['seedingStartedAt', 'seedingQualifiedAt']) {
        if (
            current[name] !== null
            && (next[name] === null || next[name] < current[name])
        ) {
            throw forbiddenAuthority(`irreversible-${name}`);
        }
    }
    if (next.legacyLedgerSeedingEnabled) {
        if (next.legacyLedgerSeedingStartedAt === null || next.seedingStartedAt === null) {
            throw forbiddenAuthority('seeding-start-required');
        }
    }
    if (next.legacyCompatibilityEnforcementEnabled) {
        if (next.seedingQualifiedAt === null || next.legacyCompatibilityEnforcedAt === null) {
            throw forbiddenAuthority('legacy-enforcement-unqualified');
        }
    }
    if (
        (
            next.targetSessionIssuanceEnabled
            || next.subjectTargetAdoptionEnabled
            || next.dualStackStartedAt !== null
        )
        && !next.legacyCompatibilityEnforcementEnabled
    ) throw forbiddenAuthority('legacy-enforcement-required-before-target');
    if (next.targetSessionIssuanceEnabled || next.subjectTargetAdoptionEnabled) {
        if (next.dualStackStartedAt === null || next.hardSunsetAt === null) {
            throw forbiddenAuthority('target-window-required');
        }
        if (
            next.hardSunsetAt.getTime() - next.dualStackStartedAt.getTime()
                !== LEGACY_SUNSET_MAXIMUM_MS
        ) throw forbiddenAuthority('sunset-must-be-fixed');
        if (serverTime < next.dualStackStartedAt) {
            throw forbiddenAuthority('target-window-inactive');
        }
    }
    if (
        next.legacyCompatibilityEnforcementEnabled
        && next.legacyIssuanceEnabled
        && !next.legacyLedgerSeedingEnabled
    ) throw forbiddenAuthority('legacy-seeding-required-during-issuance');
    if (
        next.targetSessionIssuanceEnabled
        && (!next.targetRoutesEnabled || next.targetSessionIssuanceStartedAt === null)
    ) {
        throw forbiddenAuthority('target-issuance-unqualified');
    }
    if (next.subjectTargetAdoptionEnabled && next.subjectTargetAdoptionStartedAt === null) {
        throw forbiddenAuthority('subject-adoption-start-required');
    }
    if (!next.legacyIssuanceEnabled && next.legacyStopIssuanceAt === null) {
        throw forbiddenAuthority('legacy-stop-time-required');
    }
    if (!next.legacyAcceptanceEnabled) {
        if (
            next.legacyIssuanceEnabled
            || next.legacyStopIssuanceAt === null
            || next.legacyAcceptanceDisabledAt === null
            || (
                next.legacyVerifierKeyIncidentAt === null
                && next.legacyAcceptanceDisabledAt.getTime() - next.legacyStopIssuanceAt.getTime()
                    < LEGACY_LIFETIME_MS
            )
        ) {
            throw forbiddenAuthority('legacy-aging-incomplete');
        }
    }
    if (next.hardSunsetAt !== null) {
        if (next.dualStackStartedAt === null) throw forbiddenAuthority('sunset-requires-dual-stack');
        if (next.hardSunsetAt.getTime() - next.dualStackStartedAt.getTime() !== LEGACY_SUNSET_MAXIMUM_MS) {
            throw forbiddenAuthority('sunset-must-be-fixed');
        }
        if (
            next.legacyStopIssuanceAt !== null
            && next.legacyVerifierKeyIncidentAt === null
            && next.hardSunsetAt.getTime() - next.legacyStopIssuanceAt.getTime()
                < LEGACY_LIFETIME_MS
        ) {
            throw forbiddenAuthority('legacy-stop-too-late');
        }
    }
    if (current.incidentState !== 'normal' && next.incidentState === 'normal') {
        if (
            current.incidentState !== 'recovering'
            || next.authorityGeneration !== current.authorityGeneration
            || next.globalSessionEpoch !== current.globalSessionEpoch
        ) throw forbiddenAuthority('incident-resume-requires-fenced-recovery');
        if (
            legacyAuthorityWasInScope(current, next)
            && (next.legacyIssuanceEnabled || next.legacyAcceptanceEnabled)
        ) throw forbiddenAuthority('incident-resume-requires-legacy-retirement');
    }
}

function validateTargetControlTransition(current, next, serverTime) {
    const targetIssuanceActivated = !current.targetSessionIssuanceEnabled
        && next.targetSessionIssuanceEnabled;
    const subjectAdoptionActivated = !current.subjectTargetAdoptionEnabled
        && next.subjectTargetAdoptionEnabled;
    const targetAcceptanceActivated = current.dualStackStartedAt === null
        && next.dualStackStartedAt !== null;
    if (next.targetSessionIssuanceEnabled !== next.subjectTargetAdoptionEnabled) {
        throw forbiddenAuthority('target-activation-pair-required');
    }
    const targetEvidenceAbsent = next.targetSessionIssuanceStartedAt === null
        && next.subjectTargetAdoptionStartedAt === null
        && next.dualStackStartedAt === null
        && next.hardSunsetAt === null;
    const targetEvidenceComplete = next.targetSessionIssuanceStartedAt instanceof Date
        && next.subjectTargetAdoptionStartedAt instanceof Date
        && next.dualStackStartedAt instanceof Date
        && next.hardSunsetAt instanceof Date
        && sameInstant(next.targetSessionIssuanceStartedAt, next.dualStackStartedAt)
        && sameInstant(next.subjectTargetAdoptionStartedAt, next.dualStackStartedAt)
        && next.hardSunsetAt.getTime() - next.dualStackStartedAt.getTime()
            === LEGACY_SUNSET_MAXIMUM_MS;
    if (!targetEvidenceAbsent && !targetEvidenceComplete) {
        throw forbiddenAuthority('target-activation-evidence-integrity');
    }
    if (
        targetEvidenceAbsent
        && (next.targetSessionIssuanceEnabled || next.subjectTargetAdoptionEnabled)
    ) throw forbiddenAuthority('target-window-required');
    if (
        current.targetSessionIssuanceStartedAt === null
        && current.subjectTargetAdoptionStartedAt === null
        && current.dualStackStartedAt === null
        && targetEvidenceComplete
        && (!targetIssuanceActivated || !subjectAdoptionActivated)
    ) throw forbiddenAuthority('target-activation-pair-required');
    if (
        targetIssuanceActivated
        && current.targetSessionIssuanceStartedAt === null
        && (
            !sameInstant(next.targetSessionIssuanceStartedAt, serverTime)
        )
    ) throw forbiddenAuthority('target-issuance-time-provenance');
    if (
        !targetIssuanceActivated
        && current.targetSessionIssuanceStartedAt === null
        && next.targetSessionIssuanceStartedAt !== null
    ) throw forbiddenAuthority('target-issuance-transition-required');
    if (
        subjectAdoptionActivated
        && current.subjectTargetAdoptionStartedAt === null
        && (
            !sameInstant(next.subjectTargetAdoptionStartedAt, serverTime)
        )
    ) throw forbiddenAuthority('subject-adoption-time-provenance');
    if (
        !subjectAdoptionActivated
        && current.subjectTargetAdoptionStartedAt === null
        && next.subjectTargetAdoptionStartedAt !== null
    ) throw forbiddenAuthority('subject-adoption-transition-required');
    if (targetAcceptanceActivated) {
        if (
            !sameInstant(next.dualStackStartedAt, serverTime)
            || !sameInstant(
                next.hardSunsetAt,
                new Date(serverTime.getTime() + LEGACY_SUNSET_MAXIMUM_MS),
            )
        ) throw forbiddenAuthority('target-window-time-provenance');
    } else if (
        current.dualStackStartedAt === null
        && (next.hardSunsetAt !== null || next.dualStackStartedAt !== null)
    ) {
        throw forbiddenAuthority('target-window-transition-required');
    }
}

function validateSeedingControlTransition(current, next, changes, serverTime) {
    const seedingActivated = !current.legacyLedgerSeedingEnabled
        && next.legacyLedgerSeedingEnabled;
    const firstSeedingStart = current.legacyLedgerSeedingStartedAt === null
        && next.legacyLedgerSeedingStartedAt !== null;
    const continuityAdvanced = next.seedingStartedAt instanceof Date
        && Number.isFinite(next.seedingStartedAt.getTime())
        && (
            !(current.seedingStartedAt instanceof Date)
            || next.seedingStartedAt > current.seedingStartedAt
        );
    if (
        firstSeedingStart
        && (!seedingActivated || !sameInstant(next.legacyLedgerSeedingStartedAt, serverTime))
    ) throw forbiddenAuthority('legacy-seeding-time-provenance');
    if (
        continuityAdvanced
        && (!next.legacyLedgerSeedingEnabled || !sameInstant(next.seedingStartedAt, serverTime))
    ) throw forbiddenAuthority('legacy-seeding-time-provenance');
    if (
        seedingActivated
        && (
            !(next.legacyLedgerSeedingStartedAt instanceof Date)
            || !continuityAdvanced
            || !sameInstant(next.seedingStartedAt, serverTime)
        )
    ) throw forbiddenAuthority('legacy-seeding-time-provenance');

    const qualificationAdvanced = Object.hasOwn(changes, 'seedingQualifiedAt')
        && next.seedingQualifiedAt instanceof Date
        && Number.isFinite(next.seedingQualifiedAt.getTime())
        && (
            !(current.seedingQualifiedAt instanceof Date)
            || next.seedingQualifiedAt > current.seedingQualifiedAt
        );
    const enforcementActivated = !current.legacyCompatibilityEnforcementEnabled
        && next.legacyCompatibilityEnforcementEnabled;
    if (
        qualificationAdvanced
        && !sameInstant(next.seedingQualifiedAt, serverTime)
    ) throw forbiddenAuthority('legacy-qualification-time-provenance');
    if (
        enforcementActivated
        && !sameInstant(next.legacyCompatibilityEnforcedAt, serverTime)
    ) throw forbiddenAuthority('legacy-enforcement-time-provenance');
    if (qualificationAdvanced || enforcementActivated) {
        if (
            current.incidentState !== 'normal'
            || next.incidentState !== 'normal'
            ||
            !next.legacyLedgerSeedingEnabled
            || !(current.seedingStartedAt instanceof Date)
            || !hasLiveSeedingContinuity(current, serverTime)
            || continuityAdvanced
            || serverTime.getTime() - current.seedingStartedAt.getTime() < LEGACY_LIFETIME_MS
            || !(next.seedingQualifiedAt instanceof Date)
            || next.seedingQualifiedAt.getTime() - current.seedingStartedAt.getTime()
                < LEGACY_LIFETIME_MS
            || next.seedingQualifiedAt > serverTime
        ) throw forbiddenAuthority('legacy-ledger-not-qualified');
    }
}

function validateLegacyRetirementControlTransition(
    current,
    next,
    serverTime,
    legacyKeyIncidentAdvanced,
) {
    const issuanceStopped = current.legacyIssuanceEnabled && !next.legacyIssuanceEnabled;
    const acceptanceDisabled = current.legacyAcceptanceEnabled && !next.legacyAcceptanceEnabled;
    if (issuanceStopped) {
        if (!sameInstant(next.legacyStopIssuanceAt, serverTime)) {
            throw forbiddenAuthority('legacy-stop-time-provenance');
        }
    } else if (current.legacyStopIssuanceAt === null && next.legacyStopIssuanceAt !== null) {
        throw forbiddenAuthority('legacy-stop-transition-required');
    }
    if (acceptanceDisabled) {
        if (!sameInstant(next.legacyAcceptanceDisabledAt, serverTime)) {
            throw forbiddenAuthority('legacy-acceptance-time-provenance');
        }
        if (
            !legacyKeyIncidentAdvanced
            && (
                current.legacyIssuanceEnabled
                || !(current.legacyStopIssuanceAt instanceof Date)
                || serverTime.getTime() - current.legacyStopIssuanceAt.getTime()
                    < LEGACY_LIFETIME_MS
            )
        ) throw forbiddenAuthority('legacy-final-aging-incomplete');
    } else if (
        current.legacyAcceptanceDisabledAt === null
        && next.legacyAcceptanceDisabledAt !== null
    ) {
        throw forbiddenAuthority('legacy-acceptance-transition-required');
    }
}

function verifierIncidentAdvanced(current, next, field, serverTime) {
    const currentValue = current[field];
    const nextValue = next[field];
    if (nextValue === null) {
        if (currentValue !== null) throw forbiddenAuthority(`irreversible-${field}`);
        return false;
    }
    if (
        !(nextValue instanceof Date)
        || !Number.isFinite(nextValue.getTime())
        || nextValue > serverTime
    ) throw forbiddenAuthority('verifier-key-incident-time-invalid');
    if (currentValue !== null) {
        if (!(currentValue instanceof Date) || nextValue < currentValue) {
            throw forbiddenAuthority(`irreversible-${field}`);
        }
        return nextValue > currentValue;
    }
    return true;
}

function legacyAuthorityWasInScope(current, next) {
    return [current, next].some((control) => (
        control.legacyLedgerSeedingStartedAt !== null
        || control.seedingStartedAt !== null
        || control.seedingQualifiedAt !== null
        || control.legacyCompatibilityEnforcedAt !== null
        || control.dualStackStartedAt !== null
        || control.legacyCompatibilityEnforcementEnabled
    ));
}

function binaryKey(keyId, value) {
    if (typeof keyId !== 'string' || !Buffer.isBuffer(value)) throw new TypeError('Verifier key requires key ID and Buffer');
    return `${keyId}:${value.toString('hex')}`;
}

function copyBuffer(value) {
    return Buffer.isBuffer(value) ? Buffer.from(value) : value;
}

function copyDate(value) {
    return value instanceof Date ? new Date(value.getTime()) : value;
}

function sameInstant(left, right) {
    return left instanceof Date
        && right instanceof Date
        && Number.isFinite(left.getTime())
        && Number.isFinite(right.getTime())
        && left.getTime() === right.getTime();
}

function clone(value) {
    if (value === null || value === undefined) return value;
    if (Buffer.isBuffer(value)) return Buffer.from(value);
    if (value instanceof Date) return new Date(value.getTime());
    if (Array.isArray(value)) return value.map(clone);
    if (value instanceof Map) return new Map(Array.from(value, ([key, item]) => [key, clone(item)]));
    if (typeof value === 'object') {
        return Object.fromEntries(Object.entries(value)
            .filter(([key]) => ![
                'loginLookupKeyCommitment',
                'accountMappingKeyBinding',
                'authorityKeysetBinding',
                'legacySigningKeyBinding',
            ].includes(key))
            .map(([key, item]) => [key, clone(item)]));
    }
    return value;
}

module.exports = {
    createTestSessionAuthorityBacking,
    createTestSessionAuthorityStore,
};
