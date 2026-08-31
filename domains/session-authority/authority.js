'use strict';

const { timingSafeEqual } = require('node:crypto');
const {
    NEXT_OPERATION_ROLES,
    SESSION_PHASES,
} = require('./constants');
const {
    CRYPTOGRAPHIC_PURPOSES,
    createLoginLookup,
    createOpaqueIdentifier,
    createVerifier,
    decryptPrivateValue,
    encryptPrivateValue,
    parseOpaqueIdentifier,
} = require('./cryptography');
const {
    normalizeEligibility,
    readCredentialAccount,
    readMappedAccount,
} = require('./account-authority');
const {
    authorityConflict,
    authorityUnavailable,
    forbiddenAuthority,
    invalidAuthority,
    isSessionAuthorityError,
} = require('./errors');

const RUNTIME_CONTROL_NAMES = Object.freeze([
    'durableStoreRequired',
    'targetRoutesEnabled',
    'targetSessionIssuanceEnabled',
    'legacyLedgerSeedingEnabled',
    'legacyCompatibilityEnforcementEnabled',
    'subjectTargetAdoptionEnabled',
    'protectedRoutesEnabled',
]);
const SUBJECT_REVOCATION_REASONS = new Set(['administrator', 'identifier-leak']);

function createSessionAuthority({
    store,
    accountSource,
    faceSource,
    legacyHandleAuthority,
    keys,
    runtimeControls = {},
    randomBytes,
    createSubjectId,
    createSessionId,
    createFlowId,
    createCorrelationId,
    formatLegacyAccessDate,
}) {
    validateDependencies({
        store,
        accountSource,
        faceSource,
        legacyHandleAuthority,
        keys,
        createSubjectId,
        createSessionId,
        createFlowId,
        createCorrelationId,
        formatLegacyAccessDate,
    });

    const controls = Object.freeze(Object.fromEntries(
        RUNTIME_CONTROL_NAMES.map((name) => [name, runtimeControls[name] === true]),
    ));

    async function loginTarget({ login, password, presentedIdentifier }) {
        requireRuntimeControl('targetRoutesEnabled');
        requireRuntimeControl('targetSessionIssuanceEnabled');
        requireRuntimeControl('subjectTargetAdoptionEnabled');
        requireRuntimeControl('legacyCompatibilityEnforcementEnabled');

        const { control, serverTime: observationStartedAt } = await store.readControl();
        requireAuthorityControlOpen(control);
        requireDatabaseControl(control, 'targetRoutesEnabled');
        requireDatabaseControl(control, 'targetSessionIssuanceEnabled');
        requireDatabaseControl(control, 'subjectTargetAdoptionEnabled');
        requireDatabaseControl(control, 'legacyCompatibilityEnforcementEnabled');

        const account = await readCredentialAccountFromSource(login, password);
        const faceRequired = account.faceAuthRequired;
        const subject = await resolveEligibleSubject(
            account,
            observationStartedAt,
            control.version,
        );
        const predecessor = await inspectLoginPredecessor(presentedIdentifier);
        const phase = faceRequired ? SESSION_PHASES.credentialVerified : SESSION_PHASES.authenticated;
        const issued = await issueSession({
            subjectId: subject.subjectId,
            expectedCredentialVersion: subject.credentialVersion,
            expectedCredentialFingerprintKeyId: subject.credentialFingerprintKeyId,
            expectedCredentialFingerprint: subject.credentialFingerprint,
            phase,
            faceRequired,
            registrationRequired: faceRequired && account.photoRegistrationStatus !== 'Sim',
            predecessor,
        });

        return {
            status: 200,
            body: createCurrentStatus(issued),
            issuance: createIssuance(issued),
        };
    }

    async function loginLegacyWithSeeding({ login, password }) {
        requireRuntimeControl('durableStoreRequired');
        const { control, serverTime: observationStartedAt } = await store.readControl();
        requireAuthorityControlOpen(control);
        const account = await readCredentialAccountFromSource(login, password, { legacy: true });

        const legacyBody = createLegacyLoginBody(account);
        if (account.accountStatus !== 'Ativo') return { status: 200, body: legacyBody };

        if (
            Boolean(control.legacyLedgerSeedingEnabled)
            !== controls.legacyLedgerSeedingEnabled
        ) {
            throw authorityUnavailable('legacy-seeding-gate-mismatch');
        }

        let subject;
        if (control.legacyLedgerSeedingEnabled) {
            subject = await resolveEligibleSubject(
                account,
                observationStartedAt,
                control.version,
                {
                    permitLegacyCompatibilityObservation:
                        !control.legacyCompatibilityEnforcementEnabled,
                },
            );
            if (subject.legacyAuthorityDisabledAt !== null) {
                throw authorityConflict('target-authority-established');
            }
        }

        const issuanceControl = await store.readControl();
        requireAuthorityControlOpen(issuanceControl.control);
        if (
            Boolean(issuanceControl.control.legacyLedgerSeedingEnabled)
            !== controls.legacyLedgerSeedingEnabled
        ) {
            throw authorityUnavailable('legacy-seeding-gate-mismatch');
        }

        const rawHandle = legacyHandleAuthority.createHandle(
            account.rowIndex,
            issuanceControl.serverTime.getTime(),
        );
        const metadata = legacyHandleAuthority.inspectHandle(
            rawHandle,
            issuanceControl.serverTime.getTime(),
        );
        if (!issuanceControl.control.legacyLedgerSeedingEnabled) {
            const lookup = createLoginLookup(account.exactLogin, keys.loginLookup);
            await store.admitUnboundLegacyIssuance({
                loginLookupKeyId: lookup.keyId,
                loginLookupToken: lookup.token,
                issuedAt: metadata.issuedAt,
                expiresAt: metadata.expiresAt,
            });
            legacyBody.IndexVerificado = rawHandle;
            return { status: 200, body: legacyBody };
        }

        const verified = createVerifier(
            rawHandle,
            keys.legacyCompatibility,
            CRYPTOGRAPHIC_PURPOSES.legacyCompatibility,
        );
        await store.bindLegacy({
            legacyCompatibilityId: createFlowId(),
            subjectId: subject.subjectId,
            expectedCredentialVersion: subject.credentialVersion,
            expectedCredentialFingerprintKeyId: subject.credentialFingerprintKeyId,
            expectedCredentialFingerprint: subject.credentialFingerprint,
            verifierKeyId: verified.keyId,
            verifier: verified.verifier,
            issuedAt: metadata.issuedAt,
            expiresAt: metadata.expiresAt,
        });
        legacyBody.IndexVerificado = rawHandle;
        return { status: 200, body: legacyBody };
    }

    async function registrationEnrollment(rawIdentifier) {
        const authority = await authorizeCurrent(rawIdentifier, {
            allowedPhases: [SESSION_PHASES.credentialVerified, SESSION_PHASES.registrationPending],
            revalidate: true,
        });
        if (authority.session.phase === SESSION_PHASES.registrationPending) {
            return { status: 204, body: undefined };
        }
        if (!authority.session.registrationRequired) throw forbiddenAuthority('registration-not-required');
        const rotated = await rotate(authority, SESSION_PHASES.registrationPending);
        return { status: 204, body: undefined, issuance: createIssuance(rotated) };
    }

    async function createExistingPhotoChallenge(rawIdentifier) {
        const authority = await authorizeCurrent(rawIdentifier, {
            allowedPhases: [SESSION_PHASES.credentialVerified, SESSION_PHASES.facePending],
            revalidate: true,
            resolveAccount: true,
        });
        if (authority.session.phase === SESSION_PHASES.facePending) {
            throw authorityConflict('face-challenge-active');
        }
        if (!authority.session.faceRequired || authority.session.registrationRequired) {
            throw forbiddenAuthority('face-challenge-not-allowed');
        }
        return createAndBindFaceChallenge({
            authority,
            registration: false,
            referenceImage: await callAccountSource(
                () => accountSource.downloadReferencePhoto(authority.account.rowIndex),
                'reference-photo-unavailable',
            ),
        });
    }

    async function createRegistrationChallenge(rawIdentifier, referenceImage) {
        const authority = await authorizeCurrent(rawIdentifier, {
            allowedPhases: [SESSION_PHASES.registrationPending, SESSION_PHASES.facePending],
            revalidate: true,
            resolveAccount: true,
        });
        if (authority.session.phase === SESSION_PHASES.facePending) {
            throw authorityConflict('face-challenge-active');
        }
        if (!Buffer.isBuffer(referenceImage) || referenceImage.length === 0) {
            throw forbiddenAuthority('reference-photo-required');
        }

        const reservation = await reserveFaceFlow(authority, 'enrollment-accepted');
        try {
            await accountSource.uploadReferencePhoto(authority.account.rowIndex, referenceImage);
            await accountSource.markPhotoRegistered(authority.account.rowIndex);
            return await createProviderChallengeAndBind(
                authority,
                referenceImage,
                reservation.flowId,
            );
        } catch (error) {
            await markReconciliationRequired(reservation.flowId, true);
            throw normalizeExternalFailure(error, 'registration-reconciliation-required');
        }
    }

    async function completeFace(rawIdentifier) {
        const authority = await authorizeCurrent(rawIdentifier, {
            allowedPhases: [SESSION_PHASES.facePending, SESSION_PHASES.authenticated],
            revalidate: true,
        });
        if (authority.session.phase === SESSION_PHASES.authenticated) {
            if (!authority.session.faceRequired) {
                throw forbiddenAuthority('face-completion-not-applicable');
            }
            const completed = await store.readFaceFlow({
                expectedSessionId: authority.session.sessionId,
                expectedVersion: authority.session.version,
                allowConsumed: true,
            });
            return { status: 200, body: createCurrentStatus(completed) };
        }

        const flowResult = await store.readFaceFlow({
            expectedSessionId: authority.session.sessionId,
            expectedVersion: authority.session.version,
        });
        let challengeReference;
        try {
            challengeReference = decryptPrivateValue(
                flowResult.flow.encryptedChallenge,
                keyForStoredValue(keys.faceChallengeEncryption, flowResult.flow.challengeKeyId),
                CRYPTOGRAPHIC_PURPOSES.faceChallenge,
                createFaceChallengeBinding(
                    flowResult.flow.flowId,
                    flowResult.subject.subjectId,
                    flowResult.session.sessionId,
                ),
            );
        } catch {
            throw authorityUnavailable('face-challenge-integrity');
        }

        let result;
        try {
            result = await faceSource.readLivenessSessionResult(challengeReference);
        } catch (error) {
            throw authorityUnavailable('face-provider-unavailable', { cause: error });
        }

        if (result && result.providerState === 'pending') {
            throw authorityConflict('face-result-pending');
        }
        if (!isDefinitiveFaceResult(result)) {
            throw authorityUnavailable('face-provider-response-invalid');
        }
        if (!isPassingFaceResult(result)) {
            await store.completeFaceFailure({
                expectedSessionId: authority.session.sessionId,
                expectedVersion: authority.session.version,
            });
            throw forbiddenAuthority('face-factor-failed');
        }

        const candidate = createSessionCandidate();
        const completed = await store.completeFaceSuccessAndRotate({
            expectedSessionId: authority.session.sessionId,
            expectedVersion: authority.session.version,
            sessionId: candidate.sessionId,
            verifierKeyId: candidate.verifierKeyId,
            verifier: candidate.verifier,
        });
        return {
            status: 200,
            body: createCurrentStatus(completed),
            issuance: { identifier: candidate.identifier, expiresAt: completed.session.expiresAt, serverTime: completed.serverTime },
        };
    }

    async function current(rawIdentifier) {
        const authority = await authorizeCurrent(rawIdentifier, { revalidate: true });
        return { status: 200, body: createCurrentStatus(authority) };
    }

    async function logout(rawIdentifier) {
        let parsed;
        try {
            parsed = parseOpaqueIdentifier(rawIdentifier);
        } catch {
            return { status: 204, body: undefined };
        }
        const verified = createVerifier(
            parsed,
            keys.targetVerifier,
            CRYPTOGRAPHIC_PURPOSES.targetSession,
        );
        await store.logout({ verifierKeyId: verified.keyId, verifier: verified.verifier });
        return { status: 204, body: undefined };
    }

    async function revokeAll(rawIdentifier, reason = 'user-revoke-all') {
        const authority = await authorizeCurrent(rawIdentifier, {
            allowedPhases: [SESSION_PHASES.authenticated],
            revalidate: true,
        });
        await store.revokeAll({
            expectedSessionId: authority.session.sessionId,
            expectedVersion: authority.session.version,
            reason,
        });
        return { status: 204, body: undefined };
    }

    async function revokeSubject(subjectId, reason) {
        if (typeof subjectId !== 'string' || subjectId.length === 0) {
            throw new TypeError('A subject identifier is required');
        }
        if (!SUBJECT_REVOCATION_REASONS.has(reason)) {
            throw new TypeError('The subject revocation reason is invalid');
        }
        return store.revokeSubject({ subjectId, reason });
    }

    async function revokeLeakedLegacyAuthority(subjectId) {
        if (typeof subjectId !== 'string' || subjectId.length === 0) {
            throw new TypeError('A subject identifier is required');
        }
        return store.disableLegacyAuthority({
            subjectId,
            reason: 'legacy-handle-leak',
        });
    }

    async function remapSubjectLogin({ subjectId, expectedExactLogin, newExactLogin } = {}) {
        requireRuntimeControl('durableStoreRequired');
        if (typeof subjectId !== 'string' || subjectId.length === 0) {
            throw new TypeError('A subject identifier is required');
        }
        if (
            typeof expectedExactLogin !== 'string'
            || expectedExactLogin.length === 0
            || typeof newExactLogin !== 'string'
            || newExactLogin.length === 0
            || expectedExactLogin === newExactLogin
        ) {
            throw new TypeError('Distinct exact login mappings are required');
        }

        const expectedLookup = createLoginLookup(expectedExactLogin, keys.loginLookup);
        const nextLookup = createLoginLookup(newExactLogin, keys.loginLookup);
        const encrypted = encryptPrivateValue(
            newExactLogin,
            keys.accountMappingEncryption,
            CRYPTOGRAPHIC_PURPOSES.accountMapping,
            undefined,
            createAccountMappingBinding(subjectId),
        );
        const result = await store.remapSubjectLogin({
            subjectId,
            expectedLoginLookupKeyId: expectedLookup.keyId,
            expectedLoginLookupToken: expectedLookup.token,
            loginLookupKeyId: nextLookup.keyId,
            loginLookupToken: nextLookup.token,
            encryptedAccountMapping: encrypted.ciphertext,
            accountMappingKeyId: encrypted.keyId,
        });
        if (
            !result
            || !result.subject
            || String(result.subject.subjectId).toLowerCase() !== subjectId.toLowerCase()
            || result.subject.loginLookupKeyId !== nextLookup.keyId
            || !safeEqual(result.subject.loginLookupToken, nextLookup.token)
            || !(result.serverTime instanceof Date)
            || !Number.isFinite(result.serverTime.getTime())
        ) throw authorityUnavailable('subject-mapping-integrity');
        let mappedLogin;
        try {
            mappedLogin = decryptPrivateValue(
                result.subject.encryptedAccountMapping,
                keyForStoredValue(
                    keys.accountMappingEncryption,
                    result.subject.accountMappingKeyId,
                ),
                CRYPTOGRAPHIC_PURPOSES.accountMapping,
                createAccountMappingBinding(result.subject.subjectId),
            );
        } catch {
            throw authorityUnavailable('subject-mapping-integrity');
        }
        if (mappedLogin !== newExactLogin) {
            throw authorityUnavailable('subject-mapping-conflict');
        }

        return Object.freeze({
            idempotent: result.idempotent === true,
            serverTime: new Date(result.serverTime.getTime()),
        });
    }

    async function authorizeProtected(rawIdentifier, independentlyOwnedSubjectId) {
        const authority = await authorizeCurrent(rawIdentifier, {
            allowedPhases: [SESSION_PHASES.authenticated],
            revalidate: true,
            resolveAccount: true,
        });
        if (
            independentlyOwnedSubjectId !== undefined
            && independentlyOwnedSubjectId !== authority.subject.subjectId
        ) {
            throw forbiddenAuthority('wrong-subject');
        }
        return {
            subjectId: authority.subject.subjectId,
            platformRowIndex: authority.account.rowIndex,
            sessionId: authority.session.sessionId,
        };
    }

    async function authorizeLegacy(rawHandle) {
        requireRuntimeControl('durableStoreRequired');
        const inspectionControl = await store.readControl();
        requireAuthorityControlOpen(inspectionControl.control);
        let metadata;
        try {
            metadata = legacyHandleAuthority.inspectHandle(
                rawHandle,
                inspectionControl.serverTime.getTime(),
            );
        } catch {
            throw invalidAuthority('invalid-legacy-handle');
        }
        const verified = createVerifier(
            rawHandle,
            keys.legacyCompatibility,
            CRYPTOGRAPHIC_PURPOSES.legacyCompatibility,
        );
        const authority = await store.authorizeLegacy({
            verifierKeyId: verified.keyId,
            verifier: verified.verifier,
            issuedAt: metadata.issuedAt,
            expiresAt: metadata.expiresAt,
        });
        if (
            !authority.control.legacyCompatibilityEnforcementEnabled
            && controls.legacyCompatibilityEnforcementEnabled
        ) {
            throw authorityUnavailable('legacy-enforcement-gate-mismatch');
        }
        if (authority.unbound === true) {
            if (authority.control.legacyCompatibilityEnforcementEnabled) {
                throw authorityUnavailable('legacy-authorization-integrity');
            }
            return { platformRowIndex: metadata.rowIndex };
        }
        const account = await revalidateSubject(authority, { resolveAccount: true, legacy: true });
        return {
            subjectId: authority.subject.subjectId,
            platformRowIndex: account.rowIndex,
        };
    }

    async function authorizeCurrent(rawIdentifier, options = {}) {
        requireRuntimeControl('targetRoutesEnabled');
        const parsed = parseOpaqueIdentifier(rawIdentifier);
        const verified = createVerifier(
            parsed,
            keys.targetVerifier,
            CRYPTOGRAPHIC_PURPOSES.targetSession,
        );
        let authority = await store.readSession({
            verifierKeyId: verified.keyId,
            verifier: verified.verifier,
        });
        requireDatabaseControl(authority.control, 'targetRoutesEnabled');
        if (options.allowedPhases && !options.allowedPhases.includes(authority.session.phase)) {
            throw forbiddenAuthority('wrong-phase');
        }
        if (options.revalidate || options.resolveAccount) {
            const account = await revalidateSubject(authority, options);
            if (account) authority = { ...authority, account };
        }
        return authority;
    }

    async function inspectLoginPredecessor(rawIdentifier) {
        if (rawIdentifier === undefined || rawIdentifier === null || rawIdentifier === '') return null;
        let parsed;
        try {
            parsed = parseOpaqueIdentifier(rawIdentifier);
        } catch {
            return null;
        }
        const verified = createVerifier(
            parsed,
            keys.targetVerifier,
            CRYPTOGRAPHIC_PURPOSES.targetSession,
        );
        const predecessor = await store.inspectLoginPredecessor({
            verifierKeyId: verified.keyId,
            verifier: verified.verifier,
        });
        if (predecessor.kind !== 'active') return null;
        return {
            expectedSessionId: predecessor.expectedSessionId,
            expectedVersion: predecessor.expectedVersion,
        };
    }

    async function issueSession(input) {
        const candidate = createSessionCandidate();
        const issued = await store.issueSession({
            ...input,
            sessionId: candidate.sessionId,
            verifierKeyId: candidate.verifierKeyId,
            verifier: candidate.verifier,
        });
        return { ...issued, identifier: candidate.identifier };
    }

    async function rotate(authority, phase) {
        const candidate = createSessionCandidate();
        const rotated = await store.rotateSession({
            expectedSessionId: authority.session.sessionId,
            expectedVersion: authority.session.version,
            allowedPhases: [authority.session.phase],
            phase,
            sessionId: candidate.sessionId,
            verifierKeyId: candidate.verifierKeyId,
            verifier: candidate.verifier,
        });
        return { ...rotated, identifier: candidate.identifier };
    }

    async function createAndBindFaceChallenge({ authority, registration, referenceImage }) {
        const reservation = await reserveFaceFlow(
            authority,
            registration ? 'enrollment-accepted' : 'registered',
        );
        try {
            return await createProviderChallengeAndBind(
                authority,
                referenceImage,
                reservation.flowId,
            );
        } catch (error) {
            await markReconciliationRequired(reservation.flowId, registration);
            throw normalizeExternalFailure(error, 'face-challenge-reconciliation-required');
        }
    }

    async function reserveFaceFlow(authority, registrationState) {
        const flowId = createFlowId();
        let reserved;
        try {
            reserved = await store.reserveFaceFlow({
                expectedSessionId: authority.session.sessionId,
                expectedVersion: authority.session.version,
                allowedPhases: [authority.session.phase],
                flowId,
                registrationState,
            });
        } catch (error) {
            if (
                isSessionAuthorityError(error)
                && error.reason === 'transaction-outcome-unknown'
            ) {
                await markReconciliationRequired(flowId, false);
            }
            throw error;
        }
        if (!reserved || reserved.flowId !== flowId) {
            throw authorityUnavailable('face-flow-integrity');
        }
        return { flowId };
    }

    async function createProviderChallengeAndBind(authority, referenceImage, flowId) {
        const provider = await faceSource.createLivenessSession(referenceImage, createCorrelationId());
        if (
            !provider
            || typeof provider.authToken !== 'string'
            || provider.authToken.length === 0
            || typeof provider.privateChallengeId !== 'string'
            || provider.privateChallengeId.length === 0
        ) {
            throw authorityUnavailable('face-provider-response-invalid');
        }
        const candidate = createSessionCandidate();
        const encrypted = encryptPrivateValue(
            provider.privateChallengeId,
            keys.faceChallengeEncryption,
            CRYPTOGRAPHIC_PURPOSES.faceChallenge,
            undefined,
            createFaceChallengeBinding(
                flowId,
                authority.subject.subjectId,
                candidate.sessionId,
            ),
        );
        const bound = await store.bindFaceChallengeAndRotate({
            expectedSessionId: authority.session.sessionId,
            expectedVersion: authority.session.version,
            allowedPhases: [authority.session.phase],
            sessionId: candidate.sessionId,
            verifierKeyId: candidate.verifierKeyId,
            verifier: candidate.verifier,
            flowId,
            challengeKeyId: encrypted.keyId,
            encryptedChallenge: encrypted.ciphertext,
        });
        return {
            status: 200,
            body: { Azure_Face_API_LivenessSession_authToken: provider.authToken },
            issuance: { identifier: candidate.identifier, expiresAt: bound.session.expiresAt, serverTime: bound.serverTime },
        };
    }

    async function markReconciliationRequired(flowId, registrationReconciliationRequired) {
        try {
            await store.markFaceFlowReconciliation({
                flowId,
                registrationReconciliationRequired,
            });
        } catch (error) {
            throw authorityUnavailable('face-flow-reconciliation-unavailable', { cause: error });
        }
    }

    async function revalidateSubject(authority, options) {
        const observationStartedAt = authority.serverTime;
        const due = authority.serverTime >= authority.subject.eligibilityRevalidateAt;
        if (!due && !options.resolveAccount) return null;

        const exactLogin = readVerifiedMappedLogin(authority.subject);
        const rows = await readAccountRows(options.legacy === true);
        let account;
        try {
            account = readMappedAccount(rows, exactLogin, {
                credentialFingerprintKey: keys.credentialFingerprint,
            });
        } catch (error) {
            if (due) throw normalizeExternalFailure(error, 'eligibility-source-unavailable');
            throw normalizeExternalFailure(error, 'account-mapping-unavailable');
        }

        const fingerprint = account.credentialFingerprint;
        if (
            authority.subject.credentialFingerprintKeyId !== fingerprint.keyId
            || !safeEqual(authority.subject.credentialFingerprint, fingerprint.fingerprint)
        ) {
            await store.revokeForCredentialChange({
                subjectId: authority.subject.subjectId,
                observationStartedAt,
                expectedCredentialVersion: authority.subject.credentialVersion,
                expectedCredentialFingerprintKeyId: authority.subject.credentialFingerprintKeyId,
                expectedCredentialFingerprint: authority.subject.credentialFingerprint,
                expectedControlVersion: authority.control.version,
                credentialFingerprintKeyId: fingerprint.keyId,
                credentialFingerprint: fingerprint.fingerprint,
            });
            throw invalidAuthority('credential-changed');
        }

        if (due) {
            let eligibility;
            try {
                eligibility = normalizeEligibility(account, authority.serverTime);
            } catch (error) {
                if (isSessionAuthorityError(error) && error.errorClass === 'forbidden-authority') {
                    await store.revokeForIneligibility({
                        subjectId: authority.subject.subjectId,
                        observationStartedAt,
                        expectedCredentialVersion: authority.subject.credentialVersion,
                        expectedCredentialFingerprintKeyId: authority.subject.credentialFingerprintKeyId,
                        expectedCredentialFingerprint: authority.subject.credentialFingerprint,
                        expectedControlVersion: authority.control.version,
                        eligibilityState: 'ineligible',
                        entitlementExpiresAt: error.entitlementExpiresAt || authority.serverTime,
                        reason: error.reason,
                    });
                    throw forbiddenAuthority('ineligible');
                }
                throw normalizeExternalFailure(error, 'eligibility-source-unavailable');
            }
            const updated = await store.updateEligibility({
                subjectId: authority.subject.subjectId,
                observationStartedAt,
                expectedCredentialVersion: authority.subject.credentialVersion,
                expectedCredentialFingerprintKeyId: authority.subject.credentialFingerprintKeyId,
                expectedCredentialFingerprint: authority.subject.credentialFingerprint,
                expectedControlVersion: authority.control.version,
                rowHint: account.rowIndex,
                ...eligibility,
            });
            requireEligibleUpdate(updated);
            authority.subject = updated.subject;
            authority.serverTime = updated.serverTime;
        }
        return account;
    }

    async function resolveEligibleSubject(
        account,
        observationStartedAt,
        expectedControlVersion,
        { permitLegacyCompatibilityObservation = false } = {},
    ) {
        let eligibility;
        try {
            eligibility = normalizeEligibility(account, observationStartedAt);
        } catch (error) {
            if (
                permitLegacyCompatibilityObservation
                && isSessionAuthorityError(error)
                && error.errorClass === 'forbidden-authority'
            ) {
                eligibility = {
                    eligibilityObservedAt: observationStartedAt,
                    eligibilityRevalidateAt: observationStartedAt,
                    eligibilityState: error.reason === 'entitlement-expired'
                        ? 'ineligible'
                        : 'unknown',
                    entitlementExpiresAt: error.entitlementExpiresAt || null,
                };
            } else {
            const lookup = createLoginLookup(account.exactLogin, keys.loginLookup);
            const existing = await store.readSubjectByLookup({
                loginLookupKeyId: lookup.keyId,
                loginLookupToken: lookup.token,
                expectedControlVersion,
            });
            if (existing.subject) {
                readVerifiedMappedLogin(existing.subject, account.exactLogin);
                await store.revokeForIneligibility({
                    subjectId: existing.subject.subjectId,
                    observationStartedAt,
                    expectedCredentialVersion: existing.subject.credentialVersion,
                    expectedCredentialFingerprintKeyId: existing.subject.credentialFingerprintKeyId,
                    expectedCredentialFingerprint: existing.subject.credentialFingerprint,
                    expectedControlVersion,
                    eligibilityState: 'ineligible',
                    entitlementExpiresAt: error.entitlementExpiresAt || observationStartedAt,
                    reason: error.reason,
                });
            }
            if (isSessionAuthorityError(error)) throw error;
            throw forbiddenAuthority('ineligible');
            }
        }

        const lookup = createLoginLookup(account.exactLogin, keys.loginLookup);
        const candidateSubjectId = createSubjectId();
        const encrypted = encryptPrivateValue(
            account.exactLogin,
            keys.accountMappingEncryption,
            CRYPTOGRAPHIC_PURPOSES.accountMapping,
            undefined,
            createAccountMappingBinding(candidateSubjectId),
        );
        const fingerprint = account.credentialFingerprint;
        let result = await store.createOrLoadSubject({
            subjectId: candidateSubjectId,
            loginLookupKeyId: lookup.keyId,
            loginLookupToken: lookup.token,
            accountMappingKeyId: encrypted.keyId,
            encryptedAccountMapping: encrypted.ciphertext,
            rowHint: account.rowIndex,
            credentialFingerprintKeyId: fingerprint.keyId,
            credentialFingerprint: fingerprint.fingerprint,
            expectedControlVersion,
            ...eligibility,
        });

        readVerifiedMappedLogin(result.subject, account.exactLogin);

        if (permitLegacyCompatibilityObservation) return result.subject;

        if (
            result.subject.credentialFingerprintKeyId !== fingerprint.keyId
            || !safeEqual(result.subject.credentialFingerprint, fingerprint.fingerprint)
        ) {
            result = await store.revokeForCredentialChange({
                subjectId: result.subject.subjectId,
                observationStartedAt,
                expectedCredentialVersion: result.subject.credentialVersion,
                expectedCredentialFingerprintKeyId: result.subject.credentialFingerprintKeyId,
                expectedCredentialFingerprint: result.subject.credentialFingerprint,
                expectedControlVersion,
                credentialFingerprintKeyId: fingerprint.keyId,
                credentialFingerprint: fingerprint.fingerprint,
            });
        }
        result = await store.updateEligibility({
            subjectId: result.subject.subjectId,
            observationStartedAt,
            expectedCredentialVersion: result.subject.credentialVersion,
            expectedCredentialFingerprintKeyId: result.subject.credentialFingerprintKeyId,
            expectedCredentialFingerprint: result.subject.credentialFingerprint,
            expectedControlVersion,
            rowHint: account.rowIndex,
            ...eligibility,
        });
        requireEligibleUpdate(result);
        return result.subject;
    }

    async function readCredentialAccountFromSource(login, password, options = {}) {
        const legacyLogin = options.legacy === true;
        const rows = await readAccountRows(
            legacyLogin,
            legacyLogin
                ? 'legacy-platform-data-read-failed'
                : 'eligibility-source-unavailable',
        );
        return readCredentialAccount(rows, {
            login,
            password,
            credentialFingerprintKey: keys.credentialFingerprint,
        });
    }

    async function readAccountRows(legacy = false, failureReason = 'eligibility-source-unavailable') {
        const reader = legacy && typeof accountSource.readRowsLegacy === 'function'
            ? accountSource.readRowsLegacy
            : accountSource.readRows;
        return callAccountSource(() => reader.call(accountSource), failureReason);
    }

    async function callAccountSource(operation, reason) {
        try {
            const result = await operation();
            if (!Array.isArray(result) && reason === 'eligibility-source-unavailable') {
                throw new TypeError('Account source did not return rows');
            }
            return result;
        } catch (error) {
            if (isSessionAuthorityError(error)) throw error;
            throw authorityUnavailable(reason, { cause: error });
        }
    }

    function createSessionCandidate() {
        const identifier = createOpaqueIdentifier(randomBytes);
        const verified = createVerifier(
            identifier,
            keys.targetVerifier,
            CRYPTOGRAPHIC_PURPOSES.targetSession,
        );
        return {
            identifier,
            sessionId: createSessionId(),
            verifierKeyId: verified.keyId,
            verifier: verified.verifier,
        };
    }

    function createLegacyLoginBody(account) {
        return {
            Usuário_Status_FaceID: account.faceAuthRequired ? 'Ativo' : 'Inativo',
            Usuário_Foto_Cadastrada: account.photoRegistrationStatus,
            Usuário_PrazoAcesso: formatLegacyAccessDate(account.accessDateSerial),
            Usuário_Status_Login: account.accountStatus,
        };
    }

    function readVerifiedMappedLogin(subject, expectedExactLogin) {
        let exactLogin;
        try {
            exactLogin = decryptPrivateValue(
                subject.encryptedAccountMapping,
                keyForStoredValue(keys.accountMappingEncryption, subject.accountMappingKeyId),
                CRYPTOGRAPHIC_PURPOSES.accountMapping,
                createAccountMappingBinding(subject.subjectId),
            );
        } catch {
            throw authorityUnavailable('subject-mapping-integrity');
        }
        const lookup = createLoginLookup(exactLogin, keys.loginLookup);
        if (
            subject.loginLookupKeyId !== lookup.keyId
            || !safeEqual(subject.loginLookupToken, lookup.token)
            || (expectedExactLogin !== undefined && exactLogin !== expectedExactLogin)
        ) throw authorityUnavailable('subject-mapping-integrity');
        return exactLogin;
    }

    function requireRuntimeControl(name) {
        if (controls[name] !== true) throw authorityUnavailable(`${name}-disabled`);
    }

    return Object.freeze({
        authorizeCurrent,
        authorizeLegacy,
        authorizeProtected,
        completeFace,
        createExistingPhotoChallenge,
        createRegistrationChallenge,
        current,
        loginLegacyWithSeeding,
        loginTarget,
        logout,
        remapSubjectLogin,
        registrationEnrollment,
        revokeAll,
        revokeLeakedLegacyAuthority,
        revokeSubject,
        runtimeControls: controls,
    });
}

function createFaceChallengeBinding(flowId, subjectId, sessionId) {
    const label = Buffer.from(
        'machado-session-authority\0face-challenge-binding\0v1\0',
        'utf8',
    );
    const fields = [flowId, subjectId, sessionId].map((value) => {
        if (typeof value !== 'string' || value.length === 0) {
            throw new TypeError('Face challenge binding identifiers are required');
        }
        const encoded = Buffer.from(value.toLowerCase(), 'utf8');
        const length = Buffer.allocUnsafe(4);
        length.writeUInt32BE(encoded.length);
        return Buffer.concat([length, encoded]);
    });
    return Buffer.concat([label, ...fields]);
}

function createAccountMappingBinding(subjectId) {
    if (typeof subjectId !== 'string' || subjectId.length === 0) {
        throw new TypeError('The subject mapping binding identifier is required');
    }
    return Buffer.from(
        `machado-session-authority\0account-mapping-binding\0v1\0${subjectId.toLowerCase()}`,
        'utf8',
    );
}

function createCurrentStatus(authority) {
    return {
        authenticationPhase: authority.session.phase,
        serverTime: authority.serverTime.toISOString(),
        expiresAt: authority.session.expiresAt.toISOString(),
        eligibilityRevalidateAt: authority.subject.eligibilityRevalidateAt.toISOString(),
        allowedNextOperations: allowedNextOperations(authority.session),
    };
}

function allowedNextOperations(session) {
    switch (session.phase) {
    case SESSION_PHASES.credentialVerified:
        return session.registrationRequired
            ? [NEXT_OPERATION_ROLES.registrationEnrollment]
            : [NEXT_OPERATION_ROLES.faceChallenge];
    case SESSION_PHASES.registrationPending:
        return [NEXT_OPERATION_ROLES.registrationChallenge];
    case SESSION_PHASES.facePending:
        return [NEXT_OPERATION_ROLES.faceCompletion];
    case SESSION_PHASES.authenticated:
        return [NEXT_OPERATION_ROLES.protectedLearning, NEXT_OPERATION_ROLES.revokeAll];
    default:
        return [];
    }
}

function createIssuance(authority) {
    return {
        identifier: authority.identifier,
        expiresAt: authority.session.expiresAt,
        serverTime: authority.serverTime,
    };
}

function requireDatabaseControl(control, name) {
    if (!control || control[name] !== true) throw authorityUnavailable(`${name}-not-qualified`);
}

function requireAuthorityControlOpen(control) {
    if (!control || control.incidentState !== 'normal') {
        throw authorityUnavailable('authority-incident');
    }
}

function requireEligibleUpdate(result) {
    if (result && result.eligible === false) throw forbiddenAuthority('ineligible');
}

function keyForStoredValue(activeKey, storedKeyId) {
    if (!activeKey || activeKey.keyId !== storedKeyId) throw authorityUnavailable('encryption-key-unavailable');
    return activeKey;
}

function safeEqual(left, right) {
    return Buffer.isBuffer(left)
        && Buffer.isBuffer(right)
        && left.length === right.length
        && timingSafeEqual(left, right);
}

function isDefinitiveFaceResult(result) {
    return result
        && ['realface', 'spoofface', 'uncertain'].includes(result.livenessDecision)
        && typeof result.matchDecision === 'boolean';
}

function isPassingFaceResult(result) {
    return result.livenessDecision === 'realface' && result.matchDecision === true;
}

function normalizeExternalFailure(error, reason) {
    if (isSessionAuthorityError(error)) return error;
    return authorityUnavailable(reason, { cause: error });
}

function validateDependencies(dependencies) {
    for (const name of ['store', 'accountSource', 'faceSource', 'legacyHandleAuthority', 'keys']) {
        if (!dependencies[name] || typeof dependencies[name] !== 'object') {
            throw new TypeError(`Session authority requires ${name}`);
        }
    }
    for (const name of ['createSubjectId', 'createSessionId', 'createFlowId', 'createCorrelationId', 'formatLegacyAccessDate']) {
        if (typeof dependencies[name] !== 'function') throw new TypeError(`Session authority requires ${name}`);
    }
}

module.exports = {
    RUNTIME_CONTROL_NAMES,
    allowedNextOperations,
    createCurrentStatus,
    createSessionAuthority,
};
