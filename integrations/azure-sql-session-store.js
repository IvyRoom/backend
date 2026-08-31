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
    isSessionAuthorityError,
} = require('../domains/session-authority/errors');

const ACTIVE_PHASES = Object.freeze([
    SESSION_PHASES.credentialVerified,
    SESSION_PHASES.registrationPending,
    SESSION_PHASES.facePending,
    SESSION_PHASES.authenticated,
]);
const ALL_SESSION_PHASES = Object.freeze([
    ...ACTIVE_PHASES,
    SESSION_PHASES.expired,
    SESSION_PHASES.revoked,
    SESSION_PHASES.rotatedOut,
]);
const CONTROL_INCIDENT_STATES = Object.freeze(['normal', 'suspended', 'recovering']);
const FLOW_REGISTRATION_STATES = Object.freeze([
    'not-required',
    'required',
    'enrollment-accepted',
    'registered',
    'reconciliation-required',
]);
const FLOW_CHALLENGE_STATES = Object.freeze([
    'none',
    'creating',
    'active',
    'consumed',
    'failed',
    'reconciliation-required',
]);
const LEGACY_COMPATIBILITY_STATES = Object.freeze(['active', 'revoked', 'incident']);
const KEY_ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/u;

function isValidKeyId(value) {
    return typeof value === 'string' && KEY_ID_PATTERN.test(value);
}
const ROTATABLE_KEY_PURPOSES = Object.freeze([
    ['targetVerifier', 'targetVerifier'],
    ['legacyCompatibility', 'legacyCompatibility'],
    ['credentialFingerprint', 'credentialFingerprint'],
    ['faceChallengeEncryption', 'faceChallenge'],
]);
const ACTIVE_PHASE_SQL = "('credential-verified', 'registration-pending', 'face-pending', 'authenticated')";
const POST_COMMIT_ERROR = Symbol('post-commit-error');
const AUTHORITY_DATE_PARAMETERS = new Set([
    'createdAt',
    'dualStackStartedAt',
    'eligibilityObservedAt',
    'eligibilityRevalidateAt',
    'entitlementExpiresAt',
    'expiresAt',
    'hardSunsetAt',
    'incidentRecordedAt',
    'issuedAt',
    'leaseExpiresAt',
    'legacyAcceptanceDisabledAt',
    'legacyCompatibilityEnforcedAt',
    'legacyLedgerSeedingStartedAt',
    'legacyStopIssuanceAt',
    'legacyVerifierKeyIncidentAt',
    'observationStartedAt',
    'originalIssuedAt',
    'phaseStartedAt',
    'seedingHeartbeatAt',
    'seedingLeaseExpiresAt',
    'seedingQualifiedAt',
    'seedingStartedAt',
    'serverTime',
    'subjectTargetAdoptionStartedAt',
    'targetSessionIssuanceStartedAt',
    'targetVerifierKeyIncidentAt',
]);
const AUTHORITY_UUID_PARAMETERS = new Set([
    'challengeSessionId',
    'expectedSessionId',
    'flowId',
    'legacyCompatibilityId',
    'ownerId',
    'replacementSessionId',
    'sessionId',
    'subjectId',
]);
const AUTHORITY_BOOLEAN_PARAMETERS = new Set([
    'faceRequired',
    'legacyAcceptanceEnabled',
    'legacyCompatibilityEnforcementEnabled',
    'legacyIssuanceEnabled',
    'legacyLedgerSeedingEnabled',
    'registrationReconciliationRequired',
    'registrationRequired',
    'reset',
    'subjectTargetAdoptionEnabled',
    'targetRoutesEnabled',
    'targetSessionIssuanceEnabled',
]);
const AUTHORITY_BIGINT_PARAMETERS = new Set([
    'authorityGeneration',
    'authorityGenerationSnapshot',
    'continuityVersion',
    'controlVersion',
    'expectedCredentialVersion',
    'expectedVersion',
    'globalSessionEpoch',
    'globalEpochSnapshot',
    'credentialVersionSnapshot',
    'subjectEpoch',
    'subjectEpochSnapshot',
]);
const AUTHORITY_BINARY_PARAMETERS = new Set([
    'accountMappingKeyCommitment',
    'authorityKeysetCommitment',
    'credentialFingerprint',
    'credentialFingerprintKeyCommitment',
    'expectedCredentialFingerprint',
    'legacyHandleVerifier',
    'legacyCompatibilityKeyCommitment',
    'legacySigningKeyCommitment',
    'loginLookupKeyCommitment',
    'keysetLoginLookupKeyCommitment',
    'keysetAccountMappingKeyCommitment',
    'loginLookupToken',
    'expectedLoginLookupToken',
    'faceChallengeKeyCommitment',
    'targetVerifierKeyCommitment',
    'verifier',
]);

const CONTROL_CHANGE_COLUMNS = Object.freeze({
    authorityGeneration: 'authority_generation',
    globalSessionEpoch: 'global_session_epoch',
    targetRoutesEnabled: 'target_routes_enabled',
    targetSessionIssuanceEnabled: 'target_session_issuance_enabled',
    targetSessionIssuanceStartedAt: 'target_session_issuance_started_at',
    legacyLedgerSeedingEnabled: 'legacy_ledger_seeding_enabled',
    legacyLedgerSeedingStartedAt: 'legacy_ledger_seeding_started_at',
    seedingStartedAt: 'legacy_ledger_continuous_since',
    seedingContinuityVersion: 'legacy_ledger_continuity_version',
    seedingHeartbeatOwnerId: 'legacy_ledger_heartbeat_owner_id',
    seedingHeartbeatAt: 'legacy_ledger_heartbeat_at',
    seedingLeaseExpiresAt: 'legacy_ledger_lease_expires_at',
    seedingQualifiedAt: 'legacy_ledger_qualified_at',
    legacyCompatibilityEnforcementEnabled: 'legacy_compatibility_enforcement_enabled',
    legacyCompatibilityEnforcedAt: 'legacy_compatibility_enforced_at',
    subjectTargetAdoptionEnabled: 'subject_target_adoption_enabled',
    subjectTargetAdoptionStartedAt: 'subject_target_adoption_started_at',
    dualStackStartedAt: 'target_acceptance_started_at',
    legacyIssuanceEnabled: 'legacy_handle_issuance_enabled',
    legacyStopIssuanceAt: 'legacy_handle_issuance_stopped_at',
    legacyAcceptanceEnabled: 'legacy_handle_acceptance_enabled',
    legacyAcceptanceDisabledAt: 'legacy_handle_acceptance_disabled_at',
    hardSunsetAt: 'legacy_hard_sunset_at',
    incidentState: 'incident_state',
    incidentRecordedAt: 'incident_recorded_at',
    incidentCode: 'incident_code',
    targetVerifierKeyIncidentAt: 'target_verifier_key_incident_at',
    legacyVerifierKeyIncidentAt: 'legacy_verifier_key_incident_at',
});

const CONTROL_SELECT = `
    c.control_id AS controlId,
    c.control_version AS controlVersion,
    c.authority_generation AS authorityGeneration,
    c.global_session_epoch AS globalSessionEpoch,
    c.target_routes_enabled AS targetRoutesEnabled,
    c.target_session_issuance_enabled AS targetSessionIssuanceEnabled,
    c.target_session_issuance_started_at AS targetSessionIssuanceStartedAt,
    c.legacy_ledger_seeding_enabled AS legacyLedgerSeedingEnabled,
    c.legacy_ledger_seeding_started_at AS legacyLedgerSeedingStartedAt,
    c.legacy_ledger_continuous_since AS seedingStartedAt,
    c.legacy_ledger_continuity_version AS seedingContinuityVersion,
    c.legacy_ledger_heartbeat_owner_id AS seedingHeartbeatOwnerId,
    c.legacy_ledger_heartbeat_at AS seedingHeartbeatAt,
    c.legacy_ledger_lease_expires_at AS seedingLeaseExpiresAt,
    c.legacy_ledger_qualified_at AS seedingQualifiedAt,
    c.legacy_compatibility_enforcement_enabled AS legacyCompatibilityEnforcementEnabled,
    c.legacy_compatibility_enforced_at AS legacyCompatibilityEnforcedAt,
    c.subject_target_adoption_enabled AS subjectTargetAdoptionEnabled,
    c.subject_target_adoption_started_at AS subjectTargetAdoptionStartedAt,
    c.target_acceptance_started_at AS dualStackStartedAt,
    c.legacy_handle_issuance_enabled AS legacyIssuanceEnabled,
    c.legacy_handle_issuance_stopped_at AS legacyStopIssuanceAt,
    c.legacy_handle_acceptance_enabled AS legacyAcceptanceEnabled,
    c.legacy_handle_acceptance_disabled_at AS legacyAcceptanceDisabledAt,
    c.legacy_hard_sunset_at AS hardSunsetAt,
    c.incident_state AS incidentState,
    c.incident_recorded_at AS incidentRecordedAt,
    c.incident_code AS incidentCode,
    c.target_verifier_key_incident_at AS targetVerifierKeyIncidentAt,
    c.legacy_verifier_key_incident_at AS legacyVerifierKeyIncidentAt,
    c.created_at AS createdAt,
    c.updated_at AS updatedAt`;

const SUBJECT_SELECT = `
    s.subject_id AS subjectId,
    s.login_lookup_token AS loginLookupToken,
    s.login_lookup_key_id AS loginLookupKeyId,
    s.encrypted_legacy_account_mapping AS encryptedAccountMapping,
    s.account_mapping_encryption_key_id AS accountMappingKeyId,
    s.workbook_row_hint AS rowHint,
    s.credential_version AS credentialVersion,
    s.credential_fingerprint AS credentialFingerprint,
    s.credential_fingerprint_key_id AS credentialFingerprintKeyId,
    s.subject_session_epoch AS sessionEpoch,
    s.legacy_authority_disabled_at AS legacyAuthorityDisabledAt,
    s.eligibility_state AS eligibilityState,
    s.entitlement_expires_at AS entitlementExpiresAt,
    s.eligibility_observed_at AS eligibilityObservedAt,
    s.eligibility_revalidate_at AS eligibilityRevalidateAt,
    s.created_at AS createdAt`;

const SESSION_SELECT = `
    l.session_id AS sessionId,
    l.identifier_verifier AS verifier,
    l.verifier_key_id AS verifierKeyId,
    l.subject_id AS subjectId,
    l.phase AS phase,
    l.original_issued_at AS originalIssuedAt,
    l.phase_started_at AS phaseStartedAt,
    l.absolute_expires_at AS expiresAt,
    l.face_auth_required AS faceRequired,
    l.registration_required AS registrationRequired,
    l.subject_epoch_snapshot AS subjectEpochSnapshot,
    l.credential_version_snapshot AS credentialVersionSnapshot,
    l.global_epoch_snapshot AS globalEpochSnapshot,
    l.authority_generation_snapshot AS authorityGenerationSnapshot,
    l.revoked_at AS revokedAt,
    l.revocation_reason AS revocationReason,
    l.replacement_session_id AS replacementSessionId,
    l.created_at AS createdAt`;

const FLOW_SELECT = `
    f.flow_id AS flowId,
    f.subject_id AS subjectId,
    f.current_session_id AS currentSessionId,
    f.challenge_session_id AS challengeSessionId,
    f.registration_state AS registrationState,
    f.challenge_state AS challengeState,
    f.encrypted_provider_challenge_reference AS encryptedChallenge,
    f.provider_reference_encryption_key_id AS challengeKeyId,
    f.challenge_created_at AS challengeCreatedAt,
    f.challenge_resolved_at AS consumedAt,
    f.created_at AS createdAt,
    f.updated_at AS updatedAt`;

const LEGACY_SELECT = `
    b.compatibility_id AS legacyCompatibilityId,
    b.legacy_handle_verifier AS verifier,
    b.verifier_key_id AS verifierKeyId,
    b.subject_id AS subjectId,
    b.original_issued_at AS issuedAt,
    b.original_expires_at AS expiresAt,
    b.compatibility_state AS compatibilityState,
    b.revoked_at AS revokedAt,
    b.revocation_reason AS revocationReason,
    b.incident_at AS incidentAt,
    b.incident_code AS incidentCode,
    b.created_at AS createdAt`;

function createAzureSqlSessionStore({
    sql,
    connectionString,
    options = {},
    expectedAuthorityGeneration,
    loginLookupKeyId,
    loginLookupKeyCommitment,
    accountMappingKeyBinding,
    authorityKeysetBinding,
    legacySigningKeyBinding,
} = {}) {
    validateFactoryInput(
        sql,
        connectionString,
        options,
        expectedAuthorityGeneration,
        loginLookupKeyId,
        loginLookupKeyCommitment,
        accountMappingKeyBinding,
        authorityKeysetBinding,
        legacySigningKeyBinding,
    );
    const expectedLoginLookupKeyCommitment = Buffer.from(loginLookupKeyCommitment);
    const expectedAccountMappingKeyBinding = copyKeyBinding(accountMappingKeyBinding);
    const expectedAuthorityKeysetBinding = copyAuthorityKeysetBinding(authorityKeysetBinding);
    const expectedLegacySigningKeyBinding = copyKeyBinding(legacySigningKeyBinding);
    let poolPromise;
    let poolFaulted = false;
    let continuityCompromised = false;
    let continuityObservedActive = false;

    async function getPool() {
        if (poolFaulted) throw authorityUnavailable('session-store-unavailable');
        if (!poolPromise) {
            poolPromise = connectPool(sql, connectionString, options, () => {
                poolFaulted = true;
                continuityCompromised = true;
            }).catch((error) => {
                poolPromise = undefined;
                continuityCompromised = true;
                throw normalizeDriverError(error);
            });
        }
        const pool = await poolPromise;
        if (poolFaulted) throw authorityUnavailable('session-store-unavailable');
        return pool;
    }

    async function execute(owner, statement, parameters = {}) {
        const request = createRequest(sql, owner);
        for (const [name, value] of Object.entries(parameters)) {
            const type = authorityParameterType(sql, name, value);
            if (type === undefined) request.input(name, value);
            else request.input(name, type, value);
        }
        return request.query(statement);
    }

    async function read(statement, parameters = {}) {
        try {
            return await execute(await getPool(), statement, parameters);
        } catch (error) {
            const normalized = normalizeDriverError(error);
            recordContinuityFailure(normalized);
            throw normalized;
        }
    }

    async function transact(operation, {
        fenceAuthority = true,
        fenceLookup = true,
        prelockControl = false,
    } = {}) {
        let transaction;
        let begun = false;
        let result;
        try {
            const pool = await getPool();
            transaction = createTransaction(sql, pool);
            await transaction.begin(serializableIsolation(sql));
            begun = true;
            const run = (statement, parameters) => execute(transaction, statement, parameters);
            run.expectedAuthorityGeneration = fenceAuthority
                ? expectedAuthorityGeneration
                : undefined;
            run.expectedLoginLookupKeyId = fenceLookup ? loginLookupKeyId : undefined;
            run.expectedLoginLookupKeyCommitment = fenceLookup
                ? expectedLoginLookupKeyCommitment
                : undefined;
            run.expectedAccountMappingKeyBinding = fenceLookup
                ? expectedAccountMappingKeyBinding
                : undefined;
            run.expectedAuthorityKeysetBinding = fenceLookup
                ? expectedAuthorityKeysetBinding
                : undefined;
            run.expectedLegacySigningKeyBinding = fenceLookup
                ? expectedLegacySigningKeyBinding
                : undefined;
            if (prelockControl) await selectControl(run, true);
            result = await operation(run);
            if (fenceAuthority && run.authorityGenerationFenced !== true) {
                await selectControl(run, true);
            }
        } catch (error) {
            if (begun) {
                try {
                    await transaction.rollback();
                } catch {
                    const unknown = authorityUnavailable('transaction-outcome-unknown');
                    recordContinuityFailure(unknown);
                    throw unknown;
                }
            }
            const normalized = normalizeDriverError(error);
            recordContinuityFailure(normalized);
            throw normalized;
        }

        try {
            await transaction.commit();
        } catch {
            const unknown = authorityUnavailable('transaction-outcome-unknown');
            recordContinuityFailure(unknown);
            throw unknown;
        }
        if (result && result[POST_COMMIT_ERROR]) throw result.error;
        return result;
    }

    function recordContinuityFailure(error) {
        if (
            !isSessionAuthorityError(error)
            || error.reason === 'session-store-unavailable'
            || error.reason === 'transaction-outcome-unknown'
        ) continuityCompromised = true;
    }

    function requireBoundLoginLookupKeyId(candidate) {
        if (candidate !== loginLookupKeyId) {
            throw authorityUnavailable('login-lookup-key-mismatch');
        }
    }

    async function readServerTime(run) {
        const result = await run('/* session-authority:server-time */ SELECT SYSUTCDATETIME() AS serverTime;');
        const serverTime = result.recordset && result.recordset[0] && result.recordset[0].serverTime;
        if (!isValidDate(serverTime)) throw authorityUnavailable('session-store-integrity');
        return copyDate(serverTime);
    }

    async function readControl() {
        return transact(async (run) => {
            const control = await selectControl(run, true);
            const serverTime = await readServerTime(run);
            return { control, serverTime };
        });
    }

    async function heartbeatLegacySeedingContinuity({ ownerId }) {
        validateContinuityOwnerId(ownerId);
        return transact(async (run) => {
            const current = await selectControl(run, 'update');
            const serverTime = await readServerTime(run);
            if (current.incidentState !== 'normal') {
                continuityObservedActive = false;
                continuityCompromised = true;
                return {
                    active: false,
                    owner: false,
                    reset: true,
                    control: current,
                    serverTime,
                };
            }
            if (!seedingHorizonActive(current)) {
                continuityObservedActive = false;
                continuityCompromised = false;
                return {
                    active: false,
                    owner: false,
                    reset: false,
                    control: current,
                    serverTime,
                };
            }

            const live = hasLiveSeedingContinuity(current, serverTime);
            const joining = !continuityObservedActive;
            const reset = joining || continuityCompromised || !live;
            const ownsLease = current.seedingHeartbeatOwnerId === ownerId;
            if (!reset && !ownsLease) {
                continuityObservedActive = true;
                continuityCompromised = false;
                return {
                    active: true,
                    owner: false,
                    reset: false,
                    control: current,
                    serverTime,
                };
            }

            const leaseExpiresAt = new Date(serverTime.getTime() + LEGACY_SEEDING_LEASE_MS);
            const updated = await run(`
                /* session-authority:heartbeat-legacy-seeding-continuity */
                UPDATE dbo.session_authority_control WITH (UPDLOCK, SERIALIZABLE)
                SET legacy_ledger_continuity_version = legacy_ledger_continuity_version + 1,
                    legacy_ledger_continuous_since = CASE
                        WHEN @reset = 1 THEN @serverTime
                        ELSE legacy_ledger_continuous_since
                    END,
                    legacy_ledger_heartbeat_owner_id = @ownerId,
                    legacy_ledger_heartbeat_at = @serverTime,
                    legacy_ledger_lease_expires_at = @leaseExpiresAt,
                    updated_at = @serverTime
                WHERE control_id = 1
                    AND control_version = @controlVersion
                    AND legacy_ledger_continuity_version = @continuityVersion;
            `, {
                controlVersion: current.version,
                continuityVersion: current.seedingContinuityVersion,
                leaseExpiresAt,
                ownerId,
                reset,
                serverTime,
            });
            if (affectedRows(updated) !== 1) {
                throw authorityConflict('seeding-continuity-compare-and-replace');
            }
            const control = await selectControl(run, true);
            continuityObservedActive = true;
            continuityCompromised = false;
            return {
                active: true,
                owner: true,
                reset,
                control,
                serverTime,
            };
        }, { prelockControl: false });
    }

    async function initializeLoginLookupKey({
        loginLookupKeyId: candidateKeyId,
        loginLookupKeyCommitment: candidateCommitment,
    }) {
        if (
            candidateKeyId !== loginLookupKeyId
            || !Buffer.isBuffer(candidateCommitment)
            || !candidateCommitment.equals(expectedLoginLookupKeyCommitment)
        ) throw authorityUnavailable('login-lookup-key-mismatch');
        return transact(async (run) => {
            const current = await selectControl(run, 'update');
            if (Boolean(current.loginLookupKeyInitialized)) {
                if (!Boolean(current.loginLookupKeyMatches)) {
                    throw authorityUnavailable('login-lookup-key-mismatch');
                }
                if (
                    !Boolean(current.accountMappingKeyMatches)
                    || !Boolean(current.keysetLoginLookupKeyMatches)
                    || !Boolean(current.keysetAccountMappingKeyMatches)
                    || !Boolean(current.authorityKeysetInitialized)
                    || !Boolean(current.authorityKeysetAggregateMatches)
                    || !Boolean(current.targetVerifierKeyMatches)
                    || !Boolean(current.legacyCompatibilityKeyMatches)
                    || !Boolean(current.credentialFingerprintKeyMatches)
                    || !Boolean(current.faceChallengeKeyMatches)
                    || !Boolean(current.legacySigningKeyMatches)
                ) throw authorityUnavailable('authority-keyset-mismatch');
                return {
                    control: sanitizeControlLookupFlags(current),
                    idempotent: true,
                    serverTime: await readServerTime(run),
                };
            }
            requireDormantLookupInitialization(current);
            const occupied = await run(`
                /* session-authority:initialize-login-lookup-key:empty */
                SELECT CASE WHEN
                    EXISTS (SELECT 1 FROM dbo.learning_subject WITH (UPDLOCK, HOLDLOCK))
                    OR EXISTS (SELECT 1 FROM dbo.learning_session WITH (UPDLOCK, HOLDLOCK))
                    OR EXISTS (SELECT 1 FROM dbo.learning_session_flow WITH (UPDLOCK, HOLDLOCK))
                    OR EXISTS (SELECT 1 FROM dbo.legacy_session_compatibility WITH (UPDLOCK, HOLDLOCK))
                    THEN 1 ELSE 0 END AS authorityDataExists;
            `);
            if (Boolean(exactlyOne(occupied.recordset, 'control-integrity').authorityDataExists)) {
                throw authorityUnavailable('login-lookup-key-initialization-blocked');
            }
            const serverTime = await readServerTime(run);
            const updated = await run(`
                /* session-authority:initialize-login-lookup-key */
                UPDATE dbo.session_authority_control WITH (UPDLOCK, SERIALIZABLE)
                SET login_lookup_key_id = @loginLookupKeyId,
                    login_lookup_key_commitment = @loginLookupKeyCommitment,
                    account_mapping_key_id = @accountMappingKeyId,
                    account_mapping_key_commitment = @accountMappingKeyCommitment,
                    keyset_login_lookup_key_id = @keysetLoginLookupKeyId,
                    keyset_login_lookup_key_commitment = @keysetLoginLookupKeyCommitment,
                    keyset_account_mapping_key_id = @keysetAccountMappingKeyId,
                    keyset_account_mapping_key_commitment = @keysetAccountMappingKeyCommitment,
                    target_verifier_key_id = @targetVerifierKeyId,
                    target_verifier_key_commitment = @targetVerifierKeyCommitment,
                    legacy_compatibility_key_id = @legacyCompatibilityKeyId,
                    legacy_compatibility_key_commitment = @legacyCompatibilityKeyCommitment,
                    credential_fingerprint_key_id = @credentialFingerprintKeyId,
                    credential_fingerprint_key_commitment = @credentialFingerprintKeyCommitment,
                    face_challenge_key_id = @faceChallengeKeyId,
                    face_challenge_key_commitment = @faceChallengeKeyCommitment,
                    legacy_signing_key_id = @legacySigningKeyId,
                    legacy_signing_key_commitment = @legacySigningKeyCommitment,
                    authority_keyset_commitment = @authorityKeysetCommitment,
                    control_version = control_version + 1,
                    updated_at = @serverTime
                WHERE control_id = 1
                    AND control_version = @controlVersion
                    AND login_lookup_key_id IS NULL
                    AND login_lookup_key_commitment IS NULL
                    AND authority_keyset_commitment IS NULL;
            `, {
                controlVersion: current.version,
                loginLookupKeyId,
                loginLookupKeyCommitment: expectedLoginLookupKeyCommitment,
                accountMappingKeyId: expectedAccountMappingKeyBinding.keyId,
                accountMappingKeyCommitment: expectedAccountMappingKeyBinding.commitment,
                ...keysetSqlParameters(expectedAuthorityKeysetBinding),
                legacySigningKeyId: expectedLegacySigningKeyBinding.keyId,
                legacySigningKeyCommitment: expectedLegacySigningKeyBinding.commitment,
                serverTime,
            });
            if (affectedRows(updated) !== 1) {
                throw authorityConflict('login-lookup-key-initialization-race');
            }
            run.expectedLoginLookupKeyId = loginLookupKeyId;
            run.expectedLoginLookupKeyCommitment = expectedLoginLookupKeyCommitment;
            run.expectedAccountMappingKeyBinding = expectedAccountMappingKeyBinding;
            run.expectedAuthorityKeysetBinding = expectedAuthorityKeysetBinding;
            run.expectedLegacySigningKeyBinding = expectedLegacySigningKeyBinding;
            const control = await selectControl(run, true);
            return {
                control: sanitizeControlLookupFlags(control),
                idempotent: false,
                serverTime,
            };
        }, { fenceLookup: false, prelockControl: false });
    }

    async function transitionControl({ expectedVersion, changes }) {
        validateControlChanges(changes);
        return transact(async (run) => {
            const current = await selectControl(run, 'update', {
                allowRotatableKeyMismatch: true,
            });
            const serverTime = await readServerTime(run);
            let stampedChanges = stampControlTransition(current, changes, serverTime);
            const keyRecovery = prepareAuthorityKeysetRecovery(
                current,
                stampedChanges,
                serverTime,
                expectedAuthorityGeneration,
                expectedAuthorityKeysetBinding,
                expectedLegacySigningKeyBinding,
            );
            stampedChanges = keyRecovery.changes;
            requireControlTransitionGeneration(
                current,
                stampedChanges,
                expectedAuthorityGeneration,
                keyRecovery.active,
            );
            if (current.version !== expectedVersion) throw authorityConflict('control-compare-and-replace');
            validateControlTransition(current, stampedChanges, serverTime, {
                keyRecoveryActive: keyRecovery.active,
            });
            if (current.incidentState === 'recovering' && stampedChanges.incidentState === 'normal') {
                await requireNoLiveFaceRecoveryAuthority(run);
            }
            const assignments = createControlAssignments(stampedChanges);
            const keysetAssignments = keyRecovery.active
                ? createAuthorityKeysetAssignments(
                    expectedAuthorityKeysetBinding,
                    expectedLegacySigningKeyBinding,
                    keyRecovery.changedPurposes,
                )
                : { sql: [], parameters: {} };
            const parameters = {
                expectedVersion,
                serverTime,
                ...assignments.parameters,
                ...keysetAssignments.parameters,
            };
            const result = await run(`
                /* session-authority:transition-control */
                UPDATE dbo.session_authority_control WITH (UPDLOCK, SERIALIZABLE)
                SET ${[...assignments.sql, ...keysetAssignments.sql].join(', ')},
                    control_version = control_version + 1,
                    updated_at = @serverTime
                WHERE control_id = 1
                    AND control_version = @expectedVersion;
            `, parameters);
            if (affectedRows(result) !== 1) throw authorityConflict('control-compare-and-replace');
            if (keyRecovery.active) {
                await quarantineUnresolvedFlowsForKeyRecovery(run, serverTime);
            }
            if (keyRecovery.retireLegacyBindings) {
                await incidentLegacyBindingsForKeyRecovery(run, serverTime);
            }
            const control = await selectControl(run, true);
            requireAuthorityGeneration(control, expectedAuthorityGeneration);
            return { control, serverTime };
        }, { fenceAuthority: false, prelockControl: false });
    }

    async function createOrLoadSubject(input) {
        requireBoundLoginLookupKeyId(input.loginLookupKeyId);
        validateExpectedControlVersion(input.expectedControlVersion);
        return transact(async (run) => {
            await requireExpectedNormalControl(run, input.expectedControlVersion);
            const lookup = await run(`
                /* session-authority:create-or-load-subject:lookup */
                SELECT ${SUBJECT_SELECT}
                FROM dbo.learning_subject AS s WITH (UPDLOCK, HOLDLOCK)
                WHERE s.login_lookup_key_id = @loginLookupKeyId COLLATE Latin1_General_100_BIN2
                    AND s.login_lookup_token = @loginLookupToken;
            `, pick(input, ['loginLookupKeyId', 'loginLookupToken']));
            if (lookup.recordset.length > 1) throw authorityUnavailable('subject-mapping-integrity');
            const serverTime = await readServerTime(run);
            if (lookup.recordset.length === 1) {
                return { created: false, subject: mapSubject(lookup.recordset[0]), serverTime };
            }

            await run(`
                /* session-authority:create-or-load-subject:insert */
                INSERT INTO dbo.learning_subject (
                    subject_id,
                    login_lookup_token,
                    login_lookup_key_id,
                    encrypted_legacy_account_mapping,
                    account_mapping_encryption_key_id,
                    workbook_row_hint,
                    credential_fingerprint,
                    credential_fingerprint_key_id,
                    eligibility_state,
                    entitlement_expires_at,
                    eligibility_observed_at,
                    eligibility_revalidate_at,
                    created_at
                ) VALUES (
                    @subjectId,
                    @loginLookupToken,
                    @loginLookupKeyId,
                    @encryptedAccountMapping,
                    @accountMappingKeyId,
                    @rowHint,
                    @credentialFingerprint,
                    @credentialFingerprintKeyId,
                    @eligibilityState,
                    @entitlementExpiresAt,
                    @eligibilityObservedAt,
                    @eligibilityRevalidateAt,
                    @serverTime
                );
            `, {
                ...pick(input, [
                    'subjectId',
                    'loginLookupToken',
                    'loginLookupKeyId',
                    'encryptedAccountMapping',
                    'accountMappingKeyId',
                    'rowHint',
                    'credentialFingerprint',
                    'credentialFingerprintKeyId',
                    'eligibilityState',
                    'entitlementExpiresAt',
                    'eligibilityObservedAt',
                    'eligibilityRevalidateAt',
                ]),
                serverTime,
            });
            const inserted = await selectSubject(run, input.subjectId, true);
            return { created: true, subject: inserted, serverTime };
        });
    }

    async function readSubjectByLookup({ loginLookupKeyId, loginLookupToken, expectedControlVersion }) {
        requireBoundLoginLookupKeyId(loginLookupKeyId);
        if (expectedControlVersion !== undefined) validateExpectedControlVersion(expectedControlVersion);
        return transact(async (run) => {
            if (expectedControlVersion !== undefined) {
                await requireExpectedNormalControl(run, expectedControlVersion);
            } else {
                await selectControl(run, true);
            }
            const result = await run(`
                /* session-authority:read-subject-by-lookup */
                SELECT ${SUBJECT_SELECT}
                FROM dbo.learning_subject AS s WITH (HOLDLOCK)
                WHERE s.login_lookup_key_id = @loginLookupKeyId COLLATE Latin1_General_100_BIN2
                    AND s.login_lookup_token = @loginLookupToken;
            `, { loginLookupKeyId, loginLookupToken });
            if (result.recordset.length > 1) throw authorityUnavailable('subject-mapping-integrity');
            const serverTime = await readServerTime(run);
            return {
                subject: result.recordset[0] ? mapSubject(result.recordset[0]) : null,
                serverTime,
            };
        });
    }

    async function remapSubjectLogin(input) {
        requireBoundLoginLookupKeyId(input.expectedLoginLookupKeyId);
        requireBoundLoginLookupKeyId(input.loginLookupKeyId);
        validateSubjectLoginRemap(input);
        return transact(async (run) => {
            await selectControl(run, true);
            const subject = await selectSubject(run, input.subjectId, true);
            const alreadyRemapped = subject.loginLookupKeyId === input.loginLookupKeyId
                && sameBuffer(subject.loginLookupToken, input.loginLookupToken);
            if (alreadyRemapped) {
                const serverTime = await readServerTime(run);
                return { subject, idempotent: true, serverTime };
            }
            if (
                subject.loginLookupKeyId !== input.expectedLoginLookupKeyId
                || !sameBuffer(subject.loginLookupToken, input.expectedLoginLookupToken)
            ) throw authorityUnavailable('subject-mapping-conflict');

            const target = await run(`
                /* session-authority:remap-subject-login:target */
                SELECT ${SUBJECT_SELECT}
                FROM dbo.learning_subject AS s WITH (UPDLOCK, HOLDLOCK)
                WHERE s.login_lookup_key_id = @loginLookupKeyId COLLATE Latin1_General_100_BIN2
                    AND s.login_lookup_token = @loginLookupToken;
            `, pick(input, ['loginLookupKeyId', 'loginLookupToken']));
            if (target.recordset.length !== 0) {
                throw authorityUnavailable('subject-mapping-conflict');
            }
            const serverTime = await readServerTime(run);

            const updated = await run(`
                /* session-authority:remap-subject-login:update */
                UPDATE dbo.learning_subject WITH (UPDLOCK, SERIALIZABLE)
                SET login_lookup_key_id = @loginLookupKeyId,
                    login_lookup_token = @loginLookupToken,
                    encrypted_legacy_account_mapping = @encryptedAccountMapping,
                    account_mapping_encryption_key_id = @accountMappingKeyId
                WHERE subject_id = @subjectId
                    AND login_lookup_key_id = @expectedLoginLookupKeyId COLLATE Latin1_General_100_BIN2
                    AND login_lookup_token = @expectedLoginLookupToken;
            `, pick(input, [
                'subjectId',
                'expectedLoginLookupKeyId',
                'expectedLoginLookupToken',
                'loginLookupKeyId',
                'loginLookupToken',
                'encryptedAccountMapping',
                'accountMappingKeyId',
            ]));
            if (affectedRows(updated) !== 1) {
                throw authorityUnavailable('subject-mapping-conflict');
            }
            return {
                subject: await selectSubject(run, input.subjectId, true),
                idempotent: false,
                serverTime,
            };
        });
    }

    async function updateEligibility(input) {
        return updateSubjectAuthority(input, async (run, serverTime) => {
            if (
                !isValidDate(input.entitlementExpiresAt)
                || !isValidDate(input.eligibilityRevalidateAt)
            ) {
                throw authorityUnavailable('eligibility-data-integrity');
            }
            if (input.entitlementExpiresAt <= serverTime) {
                const expired = await run(`
                    /* session-authority:update-eligibility:expired */
                    UPDATE dbo.learning_subject WITH (UPDLOCK, SERIALIZABLE)
                    SET eligibility_state = 'ineligible',
                        entitlement_expires_at = @entitlementExpiresAt,
                        eligibility_observed_at = @serverTime,
                        eligibility_revalidate_at = @serverTime,
                        subject_session_epoch = subject_session_epoch + 1
                    WHERE subject_id = @subjectId
                        AND eligibility_observed_at <= @observationStartedAt
                        AND credential_version = @expectedCredentialVersion
                        AND credential_fingerprint_key_id = @expectedCredentialFingerprintKeyId COLLATE Latin1_General_100_BIN2
                        AND credential_fingerprint = @expectedCredentialFingerprint;
                `, {
                    subjectId: input.subjectId,
                    entitlementExpiresAt: input.entitlementExpiresAt,
                    observationStartedAt: input.observationStartedAt,
                    expectedCredentialVersion: input.expectedCredentialVersion,
                    expectedCredentialFingerprintKeyId: input.expectedCredentialFingerprintKeyId,
                    expectedCredentialFingerprint: input.expectedCredentialFingerprint,
                    serverTime,
                });
                if (affectedRows(expired) !== 1) {
                    throw authorityConflict('subject-observation-compare-and-replace');
                }
                await revokeSubjectSessions(
                    run,
                    input.subjectId,
                    serverTime,
                    'entitlement-expired',
                );
                await revokeSubjectLegacyBindings(
                    run,
                    input.subjectId,
                    serverTime,
                    'entitlement-expired',
                );
                return {
                    eligible: false,
                    subject: await selectSubject(run, input.subjectId, true),
                    serverTime,
                };
            }
            const eligibilityRevalidateAt = new Date(Math.min(
                input.eligibilityRevalidateAt.getTime(),
                input.entitlementExpiresAt.getTime(),
            ));
            if (eligibilityRevalidateAt <= serverTime) {
                throw authorityUnavailable('eligibility-revalidation-required');
            }
            const result = await run(`
                /* session-authority:update-eligibility */
                UPDATE dbo.learning_subject WITH (UPDLOCK, SERIALIZABLE)
                SET workbook_row_hint = @rowHint,
                    eligibility_state = 'eligible',
                    entitlement_expires_at = @entitlementExpiresAt,
                    eligibility_observed_at = @serverTime,
                    eligibility_revalidate_at = @eligibilityRevalidateAt
                WHERE subject_id = @subjectId
                    AND eligibility_observed_at <= @observationStartedAt
                    AND credential_version = @expectedCredentialVersion
                    AND credential_fingerprint_key_id = @expectedCredentialFingerprintKeyId COLLATE Latin1_General_100_BIN2
                    AND credential_fingerprint = @expectedCredentialFingerprint;
            `, {
                subjectId: input.subjectId,
                rowHint: input.rowHint,
                entitlementExpiresAt: input.entitlementExpiresAt,
                eligibilityRevalidateAt,
                observationStartedAt: input.observationStartedAt,
                expectedCredentialVersion: input.expectedCredentialVersion,
                expectedCredentialFingerprintKeyId: input.expectedCredentialFingerprintKeyId,
                expectedCredentialFingerprint: input.expectedCredentialFingerprint,
                serverTime,
            });
            if (affectedRows(result) !== 1) {
                throw authorityConflict('subject-observation-compare-and-replace');
            }
            return { eligible: true, subject: await selectSubject(run, input.subjectId, true), serverTime };
        });
    }

    async function revokeForIneligibility(input) {
        return updateSubjectAuthority(input, async (run, serverTime) => {
            const subjectUpdate = await run(`
                /* session-authority:revoke-for-ineligibility:subject */
                UPDATE dbo.learning_subject WITH (UPDLOCK, SERIALIZABLE)
                SET eligibility_state = @eligibilityState,
                    entitlement_expires_at = @entitlementExpiresAt,
                    eligibility_observed_at = @observationStartedAt,
                    eligibility_revalidate_at = @observationStartedAt,
                    subject_session_epoch = subject_session_epoch + 1
                WHERE subject_id = @subjectId
                    AND eligibility_observed_at <= @observationStartedAt
                    AND credential_version = @expectedCredentialVersion
                    AND credential_fingerprint_key_id = @expectedCredentialFingerprintKeyId COLLATE Latin1_General_100_BIN2
                    AND credential_fingerprint = @expectedCredentialFingerprint;
            `, {
                subjectId: input.subjectId,
                eligibilityState: input.eligibilityState || 'ineligible',
                entitlementExpiresAt: input.entitlementExpiresAt,
                observationStartedAt: input.observationStartedAt,
                expectedCredentialVersion: input.expectedCredentialVersion,
                expectedCredentialFingerprintKeyId: input.expectedCredentialFingerprintKeyId,
                expectedCredentialFingerprint: input.expectedCredentialFingerprint,
            });
            if (affectedRows(subjectUpdate) !== 1) {
                throw authorityConflict('subject-observation-compare-and-replace');
            }
            const reason = input.reason || 'ineligible';
            validateRevocationReason(reason);
            await revokeSubjectSessions(run, input.subjectId, serverTime, reason);
            await revokeSubjectLegacyBindings(run, input.subjectId, serverTime, reason);
            return { subject: await selectSubject(run, input.subjectId, true), serverTime };
        });
    }

    async function revokeForCredentialChange(input) {
        return updateSubjectAuthority(input, async (run, serverTime) => {
            const subjectUpdate = await run(`
                /* session-authority:revoke-for-credential-change:subject */
                UPDATE dbo.learning_subject WITH (UPDLOCK, SERIALIZABLE)
                SET credential_version = credential_version + 1,
                    credential_fingerprint_key_id = @credentialFingerprintKeyId,
                    credential_fingerprint = @credentialFingerprint,
                    eligibility_observed_at = @observationStartedAt,
                    eligibility_revalidate_at = @observationStartedAt,
                    subject_session_epoch = subject_session_epoch + 1
                WHERE subject_id = @subjectId
                    AND eligibility_observed_at <= @observationStartedAt
                    AND credential_version = @expectedCredentialVersion
                    AND credential_fingerprint_key_id = @expectedCredentialFingerprintKeyId COLLATE Latin1_General_100_BIN2
                    AND credential_fingerprint = @expectedCredentialFingerprint;
            `, pick(input, [
                'subjectId',
                'observationStartedAt',
                'expectedCredentialVersion',
                'expectedCredentialFingerprintKeyId',
                'expectedCredentialFingerprint',
                'credentialFingerprintKeyId',
                'credentialFingerprint',
            ]));
            if (affectedRows(subjectUpdate) !== 1) {
                throw authorityConflict('subject-observation-compare-and-replace');
            }
            await revokeSubjectSessions(run, input.subjectId, serverTime, 'credential-reset');
            await revokeSubjectLegacyBindings(run, input.subjectId, serverTime, 'credential-reset');
            return { subject: await selectSubject(run, input.subjectId, true), serverTime };
        });
    }

    async function updateSubjectAuthority(input, operation) {
        if (!input || typeof input.subjectId !== 'string') throw new TypeError('subjectId is required');
        validateSubjectObservationExpectation(input);
        validateExpectedControlVersion(input.expectedControlVersion);
        return transact(async (run) => {
            await requireExpectedNormalControl(run, input.expectedControlVersion);
            await selectSubject(run, input.subjectId, true);
            const serverTime = await readServerTime(run);
            return operation(run, serverTime);
        });
    }

    async function inspectLoginPredecessor({ verifierKeyId, verifier }) {
        return transact(async (run) => {
            const authority = await selectAuthorityByVerifier(run, verifierKeyId, verifier, true);
            const serverTime = await readServerTime(run);
            if (authority.control.incidentState !== 'normal') {
                throw authorityUnavailable('authority-incident');
            }
            if (!authority.session) return { kind: 'unusable', serverTime };
            const state = evaluateAuthority(authority, serverTime);
            if (!state.active) {
                if (state.expired) await markExpired(run, authority.session.sessionId);
                if (state.ineligible) {
                    await enforceStoredIneligibility(run, authority.subject, serverTime);
                }
                if (state.unavailable) throw authorityUnavailable(state.reason);
                return { kind: 'unusable', serverTime };
            }
            return {
                kind: 'active',
                expectedSessionId: authority.session.sessionId,
                expectedVersion: authority.session.version,
                serverTime,
            };
        });
    }

    async function issueSession(input) {
        validateSessionCandidate(input);
        validateSubjectCredentialExpectation(input);
        return transact(async (run) => {
            const control = await selectControl(run, true);
            const subject = await selectSubject(run, input.subjectId, true);
            requireSubjectCredentialExpectation(subject, input);

            let predecessor;
            let predecessorSubject;
            let predecessorFlow;
            if (input.predecessor) {
                predecessor = await selectSession(run, input.predecessor.expectedSessionId, true);
                predecessorSubject = predecessor
                    ? await selectSubject(run, predecessor.subjectId, true)
                    : null;
                predecessorFlow = predecessor
                    ? await selectFlow(run, predecessor.sessionId, true, false)
                    : null;
            }
            const serverTime = await readServerTime(run);
            requireIssuanceControl(control, serverTime);
            const subjectState = evaluateSubjectEligibility(subject, serverTime);
            if (!subjectState.eligible) {
                if (subjectState.unavailable) throw authorityUnavailable(subjectState.reason);
                await enforceStoredIneligibility(run, subject, serverTime);
                return postCommitError(forbiddenAuthority('ineligible'));
            }
            requireFreshEligibility(subject, serverTime);
            if (input.predecessor) {
                requireExpectedSession(
                    predecessor,
                    input.predecessor.expectedVersion,
                    null,
                    serverTime,
                    predecessorSubject,
                    control,
                    predecessorFlow,
                );
            }

            if (subject.legacyAuthorityDisabledAt === null) {
                const cutoff = await run(`
                    /* session-authority:issue-session:disable-legacy */
                    UPDATE dbo.learning_subject
                    SET legacy_authority_disabled_at = @serverTime
                    WHERE subject_id = @subjectId
                        AND legacy_authority_disabled_at IS NULL;
                `, { subjectId: subject.subjectId, serverTime });
                if (affectedRows(cutoff) !== 1) throw authorityConflict('subject-adoption-compare-and-replace');
                subject.legacyAuthorityDisabledAt = serverTime;
            }

            await insertSession(run, input, subject, control, serverTime, {
                originalIssuedAt: serverTime,
                phaseStartedAt: serverTime,
                expiresAt: new Date(serverTime.getTime() + lifetimeForPhase(input.phase)),
            });
            if (predecessor) await rotateOut(run, predecessor.sessionId, input.sessionId, serverTime);
            return {
                session: await selectSession(run, input.sessionId, true),
                subject: await selectSubject(run, input.subjectId, true),
                control,
                serverTime,
            };
        });
    }

    async function readSession({ verifierKeyId, verifier }) {
        return transact(async (run) => {
            const authority = await selectAuthorityByVerifier(run, verifierKeyId, verifier, true);
            const serverTime = await readServerTime(run);
            if (authority.control.incidentState !== 'normal') {
                throw authorityUnavailable('authority-incident');
            }
            if (!authority.session) throw invalidAuthority('unknown-session');
            const state = evaluateAuthority(authority, serverTime);
            if (!state.active) {
                if (state.expired) {
                    await markExpired(run, authority.session.sessionId);
                    return postCommitError(invalidAuthority('expired'));
                }
                if (state.ineligible) {
                    await enforceStoredIneligibility(run, authority.subject, serverTime);
                    return postCommitError(forbiddenAuthority('ineligible'));
                }
                if (state.unavailable) throw authorityUnavailable(state.reason);
                throw invalidAuthority(state.reason);
            }
            return { ...authority, serverTime };
        });
    }

    async function rotateSession(input) {
        validateSessionCandidate(input);
        return rotateExpectedSession(input, async ({ run, serverTime, predecessor, subject, control, flow }) => {
            requireFreshEligibility(subject, serverTime);
            requireAllowedPhase(predecessor, input.allowedPhases);
            const authenticated = input.phase === SESSION_PHASES.authenticated;
            await insertSession(run, input, subject, control, serverTime, {
                originalIssuedAt: authenticated ? serverTime : predecessor.originalIssuedAt,
                phaseStartedAt: serverTime,
                expiresAt: authenticated
                    ? new Date(serverTime.getTime() + AUTHENTICATED_LIFETIME_MS)
                    : predecessor.expiresAt,
                faceRequired: predecessor.faceRequired,
                registrationRequired: predecessor.registrationRequired,
            });
            await rotateOut(run, predecessor.sessionId, input.sessionId, serverTime);
            await migrateFlow(run, predecessor.sessionId, input.sessionId, serverTime);
        });
    }

    async function reserveFaceFlow(input) {
        return transact(async (run) => {
            const context = await expectedContext(run, input);
            if (isPostCommitError(context)) return context;
            requireFreshEligibility(context.subject, context.serverTime);
            requireAllowedPhase(context.predecessor, input.allowedPhases);
            const existing = context.flow;
            if (existing) throw authorityConflict('face-challenge-active');
            await run(`
                /* session-authority:reserve-face-flow */
                INSERT INTO dbo.learning_session_flow (
                    flow_id,
                    subject_id,
                    current_session_id,
                    registration_state,
                    challenge_state,
                    created_at,
                    updated_at
                ) VALUES (
                    @flowId,
                    @subjectId,
                    @sessionId,
                    @registrationState,
                    'creating',
                    @serverTime,
                    @serverTime
                );
            `, {
                flowId: input.flowId,
                subjectId: context.predecessor.subjectId,
                sessionId: context.predecessor.sessionId,
                registrationState: databaseRegistrationState(input.registrationState),
                serverTime: context.serverTime,
            });
            return {
                flowId: input.flowId,
                session: context.predecessor,
                serverTime: context.serverTime,
            };
        });
    }

    async function markFaceFlowReconciliation({ flowId, registrationReconciliationRequired = false }) {
        return transact(async (run) => {
            await selectControl(run, true);
            await selectFlowById(run, flowId, true, true);
            const serverTime = await readServerTime(run);
            const reconciled = await run(`
                /* session-authority:mark-face-flow-reconciliation */
                UPDATE dbo.learning_session_flow WITH (UPDLOCK, SERIALIZABLE)
                SET registration_state = CASE
                        WHEN @registrationReconciliationRequired = 1
                            THEN 'reconciliation-required'
                        ELSE registration_state
                    END,
                    challenge_state = 'reconciliation-required',
                    updated_at = @serverTime
                WHERE flow_id = @flowId
                    AND challenge_state IN ('creating', 'active', 'reconciliation-required');
            `, { flowId, registrationReconciliationRequired, serverTime });
            if (affectedRows(reconciled) !== 1) {
                throw authorityUnavailable('face-flow-reconciliation-unavailable');
            }
            return { serverTime };
        });
    }

    async function bindFaceChallengeAndRotate(input) {
        validateSessionCandidate(input);
        return rotateExpectedSession(input, async ({ run, serverTime, predecessor, subject, control, flow }) => {
            requireFreshEligibility(subject, serverTime);
            requireAllowedPhase(predecessor, input.allowedPhases);
            if (
                !flow
                || flow.flowId !== input.flowId
                || flow.challengeState !== 'creating'
            ) throw authorityConflict('face-flow-not-reserved');
            await insertSession(run, { ...input, phase: SESSION_PHASES.facePending }, subject, control, serverTime, {
                originalIssuedAt: predecessor.originalIssuedAt,
                phaseStartedAt: serverTime,
                expiresAt: predecessor.expiresAt,
                faceRequired: predecessor.faceRequired,
                registrationRequired: predecessor.registrationRequired,
            });
            await rotateOut(run, predecessor.sessionId, input.sessionId, serverTime);
            const updated = await run(`
                /* session-authority:bind-face-challenge */
                UPDATE dbo.learning_session_flow
                SET current_session_id = @sessionId,
                    challenge_session_id = @sessionId,
                    registration_state = 'registered',
                    challenge_state = 'active',
                    encrypted_provider_challenge_reference = @encryptedChallenge,
                    provider_reference_encryption_key_id = @challengeKeyId,
                    challenge_created_at = @serverTime,
                    updated_at = @serverTime
                WHERE flow_id = @flowId
                    AND current_session_id = @expectedSessionId
                    AND challenge_state = 'creating';
            `, {
                flowId: flow.flowId,
                expectedSessionId: predecessor.sessionId,
                sessionId: input.sessionId,
                encryptedChallenge: input.encryptedChallenge,
                challengeKeyId: input.challengeKeyId,
                serverTime,
            });
            if (affectedRows(updated) !== 1) throw authorityConflict('face-flow-compare-and-replace');
        });
    }

    async function readFaceFlow(input) {
        return transact(async (run) => {
            const context = await expectedContext(run, input);
            if (isPostCommitError(context)) return context;
            requireFreshEligibility(context.subject, context.serverTime);
            const flow = context.flow;
            if (context.predecessor.phase === SESSION_PHASES.authenticated) {
                if (
                    input.allowConsumed !== true
                    || !context.predecessor.faceRequired
                    || !flow
                    || flow.challengeState !== 'consumed'
                    || !isValidDate(flow.consumedAt)
                ) throw forbiddenAuthority('face-completion-not-applicable');
                return {
                    session: context.predecessor,
                    flow,
                    subject: context.subject,
                    control: context.control,
                    serverTime: context.serverTime,
                };
            }
            if (context.predecessor.phase !== SESSION_PHASES.facePending) throw forbiddenAuthority('wrong-phase');
            if (!flow || flow.challengeState !== 'active' || !Buffer.isBuffer(flow.encryptedChallenge)) {
                throw authorityConflict('face-challenge-unavailable');
            }
            return {
                session: context.predecessor,
                flow,
                subject: context.subject,
                control: context.control,
                serverTime: context.serverTime,
            };
        });
    }

    async function completeFaceSuccessAndRotate(input) {
        validateSessionCandidate(input);
        return rotateExpectedSession(input, async ({ run, serverTime, predecessor, subject, control, flow }) => {
            requireFreshEligibility(subject, serverTime);
            if (predecessor.phase !== SESSION_PHASES.facePending) throw forbiddenAuthority('wrong-phase');
            if (!flow || flow.challengeState !== 'active') throw authorityConflict('face-challenge-unavailable');
            await insertSession(run, { ...input, phase: SESSION_PHASES.authenticated }, subject, control, serverTime, {
                originalIssuedAt: serverTime,
                phaseStartedAt: serverTime,
                expiresAt: new Date(serverTime.getTime() + AUTHENTICATED_LIFETIME_MS),
                faceRequired: predecessor.faceRequired,
                registrationRequired: predecessor.registrationRequired,
            });
            await rotateOut(run, predecessor.sessionId, input.sessionId, serverTime);
            const consumed = await run(`
                /* session-authority:complete-face-success */
                UPDATE dbo.learning_session_flow
                SET current_session_id = @sessionId,
                    challenge_state = 'consumed',
                    challenge_resolved_at = @serverTime,
                    updated_at = @serverTime
                WHERE flow_id = @flowId
                    AND current_session_id = @expectedSessionId
                    AND challenge_state = 'active';
            `, {
                flowId: flow.flowId,
                expectedSessionId: predecessor.sessionId,
                sessionId: input.sessionId,
                serverTime,
            });
            if (affectedRows(consumed) !== 1) throw authorityConflict('face-flow-compare-and-replace');
        });
    }

    async function completeFaceFailure(input) {
        return transact(async (run) => {
            const context = await expectedContext(run, input);
            if (isPostCommitError(context)) return context;
            if (context.predecessor.phase !== SESSION_PHASES.facePending) throw forbiddenAuthority('wrong-phase');
            const flow = context.flow;
            if (!flow || flow.challengeState !== 'active') throw authorityConflict('face-challenge-unavailable');
            const revoked = await revokeSession(run, context.predecessor.sessionId, context.serverTime, 'face-factor-failed');
            if (!revoked) throw authorityConflict('session-compare-and-replace');
            const failed = await run(`
                /* session-authority:complete-face-failure */
                UPDATE dbo.learning_session_flow
                SET challenge_state = 'failed',
                    challenge_resolved_at = @serverTime,
                    updated_at = @serverTime
                WHERE flow_id = @flowId
                    AND current_session_id = @sessionId
                    AND challenge_state = 'active';
            `, { flowId: flow.flowId, sessionId: context.predecessor.sessionId, serverTime: context.serverTime });
            if (affectedRows(failed) !== 1) throw authorityConflict('face-flow-compare-and-replace');
            return {
                session: await selectSession(run, context.predecessor.sessionId, true),
                serverTime: context.serverTime,
            };
        });
    }

    async function disableLegacyAuthority({ subjectId, reason }) {
        if (reason !== 'legacy-handle-leak') {
            throw new TypeError('Legacy cutoff reason is unsupported');
        }
        return transact(async (run) => {
            const control = await selectControl(run, true);
            if (!control.legacyCompatibilityEnforcementEnabled) {
                throw authorityUnavailable('legacy-enforcement-required');
            }
            const subject = await selectSubject(run, subjectId, true);
            const serverTime = await readServerTime(run);
            if (subject.legacyAuthorityDisabledAt === null) {
                const cutoff = await run(`
                    /* session-authority:disable-legacy-authority */
                    UPDATE dbo.learning_subject
                    SET legacy_authority_disabled_at = @serverTime
                    WHERE subject_id = @subjectId
                        AND legacy_authority_disabled_at IS NULL;
                `, { subjectId, serverTime });
                if (affectedRows(cutoff) !== 1) {
                    throw authorityConflict('subject-adoption-compare-and-replace');
                }
            }
            await revokeSubjectLegacyBindings(run, subjectId, serverTime, reason);
            return { subject: await selectSubject(run, subjectId, true), serverTime };
        });
    }

    async function logout({ verifierKeyId, verifier }) {
        return transact(async (run) => {
            const authority = await selectAuthorityByVerifier(run, verifierKeyId, verifier, true);
            const serverTime = await readServerTime(run);
            if (authority.control.incidentState !== 'normal') {
                throw authorityUnavailable('authority-incident');
            }
            if (!authority.control.targetRoutesEnabled) {
                throw authorityUnavailable('target-routes-disabled');
            }
            if (!authority.session) return { revoked: false, serverTime };
            const state = evaluateAuthority(authority, serverTime);
            if (!state.active) {
                if (state.expired) await markExpired(run, authority.session.sessionId);
                if (state.ineligible) {
                    await enforceStoredIneligibility(run, authority.subject, serverTime);
                }
                if (state.unavailable) throw authorityUnavailable(state.reason);
                return { revoked: false, serverTime };
            }
            const revoked = await revokeSession(run, authority.session.sessionId, serverTime, 'logout');
            return { revoked, serverTime };
        });
    }

    async function revokeAll(input) {
        validateRevocationReason(input.reason || 'revoke-all');
        return transact(async (run) => {
            const context = await expectedContext(run, input);
            if (isPostCommitError(context)) return context;
            if (context.predecessor.phase !== SESSION_PHASES.authenticated) throw forbiddenAuthority('wrong-phase');
            const subjectUpdate = await run(`
                /* session-authority:revoke-all:subject */
                UPDATE dbo.learning_subject
                SET subject_session_epoch = subject_session_epoch + 1
                WHERE subject_id = @subjectId
                    AND subject_session_epoch = @subjectEpoch;
            `, { subjectId: context.subject.subjectId, subjectEpoch: context.subject.sessionEpoch });
            if (affectedRows(subjectUpdate) !== 1) throw authorityConflict('subject-epoch-compare-and-replace');
            await revokeSubjectSessions(
                run,
                context.subject.subjectId,
                context.serverTime,
                input.reason || 'revoke-all',
            );
            await revokeSubjectLegacyBindings(
                run,
                context.subject.subjectId,
                context.serverTime,
                input.reason || 'revoke-all',
            );
            return {
                subject: await selectSubject(run, context.subject.subjectId, true),
                serverTime: context.serverTime,
            };
        });
    }

    async function revokeSubject({ subjectId, reason }) {
        validateRevocationReason(reason);
        return transact(async (run) => {
            const control = await selectControl(run, true);
            if (!control.legacyCompatibilityEnforcementEnabled) {
                throw authorityUnavailable('legacy-enforcement-required');
            }
            await selectSubject(run, subjectId, true);
            const serverTime = await readServerTime(run);
            const subjectUpdate = await run(`
                /* session-authority:revoke-subject:subject */
                UPDATE dbo.learning_subject WITH (UPDLOCK, SERIALIZABLE)
                SET subject_session_epoch = subject_session_epoch + 1
                WHERE subject_id = @subjectId;
            `, { subjectId });
            if (affectedRows(subjectUpdate) !== 1) throw authorityUnavailable('subject-mapping-integrity');
            await revokeSubjectSessions(run, subjectId, serverTime, reason);
            await revokeSubjectLegacyBindings(run, subjectId, serverTime, reason);
            return { subject: await selectSubject(run, subjectId, true), serverTime };
        });
    }

    async function admitUnboundLegacyIssuance(input) {
        requireBoundLoginLookupKeyId(input.loginLookupKeyId);
        validateLegacyLifetime(input.issuedAt, input.expiresAt);
        return transact(async (run) => {
            const control = await selectControl(run, true);
            const result = await run(`
                /* session-authority:admit-unbound-legacy-issuance */
                SELECT
                    s.subject_id AS subjectId,
                    s.legacy_authority_disabled_at AS legacyAuthorityDisabledAt
                FROM dbo.learning_subject AS s WITH (UPDLOCK, HOLDLOCK)
                WHERE s.login_lookup_key_id = @loginLookupKeyId COLLATE Latin1_General_100_BIN2
                    AND s.login_lookup_token = @loginLookupToken;
            `, pick(input, ['loginLookupKeyId', 'loginLookupToken']));
            if (result.recordset.length > 1) throw authorityUnavailable('subject-mapping-integrity');
            const serverTime = await readServerTime(run);
            requireUnboundLegacyIssuanceControl(control, serverTime);
            requireLegacyCandidateWithinSunset(control, input.expiresAt);
            if (input.issuedAt > serverTime || input.expiresAt <= serverTime) {
                throw authorityUnavailable('legacy-issuance-time-invalid');
            }
            const subject = result.recordset[0];
            if (subject && subject.legacyAuthorityDisabledAt !== null) {
                throw authorityConflict('target-authority-established');
            }
            return { admitted: true, control, serverTime };
        });
    }

    async function bindLegacy(input) {
        validateLegacyLifetime(input.issuedAt, input.expiresAt);
        validateSubjectCredentialExpectation(input);
        return transact(async (run) => {
            const control = await selectControl(run, true);
            const existing = await selectLegacy(run, input.verifierKeyId, input.verifier, true, false);
            const subject = await selectSubject(run, input.subjectId, true);
            const serverTime = await readServerTime(run);
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
            if (!control.legacyLedgerSeedingEnabled) throw authorityUnavailable('legacy-ledger-seeding-disabled');
            requireSubjectCredentialExpectation(subject, input);
            if (control.legacyCompatibilityEnforcementEnabled) {
                requireFreshEligibility(subject, serverTime);
            }
            if (subject.legacyAuthorityDisabledAt !== null) throw authorityConflict('target-authority-established');
            if (existing) {
                if (!sameLegacyBinding(existing, input, true)) throw authorityUnavailable('legacy-binding-integrity');
                if (existing.compatibilityState === 'incident' || existing.incidentAt !== null) {
                    throw authorityUnavailable('legacy-binding-integrity');
                }
                if (
                    existing.compatibilityState !== 'active'
                    || existing.revokedAt !== null
                ) throw authorityConflict('legacy-binding-terminal');
                return { binding: existing, idempotent: true, serverTime };
            }
            await run(`
                /* session-authority:bind-legacy */
                INSERT INTO dbo.legacy_session_compatibility (
                    compatibility_id,
                    legacy_handle_verifier,
                    verifier_key_id,
                    subject_id,
                    original_issued_at,
                    original_expires_at,
                    compatibility_state,
                    created_at
                ) VALUES (
                    @legacyCompatibilityId,
                    @verifier,
                    @verifierKeyId,
                    @subjectId,
                    @issuedAt,
                    @expiresAt,
                    'active',
                    @serverTime
                );
            `, { ...pick(input, [
                'legacyCompatibilityId',
                'verifier',
                'verifierKeyId',
                'subjectId',
                'issuedAt',
                'expiresAt',
            ]), serverTime });
            return {
                binding: await selectLegacy(run, input.verifierKeyId, input.verifier, true, true),
                idempotent: false,
                serverTime,
            };
        });
    }

    async function authorizeLegacy(input) {
        validateLegacyLifetime(input.issuedAt, input.expiresAt);
        return transact(async (run) => {
            const control = await selectControl(run, true);
            const binding = await selectLegacy(run, input.verifierKeyId, input.verifier, true, false);
            const subject = binding && control.legacyCompatibilityEnforcementEnabled
                ? await selectSubject(run, binding.subjectId, true)
                : null;
            const serverTime = await readServerTime(run);
            if (control.incidentState !== 'normal') throw authorityUnavailable('authority-incident');
            if (
                !control.legacyAcceptanceEnabled
                || (control.hardSunsetAt !== null && control.hardSunsetAt <= serverTime)
            ) throw invalidAuthority('legacy-acceptance-disabled');
            if (!binding) {
                if (control.legacyCompatibilityEnforcementEnabled) {
                    throw invalidAuthority('legacy-binding-missing');
                }
                if (input.issuedAt > serverTime || input.expiresAt <= serverTime) {
                    throw invalidAuthority('legacy-binding-terminal');
                }
                return { unbound: true, control, serverTime };
            }
            if (!sameLegacyBinding(binding, input, false)) throw authorityUnavailable('legacy-binding-integrity');
            if (!isValidDate(binding.issuedAt) || binding.issuedAt > serverTime) {
                throw authorityUnavailable('legacy-binding-integrity');
            }
            if (!isValidDate(binding.expiresAt)) throw authorityUnavailable('legacy-binding-integrity');
            if (binding.compatibilityState === 'incident') throw authorityUnavailable('legacy-binding-integrity');
            if (!control.legacyCompatibilityEnforcementEnabled) {
                if (binding.expiresAt <= serverTime) throw invalidAuthority('legacy-binding-terminal');
                if (
                    binding.compatibilityState !== 'active'
                    || binding.revokedAt !== null
                ) throw invalidAuthority('legacy-binding-terminal');
                return { unbound: true, control, serverTime };
            }
            if (binding.compatibilityState !== 'active' || binding.revokedAt !== null) {
                throw invalidAuthority('legacy-binding-terminal');
            }
            if (subject.legacyAuthorityDisabledAt !== null) throw invalidAuthority('target-authority-established');
            const subjectState = evaluateSubjectEligibility(subject, serverTime);
            const bindingExpired = binding.expiresAt <= serverTime;
            const entitlementExpired = !subjectState.eligible && subjectState.ineligible === true;
            if (
                entitlementExpired
                && (!bindingExpired || subject.entitlementExpiresAt <= binding.expiresAt)
            ) {
                await enforceStoredIneligibility(run, subject, serverTime);
                return postCommitError(forbiddenAuthority('ineligible'));
            }
            if (bindingExpired) throw invalidAuthority('legacy-binding-terminal');
            if (!subjectState.eligible) {
                if (subjectState.unavailable) throw authorityUnavailable(subjectState.reason);
                await enforceStoredIneligibility(run, subject, serverTime);
                return postCommitError(forbiddenAuthority('ineligible'));
            }
            return { binding, subject, control, serverTime };
        });
    }

    async function close() {
        const pending = poolPromise;
        poolPromise = undefined;
        if (!pending) return;
        try {
            const pool = await pending;
            await pool.close();
        } catch (error) {
            throw normalizeDriverError(error);
        }
    }

    return Object.freeze({
        admitUnboundLegacyIssuance,
        authorizeLegacy,
        bindFaceChallengeAndRotate,
        bindLegacy,
        close,
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
        revokeForCredentialChange,
        revokeForIneligibility,
        revokeSubject,
        rotateSession,
        transitionControl,
        updateEligibility,
    });

    async function rotateExpectedSession(input, transition) {
        return transact(async (run) => {
            const context = await expectedContext(run, input);
            if (isPostCommitError(context)) return context;
            await transition({ run, ...context });
            return {
                session: await selectSession(run, input.sessionId, true),
                subject: await selectSubject(run, context.subject.subjectId, true),
                control: context.control,
                serverTime: context.serverTime,
            };
        });
    }

    async function expectedContext(run, input) {
        const control = await selectControl(run, true);
        const predecessor = await selectSession(run, input.expectedSessionId, true);
        if (!predecessor || input.expectedVersion !== 1 || !ACTIVE_PHASES.includes(predecessor.phase)) {
            throw authorityConflict('session-compare-and-replace');
        }
        const subject = await selectSubject(run, predecessor.subjectId, true);
        const flow = await selectFlow(run, predecessor.sessionId, true, false);
        const serverTime = await readServerTime(run);
        const state = evaluateAuthority(
            { session: predecessor, subject, control, flow },
            serverTime,
        );
        if (state.unavailable) throw authorityUnavailable(state.reason);
        if (!state.active) {
            if (state.expired) {
                await markExpired(run, predecessor.sessionId);
                return postCommitError(invalidAuthority('expired'));
            }
            if (state.ineligible) {
                await enforceStoredIneligibility(run, subject, serverTime);
                return postCommitError(forbiddenAuthority('ineligible'));
            }
            throw invalidAuthority(state.reason);
        }
        return { serverTime, predecessor, subject, control, flow };
    }

    async function selectControl(run, locking, { allowRotatableKeyMismatch = false } = {}) {
        const hint = locking === 'update'
            ? ' WITH (UPDLOCK, HOLDLOCK)'
            : (locking ? ' WITH (HOLDLOCK)' : '');
        const result = await run(`
            /* session-authority:select-control */
            SELECT ${CONTROL_SELECT},
                CASE WHEN c.login_lookup_key_id IS NOT NULL
                    AND c.login_lookup_key_commitment IS NOT NULL THEN 1 ELSE 0 END
                    AS loginLookupKeyInitialized,
                CASE WHEN c.login_lookup_key_id = @loginLookupKeyId COLLATE Latin1_General_100_BIN2
                    AND DATALENGTH(c.login_lookup_key_id) = DATALENGTH(@loginLookupKeyId)
                    AND c.login_lookup_key_commitment = @loginLookupKeyCommitment
                    THEN 1 ELSE 0 END AS loginLookupKeyMatches,
                CASE WHEN c.account_mapping_key_id = @accountMappingKeyId COLLATE Latin1_General_100_BIN2
                    AND DATALENGTH(c.account_mapping_key_id) = DATALENGTH(@accountMappingKeyId)
                    AND c.account_mapping_key_commitment = @accountMappingKeyCommitment
                    THEN 1 ELSE 0 END AS accountMappingKeyMatches,
                CASE WHEN c.keyset_login_lookup_key_id = @keysetLoginLookupKeyId COLLATE Latin1_General_100_BIN2
                    AND DATALENGTH(c.keyset_login_lookup_key_id) = DATALENGTH(@keysetLoginLookupKeyId)
                    AND c.keyset_login_lookup_key_commitment = @keysetLoginLookupKeyCommitment
                    THEN 1 ELSE 0 END AS keysetLoginLookupKeyMatches,
                CASE WHEN c.keyset_account_mapping_key_id = @keysetAccountMappingKeyId COLLATE Latin1_General_100_BIN2
                    AND DATALENGTH(c.keyset_account_mapping_key_id) = DATALENGTH(@keysetAccountMappingKeyId)
                    AND c.keyset_account_mapping_key_commitment = @keysetAccountMappingKeyCommitment
                    THEN 1 ELSE 0 END AS keysetAccountMappingKeyMatches,
                CASE WHEN c.authority_keyset_commitment IS NOT NULL
                    AND c.keyset_login_lookup_key_id IS NOT NULL
                    AND c.keyset_account_mapping_key_id IS NOT NULL
                    AND c.target_verifier_key_id IS NOT NULL
                    AND c.legacy_compatibility_key_id IS NOT NULL
                    AND c.credential_fingerprint_key_id IS NOT NULL
                    AND c.face_challenge_key_id IS NOT NULL
                    THEN 1 ELSE 0 END AS authorityKeysetInitialized,
                CASE WHEN c.authority_keyset_commitment = @authorityKeysetCommitment
                    THEN 1 ELSE 0 END AS authorityKeysetAggregateMatches,
                CASE WHEN c.target_verifier_key_id = @targetVerifierKeyId COLLATE Latin1_General_100_BIN2
                    AND DATALENGTH(c.target_verifier_key_id) = DATALENGTH(@targetVerifierKeyId)
                    AND c.target_verifier_key_commitment = @targetVerifierKeyCommitment
                    THEN 1 ELSE 0 END AS targetVerifierKeyMatches,
                CASE WHEN c.legacy_compatibility_key_id = @legacyCompatibilityKeyId COLLATE Latin1_General_100_BIN2
                    AND DATALENGTH(c.legacy_compatibility_key_id) = DATALENGTH(@legacyCompatibilityKeyId)
                    AND c.legacy_compatibility_key_commitment = @legacyCompatibilityKeyCommitment
                    THEN 1 ELSE 0 END AS legacyCompatibilityKeyMatches,
                CASE WHEN c.credential_fingerprint_key_id = @credentialFingerprintKeyId COLLATE Latin1_General_100_BIN2
                    AND DATALENGTH(c.credential_fingerprint_key_id) = DATALENGTH(@credentialFingerprintKeyId)
                    AND c.credential_fingerprint_key_commitment = @credentialFingerprintKeyCommitment
                    THEN 1 ELSE 0 END AS credentialFingerprintKeyMatches,
                CASE WHEN c.face_challenge_key_id = @faceChallengeKeyId COLLATE Latin1_General_100_BIN2
                    AND DATALENGTH(c.face_challenge_key_id) = DATALENGTH(@faceChallengeKeyId)
                    AND c.face_challenge_key_commitment = @faceChallengeKeyCommitment
                    THEN 1 ELSE 0 END AS faceChallengeKeyMatches,
                CASE WHEN c.legacy_signing_key_id IS NOT NULL
                    AND c.legacy_signing_key_commitment IS NOT NULL
                    THEN 1 ELSE 0 END AS legacySigningKeyInitialized,
                CASE WHEN c.legacy_signing_key_id = @legacySigningKeyId COLLATE Latin1_General_100_BIN2
                    AND DATALENGTH(c.legacy_signing_key_id) = DATALENGTH(@legacySigningKeyId)
                    AND c.legacy_signing_key_commitment = @legacySigningKeyCommitment
                    THEN 1 ELSE 0 END AS legacySigningKeyMatches
            FROM dbo.session_authority_control AS c${hint}
            WHERE c.control_id = 1;
        `, {
            loginLookupKeyId: run.expectedLoginLookupKeyId || loginLookupKeyId,
            loginLookupKeyCommitment: run.expectedLoginLookupKeyCommitment
                || expectedLoginLookupKeyCommitment,
            accountMappingKeyId: (run.expectedAccountMappingKeyBinding
                || expectedAccountMappingKeyBinding).keyId,
            accountMappingKeyCommitment: (run.expectedAccountMappingKeyBinding
                || expectedAccountMappingKeyBinding).commitment,
            ...keysetSqlParameters(
                run.expectedAuthorityKeysetBinding || expectedAuthorityKeysetBinding,
            ),
            legacySigningKeyId: (run.expectedLegacySigningKeyBinding
                || expectedLegacySigningKeyBinding).keyId,
            legacySigningKeyCommitment: (run.expectedLegacySigningKeyBinding
                || expectedLegacySigningKeyBinding).commitment,
        });
        const row = exactlyOne(result.recordset, 'control-integrity');
        const control = run.expectedLoginLookupKeyId === undefined
            ? mapControl(row)
            : mapControlWithKeyFence(row, { allowRotatableKeyMismatch });
        if (run.expectedAuthorityGeneration !== undefined) {
            requireAuthorityGeneration(control, run.expectedAuthorityGeneration);
            run.authorityGenerationFenced = true;
        }
        return control;
    }

    async function requireExpectedNormalControl(run, expectedControlVersion) {
        const control = await selectControl(run, true);
        if (control.incidentState !== 'normal') throw authorityUnavailable('authority-incident');
        if (control.version !== expectedControlVersion) {
            throw authorityUnavailable('authority-control-changed');
        }
        return control;
    }

    async function selectSubject(run, subjectId, locking) {
        const result = await run(`
            /* session-authority:select-subject */
            SELECT ${SUBJECT_SELECT}
            FROM dbo.learning_subject AS s${locking ? ' WITH (UPDLOCK, HOLDLOCK)' : ''}
            WHERE s.subject_id = @subjectId;
        `, { subjectId });
        return mapSubject(exactlyOne(result.recordset, 'subject-mapping-integrity'));
    }

    async function selectSession(run, sessionId, locking) {
        const result = await run(`
            /* session-authority:select-session */
            SELECT ${SESSION_SELECT}
            FROM dbo.learning_session AS l${locking ? ' WITH (UPDLOCK, HOLDLOCK)' : ''}
            WHERE l.session_id = @sessionId;
        `, { sessionId });
        if (result.recordset.length > 1) throw authorityUnavailable('session-store-integrity');
        return result.recordset[0] ? mapSession(result.recordset[0]) : null;
    }

    async function selectAuthorityByVerifier(run, verifierKeyId, verifier, locking) {
        const control = await selectControl(run, locking);
        if (control.incidentState !== 'normal') return { session: null, subject: null, control };
        const sessionResult = await run(`
            /* session-authority:select-authority-by-verifier */
            SELECT ${SESSION_SELECT}
            FROM dbo.learning_session AS l${locking ? ' WITH (UPDLOCK, HOLDLOCK)' : ''}
            WHERE l.verifier_key_id = @verifierKeyId COLLATE Latin1_General_100_BIN2
                AND l.identifier_verifier = @verifier;
        `, { verifierKeyId, verifier });
        if (sessionResult.recordset.length > 1) throw authorityUnavailable('session-store-integrity');
        if (sessionResult.recordset.length === 0) return { session: null, subject: null, control };
        const session = mapSession(sessionResult.recordset[0]);
        const flow = await selectFlow(run, session.sessionId, locking, false);
        return {
            session,
            subject: await selectSubject(run, session.subjectId, locking),
            control,
            flow,
        };
    }

    async function selectFlow(run, sessionId, locking, required) {
        const result = await run(`
            /* session-authority:select-face-flow */
            SELECT ${FLOW_SELECT}
            FROM dbo.learning_session_flow AS f${locking ? ' WITH (UPDLOCK, HOLDLOCK)' : ''}
            WHERE f.current_session_id = @sessionId;
        `, { sessionId });
        if (result.recordset.length > 1) throw authorityUnavailable('face-flow-integrity');
        if (required && result.recordset.length !== 1) throw authorityUnavailable('face-flow-integrity');
        if (!result.recordset[0]) return null;
        const flow = mapFlow(result.recordset[0]);
        flow.challengeSession = flow.challengeSessionId === null
            ? null
            : await selectSession(run, flow.challengeSessionId, locking);
        if (flow.challengeSessionId !== null && flow.challengeSession === null) {
            throw authorityUnavailable('session-store-integrity');
        }
        return flow;
    }

    async function selectFlowById(run, flowId, locking, required) {
        const result = await run(`
            /* session-authority:select-face-flow-by-id */
            SELECT ${FLOW_SELECT}
            FROM dbo.learning_session_flow AS f${locking ? ' WITH (UPDLOCK, HOLDLOCK)' : ''}
            WHERE f.flow_id = @flowId;
        `, { flowId });
        if (result.recordset.length > 1) throw authorityUnavailable('face-flow-integrity');
        if (required && result.recordset.length !== 1) throw authorityUnavailable('face-flow-integrity');
        return result.recordset[0] ? mapFlow(result.recordset[0]) : null;
    }

    async function selectLegacy(run, verifierKeyId, verifier, locking, required) {
        const result = await run(`
            /* session-authority:select-legacy-binding */
            SELECT ${LEGACY_SELECT}
            FROM dbo.legacy_session_compatibility AS b${locking ? ' WITH (UPDLOCK, HOLDLOCK)' : ''}
            WHERE b.verifier_key_id = @verifierKeyId COLLATE Latin1_General_100_BIN2
                AND b.legacy_handle_verifier = @verifier;
        `, { verifierKeyId, verifier });
        if (result.recordset.length > 1) throw authorityUnavailable('legacy-binding-integrity');
        if (required && result.recordset.length !== 1) throw authorityUnavailable('legacy-binding-integrity');
        return result.recordset[0] ? mapLegacy(result.recordset[0]) : null;
    }

    async function insertSession(run, input, subject, control, serverTime, times) {
        await run(`
            /* session-authority:insert-session */
            INSERT INTO dbo.learning_session (
                session_id,
                identifier_verifier,
                verifier_key_id,
                subject_id,
                phase,
                original_issued_at,
                phase_started_at,
                absolute_expires_at,
                face_auth_required,
                registration_required,
                subject_epoch_snapshot,
                credential_version_snapshot,
                global_epoch_snapshot,
                authority_generation_snapshot,
                created_at
            ) VALUES (
                @sessionId,
                @verifier,
                @verifierKeyId,
                @subjectId,
                @phase,
                @originalIssuedAt,
                @phaseStartedAt,
                @expiresAt,
                @faceRequired,
                @registrationRequired,
                @subjectEpochSnapshot,
                @credentialVersionSnapshot,
                @globalEpochSnapshot,
                @authorityGenerationSnapshot,
                @serverTime
            );
        `, {
            sessionId: input.sessionId,
            verifier: input.verifier,
            verifierKeyId: input.verifierKeyId,
            subjectId: subject.subjectId,
            phase: input.phase,
            originalIssuedAt: times.originalIssuedAt,
            phaseStartedAt: times.phaseStartedAt,
            expiresAt: times.expiresAt,
            faceRequired: times.faceRequired === undefined ? input.faceRequired : times.faceRequired,
            registrationRequired: times.registrationRequired === undefined
                ? input.registrationRequired
                : times.registrationRequired,
            subjectEpochSnapshot: subject.sessionEpoch,
            credentialVersionSnapshot: subject.credentialVersion,
            globalEpochSnapshot: control.globalSessionEpoch,
            authorityGenerationSnapshot: control.authorityGeneration,
            serverTime,
        });
    }

    async function rotateOut(run, expectedSessionId, replacementSessionId, serverTime) {
        const result = await run(`
            /* session-authority:rotate-out */
            UPDATE dbo.learning_session
            SET phase = 'rotated-out',
                replacement_session_id = @replacementSessionId
            WHERE session_id = @expectedSessionId
                AND phase IN ${ACTIVE_PHASE_SQL}
                AND revoked_at IS NULL
                AND replacement_session_id IS NULL
                AND absolute_expires_at > @serverTime;
        `, { expectedSessionId, replacementSessionId, serverTime });
        if (affectedRows(result) !== 1) throw authorityConflict('session-compare-and-replace');
    }

    async function migrateFlow(run, expectedSessionId, sessionId, serverTime) {
        await run(`
            /* session-authority:migrate-face-flow */
            UPDATE dbo.learning_session_flow
            SET current_session_id = @sessionId,
                updated_at = @serverTime
            WHERE current_session_id = @expectedSessionId;
        `, { expectedSessionId, sessionId, serverTime });
    }

    async function revokeSession(run, sessionId, serverTime, reason) {
        const result = await run(`
            /* session-authority:revoke-session */
            UPDATE dbo.learning_session
            SET phase = 'revoked',
                revoked_at = @serverTime,
                revocation_reason = @reason
            WHERE session_id = @sessionId
                AND phase IN ${ACTIVE_PHASE_SQL}
                AND revoked_at IS NULL
                AND replacement_session_id IS NULL;
        `, { sessionId, serverTime, reason });
        return affectedRows(result) === 1;
    }

    async function revokeSubjectSessions(run, subjectId, serverTime, reason) {
        await run(`
            /* session-authority:revoke-subject-sessions */
            UPDATE dbo.learning_session
            SET phase = 'revoked',
                revoked_at = @serverTime,
                revocation_reason = @reason
            WHERE subject_id = @subjectId
                AND phase IN ${ACTIVE_PHASE_SQL}
                AND revoked_at IS NULL
                AND replacement_session_id IS NULL;
        `, { subjectId, serverTime, reason });
    }

    async function revokeSubjectLegacyBindings(run, subjectId, serverTime, reason) {
        await run(`
            /* session-authority:revoke-subject-legacy-bindings */
            UPDATE dbo.legacy_session_compatibility
            SET compatibility_state = 'revoked',
                revoked_at = @serverTime,
                revocation_reason = @reason
            WHERE subject_id = @subjectId
                AND compatibility_state = 'active'
                AND revoked_at IS NULL
                AND incident_at IS NULL;
        `, { subjectId, serverTime, reason });
    }

    async function quarantineUnresolvedFlowsForKeyRecovery(run, serverTime) {
        await run(`
            /* session-authority:key-recovery:quarantine-unresolved-flows */
            UPDATE dbo.learning_session_flow
            SET challenge_state = 'reconciliation-required',
                updated_at = @serverTime
            WHERE challenge_state IN ('creating', 'active', 'reconciliation-required');
        `, { serverTime });
        await run(`
            /* session-authority:key-recovery:revoke-unresolved-flow-sessions */
            UPDATE l
            SET phase = 'revoked',
                revoked_at = @serverTime,
                revocation_reason = 'key-recovery'
            FROM dbo.learning_session AS l
            INNER JOIN dbo.learning_session_flow AS f
                ON f.current_session_id = l.session_id
            WHERE f.challenge_state = 'reconciliation-required'
                AND l.phase IN ${ACTIVE_PHASE_SQL}
                AND l.revoked_at IS NULL
                AND l.replacement_session_id IS NULL;
        `, { serverTime });
    }

    async function incidentLegacyBindingsForKeyRecovery(run, serverTime) {
        await run(`
            /* session-authority:key-recovery:retire-legacy-bindings */
            UPDATE dbo.legacy_session_compatibility
            SET compatibility_state = 'incident',
                incident_at = @serverTime,
                incident_code = 'key-recovery'
            WHERE compatibility_state = 'active'
                AND revoked_at IS NULL
                AND incident_at IS NULL;
        `, { serverTime });
    }

    async function markExpired(run, sessionId) {
        await run(`
            /* session-authority:mark-expired */
            UPDATE dbo.learning_session
            SET phase = 'expired'
            WHERE session_id = @sessionId
                AND phase IN ${ACTIVE_PHASE_SQL}
                AND revoked_at IS NULL
                AND replacement_session_id IS NULL;
        `, { sessionId });
    }

    async function enforceStoredIneligibility(run, subject, serverTime) {
        if (subject.eligibilityState === 'eligible') {
            const subjectUpdate = await run(`
                /* session-authority:enforce-stored-ineligibility */
                UPDATE dbo.learning_subject
                SET eligibility_state = 'ineligible',
                    eligibility_observed_at = @serverTime,
                    eligibility_revalidate_at = @serverTime,
                    subject_session_epoch = subject_session_epoch + 1
                WHERE subject_id = @subjectId
                    AND eligibility_state = 'eligible'
                    AND subject_session_epoch = @subjectEpoch;
            `, {
                subjectId: subject.subjectId,
                subjectEpoch: subject.sessionEpoch,
                serverTime,
            });
            if (affectedRows(subjectUpdate) !== 1) {
                throw authorityConflict('subject-epoch-compare-and-replace');
            }
        }
        await revokeSubjectSessions(run, subject.subjectId, serverTime, 'entitlement-expired');
        await revokeSubjectLegacyBindings(run, subject.subjectId, serverTime, 'entitlement-expired');
    }
}

function validateFactoryInput(
    sql,
    connectionString,
    options,
    expectedAuthorityGeneration,
    loginLookupKeyId,
    loginLookupKeyCommitment,
    accountMappingKeyBinding,
    authorityKeysetBinding,
    legacySigningKeyBinding,
) {
    if (!sql || typeof sql.ConnectionPool !== 'function') throw new TypeError('An injected SQL driver is required');
    if (typeof connectionString !== 'string' || connectionString.length === 0) {
        throw new TypeError('SQL connection string must be non-empty');
    }
    if (!options || typeof options !== 'object' || Array.isArray(options)) {
        throw new TypeError('SQL options must be an object');
    }
    if (!Number.isSafeInteger(expectedAuthorityGeneration) || expectedAuthorityGeneration < 1) {
        throw new TypeError('Expected authority generation must be a positive safe integer');
    }
    if (
        typeof loginLookupKeyId !== 'string'
        || !isValidKeyId(loginLookupKeyId)
    ) throw new TypeError('Login lookup key ID must be non-empty and bounded');
    if (!Buffer.isBuffer(loginLookupKeyCommitment) || loginLookupKeyCommitment.length !== 32) {
        throw new TypeError('Login lookup key commitment must be a 32-byte Buffer');
    }
    validateKeyBinding(accountMappingKeyBinding, 'Account mapping');
    validateKeyBinding(legacySigningKeyBinding, 'Legacy signing');
    if (
        !authorityKeysetBinding
        || !Buffer.isBuffer(authorityKeysetBinding.commitment)
        || authorityKeysetBinding.commitment.length !== 32
        || !authorityKeysetBinding.purposes
    ) throw new TypeError('Authority keyset binding must be complete');
    for (const purpose of [
        'targetVerifier',
        'legacyCompatibility',
        'loginLookup',
        'credentialFingerprint',
        'accountMappingEncryption',
        'faceChallengeEncryption',
    ]) validateKeyBinding(authorityKeysetBinding.purposes[purpose], `Authority ${purpose}`);
    const lookupPurpose = authorityKeysetBinding.purposes.loginLookup;
    const accountPurpose = authorityKeysetBinding.purposes.accountMappingEncryption;
    if (
        lookupPurpose.keyId !== loginLookupKeyId
        || accountPurpose.keyId !== accountMappingKeyBinding.keyId
    ) throw new TypeError('Immutable key IDs must match the authority keyset descriptors');
}

function validateKeyBinding(binding, label) {
    if (
        !binding
        || typeof binding.keyId !== 'string'
        || !isValidKeyId(binding.keyId)
        || !Buffer.isBuffer(binding.commitment)
        || binding.commitment.length !== 32
    ) throw new TypeError(`${label} key binding must be complete`);
}

function copyKeyBinding(binding) {
    return Object.freeze({ keyId: binding.keyId, commitment: Buffer.from(binding.commitment) });
}

function copyAuthorityKeysetBinding(binding) {
    return Object.freeze({
        commitment: Buffer.from(binding.commitment),
        purposes: Object.freeze(Object.fromEntries(Object.entries(binding.purposes).map(
            ([name, value]) => [name, copyKeyBinding(value)],
        ))),
    });
}

function keysetSqlParameters(binding) {
    return {
        authorityKeysetCommitment: binding.commitment,
        keysetLoginLookupKeyId: binding.purposes.loginLookup.keyId,
        keysetLoginLookupKeyCommitment: binding.purposes.loginLookup.commitment,
        keysetAccountMappingKeyId: binding.purposes.accountMappingEncryption.keyId,
        keysetAccountMappingKeyCommitment: binding.purposes.accountMappingEncryption.commitment,
        targetVerifierKeyId: binding.purposes.targetVerifier.keyId,
        targetVerifierKeyCommitment: binding.purposes.targetVerifier.commitment,
        legacyCompatibilityKeyId: binding.purposes.legacyCompatibility.keyId,
        legacyCompatibilityKeyCommitment: binding.purposes.legacyCompatibility.commitment,
        credentialFingerprintKeyId: binding.purposes.credentialFingerprint.keyId,
        credentialFingerprintKeyCommitment: binding.purposes.credentialFingerprint.commitment,
        faceChallengeKeyId: binding.purposes.faceChallengeEncryption.keyId,
        faceChallengeKeyCommitment: binding.purposes.faceChallengeEncryption.commitment,
    };
}

function requireAuthorityGeneration(control, expectedAuthorityGeneration) {
    if (control.authorityGeneration !== expectedAuthorityGeneration) {
        throw authorityUnavailable('authority-generation-mismatch');
    }
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

async function connectPool(sql, connectionString, options, onPoolError) {
    if (typeof sql.ConnectionPool.parseConnectionString !== 'function') {
        throw new TypeError('Injected SQL driver cannot parse connection strings');
    }
    const parsed = sql.ConnectionPool.parseConnectionString(connectionString);
    const configuration = {
        ...parsed,
        ...options,
        options: {
            ...(parsed.options || {}),
            ...(options.options || {}),
            encrypt: true,
            trustServerCertificate: false,
            useUTC: true,
        },
        pool: { ...(parsed.pool || {}), ...(options.pool || {}) },
    };
    const pool = new sql.ConnectionPool(configuration);
    if (typeof pool.on !== 'function') throw new TypeError('Injected SQL pool does not expose error events');
    pool.on('error', onPoolError);
    await pool.connect();
    return pool;
}

function createRequest(sql, owner) {
    if (typeof sql.Request === 'function') return new sql.Request(owner);
    if (owner && typeof owner.request === 'function') return owner.request();
    throw new TypeError('Injected SQL driver does not expose requests');
}

function createTransaction(sql, pool) {
    if (typeof sql.Transaction === 'function') return new sql.Transaction(pool);
    if (pool && typeof pool.transaction === 'function') return pool.transaction();
    throw new TypeError('Injected SQL driver does not expose transactions');
}

function serializableIsolation(sql) {
    return sql.ISOLATION_LEVEL && sql.ISOLATION_LEVEL.SERIALIZABLE
        ? sql.ISOLATION_LEVEL.SERIALIZABLE
        : 'SERIALIZABLE';
}

function normalizeDriverError(error) {
    if (isSessionAuthorityError(error)) return error;
    return authorityUnavailable('session-store-unavailable');
}

function createControlAssignments(changes) {
    const sql = [];
    const parameters = {};
    for (const [name, value] of Object.entries(changes)) {
        const parameterName = `change_${name}`;
        sql.push(`${CONTROL_CHANGE_COLUMNS[name]} = @${parameterName}`);
        parameters[parameterName] = value;
    }
    if (sql.length === 0) sql.push('control_id = control_id');
    return { sql, parameters };
}

function prepareAuthorityKeysetRecovery(
    current,
    changes,
    serverTime,
    expectedAuthorityGeneration,
    expectedAuthorityKeysetBinding,
    expectedLegacySigningKeyBinding,
) {
    const mismatches = current.keyMismatches || Object.freeze({
        aggregate: false,
        targetVerifier: false,
        legacyCompatibility: false,
        credentialFingerprint: false,
        faceChallengeEncryption: false,
        legacySigning: false,
    });
    const changedPurposes = ROTATABLE_KEY_PURPOSES
        .filter(([purpose]) => mismatches[purpose])
        .map(([purpose]) => purpose);
    if (mismatches.legacySigning) changedPurposes.push('legacySigning');
    const keysetLeafChanges = changedPurposes.filter((purpose) => purpose !== 'legacySigning');
    const anyMismatch = mismatches.aggregate
        || keysetLeafChanges.length > 0
        || mismatches.legacySigning;
    if (!anyMismatch) {
        return {
            active: false,
            changedPurposes: [],
            changes,
            retireLegacyBindings: false,
        };
    }
    if (
        (mismatches.aggregate || keysetLeafChanges.length > 0)
        && (!mismatches.aggregate || keysetLeafChanges.length === 0)
    ) {
        throw authorityUnavailable('authority-keyset-mismatch');
    }
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
    }
    if (retireLegacyBindings && current.legacyVerifierKeyIncidentAt === null) {
        stamped.legacyVerifierKeyIncidentAt = current.incidentRecordedAt;
    }
    return {
        active: true,
        changedPurposes,
        changes: stamped,
        retireLegacyBindings,
    };
}

function createAuthorityKeysetAssignments(binding, legacySigningBinding, changedPurposes) {
    const parameters = keysetSqlParameters(binding);
    if (changedPurposes.includes('legacySigning')) {
        parameters.legacySigningKeyId = legacySigningBinding.keyId;
        parameters.legacySigningKeyCommitment = legacySigningBinding.commitment;
    }
    return {
        parameters,
        sql: [
            'target_verifier_key_id = @targetVerifierKeyId',
            'target_verifier_key_commitment = @targetVerifierKeyCommitment',
            'legacy_compatibility_key_id = @legacyCompatibilityKeyId',
            'legacy_compatibility_key_commitment = @legacyCompatibilityKeyCommitment',
            'credential_fingerprint_key_id = @credentialFingerprintKeyId',
            'credential_fingerprint_key_commitment = @credentialFingerprintKeyCommitment',
            'face_challenge_key_id = @faceChallengeKeyId',
            'face_challenge_key_commitment = @faceChallengeKeyCommitment',
            'authority_keyset_commitment = @authorityKeysetCommitment',
            ...(changedPurposes.includes('legacySigning') ? [
                'legacy_signing_key_id = @legacySigningKeyId',
                'legacy_signing_key_commitment = @legacySigningKeyCommitment',
            ] : []),
        ],
    };
}

function validateControlChanges(changes) {
    if (!changes || typeof changes !== 'object' || Array.isArray(changes)) {
        throw new TypeError('Control changes must be an object');
    }
    for (const [name, value] of Object.entries(changes)) {
        if (!CONTROL_CHANGE_COLUMNS[name]) throw new TypeError('Unsupported control field');
        if ([
            'seedingHeartbeatOwnerId',
            'seedingHeartbeatAt',
            'seedingLeaseExpiresAt',
            'seedingContinuityVersion',
        ].includes(name)) throw new TypeError('Seeding continuity fields are store-owned');
        if (
            name === 'incidentCode'
            && value !== null
            && (typeof value !== 'string' || !/^[a-z0-9][a-z0-9-]{0,127}$/.test(value))
        ) throw new TypeError('Incident code must be a privacy-safe machine value');
        if (name === 'incidentState' && !CONTROL_INCIDENT_STATES.includes(value)) {
            throw new TypeError('Incident state is invalid');
        }
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
        && isValidDate(control.seedingStartedAt)
        && isValidDate(control.seedingHeartbeatAt)
        && isValidDate(control.seedingLeaseExpiresAt)
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
    if (current.legacyLedgerSeedingEnabled !== true && nextSeedingEnabled === true) {
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
    if (
        current.legacyCompatibilityEnforcementEnabled !== true
        && nextEnforcementEnabled === true
    ) {
        if (!isValidDate(current.seedingQualifiedAt)) stamped.seedingQualifiedAt = serverTime;
        stamped.legacyCompatibilityEnforcedAt = serverTime;
    }

    const nextLegacyIssuanceEnabled = Object.hasOwn(changes, 'legacyIssuanceEnabled')
        ? changes.legacyIssuanceEnabled
        : current.legacyIssuanceEnabled;
    if (current.legacyIssuanceEnabled === true && nextLegacyIssuanceEnabled === false) {
        stamped.legacyStopIssuanceAt = serverTime;
    }
    const nextLegacyAcceptanceEnabled = Object.hasOwn(changes, 'legacyAcceptanceEnabled')
        ? changes.legacyAcceptanceEnabled
        : current.legacyAcceptanceEnabled;
    if (current.legacyAcceptanceEnabled === true && nextLegacyAcceptanceEnabled === false) {
        stamped.legacyAcceptanceDisabledAt = serverTime;
    }
    return stamped;
}

function validateControlTransition(
    current,
    changes,
    serverTime,
    { keyRecoveryActive = false } = {},
) {
    const next = { ...current, ...changes };
    if (
        next.incidentState === 'recovering'
        && current.incidentState !== 'recovering'
        && !keyRecoveryActive
    ) throw forbiddenAuthority('recovering-requires-key-recovery');
    for (const name of ['authorityGeneration', 'globalSessionEpoch']) {
        if (!Number.isSafeInteger(next[name]) || next[name] < current[name]) {
            throw forbiddenAuthority(`irreversible-${name}`);
        }
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
            || !isValidDate(next.incidentRecordedAt)
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
            next.legacyIssuanceEnabled !== false
            || next.legacyAcceptanceEnabled !== false
            || !isValidDate(next.legacyStopIssuanceAt)
            || !isValidDate(next.legacyAcceptanceDisabledAt)
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
    if (
        (
            next.targetSessionIssuanceEnabled === true
            || next.subjectTargetAdoptionEnabled === true
            || next.dualStackStartedAt !== null
        )
        && next.legacyCompatibilityEnforcementEnabled !== true
    ) throw forbiddenAuthority('legacy-enforcement-required-before-target');
    if (next.targetSessionIssuanceEnabled === true || next.subjectTargetAdoptionEnabled === true) {
        if (!isValidDate(next.dualStackStartedAt) || !isValidDate(next.hardSunsetAt)) {
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
        next.legacyCompatibilityEnforcementEnabled === true
        && next.legacyIssuanceEnabled === true
        && next.legacyLedgerSeedingEnabled !== true
    ) throw forbiddenAuthority('legacy-seeding-required-during-issuance');
    if (next.legacyAcceptanceDisabledAt !== null) {
        if (
            !isValidDate(next.legacyStopIssuanceAt)
            || !isValidDate(next.legacyAcceptanceDisabledAt)
            || (
                next.legacyVerifierKeyIncidentAt === null
                && next.legacyAcceptanceDisabledAt.getTime() - next.legacyStopIssuanceAt.getTime()
                    < LEGACY_LIFETIME_MS
            )
        ) throw forbiddenAuthority('legacy-final-aging-incomplete');
    }
    if (
        next.hardSunsetAt !== null
        && next.legacyStopIssuanceAt !== null
        && (
            !isValidDate(next.hardSunsetAt)
            || !isValidDate(next.legacyStopIssuanceAt)
            || (
                next.legacyVerifierKeyIncidentAt === null
                && next.legacyStopIssuanceAt.getTime()
                    > next.hardSunsetAt.getTime() - LEGACY_LIFETIME_MS
            )
        )
    ) throw forbiddenAuthority('legacy-stop-too-late');
    if (current.incidentState !== 'normal' && next.incidentState === 'normal') {
        if (
            current.incidentState !== 'recovering'
            ||
            next.authorityGeneration !== current.authorityGeneration
            || next.globalSessionEpoch !== current.globalSessionEpoch
        ) throw forbiddenAuthority('incident-resume-requires-fenced-recovery');
        if (
            legacyAuthorityWasInScope(current, next)
            && (next.legacyIssuanceEnabled !== false || next.legacyAcceptanceEnabled !== false)
        ) throw forbiddenAuthority('incident-resume-requires-legacy-retirement');
        // SQL proves epoch retirement; deployment control must separately install fresh verifier keys.
    }
}

async function requireNoLiveFaceRecoveryAuthority(run) {
    const result = await run(`
        /* session-authority:recovery-face-authority-empty */
        SELECT CASE WHEN EXISTS (
            SELECT 1
            FROM dbo.learning_session_flow AS f WITH (UPDLOCK, HOLDLOCK)
            INNER JOIN dbo.learning_session AS l WITH (UPDLOCK, HOLDLOCK)
                ON l.session_id = f.current_session_id
                AND l.subject_id = f.subject_id
            WHERE f.challenge_state IN ('creating', 'active', 'reconciliation-required')
                AND l.phase IN ${ACTIVE_PHASE_SQL}
        ) THEN 1 ELSE 0 END AS liveFaceAuthorityExists;
    `);
    if (Boolean(exactlyOne(result.recordset, 'control-integrity').liveFaceAuthorityExists)) {
        throw forbiddenAuthority('incident-resume-face-reconciliation-required');
    }
}

function validateTargetControlTransition(current, next, serverTime) {
    const targetIssuanceActivated = current.targetSessionIssuanceEnabled !== true
        && next.targetSessionIssuanceEnabled === true;
    const subjectAdoptionActivated = current.subjectTargetAdoptionEnabled !== true
        && next.subjectTargetAdoptionEnabled === true;
    const targetAcceptanceActivated = current.dualStackStartedAt === null
        && next.dualStackStartedAt !== null;
    if (next.targetSessionIssuanceEnabled !== next.subjectTargetAdoptionEnabled) {
        throw forbiddenAuthority('target-activation-pair-required');
    }
    const targetEvidenceAbsent = next.targetSessionIssuanceStartedAt === null
        && next.subjectTargetAdoptionStartedAt === null
        && next.dualStackStartedAt === null
        && next.hardSunsetAt === null;
    const targetEvidenceComplete = isValidDate(next.targetSessionIssuanceStartedAt)
        && isValidDate(next.subjectTargetAdoptionStartedAt)
        && isValidDate(next.dualStackStartedAt)
        && isValidDate(next.hardSunsetAt)
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
    const seedingActivated = current.legacyLedgerSeedingEnabled !== true
        && next.legacyLedgerSeedingEnabled === true;
    const firstSeedingStart = current.legacyLedgerSeedingStartedAt === null
        && next.legacyLedgerSeedingStartedAt !== null;
    const continuityAdvanced = isValidDate(next.seedingStartedAt)
        && (
            !isValidDate(current.seedingStartedAt)
            || next.seedingStartedAt > current.seedingStartedAt
        );
    if (
        firstSeedingStart
        && (!seedingActivated || !sameInstant(next.legacyLedgerSeedingStartedAt, serverTime))
    ) throw forbiddenAuthority('legacy-seeding-time-provenance');
    if (
        continuityAdvanced
        && (
            next.legacyLedgerSeedingEnabled !== true
            || !sameInstant(next.seedingStartedAt, serverTime)
        )
    ) throw forbiddenAuthority('legacy-seeding-time-provenance');
    if (
        seedingActivated
        && (
            !isValidDate(next.legacyLedgerSeedingStartedAt)
            || !continuityAdvanced
            || !sameInstant(next.seedingStartedAt, serverTime)
        )
    ) throw forbiddenAuthority('legacy-seeding-time-provenance');

    const qualificationAdvanced = Object.hasOwn(changes, 'seedingQualifiedAt')
        && isValidDate(next.seedingQualifiedAt)
        && (
            !isValidDate(current.seedingQualifiedAt)
            || next.seedingQualifiedAt > current.seedingQualifiedAt
        );
    const enforcementActivated = current.legacyCompatibilityEnforcementEnabled !== true
        && next.legacyCompatibilityEnforcementEnabled === true;
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
            next.legacyLedgerSeedingEnabled !== true
            || !isValidDate(current.seedingStartedAt)
            || !hasLiveSeedingContinuity(current, serverTime)
            || continuityAdvanced
            || serverTime.getTime() - current.seedingStartedAt.getTime() < LEGACY_LIFETIME_MS
            || !isValidDate(next.seedingQualifiedAt)
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
    const issuanceStopped = current.legacyIssuanceEnabled === true
        && next.legacyIssuanceEnabled === false;
    const acceptanceDisabled = current.legacyAcceptanceEnabled === true
        && next.legacyAcceptanceEnabled === false;
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
                current.legacyIssuanceEnabled !== false
                || !isValidDate(current.legacyStopIssuanceAt)
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
    if (!isValidDate(nextValue) || nextValue > serverTime) {
        throw forbiddenAuthority('verifier-key-incident-time-invalid');
    }
    if (currentValue !== null) {
        if (!isValidDate(currentValue) || nextValue < currentValue) {
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
        || control.legacyCompatibilityEnforcementEnabled === true
        || control.legacyVerifierKeyIncidentAt !== null
    ));
}

function authorityParameterType(sql, name, value) {
    const semanticName = name.startsWith('change_') ? name.slice('change_'.length) : name;
    if (AUTHORITY_DATE_PARAMETERS.has(semanticName) || value instanceof Date) {
        return sql.DateTime2(7);
    }
    if (AUTHORITY_UUID_PARAMETERS.has(semanticName)) return sql.UniqueIdentifier;
    if (AUTHORITY_BOOLEAN_PARAMETERS.has(semanticName) || typeof value === 'boolean') {
        return sql.Bit;
    }
    if (AUTHORITY_BIGINT_PARAMETERS.has(semanticName)) return sql.BigInt;
    if (semanticName === 'rowHint') return sql.Int;
    if (/KeyId$/.test(semanticName)) return sql.VarChar(128);
    if (semanticName === 'reason' || semanticName === 'incidentCode') return sql.VarChar(128);
    if ([
        'challengeState',
        'compatibilityState',
        'eligibilityState',
        'incidentState',
        'phase',
        'registrationState',
    ].includes(semanticName)) {
        return sql.VarChar(32);
    }
    if (AUTHORITY_BINARY_PARAMETERS.has(semanticName)) {
        return sql.Binary(Buffer.isBuffer(value) ? value.length : 32);
    }
    if (['encryptedAccountMapping', 'encryptedChallenge'].includes(semanticName)) {
        return sql.VarBinary(2048);
    }
    if (Buffer.isBuffer(value)) return sql.VarBinary(sql.MAX);
    return undefined;
}

function hasValidStoredControl(control) {
    if (![
        control.controlVersion,
        control.authorityGeneration,
        control.globalSessionEpoch,
        control.seedingContinuityVersion,
    ].every((value) => Number.isSafeInteger(value) && value >= 1)) return false;
    if (!isValidDate(control.createdAt) || !isValidDate(control.updatedAt)) return false;
    if (control.updatedAt < control.createdAt) return false;
    const incidentPair = (control.incidentRecordedAt === null && control.incidentCode === null)
        || (
            isValidDate(control.incidentRecordedAt)
            && isPrivacySafeMachineToken(control.incidentCode)
        );
    if (!incidentPair) return false;
    if (control.incidentState !== 'normal' && control.incidentRecordedAt === null) return false;

    const continuityEmpty = control.seedingHeartbeatOwnerId === null
        && control.seedingHeartbeatAt === null
        && control.seedingLeaseExpiresAt === null;
    const continuityPresent = isUuid(control.seedingHeartbeatOwnerId)
        && isValidDate(control.seedingStartedAt)
        && isValidDate(control.seedingHeartbeatAt)
        && isValidDate(control.seedingLeaseExpiresAt)
        && control.seedingHeartbeatAt >= control.seedingStartedAt
        && control.seedingLeaseExpiresAt.getTime() - control.seedingHeartbeatAt.getTime()
            === LEGACY_SEEDING_LEASE_MS;
    if (!continuityEmpty && !continuityPresent) return false;
    if (
        control.legacyLedgerSeedingEnabled
        && (
            !isValidDate(control.legacyLedgerSeedingStartedAt)
            || !isValidDate(control.seedingStartedAt)
        )
    ) return false;
    if (
        control.legacyCompatibilityEnforcementEnabled
        && (
            !isValidDate(control.legacyCompatibilityEnforcedAt)
            || !isValidDate(control.seedingQualifiedAt)
            || !isValidDate(control.seedingStartedAt)
            || control.seedingQualifiedAt.getTime() - control.seedingStartedAt.getTime()
                < LEGACY_LIFETIME_MS
        )
    ) return false;
    if (!control.legacyCompatibilityEnforcementEnabled
        && control.legacyCompatibilityEnforcedAt !== null) return false;
    if (control.targetSessionIssuanceEnabled !== control.subjectTargetAdoptionEnabled) return false;
    if (control.targetSessionIssuanceEnabled && !control.targetRoutesEnabled) return false;
    if (
        !control.legacyCompatibilityEnforcementEnabled
        && (
            control.targetSessionIssuanceEnabled
            || control.subjectTargetAdoptionEnabled
            || control.dualStackStartedAt !== null
        )
    ) return false;
    const targetWindowAbsent = control.targetSessionIssuanceStartedAt === null
        && control.subjectTargetAdoptionStartedAt === null
        && control.dualStackStartedAt === null
        && control.hardSunsetAt === null;
    const targetWindowPresent = isValidDate(control.targetSessionIssuanceStartedAt)
        && isValidDate(control.subjectTargetAdoptionStartedAt)
        && isValidDate(control.dualStackStartedAt)
        && isValidDate(control.hardSunsetAt)
        && sameInstant(control.targetSessionIssuanceStartedAt, control.dualStackStartedAt)
        && sameInstant(control.subjectTargetAdoptionStartedAt, control.dualStackStartedAt)
        && control.hardSunsetAt.getTime() - control.dualStackStartedAt.getTime()
            === LEGACY_SUNSET_MAXIMUM_MS;
    if (!targetWindowAbsent && !targetWindowPresent) return false;
    if (
        control.targetSessionIssuanceEnabled
        && !targetWindowPresent
    ) return false;
    if (
        control.legacyCompatibilityEnforcementEnabled
        && control.legacyIssuanceEnabled
        && !control.legacyLedgerSeedingEnabled
    ) return false;
    if (
        control.legacyIssuanceEnabled !== (control.legacyStopIssuanceAt === null)
    ) return false;
    if (
        control.legacyAcceptanceEnabled !== (control.legacyAcceptanceDisabledAt === null)
        || (!control.legacyAcceptanceEnabled && control.legacyIssuanceEnabled)
    ) return false;
    if (
        control.legacyAcceptanceDisabledAt !== null
        && (
            !isValidDate(control.legacyAcceptanceDisabledAt)
            || !isValidDate(control.legacyStopIssuanceAt)
            || (
                control.legacyVerifierKeyIncidentAt === null
                && control.legacyAcceptanceDisabledAt.getTime()
                    - control.legacyStopIssuanceAt.getTime() < LEGACY_LIFETIME_MS
            )
        )
    ) return false;
    if (
        targetWindowPresent
        && isValidDate(control.legacyStopIssuanceAt)
        && control.legacyVerifierKeyIncidentAt === null
        && control.legacyStopIssuanceAt
            > new Date(control.hardSunsetAt.getTime() - LEGACY_LIFETIME_MS)
    ) return false;
    if (
        control.legacyVerifierKeyIncidentAt !== null
        && (
            !isValidDate(control.legacyVerifierKeyIncidentAt)
            || control.legacyIssuanceEnabled
            || control.legacyAcceptanceEnabled
        )
    ) return false;
    if (
        control.targetVerifierKeyIncidentAt !== null
        && !isValidDate(control.targetVerifierKeyIncidentAt)
    ) return false;
    return true;
}

function mapControl(row) {
    const control = { ...row };
    if (!CONTROL_INCIDENT_STATES.includes(control.incidentState)) {
        throw authorityUnavailable('control-integrity');
    }
    for (const field of [
        'targetRoutesEnabled',
        'targetSessionIssuanceEnabled',
        'legacyLedgerSeedingEnabled',
        'legacyCompatibilityEnforcementEnabled',
        'subjectTargetAdoptionEnabled',
        'legacyIssuanceEnabled',
        'legacyAcceptanceEnabled',
    ]) {
        if (![true, false, 0, 1].includes(control[field])) {
            throw authorityUnavailable('control-integrity');
        }
        control[field] = Boolean(control[field]);
    }
    control.controlId = Number(control.controlId);
    control.controlVersion = Number(control.controlVersion);
    control.seedingContinuityVersion = Number(control.seedingContinuityVersion);
    control.authorityGeneration = Number(control.authorityGeneration);
    control.globalSessionEpoch = Number(control.globalSessionEpoch);
    control.version = control.controlVersion;
    if (!hasValidStoredControl(control)) throw authorityUnavailable('control-integrity');
    return control;
}

function mapControlWithKeyFence(row, { allowRotatableKeyMismatch = false } = {}) {
    const initialized = Boolean(row.loginLookupKeyInitialized);
    const matches = Boolean(row.loginLookupKeyMatches);
    if (!initialized) throw authorityUnavailable('login-lookup-key-uninitialized');
    if (!matches) throw authorityUnavailable('login-lookup-key-mismatch');
    if (!Boolean(row.authorityKeysetInitialized)) {
        throw authorityUnavailable('authority-keyset-uninitialized');
    }
    if (!Boolean(row.legacySigningKeyInitialized)) {
        throw authorityUnavailable('legacy-signing-key-uninitialized');
    }
    if (!Boolean(row.accountMappingKeyMatches)) {
        throw authorityUnavailable('authority-keyset-mismatch');
    }
    if (
        !Boolean(row.keysetLoginLookupKeyMatches)
        || !Boolean(row.keysetAccountMappingKeyMatches)
    ) throw authorityUnavailable('authority-keyset-mismatch');
    const mismatches = Object.freeze({
        aggregate: !Boolean(row.authorityKeysetAggregateMatches),
        targetVerifier: !Boolean(row.targetVerifierKeyMatches),
        legacyCompatibility: !Boolean(row.legacyCompatibilityKeyMatches),
        credentialFingerprint: !Boolean(row.credentialFingerprintKeyMatches),
        faceChallengeEncryption: !Boolean(row.faceChallengeKeyMatches),
        legacySigning: !Boolean(row.legacySigningKeyMatches),
    });
    const hasMismatch = Object.values(mismatches).some(Boolean);
    if (hasMismatch && !allowRotatableKeyMismatch) {
        throw authorityUnavailable('authority-keyset-mismatch');
    }
    const control = mapControl(sanitizeControlLookupFlags(row));
    Object.defineProperty(control, 'keyMismatches', {
        configurable: false,
        enumerable: false,
        value: mismatches,
        writable: false,
    });
    return control;
}

function sanitizeControlLookupFlags(control) {
    const sanitized = { ...control };
    for (const name of [
        'loginLookupKeyInitialized',
        'loginLookupKeyMatches',
        'accountMappingKeyMatches',
        'keysetLoginLookupKeyMatches',
        'keysetAccountMappingKeyMatches',
        'authorityKeysetInitialized',
        'authorityKeysetAggregateMatches',
        'targetVerifierKeyMatches',
        'legacyCompatibilityKeyMatches',
        'credentialFingerprintKeyMatches',
        'faceChallengeKeyMatches',
        'legacySigningKeyInitialized',
        'legacySigningKeyMatches',
    ]) delete sanitized[name];
    return sanitized;
}

function mapSubject(row) {
    const subject = {
        ...row,
        credentialVersion: Number(row.credentialVersion),
        sessionEpoch: Number(row.sessionEpoch),
    };
    if (
        !['eligible', 'ineligible', 'unknown'].includes(subject.eligibilityState)
        || !Number.isSafeInteger(subject.credentialVersion)
        || subject.credentialVersion < 1
        || !Number.isSafeInteger(subject.sessionEpoch)
        || subject.sessionEpoch < 1
        || !Buffer.isBuffer(subject.loginLookupToken)
        || subject.loginLookupToken.length !== 32
        || !Buffer.isBuffer(subject.credentialFingerprint)
        || subject.credentialFingerprint.length !== 32
        || !isValidKeyId(subject.loginLookupKeyId)
        || !isValidKeyId(subject.credentialFingerprintKeyId)
        || !isValidKeyId(subject.accountMappingKeyId)
        || !Buffer.isBuffer(subject.encryptedAccountMapping)
        || subject.encryptedAccountMapping.length === 0
        || !isValidDate(subject.eligibilityObservedAt)
        || !isValidDate(subject.eligibilityRevalidateAt)
        || subject.eligibilityRevalidateAt < subject.eligibilityObservedAt
        || subject.eligibilityRevalidateAt.getTime()
            - subject.eligibilityObservedAt.getTime() > 5 * 60 * 1000
    ) throw authorityUnavailable('subject-eligibility-integrity');
    return subject;
}

function mapSession(row) {
    const session = {
        ...row,
        faceRequired: Boolean(row.faceRequired),
        registrationRequired: Boolean(row.registrationRequired),
        subjectEpochSnapshot: Number(row.subjectEpochSnapshot),
        credentialVersionSnapshot: Number(row.credentialVersionSnapshot),
        globalEpochSnapshot: Number(row.globalEpochSnapshot),
        authorityGenerationSnapshot: Number(row.authorityGenerationSnapshot),
        version: ACTIVE_PHASES.includes(row.phase) ? 1 : 2,
    };
    if (!hasValidStoredSessionRecord(row, session)) {
        throw authorityUnavailable('session-store-integrity');
    }
    return session;
}

function mapFlow(row) {
    if (!hasValidStoredFlowRecord(row)) {
        throw authorityUnavailable('session-store-integrity');
    }
    return {
        ...row,
        registrationState: applicationRegistrationState(row.registrationState),
    };
}

function mapLegacy(row) {
    if (!hasValidStoredLegacyRecord(row)) {
        throw authorityUnavailable('legacy-binding-integrity');
    }
    return {
        ...row,
        incidentState: row.compatibilityState === 'active' ? 'normal' : row.compatibilityState,
    };
}

function hasValidStoredSessionRecord(raw, session) {
    if (
        !ALL_SESSION_PHASES.includes(raw.phase)
        || ![true, false, 0, 1].includes(raw.faceRequired)
        || ![true, false, 0, 1].includes(raw.registrationRequired)
        || !Buffer.isBuffer(raw.verifier)
        || raw.verifier.length !== 32
        || !isValidKeyId(raw.verifierKeyId)
        || !isUuid(raw.sessionId)
        || !isUuid(raw.subjectId)
        || !isValidDate(raw.originalIssuedAt)
        || !isValidDate(raw.phaseStartedAt)
        || !isValidDate(raw.expiresAt)
        || !isValidDate(raw.createdAt)
        || raw.originalIssuedAt > raw.phaseStartedAt
        || raw.phaseStartedAt >= raw.expiresAt
        || ![session.subjectEpochSnapshot, session.credentialVersionSnapshot,
            session.globalEpochSnapshot, session.authorityGenerationSnapshot]
            .every((value) => Number.isSafeInteger(value) && value >= 1)
    ) return false;
    const lifetime = raw.expiresAt.getTime() - raw.originalIssuedAt.getTime();
    if (
        ([SESSION_PHASES.credentialVerified, SESSION_PHASES.registrationPending,
            SESSION_PHASES.facePending].includes(raw.phase)
            && lifetime !== PROVISIONAL_LIFETIME_MS)
        || (raw.phase === SESSION_PHASES.authenticated
            && lifetime !== AUTHENTICATED_LIFETIME_MS)
        || ([SESSION_PHASES.expired, SESSION_PHASES.revoked, SESSION_PHASES.rotatedOut]
            .includes(raw.phase)
            && ![PROVISIONAL_LIFETIME_MS, AUTHENTICATED_LIFETIME_MS].includes(lifetime))
    ) return false;
    const active = ACTIVE_PHASES.includes(raw.phase);
    if (active || raw.phase === SESSION_PHASES.expired) {
        return raw.revokedAt === null
            && raw.revocationReason === null
            && raw.replacementSessionId === null;
    }
    if (raw.phase === SESSION_PHASES.revoked) {
        return isValidDate(raw.revokedAt)
            && isPrivacySafeMachineToken(raw.revocationReason)
            && raw.replacementSessionId === null;
    }
    return raw.phase === SESSION_PHASES.rotatedOut
        && raw.revokedAt === null
        && raw.revocationReason === null
        && isUuid(raw.replacementSessionId)
        && normalizedUuid(raw.replacementSessionId) !== normalizedUuid(raw.sessionId);
}

function hasValidStoredFlowRecord(row) {
    if (
        !isUuid(row.flowId)
        || !isUuid(row.subjectId)
        || !isUuid(row.currentSessionId)
        || !(row.challengeSessionId === null || isUuid(row.challengeSessionId))
        || !FLOW_REGISTRATION_STATES.includes(row.registrationState)
        || !FLOW_CHALLENGE_STATES.includes(row.challengeState)
        || !isValidDate(row.createdAt)
        || !isValidDate(row.updatedAt)
        || row.updatedAt < row.createdAt
    ) return false;
    const noReference = row.encryptedChallenge === null
        && row.challengeKeyId === null
        && row.challengeCreatedAt === null;
    const reference = Buffer.isBuffer(row.encryptedChallenge)
        && row.encryptedChallenge.length > 0
        && isValidKeyId(row.challengeKeyId)
        && isValidDate(row.challengeCreatedAt);
    if (['none', 'creating'].includes(row.challengeState)) {
        return row.challengeSessionId === null && noReference && row.consumedAt === null;
    }
    if (row.challengeState === 'active') {
        return normalizedUuid(row.challengeSessionId) === normalizedUuid(row.currentSessionId)
            && reference
            && row.consumedAt === null;
    }
    if (row.challengeState === 'consumed') {
        return reference
            && normalizedUuid(row.challengeSessionId) !== normalizedUuid(row.currentSessionId)
            && isValidDate(row.consumedAt)
            && row.consumedAt >= row.challengeCreatedAt;
    }
    if (row.challengeState === 'failed') {
        return reference
            && normalizedUuid(row.challengeSessionId) === normalizedUuid(row.currentSessionId)
            && isValidDate(row.consumedAt)
            && row.consumedAt >= row.challengeCreatedAt;
    }
    return row.challengeState === 'reconciliation-required'
        && row.consumedAt === null
        && (
            (row.challengeSessionId === null && noReference)
            || (
                normalizedUuid(row.challengeSessionId) === normalizedUuid(row.currentSessionId)
                && reference
            )
        );
}

function hasValidStoredLegacyRecord(row) {
    if (
        !LEGACY_COMPATIBILITY_STATES.includes(row.compatibilityState)
        || !isUuid(row.legacyCompatibilityId)
        || !isUuid(row.subjectId)
        || !Buffer.isBuffer(row.verifier)
        || row.verifier.length !== 32
        || !isValidKeyId(row.verifierKeyId)
        || !isValidDate(row.issuedAt)
        || !isValidDate(row.expiresAt)
        || row.expiresAt.getTime() - row.issuedAt.getTime() !== LEGACY_LIFETIME_MS
    ) return false;
    if (row.compatibilityState === 'active') {
        return row.revokedAt === null
            && row.revocationReason === null
            && row.incidentAt === null
            && row.incidentCode === null;
    }
    if (row.compatibilityState === 'revoked') {
        return isValidDate(row.revokedAt)
            && isPrivacySafeMachineToken(row.revocationReason)
            && row.incidentAt === null
            && row.incidentCode === null;
    }
    return row.revokedAt === null
        && row.revocationReason === null
        && isValidDate(row.incidentAt)
        && isPrivacySafeMachineToken(row.incidentCode);
}

function isPrivacySafeMachineToken(value) {
    return typeof value === 'string' && /^[a-z0-9][a-z0-9-]{0,127}$/.test(value);
}

function isUuid(value) {
    return typeof value === 'string'
        && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
}

function evaluateAuthority({ session, subject, control, flow = null }, serverTime) {
    if (control.incidentState !== 'normal') {
        return { active: false, unavailable: true, reason: 'authority-incident' };
    }
    if (!ACTIVE_PHASES.includes(session.phase)) return { active: false, reason: session.phase };
    if (session.revokedAt !== null) return { active: false, reason: 'revoked' };
    if (!hasValidActiveSessionLifetime(session)) {
        return { active: false, unavailable: true, reason: 'session-store-integrity' };
    }
    if (!hasValidActiveSessionEvidence(session, flow, serverTime)) {
        return { active: false, unavailable: true, reason: 'session-store-integrity' };
    }
    const subjectState = evaluateSubjectEligibility(subject, serverTime);
    if (subjectState.unavailable || subject.eligibilityState !== 'eligible') {
        return { active: false, ...subjectState };
    }
    const sessionExpired = session.expiresAt <= serverTime;
    const entitlementExpired = !subjectState.eligible;
    if (
        entitlementExpired
        && (!sessionExpired || subject.entitlementExpiresAt <= session.expiresAt)
    ) return { active: false, ...subjectState };
    if (sessionExpired) return { active: false, expired: true, reason: 'expired' };
    if (entitlementExpired) return { active: false, ...subjectState };
    if (
        session.subjectEpochSnapshot !== subject.sessionEpoch
        || session.credentialVersionSnapshot !== subject.credentialVersion
        || session.globalEpochSnapshot !== control.globalSessionEpoch
        || session.authorityGenerationSnapshot !== control.authorityGeneration
    ) return { active: false, reason: 'epoch-mismatch' };
    if (subject.eligibilityState !== 'eligible') return { active: false, reason: 'ineligible' };
    return { active: true };
}

function hasValidActiveSessionLifetime(session) {
    if (
        !isValidDate(session.originalIssuedAt)
        || !isValidDate(session.phaseStartedAt)
        || !isValidDate(session.expiresAt)
        || session.originalIssuedAt > session.phaseStartedAt
        || session.phaseStartedAt >= session.expiresAt
    ) return false;
    const expectedLifetime = session.phase === SESSION_PHASES.authenticated
        ? AUTHENTICATED_LIFETIME_MS
        : PROVISIONAL_LIFETIME_MS;
    return session.expiresAt.getTime() - session.originalIssuedAt.getTime() === expectedLifetime;
}

function hasValidActiveSessionEvidence(session, flow, serverTime) {
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

    if (flow !== null) {
        if (
            normalizedUuid(flow.subjectId) !== normalizedUuid(session.subjectId)
            || normalizedUuid(flow.currentSessionId) !== normalizedUuid(session.sessionId)
        ) return false;
    }

    if (session.phase === SESSION_PHASES.facePending) {
        return flowHasResolvedReference(flow, serverTime, 'active', false)
            && normalizedUuid(flow.challengeSessionId) === normalizedUuid(session.sessionId)
            && hasValidChallengeSessionLineage(flow, session, false);
    }
    if (session.phase === SESSION_PHASES.authenticated) {
        if (!session.faceRequired) return flow === null;
        return flowHasResolvedReference(flow, serverTime, 'consumed', true)
            && hasValidChallengeSessionLineage(flow, session, true);
    }
    if (flow === null) return true;
    return ['creating', 'reconciliation-required'].includes(flow.challengeState)
        && flow.encryptedChallenge === null
        && flow.challengeKeyId === null
        && flow.challengeCreatedAt === null
        && flow.consumedAt === null;
}

function hasValidChallengeSessionLineage(flow, currentSession, consumed) {
    const challengeSession = flow?.challengeSession;
    if (
        !challengeSession
        || normalizedUuid(challengeSession.sessionId) !== normalizedUuid(flow.challengeSessionId)
        || normalizedUuid(challengeSession.subjectId) !== normalizedUuid(flow.subjectId)
        || normalizedUuid(challengeSession.subjectId) !== normalizedUuid(currentSession.subjectId)
    ) return false;
    if (!consumed) {
        return challengeSession.phase === SESSION_PHASES.facePending
            && normalizedUuid(challengeSession.sessionId)
                === normalizedUuid(currentSession.sessionId)
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
        && normalizedUuid(challengeSession.replacementSessionId)
            === normalizedUuid(currentSession.sessionId);
}

function flowHasResolvedReference(flow, serverTime, expectedState, resolved) {
    if (
        !flow
        || flow.challengeState !== expectedState
        || flow.registrationState !== 'registered'
        || !Buffer.isBuffer(flow.encryptedChallenge)
        || flow.encryptedChallenge.length === 0
        || !isValidKeyId(flow.challengeKeyId)
        || !isValidDate(flow.challengeCreatedAt)
        || flow.challengeCreatedAt > serverTime
    ) return false;
    if (!resolved) return flow.consumedAt === null;
    return isValidDate(flow.consumedAt)
        && flow.consumedAt >= flow.challengeCreatedAt
        && flow.consumedAt <= serverTime;
}

function requireExpectedSession(
    session,
    expectedVersion,
    allowedPhases,
    serverTime,
    subject,
    control,
    flow,
) {
    if (!session || expectedVersion !== 1 || !ACTIVE_PHASES.includes(session.phase)) {
        throw authorityConflict('session-compare-and-replace');
    }
    const state = evaluateAuthority({ session, subject, control, flow }, serverTime);
    if (state.unavailable) throw authorityUnavailable(state.reason);
    if (state.ineligible) throw forbiddenAuthority('ineligible');
    if (!state.active) throw invalidAuthority(state.reason);
    if (allowedPhases) requireAllowedPhase(session, allowedPhases);
}

function requireAllowedPhase(session, allowedPhases) {
    if (!Array.isArray(allowedPhases) || !allowedPhases.includes(session.phase)) {
        throw forbiddenAuthority('wrong-phase');
    }
}

function requireIssuanceControl(control, serverTime) {
    if (control.incidentState !== 'normal') throw authorityUnavailable('authority-incident');
    if (!control.targetSessionIssuanceEnabled || !control.subjectTargetAdoptionEnabled) {
        throw authorityUnavailable('target-session-issuance-disabled');
    }
    if (
        !control.legacyCompatibilityEnforcementEnabled
        || !isValidDate(control.dualStackStartedAt)
        || !isValidDate(control.hardSunsetAt)
        || control.hardSunsetAt.getTime() - control.dualStackStartedAt.getTime()
            !== LEGACY_SUNSET_MAXIMUM_MS
    ) throw authorityUnavailable('target-session-window-unqualified');
    if (serverTime < control.dualStackStartedAt) {
        throw authorityUnavailable('target-session-window-inactive');
    }
}

function requireLegacyIssuanceControl(control, serverTime) {
    if (control.incidentState !== 'normal') throw authorityUnavailable('authority-incident');
    if (control.hardSunsetAt !== null && !isValidDate(control.hardSunsetAt)) {
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

function evaluateSubjectEligibility(subject, serverTime) {
    if (
        subject.eligibilityObservedAt > serverTime
        || subject.eligibilityRevalidateAt < subject.eligibilityObservedAt
        || subject.eligibilityRevalidateAt.getTime()
            - subject.eligibilityObservedAt.getTime() > 5 * 60 * 1000
    ) return { eligible: false, unavailable: true, reason: 'subject-eligibility-integrity' };
    if (subject.eligibilityState === 'ineligible') {
        return { eligible: false, ineligible: true, reason: 'ineligible' };
    }
    if (subject.eligibilityState !== 'eligible') {
        return { eligible: false, unavailable: true, reason: 'subject-eligibility-integrity' };
    }
    if (!isValidDate(subject.entitlementExpiresAt)) {
        return { eligible: false, unavailable: true, reason: 'subject-eligibility-integrity' };
    }
    if (subject.entitlementExpiresAt <= serverTime) {
        return { eligible: false, ineligible: true, reason: 'ineligible' };
    }
    return { eligible: true };
}

function requireFreshEligibility(subject, serverTime) {
    if (
        !isValidDate(subject.eligibilityRevalidateAt)
        || subject.eligibilityRevalidateAt <= serverTime
    ) throw authorityUnavailable('eligibility-revalidation-required');
}

function validateSessionCandidate(input) {
    if (!input || typeof input.sessionId !== 'string') throw new TypeError('sessionId is required');
    if (!Buffer.isBuffer(input.verifier)) throw new TypeError('Session verifier must be a Buffer');
    if (!isValidKeyId(input.verifierKeyId)) {
        throw new TypeError('Session verifier key ID is required');
    }
    if (!ACTIVE_PHASES.includes(input.phase) && input.phase !== undefined) {
        throw new TypeError('Session phase is invalid');
    }
}

function validateSubjectObservationExpectation(input) {
    if (!isValidDate(input.observationStartedAt)) {
        throw new TypeError('observationStartedAt must be a valid Date');
    }
    validateSubjectCredentialExpectation(input);
}

function validateSubjectCredentialExpectation(input) {
    if (!Number.isSafeInteger(input.expectedCredentialVersion) || input.expectedCredentialVersion < 1) {
        throw new TypeError('Expected credential version is required');
    }
    if (
        !isValidKeyId(input.expectedCredentialFingerprintKeyId)
    ) throw new TypeError('Expected credential fingerprint key ID is required');
    if (!Buffer.isBuffer(input.expectedCredentialFingerprint)) {
        throw new TypeError('Expected credential fingerprint must be a Buffer');
    }
}

function validateExpectedControlVersion(value) {
    if (!Number.isSafeInteger(value) || value < 1) {
        throw new TypeError('Expected control version must be a positive safe integer');
    }
}

function validateSubjectLoginRemap(input) {
    if (!input || typeof input.subjectId !== 'string' || input.subjectId.length === 0) {
        throw new TypeError('Subject ID is required');
    }
    for (const name of ['expectedLoginLookupKeyId', 'loginLookupKeyId', 'accountMappingKeyId']) {
        if (!isValidKeyId(input[name])) {
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

function validateLegacyLifetime(issuedAt, expiresAt) {
    if (
        !isValidDate(issuedAt)
        || !isValidDate(expiresAt)
        || new Date(expiresAt).getTime() - new Date(issuedAt).getTime() !== LEGACY_LIFETIME_MS
    ) throw new TypeError('Legacy issue metadata must describe one exact legacy lifetime');
}

function validateRevocationReason(reason) {
    if (typeof reason !== 'string' || !/^[a-z0-9][a-z0-9-]{0,127}$/.test(reason)) {
        throw new TypeError('Revocation reason must be a privacy-safe machine value');
    }
}

function lifetimeForPhase(phase) {
    return phase === SESSION_PHASES.authenticated ? AUTHENTICATED_LIFETIME_MS : PROVISIONAL_LIFETIME_MS;
}

function databaseRegistrationState(value) {
    if (!['enrollment-accepted', 'registered'].includes(value)) {
        throw new TypeError('Face flow registration state is invalid');
    }
    return value;
}

function applicationRegistrationState(value) {
    return value === 'required' ? 'registration-required' : value;
}

function sameLegacyBinding(binding, input, requireSubject) {
    return (!requireSubject || normalizedUuid(binding.subjectId) === normalizedUuid(input.subjectId))
        && binding.verifierKeyId === input.verifierKeyId
        && sameInstant(binding.issuedAt, input.issuedAt)
        && sameInstant(binding.expiresAt, input.expiresAt);
}

function normalizedUuid(value) {
    return typeof value === 'string' ? value.toLowerCase() : value;
}

function sameInstant(left, right) {
    return isValidDate(left) && isValidDate(right) && new Date(left).getTime() === new Date(right).getTime();
}

function exactlyOne(rows, reason) {
    if (!Array.isArray(rows) || rows.length !== 1) throw authorityUnavailable(reason);
    return rows[0];
}

function affectedRows(result) {
    return Array.isArray(result.rowsAffected)
        ? result.rowsAffected.reduce((sum, count) => sum + Number(count || 0), 0)
        : 0;
}

function pick(source, names) {
    return Object.fromEntries(names.map((name) => [name, source[name]]));
}

function requiredDate(value) {
    if (!isValidDate(value)) throw authorityUnavailable('session-store-integrity');
    return copyDate(value);
}

function isValidDate(value) {
    return value instanceof Date && Number.isFinite(value.getTime());
}

function copyDate(value) {
    return new Date(value.getTime());
}

function postCommitError(error) {
    return { [POST_COMMIT_ERROR]: true, error };
}

function isPostCommitError(value) {
    return Boolean(value && value[POST_COMMIT_ERROR]);
}

module.exports = { createAzureSqlSessionStore };
