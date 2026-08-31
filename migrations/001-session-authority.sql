SET NOCOUNT ON;
SET XACT_ABORT ON;

-- Forward-only session-authority foundation. This migration deliberately has
-- no down path: revocation and migration-cutoff evidence must not be discarded.
BEGIN TRY
    BEGIN TRANSACTION;

    IF OBJECT_ID(N'dbo.learning_subject', N'U') IS NOT NULL
        OR OBJECT_ID(N'dbo.learning_session', N'U') IS NOT NULL
        OR OBJECT_ID(N'dbo.learning_session_flow', N'U') IS NOT NULL
        OR OBJECT_ID(N'dbo.legacy_session_compatibility', N'U') IS NOT NULL
        OR OBJECT_ID(N'dbo.session_authority_control', N'U') IS NOT NULL
    BEGIN
        THROW 51000, 'Session-authority schema already exists; refusing to replace authority evidence.', 1;
    END;

    CREATE TABLE dbo.learning_subject (
        subject_id uniqueidentifier NOT NULL
            CONSTRAINT DF_learning_subject_subject_id DEFAULT NEWID(),
        login_lookup_token binary(32) NOT NULL,
        login_lookup_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NOT NULL,
        encrypted_legacy_account_mapping varbinary(2048) NOT NULL,
        account_mapping_encryption_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NOT NULL,
        workbook_row_hint int NULL,
        credential_version bigint NOT NULL
            CONSTRAINT DF_learning_subject_credential_version DEFAULT (1),
        credential_fingerprint binary(32) NOT NULL,
        credential_fingerprint_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NOT NULL,
        subject_session_epoch bigint NOT NULL
            CONSTRAINT DF_learning_subject_session_epoch DEFAULT (1),
        legacy_authority_disabled_at datetime2(7) NULL,
        eligibility_state varchar(16) COLLATE Latin1_General_100_BIN2 NOT NULL,
        entitlement_expires_at datetime2(7) NULL,
        eligibility_observed_at datetime2(7) NOT NULL,
        eligibility_revalidate_at datetime2(7) NOT NULL,
        created_at datetime2(7) NOT NULL
            CONSTRAINT DF_learning_subject_created_at DEFAULT SYSUTCDATETIME(),
        CONSTRAINT PK_learning_subject PRIMARY KEY (subject_id),
        CONSTRAINT UQ_learning_subject_login_lookup
            UNIQUE (login_lookup_key_id, login_lookup_token),
        CONSTRAINT CK_learning_subject_key_ids CHECK (
            login_lookup_key_id COLLATE Latin1_General_100_BIN2 LIKE '[A-Za-z0-9]%'
            AND login_lookup_key_id COLLATE Latin1_General_100_BIN2
                NOT LIKE '%[^A-Za-z0-9._:-]%'
            AND DATALENGTH(login_lookup_key_id) = LEN(login_lookup_key_id)
            AND account_mapping_encryption_key_id COLLATE Latin1_General_100_BIN2
                LIKE '[A-Za-z0-9]%'
            AND account_mapping_encryption_key_id COLLATE Latin1_General_100_BIN2
                NOT LIKE '%[^A-Za-z0-9._:-]%'
            AND DATALENGTH(account_mapping_encryption_key_id)
                = LEN(account_mapping_encryption_key_id)
            AND credential_fingerprint_key_id COLLATE Latin1_General_100_BIN2
                LIKE '[A-Za-z0-9]%'
            AND credential_fingerprint_key_id COLLATE Latin1_General_100_BIN2
                NOT LIKE '%[^A-Za-z0-9._:-]%'
            AND DATALENGTH(credential_fingerprint_key_id)
                = LEN(credential_fingerprint_key_id)
        ),
        CONSTRAINT CK_learning_subject_row_hint CHECK (
            workbook_row_hint IS NULL OR workbook_row_hint >= 0
        ),
        CONSTRAINT CK_learning_subject_versions CHECK (
            credential_version BETWEEN 1 AND 9007199254740991
            AND subject_session_epoch BETWEEN 1 AND 9007199254740991
        ),
        CONSTRAINT CK_learning_subject_eligibility_state CHECK (
            eligibility_state IN ('eligible', 'ineligible', 'unknown')
        ),
        CONSTRAINT CK_learning_subject_eligibility_window CHECK (
            eligibility_revalidate_at >= eligibility_observed_at
            AND eligibility_revalidate_at <= DATEADD(minute, 5, eligibility_observed_at)
        ),
        CONSTRAINT CK_learning_subject_cutoff_time CHECK (
            legacy_authority_disabled_at IS NULL
            OR legacy_authority_disabled_at >= created_at
        )
    );

    CREATE TABLE dbo.learning_session (
        session_id uniqueidentifier NOT NULL
            CONSTRAINT DF_learning_session_session_id DEFAULT NEWID(),
        identifier_verifier binary(32) NOT NULL,
        verifier_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NOT NULL,
        subject_id uniqueidentifier NOT NULL,
        phase varchar(24) COLLATE Latin1_General_100_BIN2 NOT NULL,
        original_issued_at datetime2(7) NOT NULL
            CONSTRAINT DF_learning_session_original_issued_at DEFAULT SYSUTCDATETIME(),
        phase_started_at datetime2(7) NOT NULL
            CONSTRAINT DF_learning_session_phase_started_at DEFAULT SYSUTCDATETIME(),
        absolute_expires_at datetime2(7) NOT NULL,
        face_auth_required bit NOT NULL,
        registration_required bit NOT NULL,
        subject_epoch_snapshot bigint NOT NULL,
        credential_version_snapshot bigint NOT NULL,
        global_epoch_snapshot bigint NOT NULL,
        authority_generation_snapshot bigint NOT NULL,
        revoked_at datetime2(7) NULL,
        revocation_reason varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        replacement_session_id uniqueidentifier NULL,
        created_at datetime2(7) NOT NULL
            CONSTRAINT DF_learning_session_created_at DEFAULT SYSUTCDATETIME(),
        CONSTRAINT PK_learning_session PRIMARY KEY (session_id),
        CONSTRAINT UQ_learning_session_identifier_verifier UNIQUE (identifier_verifier),
        CONSTRAINT UQ_learning_session_id_subject UNIQUE (session_id, subject_id),
        CONSTRAINT FK_learning_session_subject FOREIGN KEY (subject_id)
            REFERENCES dbo.learning_subject (subject_id),
        CONSTRAINT FK_learning_session_replacement FOREIGN KEY (replacement_session_id)
            REFERENCES dbo.learning_session (session_id),
        CONSTRAINT CK_learning_session_verifier_key_id CHECK (
            verifier_key_id COLLATE Latin1_General_100_BIN2 LIKE '[A-Za-z0-9]%'
            AND verifier_key_id COLLATE Latin1_General_100_BIN2
                NOT LIKE '%[^A-Za-z0-9._:-]%'
            AND DATALENGTH(verifier_key_id) = LEN(verifier_key_id)
        ),
        CONSTRAINT CK_learning_session_phase CHECK (
            phase IN (
                'credential-verified',
                'registration-pending',
                'face-pending',
                'authenticated',
                'expired',
                'revoked',
                'rotated-out'
            )
        ),
        CONSTRAINT CK_learning_session_time_order CHECK (
            original_issued_at <= phase_started_at
            AND phase_started_at < absolute_expires_at
            AND created_at = phase_started_at
            AND (
                (
                    phase IN ('credential-verified', 'registration-pending', 'face-pending')
                    AND absolute_expires_at = DATEADD(minute, 20, original_issued_at)
                )
                OR
                (
                    phase = 'authenticated'
                    AND absolute_expires_at = DATEADD(hour, 4, original_issued_at)
                )
                OR
                (
                    phase IN ('expired', 'revoked', 'rotated-out')
                    AND absolute_expires_at IN (
                        DATEADD(minute, 20, original_issued_at),
                        DATEADD(hour, 4, original_issued_at)
                    )
                )
            )
        ),
        CONSTRAINT CK_learning_session_epoch_snapshots CHECK (
            subject_epoch_snapshot BETWEEN 1 AND 9007199254740991
            AND credential_version_snapshot BETWEEN 1 AND 9007199254740991
            AND global_epoch_snapshot BETWEEN 1 AND 9007199254740991
            AND authority_generation_snapshot BETWEEN 1 AND 9007199254740991
        ),
        CONSTRAINT CK_learning_session_registration_policy CHECK (
            registration_required = 0 OR face_auth_required = 1
        ),
        CONSTRAINT CK_learning_session_active_phase_policy CHECK (
            phase NOT IN ('credential-verified', 'registration-pending', 'face-pending')
            OR (
                face_auth_required = 1
                AND (phase <> 'registration-pending' OR registration_required = 1)
            )
        ),
        CONSTRAINT CK_learning_session_revocation_evidence CHECK (
            (
                phase = 'revoked'
                AND revoked_at IS NOT NULL
                AND revocation_reason IS NOT NULL
                AND revocation_reason COLLATE Latin1_General_100_BIN2 LIKE '[a-z0-9]%'
                AND revocation_reason COLLATE Latin1_General_100_BIN2 NOT LIKE '%[^a-z0-9-]%'
            )
            OR
            (
                phase <> 'revoked'
                AND revoked_at IS NULL
                AND revocation_reason IS NULL
            )
        ),
        CONSTRAINT CK_learning_session_replacement_evidence CHECK (
            (
                phase = 'rotated-out'
                AND replacement_session_id IS NOT NULL
                AND replacement_session_id <> session_id
            )
            OR
            (
                phase <> 'rotated-out'
                AND replacement_session_id IS NULL
            )
        )
    );

    CREATE UNIQUE INDEX UX_learning_session_replacement
        ON dbo.learning_session (replacement_session_id)
        WHERE replacement_session_id IS NOT NULL;

    CREATE INDEX IX_learning_session_subject_authority
        ON dbo.learning_session (subject_id, phase, absolute_expires_at);

    CREATE TABLE dbo.learning_session_flow (
        flow_id uniqueidentifier NOT NULL
            CONSTRAINT DF_learning_session_flow_flow_id DEFAULT NEWID(),
        subject_id uniqueidentifier NOT NULL,
        current_session_id uniqueidentifier NOT NULL,
        challenge_session_id uniqueidentifier NULL,
        registration_state varchar(32) COLLATE Latin1_General_100_BIN2 NOT NULL,
        challenge_state varchar(32) COLLATE Latin1_General_100_BIN2 NOT NULL
            CONSTRAINT DF_learning_session_flow_challenge_state DEFAULT ('none'),
        encrypted_provider_challenge_reference varbinary(2048) NULL,
        provider_reference_encryption_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        challenge_created_at datetime2(7) NULL,
        challenge_resolved_at datetime2(7) NULL,
        created_at datetime2(7) NOT NULL
            CONSTRAINT DF_learning_session_flow_created_at DEFAULT SYSUTCDATETIME(),
        updated_at datetime2(7) NOT NULL
            CONSTRAINT DF_learning_session_flow_updated_at DEFAULT SYSUTCDATETIME(),
        CONSTRAINT PK_learning_session_flow PRIMARY KEY (flow_id),
        CONSTRAINT UQ_learning_session_flow_current_session UNIQUE (current_session_id),
        CONSTRAINT FK_learning_session_flow_session_subject
            FOREIGN KEY (current_session_id, subject_id)
            REFERENCES dbo.learning_session (session_id, subject_id),
        CONSTRAINT FK_learning_session_flow_challenge_session_subject
            FOREIGN KEY (challenge_session_id, subject_id)
            REFERENCES dbo.learning_session (session_id, subject_id),
        CONSTRAINT CK_learning_session_flow_registration_state CHECK (
            registration_state IN (
                'not-required',
                'required',
                'enrollment-accepted',
                'registered',
                'reconciliation-required'
            )
        ),
        CONSTRAINT CK_learning_session_flow_challenge_state CHECK (
            challenge_state IN (
                'none',
                'creating',
                'active',
                'consumed',
                'failed',
                'reconciliation-required'
            )
        ),
        CONSTRAINT CK_learning_session_flow_reference_pair CHECK (
            (
                encrypted_provider_challenge_reference IS NULL
                AND provider_reference_encryption_key_id IS NULL
                AND challenge_created_at IS NULL
            )
            OR
            (
                encrypted_provider_challenge_reference IS NOT NULL
                AND provider_reference_encryption_key_id IS NOT NULL
                AND provider_reference_encryption_key_id COLLATE Latin1_General_100_BIN2
                    LIKE '[A-Za-z0-9]%'
                AND provider_reference_encryption_key_id COLLATE Latin1_General_100_BIN2
                    NOT LIKE '%[^A-Za-z0-9._:-]%'
                AND DATALENGTH(provider_reference_encryption_key_id)
                    = LEN(provider_reference_encryption_key_id)
                AND challenge_created_at IS NOT NULL
            )
        ),
        CONSTRAINT CK_learning_session_flow_challenge_evidence CHECK (
            (
                challenge_state IN ('none', 'creating')
                AND challenge_session_id IS NULL
                AND encrypted_provider_challenge_reference IS NULL
                AND challenge_resolved_at IS NULL
            )
            OR
            (
                challenge_state = 'active'
                AND challenge_session_id = current_session_id
                AND encrypted_provider_challenge_reference IS NOT NULL
                AND challenge_resolved_at IS NULL
            )
            OR
            (
                challenge_state = 'consumed'
                AND challenge_session_id IS NOT NULL
                AND challenge_session_id <> current_session_id
                AND encrypted_provider_challenge_reference IS NOT NULL
                AND challenge_resolved_at IS NOT NULL
                AND challenge_resolved_at >= challenge_created_at
            )
            OR
            (
                challenge_state = 'failed'
                AND challenge_session_id = current_session_id
                AND encrypted_provider_challenge_reference IS NOT NULL
                AND challenge_resolved_at IS NOT NULL
                AND challenge_resolved_at >= challenge_created_at
            )
            OR
            (
                challenge_state = 'reconciliation-required'
                AND challenge_resolved_at IS NULL
            )
        ),
        CONSTRAINT CK_learning_session_flow_update_time CHECK (updated_at >= created_at)
    );

    CREATE UNIQUE INDEX UX_learning_session_flow_challenge_session
        ON dbo.learning_session_flow (challenge_session_id)
        WHERE challenge_session_id IS NOT NULL;

    CREATE TABLE dbo.legacy_session_compatibility (
        compatibility_id uniqueidentifier NOT NULL
            CONSTRAINT DF_legacy_session_compatibility_id DEFAULT NEWID(),
        legacy_handle_verifier binary(32) NOT NULL,
        verifier_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NOT NULL,
        subject_id uniqueidentifier NOT NULL,
        original_issued_at datetime2(7) NOT NULL,
        original_expires_at datetime2(7) NOT NULL,
        compatibility_state varchar(16) COLLATE Latin1_General_100_BIN2 NOT NULL
            CONSTRAINT DF_legacy_session_compatibility_state DEFAULT ('active'),
        revoked_at datetime2(7) NULL,
        revocation_reason varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        incident_at datetime2(7) NULL,
        incident_code varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        created_at datetime2(7) NOT NULL
            CONSTRAINT DF_legacy_session_compatibility_created_at DEFAULT SYSUTCDATETIME(),
        CONSTRAINT PK_legacy_session_compatibility PRIMARY KEY (compatibility_id),
        CONSTRAINT UQ_legacy_session_compatibility_verifier UNIQUE (legacy_handle_verifier),
        CONSTRAINT FK_legacy_session_compatibility_subject FOREIGN KEY (subject_id)
            REFERENCES dbo.learning_subject (subject_id),
        CONSTRAINT CK_legacy_session_compatibility_key_id CHECK (
            verifier_key_id COLLATE Latin1_General_100_BIN2 LIKE '[A-Za-z0-9]%'
            AND verifier_key_id COLLATE Latin1_General_100_BIN2
                NOT LIKE '%[^A-Za-z0-9._:-]%'
            AND DATALENGTH(verifier_key_id) = LEN(verifier_key_id)
        ),
        CONSTRAINT CK_legacy_session_compatibility_issue_window CHECK (
            original_expires_at = DATEADD(hour, 4, original_issued_at)
        ),
        CONSTRAINT CK_legacy_session_compatibility_state CHECK (
            compatibility_state IN ('active', 'revoked', 'incident')
        ),
        CONSTRAINT CK_legacy_session_compatibility_state_evidence CHECK (
            (
                compatibility_state = 'active'
                AND revoked_at IS NULL
                AND revocation_reason IS NULL
                AND incident_at IS NULL
                AND incident_code IS NULL
            )
            OR
            (
                compatibility_state = 'revoked'
                AND revoked_at IS NOT NULL
                AND revocation_reason IS NOT NULL
                AND revocation_reason COLLATE Latin1_General_100_BIN2 LIKE '[a-z0-9]%'
                AND revocation_reason COLLATE Latin1_General_100_BIN2 NOT LIKE '%[^a-z0-9-]%'
                AND incident_at IS NULL
                AND incident_code IS NULL
            )
            OR
            (
                compatibility_state = 'incident'
                AND revoked_at IS NULL
                AND revocation_reason IS NULL
                AND incident_at IS NOT NULL
                AND incident_code IS NOT NULL
                AND incident_code COLLATE Latin1_General_100_BIN2 LIKE '[a-z0-9]%'
                AND incident_code COLLATE Latin1_General_100_BIN2 NOT LIKE '%[^a-z0-9-]%'
            )
        )
    );

    CREATE INDEX IX_legacy_session_compatibility_subject_authority
        ON dbo.legacy_session_compatibility (
            subject_id,
            compatibility_state,
            original_expires_at
        );

    CREATE TABLE dbo.session_authority_control (
        control_id tinyint NOT NULL,
        control_version bigint NOT NULL
            CONSTRAINT DF_session_authority_control_version DEFAULT (1),
        authority_generation bigint NOT NULL
            CONSTRAINT DF_session_authority_control_generation DEFAULT (1),
        global_session_epoch bigint NOT NULL
            CONSTRAINT DF_session_authority_control_epoch DEFAULT (1),
        login_lookup_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        login_lookup_key_commitment binary(32) NULL,
        account_mapping_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        account_mapping_key_commitment binary(32) NULL,
        keyset_login_lookup_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        keyset_login_lookup_key_commitment binary(32) NULL,
        keyset_account_mapping_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        keyset_account_mapping_key_commitment binary(32) NULL,
        target_verifier_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        target_verifier_key_commitment binary(32) NULL,
        legacy_compatibility_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        legacy_compatibility_key_commitment binary(32) NULL,
        credential_fingerprint_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        credential_fingerprint_key_commitment binary(32) NULL,
        face_challenge_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        face_challenge_key_commitment binary(32) NULL,
        legacy_signing_key_id varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        legacy_signing_key_commitment binary(32) NULL,
        authority_keyset_commitment binary(32) NULL,
        target_routes_enabled bit NOT NULL
            CONSTRAINT DF_session_authority_control_target_routes DEFAULT (0),
        target_session_issuance_enabled bit NOT NULL
            CONSTRAINT DF_session_authority_control_target_issuance DEFAULT (0),
        target_session_issuance_started_at datetime2(7) NULL,
        legacy_ledger_seeding_enabled bit NOT NULL
            CONSTRAINT DF_session_authority_control_seeding DEFAULT (0),
        legacy_ledger_seeding_started_at datetime2(7) NULL,
        legacy_ledger_continuous_since datetime2(7) NULL,
        legacy_ledger_continuity_version bigint NOT NULL
            CONSTRAINT DF_session_authority_control_continuity_version DEFAULT (1),
        legacy_ledger_heartbeat_owner_id uniqueidentifier NULL,
        legacy_ledger_heartbeat_at datetime2(7) NULL,
        legacy_ledger_lease_expires_at datetime2(7) NULL,
        legacy_ledger_qualified_at datetime2(7) NULL,
        legacy_compatibility_enforcement_enabled bit NOT NULL
            CONSTRAINT DF_session_authority_control_enforcement DEFAULT (0),
        legacy_compatibility_enforced_at datetime2(7) NULL,
        subject_target_adoption_enabled bit NOT NULL
            CONSTRAINT DF_session_authority_control_subject_adoption DEFAULT (0),
        subject_target_adoption_started_at datetime2(7) NULL,
        target_acceptance_started_at datetime2(7) NULL,
        legacy_handle_issuance_enabled bit NOT NULL
            CONSTRAINT DF_session_authority_control_legacy_issuance DEFAULT (1),
        legacy_handle_issuance_stopped_at datetime2(7) NULL,
        legacy_handle_acceptance_enabled bit NOT NULL
            CONSTRAINT DF_session_authority_control_legacy_acceptance DEFAULT (1),
        legacy_handle_acceptance_disabled_at datetime2(7) NULL,
        legacy_hard_sunset_at datetime2(7) NULL,
        incident_state varchar(16) COLLATE Latin1_General_100_BIN2 NOT NULL
            CONSTRAINT DF_session_authority_control_incident_state DEFAULT ('normal'),
        incident_recorded_at datetime2(7) NULL,
        incident_code varchar(128) COLLATE Latin1_General_100_BIN2 NULL,
        target_verifier_key_incident_at datetime2(7) NULL,
        legacy_verifier_key_incident_at datetime2(7) NULL,
        created_at datetime2(7) NOT NULL
            CONSTRAINT DF_session_authority_control_created_at DEFAULT SYSUTCDATETIME(),
        updated_at datetime2(7) NOT NULL
            CONSTRAINT DF_session_authority_control_updated_at DEFAULT SYSUTCDATETIME(),
        CONSTRAINT PK_session_authority_control PRIMARY KEY (control_id),
        CONSTRAINT CK_session_authority_control_singleton CHECK (control_id = 1),
        CONSTRAINT CK_session_authority_control_versions CHECK (
            control_version BETWEEN 1 AND 9007199254740991
            AND authority_generation BETWEEN 1 AND 9007199254740991
            AND global_session_epoch BETWEEN 1 AND 9007199254740991
            AND legacy_ledger_continuity_version BETWEEN 1 AND 9007199254740991
        ),
        CONSTRAINT CK_session_authority_control_login_lookup_key CHECK (
            (
                login_lookup_key_id IS NULL
                AND login_lookup_key_commitment IS NULL
            )
            OR (
                login_lookup_key_id IS NOT NULL
                AND LEN(login_lookup_key_id) BETWEEN 1 AND 128
                AND login_lookup_key_commitment IS NOT NULL
                AND DATALENGTH(login_lookup_key_commitment) = 32
            )
        ),
        CONSTRAINT CK_session_authority_control_authority_keyset CHECK (
            (
                login_lookup_key_id IS NULL
                AND login_lookup_key_commitment IS NULL
                AND account_mapping_key_id IS NULL
                AND account_mapping_key_commitment IS NULL
                AND keyset_login_lookup_key_id IS NULL
                AND keyset_login_lookup_key_commitment IS NULL
                AND keyset_account_mapping_key_id IS NULL
                AND keyset_account_mapping_key_commitment IS NULL
                AND target_verifier_key_id IS NULL
                AND target_verifier_key_commitment IS NULL
                AND legacy_compatibility_key_id IS NULL
                AND legacy_compatibility_key_commitment IS NULL
                AND credential_fingerprint_key_id IS NULL
                AND credential_fingerprint_key_commitment IS NULL
                AND face_challenge_key_id IS NULL
                AND face_challenge_key_commitment IS NULL
                AND legacy_signing_key_id IS NULL
                AND legacy_signing_key_commitment IS NULL
                AND authority_keyset_commitment IS NULL
            )
            OR (
                login_lookup_key_id IS NOT NULL
                AND login_lookup_key_commitment IS NOT NULL
                AND account_mapping_key_id IS NOT NULL
                AND LEN(account_mapping_key_id) BETWEEN 1 AND 128
                AND account_mapping_key_commitment IS NOT NULL
                AND keyset_login_lookup_key_id IS NOT NULL
                AND LEN(keyset_login_lookup_key_id) BETWEEN 1 AND 128
                AND keyset_login_lookup_key_commitment IS NOT NULL
                AND keyset_account_mapping_key_id IS NOT NULL
                AND LEN(keyset_account_mapping_key_id) BETWEEN 1 AND 128
                AND keyset_account_mapping_key_commitment IS NOT NULL
                AND target_verifier_key_id IS NOT NULL
                AND LEN(target_verifier_key_id) BETWEEN 1 AND 128
                AND target_verifier_key_commitment IS NOT NULL
                AND legacy_compatibility_key_id IS NOT NULL
                AND LEN(legacy_compatibility_key_id) BETWEEN 1 AND 128
                AND legacy_compatibility_key_commitment IS NOT NULL
                AND credential_fingerprint_key_id IS NOT NULL
                AND LEN(credential_fingerprint_key_id) BETWEEN 1 AND 128
                AND credential_fingerprint_key_commitment IS NOT NULL
                AND face_challenge_key_id IS NOT NULL
                AND LEN(face_challenge_key_id) BETWEEN 1 AND 128
                AND face_challenge_key_commitment IS NOT NULL
                AND legacy_signing_key_id IS NOT NULL
                AND LEN(legacy_signing_key_id) BETWEEN 1 AND 128
                AND legacy_signing_key_commitment IS NOT NULL
                AND authority_keyset_commitment IS NOT NULL
            )
        ),
        CONSTRAINT CK_session_authority_control_key_id_tokens CHECK (
            login_lookup_key_id IS NULL
            OR (
                login_lookup_key_id COLLATE Latin1_General_100_BIN2 LIKE '[A-Za-z0-9]%'
                AND login_lookup_key_id COLLATE Latin1_General_100_BIN2
                    NOT LIKE '%[^A-Za-z0-9._:-]%'
                AND DATALENGTH(login_lookup_key_id) = LEN(login_lookup_key_id)
                AND account_mapping_key_id COLLATE Latin1_General_100_BIN2
                    LIKE '[A-Za-z0-9]%'
                AND account_mapping_key_id COLLATE Latin1_General_100_BIN2
                    NOT LIKE '%[^A-Za-z0-9._:-]%'
                AND DATALENGTH(account_mapping_key_id) = LEN(account_mapping_key_id)
                AND keyset_login_lookup_key_id COLLATE Latin1_General_100_BIN2
                    LIKE '[A-Za-z0-9]%'
                AND keyset_login_lookup_key_id COLLATE Latin1_General_100_BIN2
                    NOT LIKE '%[^A-Za-z0-9._:-]%'
                AND DATALENGTH(keyset_login_lookup_key_id)
                    = LEN(keyset_login_lookup_key_id)
                AND keyset_account_mapping_key_id COLLATE Latin1_General_100_BIN2
                    LIKE '[A-Za-z0-9]%'
                AND keyset_account_mapping_key_id COLLATE Latin1_General_100_BIN2
                    NOT LIKE '%[^A-Za-z0-9._:-]%'
                AND DATALENGTH(keyset_account_mapping_key_id)
                    = LEN(keyset_account_mapping_key_id)
                AND target_verifier_key_id COLLATE Latin1_General_100_BIN2
                    LIKE '[A-Za-z0-9]%'
                AND target_verifier_key_id COLLATE Latin1_General_100_BIN2
                    NOT LIKE '%[^A-Za-z0-9._:-]%'
                AND DATALENGTH(target_verifier_key_id) = LEN(target_verifier_key_id)
                AND legacy_compatibility_key_id COLLATE Latin1_General_100_BIN2
                    LIKE '[A-Za-z0-9]%'
                AND legacy_compatibility_key_id COLLATE Latin1_General_100_BIN2
                    NOT LIKE '%[^A-Za-z0-9._:-]%'
                AND DATALENGTH(legacy_compatibility_key_id)
                    = LEN(legacy_compatibility_key_id)
                AND credential_fingerprint_key_id COLLATE Latin1_General_100_BIN2
                    LIKE '[A-Za-z0-9]%'
                AND credential_fingerprint_key_id COLLATE Latin1_General_100_BIN2
                    NOT LIKE '%[^A-Za-z0-9._:-]%'
                AND DATALENGTH(credential_fingerprint_key_id)
                    = LEN(credential_fingerprint_key_id)
                AND face_challenge_key_id COLLATE Latin1_General_100_BIN2
                    LIKE '[A-Za-z0-9]%'
                AND face_challenge_key_id COLLATE Latin1_General_100_BIN2
                    NOT LIKE '%[^A-Za-z0-9._:-]%'
                AND DATALENGTH(face_challenge_key_id) = LEN(face_challenge_key_id)
                AND legacy_signing_key_id COLLATE Latin1_General_100_BIN2
                    LIKE '[A-Za-z0-9]%'
                AND legacy_signing_key_id COLLATE Latin1_General_100_BIN2
                    NOT LIKE '%[^A-Za-z0-9._:-]%'
                AND DATALENGTH(legacy_signing_key_id) = LEN(legacy_signing_key_id)
            )
        ),
        CONSTRAINT CK_session_authority_control_continuity_lease CHECK (
            (
                legacy_ledger_heartbeat_owner_id IS NULL
                AND legacy_ledger_heartbeat_at IS NULL
                AND legacy_ledger_lease_expires_at IS NULL
            )
            OR (
                legacy_ledger_heartbeat_owner_id IS NOT NULL
                AND legacy_ledger_heartbeat_at IS NOT NULL
                AND legacy_ledger_lease_expires_at IS NOT NULL
                AND legacy_ledger_continuous_since IS NOT NULL
                AND legacy_ledger_heartbeat_at >= legacy_ledger_continuous_since
                AND legacy_ledger_lease_expires_at
                    = DATEADD(second, 120, legacy_ledger_heartbeat_at)
            )
        ),
        CONSTRAINT CK_session_authority_control_target_issuance CHECK (
            target_session_issuance_enabled = 0
            OR (
                target_routes_enabled = 1
                AND target_session_issuance_started_at IS NOT NULL
            )
        ),
        CONSTRAINT CK_session_authority_control_seeding CHECK (
            legacy_ledger_seeding_enabled = 0
            OR (
                legacy_ledger_seeding_started_at IS NOT NULL
                AND legacy_ledger_continuous_since IS NOT NULL
            )
        ),
        CONSTRAINT CK_session_authority_control_enforcement CHECK (
            (
                legacy_compatibility_enforcement_enabled = 0
                AND legacy_compatibility_enforced_at IS NULL
            )
            OR
            (
                legacy_compatibility_enforcement_enabled = 1
                AND legacy_compatibility_enforced_at IS NOT NULL
                AND legacy_ledger_qualified_at IS NOT NULL
                AND legacy_ledger_continuous_since IS NOT NULL
                AND legacy_ledger_qualified_at
                    >= DATEADD(hour, 4, legacy_ledger_continuous_since)
            )
        ),
        CONSTRAINT CK_session_authority_control_adoption CHECK (
            subject_target_adoption_enabled = 0
            OR subject_target_adoption_started_at IS NOT NULL
        ),
        CONSTRAINT CK_session_authority_control_target_after_enforcement CHECK (
            legacy_compatibility_enforcement_enabled = 1
            OR (
                target_session_issuance_enabled = 0
                AND subject_target_adoption_enabled = 0
                AND target_acceptance_started_at IS NULL
            )
        ),
        CONSTRAINT CK_session_authority_control_target_window CHECK (
            (
                target_session_issuance_enabled = subject_target_adoption_enabled
                AND target_session_issuance_started_at IS NULL
                AND subject_target_adoption_started_at IS NULL
                AND target_acceptance_started_at IS NULL
                AND legacy_hard_sunset_at IS NULL
                AND target_session_issuance_enabled = 0
            )
            OR (
                target_session_issuance_enabled = subject_target_adoption_enabled
                AND target_session_issuance_started_at IS NOT NULL
                AND subject_target_adoption_started_at = target_session_issuance_started_at
                AND target_acceptance_started_at IS NOT NULL
                AND target_acceptance_started_at = target_session_issuance_started_at
                AND legacy_hard_sunset_at IS NOT NULL
            )
        ),
        CONSTRAINT CK_session_authority_control_seeding_during_legacy_issuance CHECK (
            legacy_compatibility_enforcement_enabled = 0
            OR legacy_handle_issuance_enabled = 0
            OR legacy_ledger_seeding_enabled = 1
        ),
        CONSTRAINT CK_session_authority_control_legacy_issuance CHECK (
            (
                legacy_handle_issuance_enabled = 1
                AND legacy_handle_issuance_stopped_at IS NULL
            )
            OR
            (
                legacy_handle_issuance_enabled = 0
                AND legacy_handle_issuance_stopped_at IS NOT NULL
            )
        ),
        CONSTRAINT CK_session_authority_control_legacy_acceptance CHECK (
            (
                legacy_handle_acceptance_enabled = 1
                AND legacy_handle_acceptance_disabled_at IS NULL
            )
            OR
            (
                legacy_handle_acceptance_enabled = 0
                AND legacy_handle_acceptance_disabled_at IS NOT NULL
                AND legacy_handle_issuance_enabled = 0
            )
        ),
        CONSTRAINT CK_session_authority_control_sunset CHECK (
            (
                target_acceptance_started_at IS NULL
                AND legacy_hard_sunset_at IS NULL
            )
            OR
            (
                target_acceptance_started_at IS NOT NULL
                AND legacy_hard_sunset_at = DATEADD(day, 7, target_acceptance_started_at)
            )
        ),
        CONSTRAINT CK_session_authority_control_stop_order CHECK (
            legacy_handle_issuance_stopped_at IS NULL
            OR legacy_handle_acceptance_disabled_at IS NULL
            OR legacy_handle_acceptance_disabled_at >= legacy_handle_issuance_stopped_at
        ),
        CONSTRAINT CK_session_authority_control_final_aging CHECK (
            legacy_handle_acceptance_disabled_at IS NULL
            OR (
                legacy_handle_issuance_stopped_at IS NOT NULL
                AND (
                    legacy_verifier_key_incident_at IS NOT NULL
                    OR legacy_handle_acceptance_disabled_at
                        >= DATEADD(hour, 4, legacy_handle_issuance_stopped_at)
                )
            )
        ),
        CONSTRAINT CK_session_authority_control_stop_before_sunset CHECK (
            legacy_hard_sunset_at IS NULL
            OR legacy_handle_issuance_stopped_at IS NULL
            OR legacy_verifier_key_incident_at IS NOT NULL
            OR legacy_handle_issuance_stopped_at
                <= DATEADD(hour, -4, legacy_hard_sunset_at)
        ),
        CONSTRAINT CK_session_authority_control_incident CHECK (
            incident_state IN ('normal', 'suspended', 'recovering')
            AND (
                incident_state = 'normal'
                OR (
                    incident_recorded_at IS NOT NULL
                    AND incident_code IS NOT NULL
                    AND LEN(incident_code) > 0
                )
            )
            AND (
                (incident_recorded_at IS NULL AND incident_code IS NULL)
                OR (incident_recorded_at IS NOT NULL AND incident_code IS NOT NULL)
            )
        ),
        CONSTRAINT CK_session_authority_control_incident_code CHECK (
            incident_code IS NULL
            OR (
                DATALENGTH(incident_code) BETWEEN 1 AND 128
                AND DATALENGTH(incident_code) = LEN(incident_code)
                AND incident_code COLLATE Latin1_General_100_BIN2 LIKE '[a-z0-9]%'
                AND incident_code COLLATE Latin1_General_100_BIN2 NOT LIKE '%[^a-z0-9-]%'
            )
        ),
        CONSTRAINT CK_session_authority_control_update_time CHECK (updated_at >= created_at)
    );

    INSERT INTO dbo.session_authority_control (
        control_id,
        control_version,
        authority_generation,
        global_session_epoch,
        target_routes_enabled,
        target_session_issuance_enabled,
        legacy_ledger_seeding_enabled,
        legacy_compatibility_enforcement_enabled,
        subject_target_adoption_enabled,
        legacy_handle_issuance_enabled,
        legacy_handle_acceptance_enabled,
        incident_state
    )
    VALUES (1, 1, 1, 1, 0, 0, 0, 0, 0, 1, 1, 'normal');

    EXEC sys.sp_executesql N'
        CREATE TRIGGER dbo.trg_learning_subject_authority_evidence
        ON dbo.learning_subject
        AFTER UPDATE
        AS
        BEGIN
            SET NOCOUNT ON;

            IF UPDATE(subject_id) OR UPDATE(created_at)
                THROW 51001, ''Subject identity and creation time are immutable.'', 1;

            IF EXISTS (
                SELECT 1
                FROM inserted AS i
                INNER JOIN deleted AS d ON d.subject_id = i.subject_id
                WHERE i.credential_version < d.credential_version
                    OR i.subject_session_epoch < d.subject_session_epoch
                    OR (
                        d.legacy_authority_disabled_at IS NOT NULL
                        AND (
                            i.legacy_authority_disabled_at IS NULL
                            OR i.legacy_authority_disabled_at <> d.legacy_authority_disabled_at
                        )
                    )
            )
                THROW 51002, ''Subject versions and legacy cutoffs only move forward.'', 1;
        END;
    ';

    EXEC sys.sp_executesql N'
        CREATE TRIGGER dbo.trg_learning_session_authority_evidence
        ON dbo.learning_session
        AFTER UPDATE
        AS
        BEGIN
            SET NOCOUNT ON;

            IF UPDATE(session_id)
                OR UPDATE(identifier_verifier)
                OR UPDATE(verifier_key_id)
                OR UPDATE(subject_id)
                OR UPDATE(original_issued_at)
                OR UPDATE(phase_started_at)
                OR UPDATE(absolute_expires_at)
                OR UPDATE(face_auth_required)
                OR UPDATE(registration_required)
                OR UPDATE(subject_epoch_snapshot)
                OR UPDATE(credential_version_snapshot)
                OR UPDATE(global_epoch_snapshot)
                OR UPDATE(authority_generation_snapshot)
                OR UPDATE(created_at)
                THROW 51003, ''Session identity, authority snapshots, and deadlines are immutable.'', 1;

            IF EXISTS (
                SELECT 1
                FROM inserted AS i
                INNER JOIN deleted AS d ON d.session_id = i.session_id
                WHERE (
                    d.phase IN (''expired'', ''revoked'', ''rotated-out'')
                    AND i.phase <> d.phase
                )
                OR (
                    d.phase NOT IN (''expired'', ''revoked'', ''rotated-out'')
                    AND i.phase <> d.phase
                    AND i.phase NOT IN (''expired'', ''revoked'', ''rotated-out'')
                )
                OR (
                    d.revoked_at IS NOT NULL
                    AND (
                        i.revoked_at IS NULL
                        OR i.revoked_at <> d.revoked_at
                        OR ISNULL(i.revocation_reason, '''') <> ISNULL(d.revocation_reason, '''')
                    )
                )
                OR (
                    d.replacement_session_id IS NOT NULL
                    AND (
                        i.replacement_session_id IS NULL
                        OR i.replacement_session_id <> d.replacement_session_id
                    )
                )
            )
                THROW 51004, ''Terminal session and rotation evidence cannot be changed.'', 1;
        END;
    ';

    EXEC sys.sp_executesql N'
        CREATE TRIGGER dbo.trg_learning_session_flow_binding
        ON dbo.learning_session_flow
        AFTER INSERT, UPDATE
        AS
        BEGIN
            SET NOCOUNT ON;

            IF EXISTS (SELECT 1 FROM deleted)
                AND (UPDATE(flow_id) OR UPDATE(subject_id) OR UPDATE(created_at))
                THROW 51005, ''Flow identity and subject binding are immutable.'', 1;

            IF EXISTS (
                SELECT 1
                FROM inserted AS i
                INNER JOIN deleted AS d ON d.flow_id = i.flow_id
                WHERE d.encrypted_provider_challenge_reference IS NOT NULL
                    AND (
                        i.encrypted_provider_challenge_reference IS NULL
                        OR i.encrypted_provider_challenge_reference <> d.encrypted_provider_challenge_reference
                        OR ISNULL(i.provider_reference_encryption_key_id, '''')
                            <> ISNULL(d.provider_reference_encryption_key_id, '''')
                        OR i.challenge_created_at <> d.challenge_created_at
                        OR i.challenge_session_id IS NULL
                        OR i.challenge_session_id <> d.challenge_session_id
                    )
                OR (
                    d.challenge_session_id IS NOT NULL
                    AND (
                        i.challenge_session_id IS NULL
                        OR i.challenge_session_id <> d.challenge_session_id
                    )
                )
                OR (
                    d.challenge_state IN (''consumed'', ''failed'')
                    AND (
                        i.challenge_state <> d.challenge_state
                        OR i.challenge_resolved_at <> d.challenge_resolved_at
                    )
                )
            )
                THROW 51006, ''Bound or resolved Face challenge evidence cannot be replaced.'', 1;

            IF EXISTS (
                SELECT 1
                FROM inserted AS i
                LEFT JOIN dbo.learning_session AS challenge
                    ON challenge.session_id = i.challenge_session_id
                    AND challenge.subject_id = i.subject_id
                LEFT JOIN dbo.learning_session AS current_session
                    ON current_session.session_id = i.current_session_id
                    AND current_session.subject_id = i.subject_id
                WHERE (
                    i.challenge_state = ''active''
                    AND (
                        challenge.session_id IS NULL
                        OR challenge.session_id <> current_session.session_id
                        OR challenge.phase <> ''face-pending''
                        OR challenge.face_auth_required <> 1
                        OR challenge.registration_required <> current_session.registration_required
                        OR challenge.absolute_expires_at
                            <> DATEADD(minute, 20, challenge.original_issued_at)
                        OR i.challenge_created_at <> challenge.phase_started_at
                    )
                )
                OR (
                    i.challenge_state = ''consumed''
                    AND (
                        challenge.session_id IS NULL
                        OR current_session.session_id IS NULL
                        OR challenge.phase <> ''rotated-out''
                        OR challenge.face_auth_required <> 1
                        OR current_session.phase <> ''authenticated''
                        OR current_session.face_auth_required <> 1
                        OR challenge.registration_required <> current_session.registration_required
                        OR challenge.subject_epoch_snapshot <> current_session.subject_epoch_snapshot
                        OR challenge.credential_version_snapshot
                            <> current_session.credential_version_snapshot
                        OR challenge.global_epoch_snapshot <> current_session.global_epoch_snapshot
                        OR challenge.authority_generation_snapshot
                            <> current_session.authority_generation_snapshot
                        OR challenge.absolute_expires_at
                            <> DATEADD(minute, 20, challenge.original_issued_at)
                        OR i.challenge_created_at <> challenge.phase_started_at
                        OR challenge.replacement_session_id <> current_session.session_id
                        OR current_session.original_issued_at <> i.challenge_resolved_at
                        OR current_session.phase_started_at <> i.challenge_resolved_at
                        OR current_session.created_at <> i.challenge_resolved_at
                    )
                )
            )
                THROW 51014, ''Face challenge lineage is inconsistent.'', 1;
        END;
    ';

    EXEC sys.sp_executesql N'
        CREATE TRIGGER dbo.trg_legacy_session_compatibility_binding
        ON dbo.legacy_session_compatibility
        AFTER UPDATE
        AS
        BEGIN
            SET NOCOUNT ON;

            IF UPDATE(compatibility_id)
                OR UPDATE(legacy_handle_verifier)
                OR UPDATE(verifier_key_id)
                OR UPDATE(subject_id)
                OR UPDATE(original_issued_at)
                OR UPDATE(original_expires_at)
                OR UPDATE(created_at)
                THROW 51007, ''Legacy verifier bindings and issue metadata are immutable.'', 1;

            IF EXISTS (
                SELECT 1
                FROM inserted AS i
                INNER JOIN deleted AS d ON d.compatibility_id = i.compatibility_id
                WHERE d.compatibility_state IN (''revoked'', ''incident'')
                    AND (
                        i.compatibility_state <> d.compatibility_state
                        OR ISNULL(i.revoked_at, CONVERT(datetime2(7), ''0001-01-01T00:00:00''))
                            <> ISNULL(d.revoked_at, CONVERT(datetime2(7), ''0001-01-01T00:00:00''))
                        OR ISNULL(i.revocation_reason, '''') <> ISNULL(d.revocation_reason, '''')
                        OR ISNULL(i.incident_at, CONVERT(datetime2(7), ''0001-01-01T00:00:00''))
                            <> ISNULL(d.incident_at, CONVERT(datetime2(7), ''0001-01-01T00:00:00''))
                        OR ISNULL(i.incident_code, '''') <> ISNULL(d.incident_code, '''')
                    )
            )
                THROW 51008, ''Terminal legacy binding evidence cannot be changed.'', 1;
        END;
    ';

    EXEC sys.sp_executesql N'
        CREATE TRIGGER dbo.trg_session_authority_control_irreversible
        ON dbo.session_authority_control
        AFTER UPDATE
        AS
        BEGIN
            SET NOCOUNT ON;

            DECLARE @continuity_only bit = 0;
            IF UPDATE(legacy_ledger_continuity_version)
                AND NOT UPDATE(control_version)
                AND NOT UPDATE(authority_generation)
                AND NOT UPDATE(global_session_epoch)
                AND NOT UPDATE(login_lookup_key_id)
                AND NOT UPDATE(login_lookup_key_commitment)
                AND NOT UPDATE(account_mapping_key_id)
                AND NOT UPDATE(account_mapping_key_commitment)
                AND NOT UPDATE(keyset_login_lookup_key_id)
                AND NOT UPDATE(keyset_login_lookup_key_commitment)
                AND NOT UPDATE(keyset_account_mapping_key_id)
                AND NOT UPDATE(keyset_account_mapping_key_commitment)
                AND NOT UPDATE(target_verifier_key_id)
                AND NOT UPDATE(target_verifier_key_commitment)
                AND NOT UPDATE(legacy_compatibility_key_id)
                AND NOT UPDATE(legacy_compatibility_key_commitment)
                AND NOT UPDATE(credential_fingerprint_key_id)
                AND NOT UPDATE(credential_fingerprint_key_commitment)
                AND NOT UPDATE(face_challenge_key_id)
                AND NOT UPDATE(face_challenge_key_commitment)
                AND NOT UPDATE(legacy_signing_key_id)
                AND NOT UPDATE(legacy_signing_key_commitment)
                AND NOT UPDATE(authority_keyset_commitment)
                AND NOT UPDATE(target_routes_enabled)
                AND NOT UPDATE(target_session_issuance_enabled)
                AND NOT UPDATE(target_session_issuance_started_at)
                AND NOT UPDATE(legacy_ledger_seeding_enabled)
                AND NOT UPDATE(legacy_ledger_seeding_started_at)
                AND NOT UPDATE(legacy_ledger_qualified_at)
                AND NOT UPDATE(legacy_compatibility_enforcement_enabled)
                AND NOT UPDATE(legacy_compatibility_enforced_at)
                AND NOT UPDATE(subject_target_adoption_enabled)
                AND NOT UPDATE(subject_target_adoption_started_at)
                AND NOT UPDATE(target_acceptance_started_at)
                AND NOT UPDATE(legacy_handle_issuance_enabled)
                AND NOT UPDATE(legacy_handle_issuance_stopped_at)
                AND NOT UPDATE(legacy_handle_acceptance_enabled)
                AND NOT UPDATE(legacy_handle_acceptance_disabled_at)
                AND NOT UPDATE(legacy_hard_sunset_at)
                AND NOT UPDATE(incident_state)
                AND NOT UPDATE(incident_recorded_at)
                AND NOT UPDATE(incident_code)
                AND NOT UPDATE(target_verifier_key_incident_at)
                AND NOT UPDATE(legacy_verifier_key_incident_at)
                SET @continuity_only = 1;

            IF UPDATE(control_id) OR UPDATE(created_at)
                THROW 51009, ''The authority-control identity and creation time are immutable.'', 1;

            IF UPDATE(login_lookup_key_id)
                OR UPDATE(login_lookup_key_commitment)
                OR UPDATE(account_mapping_key_id)
                OR UPDATE(account_mapping_key_commitment)
                OR UPDATE(keyset_login_lookup_key_id)
                OR UPDATE(keyset_login_lookup_key_commitment)
                OR UPDATE(keyset_account_mapping_key_id)
                OR UPDATE(keyset_account_mapping_key_commitment)
                OR UPDATE(target_verifier_key_id)
                OR UPDATE(target_verifier_key_commitment)
                OR UPDATE(legacy_compatibility_key_id)
                OR UPDATE(legacy_compatibility_key_commitment)
                OR UPDATE(credential_fingerprint_key_id)
                OR UPDATE(credential_fingerprint_key_commitment)
                OR UPDATE(face_challenge_key_id)
                OR UPDATE(face_challenge_key_commitment)
                OR UPDATE(legacy_signing_key_id)
                OR UPDATE(legacy_signing_key_commitment)
                OR UPDATE(authority_keyset_commitment)
            BEGIN
                IF EXISTS (
                        SELECT 1
                        FROM inserted AS i
                        INNER JOIN deleted AS d ON d.control_id = i.control_id
                        WHERE d.login_lookup_key_id IS NULL
                            AND (
                                NOT (
                                    UPDATE(login_lookup_key_id)
                                    AND UPDATE(login_lookup_key_commitment)
                                    AND UPDATE(account_mapping_key_id)
                                    AND UPDATE(account_mapping_key_commitment)
                                    AND UPDATE(keyset_login_lookup_key_id)
                                    AND UPDATE(keyset_login_lookup_key_commitment)
                                    AND UPDATE(keyset_account_mapping_key_id)
                                    AND UPDATE(keyset_account_mapping_key_commitment)
                                    AND UPDATE(target_verifier_key_id)
                                    AND UPDATE(target_verifier_key_commitment)
                                    AND UPDATE(legacy_compatibility_key_id)
                                    AND UPDATE(legacy_compatibility_key_commitment)
                                    AND UPDATE(credential_fingerprint_key_id)
                                    AND UPDATE(credential_fingerprint_key_commitment)
                                    AND UPDATE(face_challenge_key_id)
                                    AND UPDATE(face_challenge_key_commitment)
                                    AND UPDATE(legacy_signing_key_id)
                                    AND UPDATE(legacy_signing_key_commitment)
                                    AND UPDATE(authority_keyset_commitment)
                                )
                            OR d.login_lookup_key_commitment IS NOT NULL
                            OR i.login_lookup_key_id IS NULL
                            OR i.login_lookup_key_commitment IS NULL
                            OR i.account_mapping_key_id IS NULL
                            OR i.account_mapping_key_commitment IS NULL
                            OR i.keyset_login_lookup_key_id IS NULL
                            OR i.keyset_login_lookup_key_commitment IS NULL
                            OR i.keyset_account_mapping_key_id IS NULL
                            OR i.keyset_account_mapping_key_commitment IS NULL
                            OR i.target_verifier_key_id IS NULL
                            OR i.target_verifier_key_commitment IS NULL
                            OR i.legacy_compatibility_key_id IS NULL
                            OR i.legacy_compatibility_key_commitment IS NULL
                            OR i.credential_fingerprint_key_id IS NULL
                            OR i.credential_fingerprint_key_commitment IS NULL
                            OR i.face_challenge_key_id IS NULL
                            OR i.face_challenge_key_commitment IS NULL
                            OR i.legacy_signing_key_id IS NULL
                            OR i.legacy_signing_key_commitment IS NULL
                            OR i.authority_keyset_commitment IS NULL
                            OR i.target_routes_enabled <> 0
                            OR i.target_session_issuance_enabled <> 0
                            OR i.target_session_issuance_started_at IS NOT NULL
                            OR i.legacy_ledger_seeding_enabled <> 0
                            OR i.legacy_ledger_seeding_started_at IS NOT NULL
                            OR i.legacy_ledger_continuous_since IS NOT NULL
                            OR i.legacy_ledger_qualified_at IS NOT NULL
                            OR i.legacy_ledger_heartbeat_owner_id IS NOT NULL
                            OR i.legacy_ledger_heartbeat_at IS NOT NULL
                            OR i.legacy_ledger_lease_expires_at IS NOT NULL
                            OR i.legacy_compatibility_enforcement_enabled <> 0
                            OR i.legacy_compatibility_enforced_at IS NOT NULL
                            OR i.subject_target_adoption_enabled <> 0
                            OR i.subject_target_adoption_started_at IS NOT NULL
                            OR i.target_acceptance_started_at IS NOT NULL
                            OR i.legacy_handle_issuance_enabled <> 1
                            OR i.legacy_handle_issuance_stopped_at IS NOT NULL
                            OR i.legacy_handle_acceptance_enabled <> 1
                            OR i.legacy_handle_acceptance_disabled_at IS NOT NULL
                            OR i.legacy_hard_sunset_at IS NOT NULL
                            OR i.incident_state <> ''normal''
                            OR i.incident_recorded_at IS NOT NULL
                            OR i.incident_code IS NOT NULL
                            OR i.target_verifier_key_incident_at IS NOT NULL
                            OR i.legacy_verifier_key_incident_at IS NOT NULL
                            )
                    )
                    OR (
                        EXISTS (
                            SELECT 1
                            FROM inserted AS i
                            INNER JOIN deleted AS d ON d.control_id = i.control_id
                            WHERE d.login_lookup_key_id IS NULL
                        )
                        AND (
                            EXISTS (SELECT 1 FROM dbo.learning_subject)
                            OR EXISTS (SELECT 1 FROM dbo.learning_session)
                            OR EXISTS (SELECT 1 FROM dbo.learning_session_flow)
                            OR EXISTS (SELECT 1 FROM dbo.legacy_session_compatibility)
                        )
                    )
                    THROW 51011, ''Login-lookup key binding requires an empty dormant authority.'', 1;

                IF EXISTS (
                    SELECT 1
                    FROM inserted AS i
                    INNER JOIN deleted AS d ON d.control_id = i.control_id
                    WHERE d.login_lookup_key_id IS NOT NULL
                        AND (
                            i.login_lookup_key_id <> d.login_lookup_key_id
                            OR i.login_lookup_key_commitment <> d.login_lookup_key_commitment
                            OR i.account_mapping_key_id <> d.account_mapping_key_id
                            OR i.account_mapping_key_commitment <> d.account_mapping_key_commitment
                            OR i.keyset_login_lookup_key_id <> d.keyset_login_lookup_key_id
                            OR i.keyset_login_lookup_key_commitment
                                <> d.keyset_login_lookup_key_commitment
                            OR i.keyset_account_mapping_key_id
                                <> d.keyset_account_mapping_key_id
                            OR i.keyset_account_mapping_key_commitment
                                <> d.keyset_account_mapping_key_commitment
                            OR i.incident_state <> ''recovering''
                            OR d.incident_state <> ''suspended''
                            OR i.authority_generation <> d.authority_generation + 1
                            OR i.global_session_epoch <= d.global_session_epoch
                            OR i.target_session_issuance_enabled <> 0
                            OR i.subject_target_adoption_enabled <> 0
                            OR (
                                i.authority_keyset_commitment = d.authority_keyset_commitment
                                AND NOT (
                                    i.target_verifier_key_id = d.target_verifier_key_id
                                    AND i.target_verifier_key_commitment
                                        = d.target_verifier_key_commitment
                                    AND i.legacy_compatibility_key_id
                                        = d.legacy_compatibility_key_id
                                    AND i.legacy_compatibility_key_commitment
                                        = d.legacy_compatibility_key_commitment
                                    AND i.credential_fingerprint_key_id
                                        = d.credential_fingerprint_key_id
                                    AND i.credential_fingerprint_key_commitment
                                        = d.credential_fingerprint_key_commitment
                                    AND i.face_challenge_key_id = d.face_challenge_key_id
                                    AND i.face_challenge_key_commitment
                                        = d.face_challenge_key_commitment
                                )
                            )
                            OR (
                                i.authority_keyset_commitment <> d.authority_keyset_commitment
                                AND (
                                    i.target_verifier_key_id = d.target_verifier_key_id
                                    AND i.target_verifier_key_commitment
                                        = d.target_verifier_key_commitment
                                    AND i.legacy_compatibility_key_id
                                        = d.legacy_compatibility_key_id
                                    AND i.legacy_compatibility_key_commitment
                                        = d.legacy_compatibility_key_commitment
                                    AND i.credential_fingerprint_key_id
                                        = d.credential_fingerprint_key_id
                                    AND i.credential_fingerprint_key_commitment
                                        = d.credential_fingerprint_key_commitment
                                    AND i.face_challenge_key_id = d.face_challenge_key_id
                                    AND i.face_challenge_key_commitment
                                        = d.face_challenge_key_commitment
                                )
                            )
                            OR (
                                i.target_verifier_key_id = d.target_verifier_key_id
                                AND i.target_verifier_key_commitment
                                    = d.target_verifier_key_commitment
                                AND i.legacy_compatibility_key_id
                                    = d.legacy_compatibility_key_id
                                AND i.legacy_compatibility_key_commitment
                                    = d.legacy_compatibility_key_commitment
                                AND i.credential_fingerprint_key_id
                                    = d.credential_fingerprint_key_id
                                AND i.credential_fingerprint_key_commitment
                                    = d.credential_fingerprint_key_commitment
                                AND i.face_challenge_key_id = d.face_challenge_key_id
                                AND i.face_challenge_key_commitment
                                    = d.face_challenge_key_commitment
                                AND i.legacy_signing_key_id = d.legacy_signing_key_id
                                AND i.legacy_signing_key_commitment
                                    = d.legacy_signing_key_commitment
                            )
                        )
                )
                    THROW 51013, ''Authority key changes require fenced incident recovery.'', 1;
            END;

            IF EXISTS (
                SELECT 1
                FROM inserted AS i
                INNER JOIN deleted AS d ON d.control_id = i.control_id
                WHERE (
                        @continuity_only = 0
                        AND i.control_version <> d.control_version + 1
                    )
                    OR (
                        @continuity_only = 1
                        AND (
                            i.control_version <> d.control_version
                            OR i.legacy_ledger_continuity_version
                                <> d.legacy_ledger_continuity_version + 1
                            OR i.legacy_ledger_seeding_enabled <> 1
                            OR i.legacy_compatibility_enforcement_enabled <> 0
                            OR i.legacy_handle_issuance_enabled <> 1
                            OR i.legacy_ledger_heartbeat_owner_id IS NULL
                            OR i.legacy_ledger_heartbeat_at <> i.updated_at
                            OR i.legacy_ledger_lease_expires_at
                                <> DATEADD(second, 120, i.legacy_ledger_heartbeat_at)
                            OR (
                                (
                                    d.legacy_ledger_heartbeat_owner_id IS NULL
                                    OR d.legacy_ledger_lease_expires_at IS NULL
                                    OR d.legacy_ledger_lease_expires_at <= i.updated_at
                                    OR i.legacy_ledger_heartbeat_owner_id
                                        <> d.legacy_ledger_heartbeat_owner_id
                                )
                                AND i.legacy_ledger_continuous_since <> i.updated_at
                            )
                            OR (
                                i.legacy_ledger_continuous_since
                                    <> d.legacy_ledger_continuous_since
                                AND i.legacy_ledger_continuous_since <> i.updated_at
                            )
                        )
                    )
                    OR (
                        @continuity_only = 0
                        AND i.legacy_ledger_continuity_version
                            <> d.legacy_ledger_continuity_version
                        AND NOT (
                            d.incident_state = ''normal''
                            AND i.incident_state <> ''normal''
                            AND d.legacy_ledger_seeding_enabled = 1
                            AND d.legacy_compatibility_enforcement_enabled = 0
                            AND d.legacy_handle_issuance_enabled = 1
                            AND i.legacy_ledger_continuity_version
                                = d.legacy_ledger_continuity_version + 1
                            AND i.legacy_ledger_continuous_since = i.updated_at
                            AND i.legacy_ledger_heartbeat_owner_id IS NULL
                            AND i.legacy_ledger_heartbeat_at IS NULL
                            AND i.legacy_ledger_lease_expires_at IS NULL
                        )
                    )
                    OR (
                        @continuity_only = 0
                        AND (
                            ISNULL(CONVERT(varchar(36), i.legacy_ledger_heartbeat_owner_id), '''')
                                <> ISNULL(CONVERT(varchar(36), d.legacy_ledger_heartbeat_owner_id), '''')
                            OR ISNULL(i.legacy_ledger_heartbeat_at, CONVERT(datetime2(7), ''0001-01-01T00:00:00''))
                                <> ISNULL(d.legacy_ledger_heartbeat_at, CONVERT(datetime2(7), ''0001-01-01T00:00:00''))
                            OR ISNULL(i.legacy_ledger_lease_expires_at, CONVERT(datetime2(7), ''0001-01-01T00:00:00''))
                                <> ISNULL(d.legacy_ledger_lease_expires_at, CONVERT(datetime2(7), ''0001-01-01T00:00:00''))
                        )
                        AND NOT (
                            i.legacy_ledger_continuous_since = i.updated_at
                            AND (
                                d.legacy_ledger_continuous_since IS NULL
                                OR i.legacy_ledger_continuous_since
                                    > d.legacy_ledger_continuous_since
                                OR (
                                    d.incident_state = ''normal''
                                    AND i.incident_state <> ''normal''
                                    AND d.legacy_ledger_seeding_enabled = 1
                                    AND d.legacy_compatibility_enforcement_enabled = 0
                                    AND d.legacy_handle_issuance_enabled = 1
                                )
                            )
                            AND i.legacy_ledger_heartbeat_owner_id IS NULL
                            AND i.legacy_ledger_heartbeat_at IS NULL
                            AND i.legacy_ledger_lease_expires_at IS NULL
                        )
                    )
                    OR i.authority_generation < d.authority_generation
                    OR i.global_session_epoch < d.global_session_epoch
                    OR (
                        d.legacy_compatibility_enforcement_enabled = 1
                        AND i.legacy_compatibility_enforcement_enabled = 0
                    )
                    OR (
                        d.legacy_handle_issuance_enabled = 0
                        AND i.legacy_handle_issuance_enabled = 1
                    )
                    OR (
                        d.legacy_handle_acceptance_enabled = 0
                        AND i.legacy_handle_acceptance_enabled = 1
                    )
                    OR (
                        i.legacy_compatibility_enforcement_enabled = 0
                        AND (
                            i.target_session_issuance_enabled = 1
                            OR i.subject_target_adoption_enabled = 1
                            OR i.target_acceptance_started_at IS NOT NULL
                        )
                    )
                    OR (
                        i.legacy_compatibility_enforcement_enabled = 1
                        AND i.legacy_handle_issuance_enabled = 1
                        AND i.legacy_ledger_seeding_enabled = 0
                    )
                    OR (
                        (
                            i.target_session_issuance_enabled = 1
                            OR i.subject_target_adoption_enabled = 1
                        )
                        AND (
                            i.target_acceptance_started_at IS NULL
                            OR i.legacy_hard_sunset_at IS NULL
                            OR SYSUTCDATETIME() < i.target_acceptance_started_at
                        )
                    )
                    OR (
                        d.target_session_issuance_enabled = 0
                        AND i.target_session_issuance_enabled = 1
                        AND d.target_session_issuance_started_at IS NULL
                        AND i.target_session_issuance_started_at <> i.updated_at
                    )
                    OR (
                        d.target_session_issuance_started_at IS NULL
                        AND i.target_session_issuance_started_at IS NOT NULL
                        AND NOT (
                            d.target_session_issuance_enabled = 0
                            AND i.target_session_issuance_enabled = 1
                        )
                    )
                    OR (
                        d.subject_target_adoption_enabled = 0
                        AND i.subject_target_adoption_enabled = 1
                        AND d.subject_target_adoption_started_at IS NULL
                        AND i.subject_target_adoption_started_at <> i.updated_at
                    )
                    OR (
                        d.subject_target_adoption_started_at IS NULL
                        AND i.subject_target_adoption_started_at IS NOT NULL
                        AND NOT (
                            d.subject_target_adoption_enabled = 0
                            AND i.subject_target_adoption_enabled = 1
                        )
                    )
                    OR (
                        d.target_acceptance_started_at IS NULL
                        AND i.target_acceptance_started_at IS NOT NULL
                        AND (
                            i.target_acceptance_started_at <> i.updated_at
                            OR d.legacy_hard_sunset_at IS NOT NULL
                            OR i.legacy_hard_sunset_at <> DATEADD(day, 7, i.updated_at)
                        )
                    )
                    OR (
                        d.target_acceptance_started_at IS NULL
                        AND i.target_acceptance_started_at IS NULL
                        AND d.legacy_hard_sunset_at IS NULL
                        AND i.legacy_hard_sunset_at IS NOT NULL
                    )
                    OR (
                        d.legacy_ledger_seeding_enabled = 0
                        AND i.legacy_ledger_seeding_enabled = 1
                        AND (
                            i.legacy_ledger_seeding_started_at IS NULL
                            OR i.legacy_ledger_continuous_since IS NULL
                            OR i.legacy_ledger_continuous_since <> i.updated_at
                            OR (
                                d.legacy_ledger_seeding_started_at IS NULL
                                AND i.legacy_ledger_seeding_started_at <> i.updated_at
                            )
                        )
                    )
                    OR (
                        d.legacy_ledger_seeding_started_at IS NULL
                        AND i.legacy_ledger_seeding_started_at IS NOT NULL
                        AND NOT (
                            d.legacy_ledger_seeding_enabled = 0
                            AND i.legacy_ledger_seeding_enabled = 1
                            AND i.legacy_ledger_seeding_started_at = i.updated_at
                        )
                    )
                    OR (
                        (
                            (d.legacy_ledger_continuous_since IS NULL
                                AND i.legacy_ledger_continuous_since IS NOT NULL)
                            OR i.legacy_ledger_continuous_since
                                > d.legacy_ledger_continuous_since
                        )
                        AND (
                            i.legacy_ledger_seeding_enabled <> 1
                            OR i.legacy_ledger_continuous_since <> i.updated_at
                        )
                    )
                    OR (
                        (
                            (d.legacy_ledger_qualified_at IS NULL
                                AND i.legacy_ledger_qualified_at IS NOT NULL)
                            OR i.legacy_ledger_qualified_at > d.legacy_ledger_qualified_at
                        )
                        AND (
                            d.incident_state <> ''normal''
                            OR i.incident_state <> ''normal''
                            OR i.legacy_ledger_qualified_at <> i.updated_at
                            OR d.legacy_ledger_continuous_since IS NULL
                            OR i.legacy_ledger_continuous_since
                                <> d.legacy_ledger_continuous_since
                            OR d.legacy_ledger_heartbeat_owner_id IS NULL
                            OR d.legacy_ledger_heartbeat_at IS NULL
                            OR d.legacy_ledger_heartbeat_at
                                < d.legacy_ledger_continuous_since
                            OR d.legacy_ledger_lease_expires_at IS NULL
                            OR d.legacy_ledger_lease_expires_at <= i.updated_at
                            OR i.legacy_ledger_seeding_enabled <> 1
                            OR i.updated_at < DATEADD(
                                hour,
                                4,
                                d.legacy_ledger_continuous_since
                            )
                        )
                    )
                    OR (
                        d.legacy_compatibility_enforcement_enabled = 0
                        AND i.legacy_compatibility_enforcement_enabled = 1
                        AND (
                            d.incident_state <> ''normal''
                            OR i.incident_state <> ''normal''
                            OR i.legacy_compatibility_enforced_at <> i.updated_at
                            OR d.legacy_ledger_continuous_since IS NULL
                            OR i.legacy_ledger_continuous_since
                                <> d.legacy_ledger_continuous_since
                            OR d.legacy_ledger_heartbeat_owner_id IS NULL
                            OR d.legacy_ledger_heartbeat_at IS NULL
                            OR d.legacy_ledger_heartbeat_at
                                < d.legacy_ledger_continuous_since
                            OR d.legacy_ledger_lease_expires_at IS NULL
                            OR d.legacy_ledger_lease_expires_at <= i.updated_at
                            OR i.legacy_ledger_seeding_enabled <> 1
                            OR i.updated_at < DATEADD(
                                hour,
                                4,
                                d.legacy_ledger_continuous_since
                            )
                        )
                    )
                    OR (
                        d.legacy_handle_issuance_enabled = 1
                        AND i.legacy_handle_issuance_enabled = 0
                        AND (
                            i.legacy_handle_issuance_stopped_at IS NULL
                            OR i.legacy_handle_issuance_stopped_at <> i.updated_at
                        )
                    )
                    OR (
                        d.legacy_handle_issuance_stopped_at IS NULL
                        AND i.legacy_handle_issuance_stopped_at IS NOT NULL
                        AND NOT (
                            d.legacy_handle_issuance_enabled = 1
                            AND i.legacy_handle_issuance_enabled = 0
                        )
                    )
                    OR (
                        d.legacy_handle_acceptance_enabled = 1
                        AND i.legacy_handle_acceptance_enabled = 0
                        AND (
                            i.legacy_handle_acceptance_disabled_at IS NULL
                            OR i.legacy_handle_acceptance_disabled_at <> i.updated_at
                            OR (
                                NOT (
                                    (
                                        d.legacy_verifier_key_incident_at IS NULL
                                        AND i.legacy_verifier_key_incident_at IS NOT NULL
                                    )
                                    OR (
                                        d.legacy_verifier_key_incident_at IS NOT NULL
                                        AND i.legacy_verifier_key_incident_at
                                            > d.legacy_verifier_key_incident_at
                                    )
                                )
                                AND (
                                    d.legacy_handle_issuance_enabled <> 0
                                    OR d.legacy_handle_issuance_stopped_at IS NULL
                                    OR i.updated_at < DATEADD(
                                        hour,
                                        4,
                                        d.legacy_handle_issuance_stopped_at
                                    )
                                )
                            )
                        )
                    )
                    OR (
                        d.legacy_handle_acceptance_disabled_at IS NULL
                        AND i.legacy_handle_acceptance_disabled_at IS NOT NULL
                        AND NOT (
                            d.legacy_handle_acceptance_enabled = 1
                            AND i.legacy_handle_acceptance_enabled = 0
                        )
                    )
                    OR (
                        d.legacy_ledger_seeding_started_at IS NOT NULL
                        AND (
                            i.legacy_ledger_seeding_started_at IS NULL
                            OR i.legacy_ledger_seeding_started_at <> d.legacy_ledger_seeding_started_at
                        )
                    )
                    OR (
                        d.legacy_ledger_continuous_since IS NOT NULL
                        AND (
                            i.legacy_ledger_continuous_since IS NULL
                            OR i.legacy_ledger_continuous_since < d.legacy_ledger_continuous_since
                        )
                    )
                    OR (
                        d.legacy_ledger_qualified_at IS NOT NULL
                        AND (
                            i.legacy_ledger_qualified_at IS NULL
                            OR i.legacy_ledger_qualified_at < d.legacy_ledger_qualified_at
                        )
                    )
                    OR (
                        d.legacy_compatibility_enforced_at IS NOT NULL
                        AND (
                            i.legacy_compatibility_enforced_at IS NULL
                            OR i.legacy_compatibility_enforced_at
                                <> d.legacy_compatibility_enforced_at
                        )
                    )
                    OR (
                        d.target_session_issuance_started_at IS NOT NULL
                        AND (
                            i.target_session_issuance_started_at IS NULL
                            OR i.target_session_issuance_started_at <> d.target_session_issuance_started_at
                        )
                    )
                    OR (
                        d.subject_target_adoption_started_at IS NOT NULL
                        AND (
                            i.subject_target_adoption_started_at IS NULL
                            OR i.subject_target_adoption_started_at <> d.subject_target_adoption_started_at
                        )
                    )
                    OR (
                        d.target_acceptance_started_at IS NOT NULL
                        AND (
                            i.target_acceptance_started_at IS NULL
                            OR i.target_acceptance_started_at <> d.target_acceptance_started_at
                        )
                    )
                    OR (
                        d.legacy_handle_issuance_stopped_at IS NOT NULL
                        AND (
                            i.legacy_handle_issuance_stopped_at IS NULL
                            OR i.legacy_handle_issuance_stopped_at <> d.legacy_handle_issuance_stopped_at
                        )
                    )
                    OR (
                        d.legacy_handle_acceptance_disabled_at IS NOT NULL
                        AND (
                            i.legacy_handle_acceptance_disabled_at IS NULL
                            OR i.legacy_handle_acceptance_disabled_at <> d.legacy_handle_acceptance_disabled_at
                        )
                    )
                    OR (
                        d.legacy_hard_sunset_at IS NOT NULL
                        AND (
                            i.legacy_hard_sunset_at IS NULL
                            OR i.legacy_hard_sunset_at <> d.legacy_hard_sunset_at
                        )
                    )
                    OR (
                        d.target_verifier_key_incident_at IS NOT NULL
                        AND (
                            i.target_verifier_key_incident_at IS NULL
                            OR i.target_verifier_key_incident_at < d.target_verifier_key_incident_at
                        )
                    )
                    OR (
                        d.legacy_verifier_key_incident_at IS NOT NULL
                        AND (
                            i.legacy_verifier_key_incident_at IS NULL
                            OR i.legacy_verifier_key_incident_at < d.legacy_verifier_key_incident_at
                        )
                    )
                    OR (
                        (
                            (d.target_verifier_key_incident_at IS NULL
                                AND i.target_verifier_key_incident_at IS NOT NULL)
                            OR i.target_verifier_key_incident_at > d.target_verifier_key_incident_at
                        )
                        AND (
                            i.target_verifier_key_incident_at > SYSUTCDATETIME()
                            OR i.incident_state = ''normal''
                            OR i.incident_recorded_at IS NULL
                            OR i.incident_recorded_at < i.target_verifier_key_incident_at
                            OR i.incident_recorded_at > SYSUTCDATETIME()
                        )
                    )
                    OR (
                        (
                            (d.legacy_verifier_key_incident_at IS NULL
                                AND i.legacy_verifier_key_incident_at IS NOT NULL)
                            OR i.legacy_verifier_key_incident_at > d.legacy_verifier_key_incident_at
                        )
                        AND (
                            i.legacy_verifier_key_incident_at > SYSUTCDATETIME()
                            OR i.incident_state = ''normal''
                            OR i.incident_recorded_at IS NULL
                            OR i.incident_recorded_at < i.legacy_verifier_key_incident_at
                            OR i.incident_recorded_at > SYSUTCDATETIME()
                            OR i.legacy_handle_issuance_enabled <> 0
                            OR i.legacy_handle_acceptance_enabled <> 0
                            OR i.legacy_handle_issuance_stopped_at IS NULL
                            OR i.legacy_handle_acceptance_disabled_at IS NULL
                            OR i.legacy_handle_issuance_stopped_at
                                > i.legacy_handle_acceptance_disabled_at
                            OR i.legacy_handle_issuance_stopped_at > SYSUTCDATETIME()
                            OR i.legacy_handle_acceptance_disabled_at > SYSUTCDATETIME()
                        )
                    )
                    OR (
                        d.incident_recorded_at IS NOT NULL
                        AND (
                            i.incident_recorded_at IS NULL
                            OR i.incident_recorded_at < d.incident_recorded_at
                        )
                    )
                    OR (
                        d.incident_state <> ''normal''
                        AND i.incident_state = ''normal''
                        AND (
                            d.incident_state <> ''recovering''
                            OR i.authority_generation <> d.authority_generation
                            OR i.global_session_epoch <> d.global_session_epoch
                            OR EXISTS (
                                SELECT 1
                                FROM dbo.learning_session_flow AS f
                                INNER JOIN dbo.learning_session AS l
                                    ON l.session_id = f.current_session_id
                                    AND l.subject_id = f.subject_id
                                WHERE f.challenge_state IN (
                                    ''creating'',
                                    ''active'',
                                    ''reconciliation-required''
                                )
                                    AND l.phase IN (
                                        ''credential-verified'',
                                        ''registration-pending'',
                                        ''face-pending'',
                                        ''authenticated''
                                    )
                            )
                            OR (
                                (
                                    d.legacy_ledger_seeding_started_at IS NOT NULL
                                    OR d.legacy_ledger_continuous_since IS NOT NULL
                                    OR d.legacy_ledger_qualified_at IS NOT NULL
                                    OR d.legacy_compatibility_enforced_at IS NOT NULL
                                    OR d.target_acceptance_started_at IS NOT NULL
                                    OR d.legacy_verifier_key_incident_at IS NOT NULL
                                    OR i.legacy_compatibility_enforcement_enabled = 1
                                )
                                AND (
                                    i.legacy_handle_issuance_enabled <> 0
                                    OR i.legacy_handle_acceptance_enabled <> 0
                                )
                            )
                        )
                    )
                    OR (
                        d.incident_state <> ''recovering''
                        AND i.incident_state = ''recovering''
                        AND NOT (
                            d.incident_state = ''suspended''
                            AND i.authority_generation = d.authority_generation + 1
                            AND i.global_session_epoch > d.global_session_epoch
                            AND i.target_session_issuance_enabled = 0
                            AND i.subject_target_adoption_enabled = 0
                            AND (
                                i.target_verifier_key_id <> d.target_verifier_key_id
                                OR i.target_verifier_key_commitment
                                    <> d.target_verifier_key_commitment
                                OR i.legacy_compatibility_key_id
                                    <> d.legacy_compatibility_key_id
                                OR i.legacy_compatibility_key_commitment
                                    <> d.legacy_compatibility_key_commitment
                                OR i.credential_fingerprint_key_id
                                    <> d.credential_fingerprint_key_id
                                OR i.credential_fingerprint_key_commitment
                                    <> d.credential_fingerprint_key_commitment
                                OR i.face_challenge_key_id <> d.face_challenge_key_id
                                OR i.face_challenge_key_commitment
                                    <> d.face_challenge_key_commitment
                                OR i.legacy_signing_key_id <> d.legacy_signing_key_id
                                OR i.legacy_signing_key_commitment
                                    <> d.legacy_signing_key_commitment
                            )
                        )
                    )
                    OR (
                        i.authority_generation > d.authority_generation
                        AND NOT (
                            d.incident_state = ''suspended''
                            AND i.incident_state = ''recovering''
                            AND i.authority_generation = d.authority_generation + 1
                            AND i.global_session_epoch > d.global_session_epoch
                            AND i.target_session_issuance_enabled = 0
                            AND i.subject_target_adoption_enabled = 0
                        )
                    )
                    OR i.updated_at < d.updated_at
            )
                THROW 51010, ''Authority epochs, cutovers, sunsets, and incident evidence only move forward.'', 1;
        END;
    ';

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
        ROLLBACK TRANSACTION;
    THROW;
END CATCH;
