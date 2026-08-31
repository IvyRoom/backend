'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const MIGRATION_PATH = path.join(
    __dirname,
    '..',
    'migrations',
    '001-session-authority.sql',
);
const SQL = fs.readFileSync(MIGRATION_PATH, 'utf8');
const COMPACT_SQL = SQL.replace(/\s+/g, ' ').trim();

function tableBody(tableName) {
    const marker = `CREATE TABLE dbo.${tableName} (`;
    const start = SQL.indexOf(marker);
    assert.notEqual(start, -1, `missing ${tableName}`);

    let depth = 1;
    let inString = false;
    for (let index = start + marker.length; index < SQL.length; index += 1) {
        const character = SQL[index];
        const nextCharacter = SQL[index + 1];

        if (character === "'") {
            if (inString && nextCharacter === "'") {
                index += 1;
                continue;
            }
            inString = !inString;
            continue;
        }
        if (inString) continue;
        if (character === '(') depth += 1;
        if (character === ')') depth -= 1;
        if (depth === 0) return SQL.slice(start + marker.length, index);
    }

    assert.fail(`unterminated ${tableName} definition`);
}

test('defines exactly the five approved authority records in one forward-only migration', () => {
    const tableNames = [...SQL.matchAll(/CREATE TABLE dbo\.([a-z_]+)\s*\(/g)]
        .map((match) => match[1]);

    assert.deepEqual(tableNames, [
        'learning_subject',
        'learning_session',
        'learning_session_flow',
        'legacy_session_compatibility',
        'session_authority_control',
    ]);
    assert.match(SQL, /SET XACT_ABORT ON;/);
    assert.match(SQL, /BEGIN TRANSACTION;/);
    assert.match(SQL, /COMMIT TRANSACTION;/);
    assert.match(SQL, /ROLLBACK TRANSACTION;/);
    assert.doesNotMatch(SQL, /^\s*GO\s*$/im);
    assert.doesNotMatch(SQL, /\bDROP\s+(?:TABLE|COLUMN|CONSTRAINT|INDEX)\b/i);
    assert.doesNotMatch(SQL, /\bTRUNCATE\s+TABLE\b/i);
    assert.doesNotMatch(SQL, /\bDELETE\s+FROM\b/i);
    assert.doesNotMatch(SQL, /\bON\s+DELETE\s+CASCADE\b/i);
});

test('models stable subjects without plaintext account or credential authority', () => {
    const subject = tableBody('learning_subject');

    assert.match(subject, /subject_id uniqueidentifier NOT NULL/);
    assert.match(subject, /login_lookup_token binary\(32\) NOT NULL/);
    assert.match(
        subject,
        /login_lookup_key_id varchar\(128\) COLLATE Latin1_General_100_BIN2 NOT NULL/,
    );
    assert.match(subject, /encrypted_legacy_account_mapping varbinary\(2048\) NOT NULL/);
    assert.match(
        subject,
        /account_mapping_encryption_key_id varchar\(128\) COLLATE Latin1_General_100_BIN2 NOT NULL/,
    );
    assert.match(subject, /credential_fingerprint binary\(32\) NOT NULL/);
    assert.match(
        subject,
        /credential_fingerprint_key_id varchar\(128\) COLLATE Latin1_General_100_BIN2 NOT NULL/,
    );
    assert.match(
        subject,
        /UNIQUE \(login_lookup_key_id, login_lookup_token\)/,
    );
    assert.match(subject, /subject_session_epoch bigint NOT NULL/);
    assert.match(subject, /legacy_authority_disabled_at datetime2\(7\) NULL/);
    assert.match(subject, /eligibility_state IN \('eligible', 'ineligible', 'unknown'\)/);
    assert.match(subject, /entitlement_expires_at datetime2\(7\) NULL/);
    assert.match(subject, /eligibility_observed_at datetime2\(7\) NOT NULL/);
    assert.match(subject, /eligibility_revalidate_at datetime2\(7\) NOT NULL/);
    assert.match(
        subject,
        /eligibility_revalidate_at <= DATEADD\(minute, 5, eligibility_observed_at\)/,
    );
    assert.doesNotMatch(subject, /^\s*(?:login|credential|password|account_name)\s+/im);
});

test('stores only fixed-size target verifiers and constrains every persisted phase', () => {
    const session = tableBody('learning_session');

    assert.match(session, /identifier_verifier binary\(32\) NOT NULL/);
    assert.match(
        session,
        /verifier_key_id varchar\(128\) COLLATE Latin1_General_100_BIN2 NOT NULL/,
    );
    assert.match(session, /UNIQUE \(identifier_verifier\)/);
    assert.match(session, /FOREIGN KEY \(subject_id\)\s+REFERENCES dbo\.learning_subject \(subject_id\)/);
    assert.match(session, /original_issued_at datetime2\(7\) NOT NULL/);
    assert.match(session, /phase_started_at datetime2\(7\) NOT NULL/);
    assert.match(session, /absolute_expires_at datetime2\(7\) NOT NULL/);
    assert.match(session, /face_auth_required bit NOT NULL/);
    assert.match(session, /registration_required bit NOT NULL/);
    assert.match(session, /registration_required = 0 OR face_auth_required = 1/);
    assert.match(session, /subject_epoch_snapshot bigint NOT NULL/);
    assert.match(session, /credential_version_snapshot bigint NOT NULL/);
    assert.match(session, /global_epoch_snapshot bigint NOT NULL/);
    assert.match(session, /authority_generation_snapshot bigint NOT NULL/);
    assert.match(session, /DATEADD\(minute, 20, original_issued_at\)/);
    assert.match(session, /DATEADD\(hour, 4, original_issued_at\)/);
    assert.match(
        session,
        /phase IN \('credential-verified', 'registration-pending', 'face-pending'\)[\s\S]*absolute_expires_at = DATEADD\(minute, 20, original_issued_at\)/,
    );
    assert.match(
        session,
        /phase = 'authenticated'[\s\S]*absolute_expires_at = DATEADD\(hour, 4, original_issued_at\)/,
    );
    assert.match(
        session,
        /phase IN \('expired', 'revoked', 'rotated-out'\)[\s\S]*absolute_expires_at IN/,
    );

    for (const phase of [
        'credential-verified',
        'registration-pending',
        'face-pending',
        'authenticated',
        'expired',
        'revoked',
        'rotated-out',
    ]) {
        assert.match(session, new RegExp(`'${phase}'`));
    }
    assert.doesNotMatch(session, /'anonymous'/);
    assert.match(session, /phase = 'revoked'[\s\S]*revoked_at IS NOT NULL/);
    assert.match(session, /phase = 'rotated-out'[\s\S]*replacement_session_id IS NOT NULL/);
});

test('binds one encrypted Face flow and one immutable legacy verifier to a subject', () => {
    const flow = tableBody('learning_session_flow');
    const legacy = tableBody('legacy_session_compatibility');

    assert.match(flow, /UNIQUE \(current_session_id\)/);
    assert.match(
        flow,
        /FOREIGN KEY \(current_session_id, subject_id\)\s+REFERENCES dbo\.learning_session \(session_id, subject_id\)/,
    );
    assert.match(flow, /encrypted_provider_challenge_reference varbinary\(2048\) NULL/);
    assert.match(
        flow,
        /provider_reference_encryption_key_id varchar\(128\) COLLATE Latin1_General_100_BIN2 NULL/,
    );
    assert.match(flow, /challenge_state IN \(/);
    assert.match(flow, /'active'/);
    assert.match(flow, /'consumed'/);
    assert.match(flow, /'failed'/);
    assert.match(flow, /'reconciliation-required'/);
    assert.doesNotMatch(flow, /^\s*(?:client_verdict|provider_session_id)\s+/im);

    assert.match(legacy, /legacy_handle_verifier binary\(32\) NOT NULL/);
    assert.match(
        legacy,
        /verifier_key_id varchar\(128\) COLLATE Latin1_General_100_BIN2 NOT NULL/,
    );
    assert.match(legacy, /UNIQUE \(legacy_handle_verifier\)/);
    assert.match(legacy, /FOREIGN KEY \(subject_id\)\s+REFERENCES dbo\.learning_subject \(subject_id\)/);
    assert.match(legacy, /original_issued_at datetime2\(7\) NOT NULL/);
    assert.match(legacy, /original_expires_at datetime2\(7\) NOT NULL/);
    assert.match(legacy, /original_expires_at = DATEADD\(hour, 4, original_issued_at\)/);
    assert.match(legacy, /compatibility_state IN \('active', 'revoked', 'incident'\)/);
    assert.doesNotMatch(legacy, /^\s*(?:legacy_handle|indexverificado|workbook_row_hint)\s+/im);
});

test('installs dormant target controls while preserving current legacy behavior', () => {
    const control = tableBody('session_authority_control');

    for (const dormantControl of [
        'target_routes_enabled',
        'target_session_issuance_enabled',
        'legacy_ledger_seeding_enabled',
        'legacy_compatibility_enforcement_enabled',
        'subject_target_adoption_enabled',
    ]) {
        assert.match(
            control,
            new RegExp(
                `${dormantControl} bit NOT NULL\\s+`
                    + 'CONSTRAINT [A-Za-z0-9_]+ DEFAULT \\(0\\)',
            ),
        );
    }
    for (const enabledLegacyControl of [
        'legacy_handle_issuance_enabled',
        'legacy_handle_acceptance_enabled',
    ]) {
        assert.match(
            control,
            new RegExp(
                `${enabledLegacyControl} bit NOT NULL\\s+`
                    + 'CONSTRAINT [A-Za-z0-9_]+ DEFAULT \\(1\\)',
            ),
        );
    }

    assert.match(control, /control_version bigint NOT NULL/);
    assert.match(control, /authority_generation bigint NOT NULL/);
    assert.match(control, /global_session_epoch bigint NOT NULL/);
    assert.match(
        control,
        /login_lookup_key_id varchar\(128\) COLLATE Latin1_General_100_BIN2 NULL/,
    );
    assert.match(control, /login_lookup_key_commitment binary\(32\) NULL/);
    assert.match(control, /legacy_hard_sunset_at datetime2\(7\) NULL/);
    assert.match(control, /legacy_hard_sunset_at = DATEADD\(day, 7, target_acceptance_started_at\)/);
    assert.match(
        control,
        /legacy_handle_acceptance_disabled_at\s+>= DATEADD\(hour, 4, legacy_handle_issuance_stopped_at\)/,
    );
    assert.match(
        control,
        /legacy_handle_issuance_stopped_at\s+<= DATEADD\(hour, -4, legacy_hard_sunset_at\)/,
    );
    assert.match(
        control,
        /CK_session_authority_control_final_aging[\s\S]*legacy_verifier_key_incident_at IS NOT NULL[\s\S]*DATEADD\(hour, 4, legacy_handle_issuance_stopped_at\)/,
    );
    assert.match(
        control,
        /CK_session_authority_control_target_after_enforcement[\s\S]*legacy_compatibility_enforcement_enabled = 1[\s\S]*target_session_issuance_enabled = 0[\s\S]*subject_target_adoption_enabled = 0[\s\S]*target_acceptance_started_at IS NULL/,
    );
    assert.match(
        control,
        /CK_session_authority_control_target_window[\s\S]*target_session_issuance_enabled = subject_target_adoption_enabled[\s\S]*target_acceptance_started_at IS NULL[\s\S]*target_session_issuance_enabled = 0[\s\S]*target_acceptance_started_at IS NOT NULL[\s\S]*legacy_hard_sunset_at IS NOT NULL/,
    );
    assert.match(
        control,
        /CK_session_authority_control_seeding_during_legacy_issuance[\s\S]*legacy_compatibility_enforcement_enabled = 0[\s\S]*legacy_handle_issuance_enabled = 0[\s\S]*legacy_ledger_seeding_enabled = 1/,
    );
    assert.match(control, /target_verifier_key_incident_at datetime2\(7\) NULL/);
    assert.match(control, /legacy_verifier_key_incident_at datetime2\(7\) NULL/);
    assert.match(
        control,
        /CK_session_authority_control_incident_code[\s\S]*DATALENGTH\(incident_code\) BETWEEN 1 AND 128[\s\S]*Latin1_General_100_BIN2 LIKE '\[a-z0-9\]%'[\s\S]*NOT LIKE '%\[\^a-z0-9-\]%'/,
    );
    assert.match(
        COMPACT_SQL,
        /INSERT INTO dbo\.session_authority_control \([\s\S]*?\) VALUES \(1, 1, 1, 1, 0, 0, 0, 0, 0, 1, 1, 'normal'\);/,
    );
});

test('uses SQL UTC instants and database guards for irreversible evidence', () => {
    assert.match(SQL, /datetime2\(7\)/);
    assert.match(SQL, /DEFAULT SYSUTCDATETIME\(\)/);
    assert.doesNotMatch(SQL, /\bGETDATE\s*\(/i);
    assert.doesNotMatch(SQL, /\bSYSDATETIME\s*\(/i);
    assert.doesNotMatch(SQL, /\bCURRENT_TIMESTAMP\b/i);

    for (const triggerName of [
        'trg_learning_subject_authority_evidence',
        'trg_learning_session_authority_evidence',
        'trg_learning_session_flow_binding',
        'trg_legacy_session_compatibility_binding',
        'trg_session_authority_control_irreversible',
    ]) {
        assert.match(SQL, new RegExp(`CREATE TRIGGER dbo\\.${triggerName}`));
    }
    assert.match(SQL, /legacy cutoffs only move forward/);
    assert.match(SQL, /Terminal session and rotation evidence cannot be changed/);
    assert.match(SQL, /OR UPDATE\(registration_required\)/);
    assert.match(SQL, /Legacy verifier bindings and issue metadata are immutable/);
    assert.match(SQL, /legacy_handle_issuance_enabled = 0[\s\S]*legacy_handle_issuance_enabled = 1/);
    assert.match(SQL, /legacy_handle_acceptance_enabled = 0[\s\S]*legacy_handle_acceptance_enabled = 1/);
    assert.match(SQL, /legacy_hard_sunset_at IS NOT NULL[\s\S]*legacy_hard_sunset_at IS NULL/);
    assert.match(SQL, /i\.control_version <> d\.control_version \+ 1/);
    assert.match(
        SQL,
        /d\.incident_state <> ''normal''[\s\S]*i\.incident_state = ''normal''[\s\S]*d\.incident_state <> ''recovering''[\s\S]*i\.authority_generation <> d\.authority_generation[\s\S]*i\.global_session_epoch <> d\.global_session_epoch/,
    );
    assert.match(
        SQL,
        /legacy_ledger_seeding_started_at IS NOT NULL[\s\S]*legacy_handle_issuance_enabled <> 0[\s\S]*legacy_handle_acceptance_enabled <> 0/,
    );
    assert.match(
        SQL,
        /i\.legacy_compatibility_enforcement_enabled = 0[\s\S]*i\.target_session_issuance_enabled = 1[\s\S]*i\.subject_target_adoption_enabled = 1[\s\S]*i\.target_acceptance_started_at IS NOT NULL/,
    );
    assert.match(
        SQL,
        /i\.legacy_compatibility_enforcement_enabled = 1[\s\S]*i\.legacy_handle_issuance_enabled = 1[\s\S]*i\.legacy_ledger_seeding_enabled = 0/,
    );
    assert.match(
        SQL,
        /SYSUTCDATETIME\(\) < i\.target_acceptance_started_at/,
    );
    assert.match(
        COMPACT_SQL,
        /d\.target_session_issuance_enabled = 0 AND i\.target_session_issuance_enabled = 1 AND d\.target_session_issuance_started_at IS NULL AND i\.target_session_issuance_started_at <> i\.updated_at/,
    );
    assert.match(
        COMPACT_SQL,
        /d\.subject_target_adoption_enabled = 0 AND i\.subject_target_adoption_enabled = 1 AND d\.subject_target_adoption_started_at IS NULL AND i\.subject_target_adoption_started_at <> i\.updated_at/,
    );
    assert.match(
        COMPACT_SQL,
        /d\.target_acceptance_started_at IS NULL AND i\.target_acceptance_started_at IS NOT NULL[\s\S]*i\.target_acceptance_started_at <> i\.updated_at[\s\S]*i\.legacy_hard_sunset_at <> DATEADD\(day, 7, i\.updated_at\)/,
    );
    assert.match(
        COMPACT_SQL,
        /d\.legacy_ledger_seeding_enabled = 0 AND i\.legacy_ledger_seeding_enabled = 1[\s\S]*i\.legacy_ledger_continuous_since <> i\.updated_at[\s\S]*i\.legacy_ledger_seeding_started_at <> i\.updated_at/,
    );
    assert.match(
        COMPACT_SQL,
        /i\.legacy_ledger_qualified_at <> i\.updated_at[\s\S]*d\.legacy_ledger_continuous_since IS NULL[\s\S]*i\.updated_at < DATEADD\( hour, 4, d\.legacy_ledger_continuous_since \)/,
    );
    assert.match(
        COMPACT_SQL,
        /d\.legacy_handle_issuance_enabled = 1 AND i\.legacy_handle_issuance_enabled = 0[\s\S]*i\.legacy_handle_issuance_stopped_at <> i\.updated_at/,
    );
    assert.match(
        COMPACT_SQL,
        /d\.legacy_handle_acceptance_enabled = 1 AND i\.legacy_handle_acceptance_enabled = 0[\s\S]*d\.legacy_handle_issuance_stopped_at IS NULL[\s\S]*i\.updated_at < DATEADD\( hour, 4, d\.legacy_handle_issuance_stopped_at \)/,
    );
    assert.doesNotMatch(SQL, /SYSUTCDATETIME\(\) >= i\.legacy_hard_sunset_at/);
    assert.match(
        SQL,
        /d\.target_verifier_key_incident_at IS NULL[\s\S]*i\.target_verifier_key_incident_at IS NOT NULL[\s\S]*i\.incident_state = ''normal''[\s\S]*i\.incident_recorded_at < i\.target_verifier_key_incident_at/,
    );
    assert.match(
        SQL,
        /d\.legacy_verifier_key_incident_at IS NULL[\s\S]*i\.legacy_verifier_key_incident_at IS NOT NULL[\s\S]*i\.legacy_handle_issuance_enabled <> 0[\s\S]*i\.legacy_handle_acceptance_enabled <> 0[\s\S]*i\.legacy_handle_issuance_stopped_at IS NULL[\s\S]*i\.legacy_handle_acceptance_disabled_at IS NULL/,
    );
    assert.match(
        SQL,
        /i\.authority_generation > d\.authority_generation[\s\S]*d\.incident_state = ''suspended''[\s\S]*i\.incident_state = ''recovering''[\s\S]*i\.global_session_epoch > d\.global_session_epoch/,
    );
});

test('contains schema metadata only, with no raw authority or configuration values', () => {
    assert.doesNotMatch(SQL, /^\s*(?:raw_identifier|session_identifier|indexverificado)\s+/im);
    assert.doesNotMatch(SQL, /^\s*(?:connection_string|database_password|secret_key)\s+/im);
    assert.doesNotMatch(SQL, /__Host-machado-session=/);
    assert.doesNotMatch(SQL, /AccountKey=/i);
    assert.doesNotMatch(SQL, /Server=tcp:/i);
});
