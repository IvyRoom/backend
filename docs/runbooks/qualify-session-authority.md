# Qualify and activate the session authority

## Current boundary

The backend implementation is intentionally dormant. No Azure SQL database,
secret, application setting, production ledger, or target session was created
or changed by the implementation task. The existing F1 App Service endpoint,
default TLS, and hostname remain unchanged. The durable-store latch, all six
runtime rollout controls, and every database activation control default off.
The compatibility issuance and acceptance booleans deliberately start open so applying the inert schema
cannot itself stop current clients; they gain authority only after the latch is
separately approved and enabled. The migration is forward-only and is never
applied by application startup or deployment.

Production activation is blocked until Lucas explicitly approves a concrete
infrastructure and rollout proposal and every inert qualification below passes.
Process memory is test-only and is never a production fallback. The current
four-hour `IndexVerificado` behavior remains authoritative while blocked.

## Approval packet required before any external mutation

Before provisioning or changing anything, prepare one reviewable packet that
names all of the following exactly:

- Azure subscription and resource group;
- Azure SQL logical server and database resource names;
- Azure region and confirmation that it matches the intended data boundary;
- Azure SQL Database Basic SKU, storage/capacity limits, connection limits, and
  current monthly cost from the official Azure calculator;
- private/public network exposure, firewall rules, TLS requirements, database
  authentication, and the least-privilege migration/runtime principals;
- the existing `plataforma-backend-v3.azurewebsites.net` endpoint and F1 App
  Service plan, including its zero added monthly plan cost, no SLA, daily CPU
  quota, memory/data limits, and the capacity threshold that would require a
  separately approved plan change;
- supported Edge versions/profiles and the qualification matrix for host-only
  partitioned cookies with ordinary third-party cookies blocked;
- secret owner, App Service setting owner, rotation procedure, and the six
  distinct authority key IDs without including any key or connection value;
- exact migration, qualification, ledger-seeding, adoption, monitoring, and
  rollback steps; and
- resources/settings to remove or restore during rollback, with the point after
  which per-subject adoption, stop-issuance, acceptance disablement, and sunset
  are irreversible.

Obtain Lucas's explicit approval of that exact packet before using Azure, DNS,
TLS, App Service, or secret-management write operations. Re-price or re-approve
if the resource, region, SKU, cost, security boundary, or rollback plan changes.

## Safe configuration model

The composition latch is
`SESSION_AUTHORITY_DURABLE_STORE_REQUIRED`. Every rollout control requires it.
While false, startup constructs no SQL authority and preserves the exact current
14-route composition. Once production ledger seeding starts, configuration
ownership must treat the latch, SQL setting, and required keys as one-way: they
may be rotated through the incident procedure but never removed to select the
unchecked legacy path.

Rollout controls are separate environment settings:

1. `SESSION_AUTHORITY_TARGET_ROUTES_ENABLED`
2. `SESSION_AUTHORITY_TARGET_ISSUANCE_ENABLED`
3. `SESSION_AUTHORITY_LEGACY_SEEDING_ENABLED`
4. `SESSION_AUTHORITY_LEGACY_ENFORCEMENT_ENABLED`
5. `SESSION_AUTHORITY_SUBJECT_ADOPTION_ENABLED`
6. `SESSION_AUTHORITY_PROTECTED_ROUTES_ENABLED`

`SESSION_AUTHORITY_PARTITIONED_COOKIE_TOPOLOGY_QUALIFIED` is independent
evidence for the topology gate. The retired first-party latch name is ignored.
A request header, cookie, URL value, or other browser-controlled input cannot
enable a disabled control. Runtime activation also requires its corresponding
central database state; disagreement fails closed.

With the latch enabled, central SQL admission owns legacy issuance and
authorization even when a rollout permission is false. Before seeding, the
serialized admission preserves the existing signed-handle behavior. During
seeding, both runtime and database seeding gates must agree. After qualified
enforcement, the irreversible database state wins over a rolled-back runtime
flag and missing bindings never fall through to row-derived authority. Target
issuance/adoption cannot precede qualified enforcement.

The store uses `SESSION_AUTHORITY_SQL_CONNECTION_STRING` plus bounded connection,
request, pool-size, and idle-timeout settings documented in the README. The SQL
transport always enables encryption and rejects untrusted server certificates.
Each deployment also requires `SESSION_AUTHORITY_EXPECTED_GENERATION`; every
authority transaction fails closed unless that instance value equals the
singleton database generation. Key recovery is deliberately two-step. While
authority is suspended, an old-generation recovery transaction installs only
permitted replacement bindings, advances both authority generation and global
epoch, and leaves the control record in `recovering`. A separate transaction
running with the complete replacement configuration may resume `normal`
authority. Old instances and partial/mixed keysets therefore remain `503` and
cannot mint records under retired material.

Six purpose-specific key pairs are required for target verification, legacy
compatibility, login lookup, credential fingerprinting, account-mapping
encryption, and Face-reference encryption. They must also differ from the
legacy signed-handle key. The latter keeps its existing
`PLATFORM_ROW_AUTHORIZATION_KEY_BASE64` material and adds the non-secret
`SESSION_AUTHORITY_LEGACY_SIGNING_KEY_ID` descriptor whenever the durable latch
is enabled. Its ID and material must differ from every session-purpose key.
Never place key material in commands, output, fixtures, snapshots, logs,
documentation, or pull-request text.

The migration leaves the complete central key-binding set uninitialized: four
rotatable purpose IDs/leaf commitments, independent domain-framed login and
account-mapping bindings, the canonical aggregate over all six purposes, and
the independent legacy-signing binding. The explicit store initialization primitive may bind the
configured set only while every rollout/evidence field is dormant and all four
non-control authority-data tables are empty; an exact retry is idempotent. It is not called
by server startup or a request. Once bound, every authority transaction rejects
an uninitialized, internally inconsistent, or mismatched set before data
access. No production initialization is authorized by this runbook.
Login-lookup and account-mapping key rotation are blocked in this milestone:
keep authority suspended until a future purpose-built rekey migration preserves
every stable subject and exact mapping. Never swap either key through ordinary
application configuration. A permitted verifier, fingerprint, Face-reference,
or legacy-signing change still requires the two-step recovery above; a
legacy-signing change permanently disables legacy issuance and acceptance, and
every key recovery quarantines unresolved Face flows and revokes their active
provisional sessions before resume.

## Inert durable-store qualification

Use only invented UUIDs, generated verifier bytes, fake UTC clocks, synthetic
accounts, and a database that contains no production learner or Face data.
Apply `migrations/001-session-authority.sql` with the migration principal, then
prove and record:

- exactly the five selected tables, constraints, indexes, singleton control
  row, immutable-evidence triggers, dormant defaults, and empty-only
  complete key-binding initialization;
- exact rotatable-purpose, independent identity-binding, aggregate, and
  legacy-signing ID/commitment agreement
  across instances, including same-ID/different-material and
  different-ID/same-material rejection before any authority read or write;
- concurrent create-or-load subject uniqueness and verifier uniqueness;
- multi-instance same-predecessor rotation, elevation, logout, revoke-all,
  administrator revocation, and stale-loser behavior under serializable
  transactions;
- SQL `SYSUTCDATETIME()` expiry and five-minute eligibility semantics;
- pool maximum, Basic-tier connection/capacity limits, request/connection
  timeouts, pool faults, store outage, rollback failure, and uncertain commit;
- backup and restore into an isolated target, followed by the suspended
  old-generation-to-recovering-new-generation transition, global-epoch advance,
  permitted key replacement or legacy retirement, and separately fenced
  all-instance resume;
- migration-principal versus runtime-principal permissions, including proof
  that runtime cannot alter schema or read unrelated application data; and
- secret/application-setting ownership and permitted verifier-key rotation
  without revealing values; record login-lookup and account-mapping rekeying as
  blocked rather than substituting an online configuration swap.

Any failed or interrupted proof leaves all activation controls off. Do not
substitute a file, cache, workbook, signed token, or process-memory store.

## Partitioned-cookie topology qualification

Before target issuance, verify the existing default-TLS App Service endpoint
`https://plataforma-backend-v3.azurewebsites.net` in every supported Edge
profile: Stable, Extended Stable, InPrivate, and the supported tracking-
prevention configuration with ordinary third-party cookies blocked. With only
synthetic accounts and denied production side effects, prove that the host-only
`Secure; HttpOnly; SameSite=None; Partitioned` cookie survives the complete
credentialed flow when top-level `https://machadogestao.com` is the partition
key, is unavailable under an unrelated synthetic top-level site, rotates on
fresh login, and remains unreadable to JavaScript. Also verify exact
credentialed CORS, `Origin`, custom-header/CSRF, no-store, validator suppression,
`Vary`, referrer, preflight, and cookie-mutation boundaries. Any failed browser
profile leaves target issuance and protected-route adoption off.

## Ordered rollout and rollback boundaries

1. Deploy the code with the durable-store latch and every rollout control off;
   confirm the exact 14-route legacy inventory and unchanged frontend
   fingerprint. In the separately approved inert qualification boundary,
   initialize the complete key-binding set before any authority record exists;
   production initialization remains blocked without that approval.
2. After explicit approval, qualify the SQL schema and partitioned-cookie
   topology with inert data. Then enable the durable-store latch with every rollout permission still
   off and verify central admission plus the same 14-route client behavior.
   Enabling dormant routes alone must not permit issuance.
3. If separately approved, start legacy-ledger seeding while all client-visible
   behavior remains unchanged. Any failed write or continuity gap advances the
   continuous-since instant and restarts the full four-hour horizon.
4. Only after a continuous four-hour horizon may central legacy enforcement be
   enabled. Missing bindings then fail `401`; integrity/store failures fail
   `503` and never fall back to the signed row.
5. Target issuance, protected-route adoption, per-subject adoption,
   partitioned-cookie topology proof, and the frontend release must be
   coordinated. The implementation task
   does not authorize this step.
6. Global stop-issuance begins only with at least four hours remaining before
   the fixed seven-day sunset. Existing handles age for a full four hours before
   central acceptance disablement. Neither boundary can be reversed.

Before ledger seeding, rollback may remove a never-authoritative latch under the
approved plan. Once production seeding begins, rollback must retain the latch,
store, keys, and stronger central state; it may disable only reversible rollout
permissions. After a subject adopts target authority, its legacy cutoff is
irreversible. After global stop-issuance or acceptance disablement, legacy
authority cannot be restored. Store/key incidents fail authority closed,
retire affected keys outside restored SQL, advance the required epochs, and
resume only through the separately fenced `recovering`-to-`normal` transition
after every instance acknowledges the recovered generation. Neither
`suspended` nor any same-generation state may return directly to `normal`.
Suspension/recovery invalidates the seeding lease and continuous-since evidence;
the first healthy normal-state heartbeat begins a fresh four-hour horizon.
