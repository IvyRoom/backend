# Session authority target decision

- **Status:** approved target contract; not implemented
- **Decision owner:** Machado backend
- **Decision scope:** Topic 05 · Session authority and logout
- **Source bases:** backend `00bac84c7ef9d9a1aaa014719043451a1362602c`;
  frontend `1afd814adc02fa2da5ab4c55c1eeb6ebb5bb05b7`

This ADR defines the future authority boundary for revocable learning-platform
sessions across the backend and the
[`sistemas` consumer](https://github.com/IvyRoom/sistemas/blob/d7e220ee60bfb5aceb98094f72fb7d8bc5ead727/docs/learning-platform-contracts.md#approved-future-session-authority-consumer-decision).
It is a target decision, not a description of deployed behavior. Publishing it
does not create a session store, set a cookie, change CORS, protect a route,
revoke a handle, alter logout, or remediate any current runtime risk.

The source-observed compatibility contract remains
[`api-contracts.md`](api-contracts.md). Where this ADR says **current**, it
characterizes the pinned source bases above. Where it says **target**, it states
acceptance criteria for later Topic 05 implementation tasks.

## Decision

The backend will be the sole authority for:

- whether a session exists;
- the authenticated subject;
- the current authentication phase and its allowed operations;
- server-computed expiry;
- per-session and subject-wide revocation;
- account eligibility and access-entitlement enforcement; and
- atomic privilege elevation and credential rotation.

The browser will hold a session identifier only in a host-only `Secure`,
`HttpOnly` cookie. Web Storage will remain presentation and navigation state
only. No browser key, workbook row position, client-reported Face result,
client clock, URL value, or request-body identity will grant backend authority.

The selected production design is:

1. a verified first-party API origin at `https://api.machadogestao.com`;
2. an opaque, cryptographically random 256-bit session identifier;
3. a backend-owned durable session record in Azure SQL Database Basic;
4. a stable backend subject identifier independent of workbook row position;
5. a 20-minute non-extending provisional lifetime;
6. a four-hour non-extending authenticated absolute lifetime and no idle
   timeout;
7. identifier rotation for every externally visible phase transition, with the
   original provisional deadline preserved until authenticated promotion;
8. server-bound Face verification as the only Face-enabled promotion path;
9. one cookie session shared by tabs in a browser profile, with concurrent
   sessions on other devices allowed; and
10. a centrally committed revocation decision read on every protected request.

No fallback may replace this design with a third-party cookie on the current
App Service hostname, a bearer token in `sessionStorage` or `localStorage`, a
signed workbook-row token, or process-memory session state.

## Source-observed current contract

The current backend has no server-side session record. After an exact active
credential match, `POST /plataforma_v2/login-FaceID` returns a four-hour signed
authorization handle in the legacy `IndexVerificado` field. The value is
consumer-opaque, signed, self-contained, replayable, and row-bearing; opaque
does not mean random, encrypted, or confidential. Its signature prevents row
tampering, but its authority is not revocable per session and is not rechecked
against current account eligibility.

Exactly these five routes read it from JSON or multipart bodies:

- `POST /plataforma_v2/CadastroFoto_e_FaceID`;
- `POST /plataforma_v2/FaceID`;
- `POST /plataforma_v2/refresh`;
- `POST /plataforma_v2/updates`; and
- `POST /plataforma_v2/processa-feedback`.

Login, Face-result lookup, and status report are currently public. They are not
the only routes without route-specific authorization; the complete inventory
appears below. There is currently no platform-session cookie, platform
`Authorization` header, handle rotation, per-session revocation, current-session
API, logout API, revoke-all API, credentialed CORS, Origin/CSRF enforcement, or
explicit session response-cache policy.

The current frontend has exactly these seven `sessionStorage` keys and no
`removeItem()` or `clear()` call:

- `IndexVerificado`;
- `Usuário_Foto_Cadastrada`;
- `Horário-Encerramento-Sessão`;
- `Usuário_Logado`;
- `Usuário_Autorização_Cadastro`;
- `Origem_Aviso_Dispositivo`; and
- `TempoSessão_Segundos`.

Explicit logout and timer expiry write only `Usuário_Logado=Não` and navigate.
The logged flag, registration flag, browser deadline, photo mirror, and warning
origin are forgeable UI state. Refresh returns workbook login status but the
frontend does not enforce it. A protected `401 {}` is presented as the generic
frontend failure instead of an explicit invalid-session transition. No page
handles BFCache restoration.

Current evidence:

- ordered route registration: [`app.js` lines 77-90](../app.js#L77-L90);
- signed-handle implementation:
  [`platform-row-authorization.js` lines 5-187](../platform-row-authorization.js#L5-L187);
- current learning-platform handlers:
  [`domains/learning-platform.js` lines 19-184](../domains/learning-platform.js#L19-L184);
- current API characterization: [signed authorization](api-contracts.md#signed-platform-row-authorization)
  and [route inventory](api-contracts.md#route-inventory); and
- pinned frontend evidence: [storage and API](https://github.com/IvyRoom/sistemas/blob/1afd814adc02fa2da5ab4c55c1eeb6ebb5bb05b7/docs/learning-platform-contracts.md#L825-L914),
  [logout/expiry](https://github.com/IvyRoom/sistemas/blob/1afd814adc02fa2da5ab4c55c1eeb6ebb5bb05b7/docs/learning-platform-contracts.md#L1221-L1243),
  and [session risks](https://github.com/IvyRoom/sistemas/blob/1afd814adc02fa2da5ab4c55c1eeb6ebb5bb05b7/docs/learning-platform-contracts.md#L1822-L1841).

`GATE-01` device/browser admission and `GATE-02` minimum-viewport admission
remain prerequisites ahead of every Topic 05 runtime change. This ADR neither
changes their precedence nor treats either browser gate as authentication.

## Trust and ownership boundaries

### Authoritative subject

The target subject is an immutable backend-generated UUID named `subject_id`.
It is not a workbook row, email address, display name, certificate identifier,
Face session identifier, or browser value. Every session record references one
`subject_id`, and handlers derive their subject only from the validated session
record.

During workbook migration, the backend may own a private adapter that maps the
stable subject to an encrypted exact legacy login value plus a mutable row hint.
Initial subject provisioning requires exactly one match under the current
strict, untrimmed login semantics. It also computes a keyed HMAC lookup token
with a dedicated identity-mapping key. A SQL unique constraint on that token
and one transactionally idempotent create-or-load operation guarantee that
concurrent first logins for the same exact account return one `subject_id`
rather than minting duplicates. The lookup-token key ID and encrypted-value key
ID are stored; neither the plaintext login nor either key is stored in the
token column.

The adapter must re-find and verify the exact encrypted value before a
row-scoped operation; insertion, deletion, or movement of a workbook row must
not change the subject. A controlled login change updates the encrypted value
and unique lookup token for the same `subject_id` in one transaction and never
mints a new subject merely because login text or row position changed. Missing,
duplicate, ambiguous, or unique-constraint-conflicting mappings fail closed and
block session issuance until reconciled. The login value and row hint never
leave the backend and never become session authority.

Client-supplied name, email, row, progress, or other identity fields remain
compatibility input until their own domain-authority milestones. They may not
select the session subject. If an operation has an independently identified
resource subject and it differs from the session subject, the target backend
returns the wrong-subject failure defined below without performing domain work.

### Durable production authority

Azure SQL Database Basic is the selected production session and revocation
store. The backend session-authority module is its sole application owner. All
backend instances use the same database and SQL UTC time; process memory may be
used only by deterministic tests and never as production authority.

The minimum owned schema is:

| Record | Required authority data |
| --- | --- |
| `learning_subject` | immutable `subject_id`; unique keyed legacy-login lookup token and key ID; encrypted exact legacy-account mapping and key ID; mutable row hint; credential version; keyed credential fingerprint and fingerprint-key ID; subject session epoch; irreversible `legacy_authority_disabled_at`; normalized eligibility state; entitlement expiry; eligibility observation and revalidation instants |
| `learning_session` | HMAC-SHA-256 identifier verifier and verifier-key ID; `subject_id`; phase; original issue time; phase start; absolute expiry; Face-requirement snapshot captured at credential validation; subject, credential, and global epoch snapshots; revocation time/reason; replacement relation |
| `learning_session_flow` | current provisional session; registration state; one private provider Face challenge reference; creation/consumption state; no client assertion of result |
| `legacy_session_compatibility` | unique HMAC-SHA-256 verifier of the complete signed legacy handle and dedicated key ID; immutable bound `subject_id`; original issue/expiry instants; revocation/incident state; never the raw handle or a mutable row as authority |
| `session_authority_control` | global session epoch; legacy-ledger seeding/enforcement state; legacy-handle issuance/acceptance flags and hard sunset; incident state |

Of the browser identifier, only an HMAC-SHA-256 verifier and its non-secret key
ID are persisted. The verifier key is owned by approved secret management and
is separate from the legacy signed-handle key. The raw identifier exists
transiently while setting the cookie and is never written to SQL, a URL, a
request/response body, a diagnostic, a fixture, a snapshot, or a log. A
separate, non-authorizing trace identifier supports diagnostics.

The bounded legacy ledger uses a second dedicated HMAC verifier key. Legacy
login computes the verifier over the complete signed handle, binds it to the
already resolved stable subject in the same SQL transaction, and returns the
raw handle only after that commit. The raw legacy value is never persisted or
logged. The target identifier verifier, legacy compatibility verifier,
credential fingerprint, and login lookup token all use distinct keys and key
IDs.

Current legacy issuance is deterministic for a row within one second. Ledger
insert is therefore idempotent: when the verifier already has the identical
`subject_id`, issue/expiry instants, and verifier-key ID, login returns the same
handle successfully and does not create a second row. A differing subject or
metadata for that verifier, multiple rows despite the unique constraint, or
corrupt/unreadable state is a `503` integrity incident.

Session lookup, phase change, identifier rotation, challenge consumption,
privilege elevation, per-session revocation, revoke-all, and epoch changes use
SQL transactions. Positive authorization results are not cached in process
memory. A committed old verifier is invalid immediately for every request that
has not already completed authorization.

### Required roadmap dependency

`Implement revocable sessions` cannot start production implementation until a
narrow session slice of Topic 12 is moved ahead of it:

1. model and review the four records above as the session subset of **Model the
   relational target**;
2. provision and secure Azure SQL Database Basic as the session subset of
   **Provision Azure SQL Basic**;
3. prove multi-instance transactions, backup/restore, UTC behavior, connection
   limits, capacity, outage handling, and least-privilege ownership with inert
   data; and
4. install secret/configuration ownership without exposing a connection string.

This dependency does not mark the broader relational-data foundation or any
workbook migration complete. It explicitly reorders only the production
session-authority slice. If the store is not ready, Topic 05 runtime work stops;
there is no process-memory or signed-handle production substitute.

## Session states and transitions

`expired`, `revoked`, and `rotated-out` are terminal records, not usable
capabilities.
`anonymous` means there is no usable session record for the request.

| State | Meaning | Authority |
| --- | --- | --- |
| `anonymous` | No cookie, malformed cookie, unknown verifier, or no active record | Public/session-free operations only |
| `credential-verified` | Valid credentials and eligibility were checked; a required registration or Face step has not begun | Current-session inspection, logout, registration enrollment when required, or existing-photo Face challenge creation only |
| `registration-pending` | The backend accepted registration enrollment for this subject | Registration upload/reconciliation, one Face challenge creation, current-session inspection, and logout only |
| `face-pending` | One backend-created provider challenge is bound to this subject and provisional session | Face completion for that bound challenge, current-session inspection, and logout only |
| `authenticated` | Credentials and every backend-required factor succeeded and an authenticated session was created | Protected learning operations, current-session inspection, current-session logout, and revoke-all |
| `expired` | Server time reached the provisional, authenticated, or entitlement deadline | No protected or provisional operation; status rejects without mutating the cookie |
| `revoked` | Logout, eligibility, reset, administrator, incident, or failed-factor policy ended authority | No protected or provisional operation; status rejects without mutating the cookie |
| `rotated-out` | A replacement identifier committed for an allowed phase change | No operation; status rejects the predecessor immediately, does not mutate the cookie, and follows no replacement link for client recovery |

Target transitions are exact:

| From | Event and backend proof | To | Identifier/time effect |
| --- | --- | --- | --- |
| `anonymous` | Credentials match one subject, account is eligible, and the backend reads exact `FACEID = Ativo` | `credential-verified` | Issue a new identifier; start the 20-minute provisional clock; capture that Face is required and whether registration is required |
| `anonymous` | Credentials match one subject, account is eligible, and the backend reads exact `FACEID = Inativo` | `authenticated` | Issue an authenticated identifier regardless of photo-registration state; capture that Face was not required and start the four-hour clock |
| `credential-verified` | Backend accepts required registration enrollment | `registration-pending` | Rotate identifier; preserve the original provisional deadline |
| `credential-verified` | Backend creates and privately binds an existing-photo Face challenge | `face-pending` | Rotate identifier; preserve the original provisional deadline |
| `registration-pending` | Required registration state is reconciled and the backend creates and privately binds a Face challenge | `face-pending` | Rotate identifier; preserve the original provisional deadline |
| `face-pending` | Backend reads the bound provider result and verifies the required liveness and subject match | `authenticated` | In one transaction consume the challenge, invalidate the provisional verifier, issue a new identifier, and start a new four-hour clock |
| `face-pending` | Backend verifies a failed factor result | `revoked` | Atomically revoke the active presented verifier without mutating the browser cookie; fresh credentials are required |
| Any active state | Its applicable server deadline is reached | `expired` | Reject; do not extend or rotate |
| Any active state | A revocation trigger commits | `revoked` | Reject; do not fall back to another presented credential |

Whenever a rotating transition creates the new-state record, its predecessor
becomes `rotated-out` in the same transaction. The replacement relation is for
server audit/reconciliation only and is never followed to recover or authorize
a client presenting the old identifier.

An internal step that is completed within one transaction before any
identifier is exposed need not perform a meaningless extra rotation. Every
phase change observable to the browser rotates. Successful authenticated
promotion is the only transition that restarts a clock.

If rotation commits but the browser does not receive `Set-Cookie`, the old
identifier remains invalid and the user must authenticate again. The backend
must not revive the old identifier or persist recoverable raw identifiers to
mask that availability failure.

## Permission matrix

Public/session-free operations ignore any incidental session cookie. `status`
and `logout` below mean the target session APIs, not the current workbook
refresh or browser-only logout.

Credential login never derives subject or permission from an incidental
cookie. Only after new credentials and eligibility succeed may its transaction
use a presented active current-session verifier as the predecessor to
revoke/replace that browser profile's prior session. An absent, malformed,
unknown, expired, revoked, or rotated-out cookie is not an authority-bearing
predecessor: with the store available, fresh valid credentials may create and
set a new session over it. An unavailable lookup is `503`, not a bypass.

| Operation | Anonymous | Credential verified | Registration pending | Face pending | Authenticated | Expired/revoked/rotated-out |
| --- | --- | --- | --- | --- | --- | --- |
| Public status report, client intake, quote, Conecta, access release, Face-result compatibility lookup, DRM URL, certificate validation, warning pages | Allow | Allow | Allow | Allow | Allow | Allow |
| Credential login / replace current profile session | Allow | Allow | Allow | Allow | Allow | Allow |
| Current-session status | `401` | Allow | Allow | Allow | Allow | `401` |
| Current-session logout | Idempotent `204` | Allow | Allow | Allow | Allow | Idempotent `204` |
| Registration enrollment | Deny | Allow only when backend says registration is required | Idempotent success | Deny | Deny | Deny |
| Registration upload/reconciliation | Deny | Deny | Allow | Deny | Deny | Deny |
| Existing-photo Face challenge creation | Deny | Allow only when backend says registration is complete and Face is required | Deny | `409`; no second active challenge | Deny | Deny |
| Registration Face challenge creation | Deny | Deny | Allow | `409`; no second active challenge | Deny | Deny |
| Bound Face completion | Deny | Deny | Deny | Allow | `200` current-state response after successful promotion | Deny |
| Study refresh/data read | Deny | Deny | Deny | Deny | Allow | Deny |
| Progress, assessment, feedback, and other protected writes | Deny | Deny | Deny | Deny | Allow | Deny |
| Revoke all subject sessions | Deny | Deny | Deny | Deny | Allow | Deny |

A provisional capability is not a weak authenticated session. It cannot read
study state, submit progress or assessment data, append feedback, create a
certificate, or authorize any other authenticated operation.

## Identifier, transport, and first-party topology

### Identifier and cookie

The raw identifier is 32 bytes from a cryptographically secure random source
and is encoded as unpadded base64url only for cookie transport. It contains no
subject, phase, time, row, or other meaning. The target cookie is exactly:

```text
__Host-machado-session=<opaque value>; Path=/; Secure; HttpOnly; SameSite=Strict
```

`Domain` is absent, making the cookie host-only. `Max-Age` and `Expires` match
the current record only as browser cleanup hints; SQL server time and the
session record remain authoritative.

Only successful identifier issuance or rotation emits `Set-Cookie`. Logout,
revoke-all, definitive Face failure, ineligibility, expiry, malformed/unknown
credentials, terminal records, repeated requests, store failures, and
compare-and-replace losers never emit a deletion or other cookie mutation. A
revoked browser value is inert backend-side until its original browser expiry
or until a later successful credential login overwrites it. This is deliberate:
HTTP cannot conditionally delete only the predecessor carried by a request, so
a delayed deletion response could otherwise erase a newer login cookie.

The cookie is accepted and set only by `api.machadogestao.com`. There is no
redirect, mirrored cookie, or fallback to an `azurewebsites.net` host. The
identifier never appears in Web Storage, JavaScript, URLs, API bodies, public
diagnostics, or unredacted telemetry.

### Infrastructure prerequisite

Before any target cookie is issued, operations must prove that
`api.machadogestao.com` has controlled DNS, a valid managed TLS certificate,
and an App Service custom-hostname binding for the intended production backend.
They must qualify first-party behavior in supported browsers. Until that proof
exists, the target is blocked; current cross-site fetches do not become safe
cookie sessions merely by adding `HttpOnly`.

At adoption, this verified hostname becomes the sole shared production API
origin for all eight current frontend consumers. Only learning session and
protected consumers opt into credentials and the session request header;
public consumers continue to omit credentials and remain session-free. The
frontend does not create a second competing backend-origin constant.

### CORS, Origin, and CSRF

Production session and protected responses use this exact browser boundary:

- frontend fetches use `credentials: "include"`;
- `Access-Control-Allow-Origin` is the exact
  `https://machadogestao.com` origin, never `*` and never reflected from an
  arbitrary request;
- `Access-Control-Allow-Credentials: true` is present for session/protected
  routes;
- allowed methods are only the methods owned by the matching route;
- allowed request headers are `Content-Type` and
  `X-Machado-Session-Request` for the target session surface;
- every cookie-authenticated request requires
  `X-Machado-Session-Request: 1`;
- every unsafe request requires the exact
  `Origin: https://machadogestao.com`; a missing, `null`, or different Origin
  is `403` before body or domain work;
- preflight responses use `Vary: Origin, Access-Control-Request-Method,
  Access-Control-Request-Headers`; actual responses use `Vary: Origin, Cookie`;
  and
- public routes remain session-free and do not derive authority from an
  incidental cookie.

SameSite is defense in depth, not the only CSRF control. Exact Origin and the
preflighted custom header prevent another same-site or cross-site origin from
using the cookie. Preview and local validation use invented origins and
synthetic transports; they never credential a request to production.

### Cache and response rules

Every session, credential, Face-challenge, Face-completion, and protected
response—including failures—sets `Cache-Control: no-store` and emits no `ETag`
or `Last-Modified`. Session responses also set `Pragma: no-cache`,
`Expires: 0`, and `Referrer-Policy: no-referrer`. A response body may contain
phase, server time, absolute expiry, and the next allowed operation, but never
the application session identifier, its verifier, a workbook row, or the
private provider challenge identifier.

## Server time and eligibility

All timestamps are UTC instants computed and compared by the backend/store.
Browser clocks and `Horário-Encerramento-Sessão` have no authority effect.

| Clock | Target decision |
| --- | --- |
| Provisional absolute lifetime | 20 minutes from successful credential verification |
| Authenticated absolute lifetime | Four hours from successful authenticated creation; the current duration is retained |
| Idle timeout | None in Topic 05 |
| Extension | No request, refresh, activity, registration step, or Face retry extends a deadline |
| Provisional rotation | Rotate the identifier but preserve the original provisional expiry |
| Authenticated elevation | Rotate the identifier and start a new four-hour authenticated clock |
| Entitlement limit | Effective session expiry is the earlier of its absolute deadline and the normalized account-entitlement expiry |

### Face-requirement account policy

While `BD - PLATAFORMA` remains the account-policy adapter, its existing
`FACEID` cell at platform-row index `4` is the single policy source for whether
a fresh credential login requires Face. The backend reads and interprets that
cell; the current response projection `Usuário_Status_FaceID`, a client request,
Web Storage, or any other browser state never selects or overrides the policy.

The accepted values and effects are exact and case-sensitive:

- `Ativo` requires the scoped registration/Face flow and backend-bound
  successful Face completion before an authenticated session can exist;
- `Inativo` waives that factor for this credential login, so valid credentials
  plus account eligibility create an authenticated four-hour session directly,
  regardless of photo-registration state; and
- a missing, blank, unreadable, or different value is an authority-data
  configuration failure. It fails closed as `503`, issues neither a target
  identifier nor a legacy handle, emits no `Set-Cookie`, and leaves any
  previously valid cookie session unchanged.

The selected value is captured in the newly issued session and copied unchanged
through every provisional rotation and authenticated promotion. An
`Ativo`/`Inativo` workbook edit applies only to a fresh credential login and is
not part of the five-minute eligibility-revalidation loop. It does not promote
an existing provisional session or revoke, downgrade, or upgrade an existing
authenticated session. A user blocked in a provisional Face flow must submit
fresh credentials after `Ativo` becomes `Inativo`; successful direct
authentication replaces the profile's presented provisional session and starts
a new four-hour clock. An authenticated session issued while Face was not
required continues until logout, revoke-all, absolute or entitlement expiry,
or another already defined revocation trigger. A later `Inativo` to `Ativo`
change governs the next login only.

Topic 05 requires no dedicated Face-policy audit log, actor/reason field,
time-bounded override, or management UI under the current single-operator
model. Existing workbook access control and ordinary file history are
operational safeguards, not session-authority records. If account-policy
ownership later expands or compliance requirements change, a separate
milestone may reconsider that decision without blocking Topic 05.

When account authority migrates from the workbook, this policy moves once to a
backend-owned boolean such as `face_auth_required`. The cutover switches the
read authority and stops consulting `BD - PLATAFORMA.FACEID` for new logins;
there is no workbook/SQL dual authority, precedence rule, or fallback window.

Credential validation synchronously reads current backend-owned account input,
normalizes entitlement expiry to a UTC instant, and records an eligibility
observation. Invalid, ambiguous, inactive, or already expired eligibility
fails closed without a session.

While the workbook remains the account adapter, its numeric Excel access-date
cell is interpreted as an inclusive `America/Sao_Paulo` civil date. Authority
ends at the exclusive start of the following civil date in that timezone,
converted once to UTC and stored as `entitlement_expires_at`. A nonnumeric,
nonfinite, out-of-range, or otherwise unparseable value is ineligible; locale
display text is never parsed back into authority. The exact `Ativo` account
status and that entitlement instant must both pass.

Every authority-bearing request and phase transition reads the session record.
A known entitlement expiry is compared with server time before provisional as
well as authenticated work. Workbook-driven eligibility may be reused only
until `eligibility_revalidate_at`, at most five minutes after the prior
successful observation. At or after that instant, authorization synchronously
revalidates before registration, Face challenge/completion, authenticated
promotion, or protected domain work. Revalidation failure is a `503`
availability failure; stale eligibility is never served past five minutes.

This is the exact maximum propagation claim for manual workbook deactivation
or an earlier workbook entitlement change: no provisional transition,
authenticated promotion, or protected operation may begin more than five
minutes after the last successful eligible observation without a fresh
successful revalidation. It is not a claim of immediate workbook notification.
A future backend-owned deactivation or credential-reset command must also
increment the subject epoch/credential version and revoke affected sessions in
the same SQL transaction before reporting success.

Until credential migration supplies a native version, workbook revalidation
also derives a keyed HMAC fingerprint of the current credential cell with a
dedicated secret and compares it with the subject's prior fingerprint. The raw
credential is never persisted in SQL or diagnostics. A changed fingerprint is
a credential-reset trigger: increment the credential version and revoke all
subject sessions before any further provisional or authenticated work. Missing
or ambiguous credential input fails as an eligibility dependency failure.

## Tabs, devices, and concurrency

- The host-only cookie is shared by all tabs in one browser profile. Topic 05
  does not create tab-local authenticated sessions.
- A successful credential login in that profile conditionally
  revokes/replaces an active session currently named by its cookie. When two
  requests present the same active predecessor, the SQL transaction compares
  that predecessor; one concurrent winner may set a new cookie and a stale
  loser returns `409` without `Set-Cookie`. All tabs then observe the winning
  shared cookie.
- Failed credential validation leaves an existing valid cookie session
  unchanged and issues no replacement.
- Logout in one tab revokes the shared current session and logs out every tab
  using it. The inert cookie remains until its original browser expiry or a
  successful login overwrites it; a stale tab cannot restore authority from
  Web Storage.
- Concurrent sessions on other browser profiles/devices are allowed. A new
  login does not revoke those older sessions.
- `DELETE /plataforma_v2/sessions/current` revokes only the current session.
- `DELETE /plataforma_v2/sessions` is an authenticated revoke-all operation;
  it increments the subject session epoch, revokes every subject session
  including the caller, and does not mutate the caller cookie.
- There is no target fixed device-count cap in Topic 05. Abuse controls remain
  a separate API-perimeter concern and cannot weaken the session checks.

Every phase rotation uses the same compare-and-replace rule. Every non-issuance
response has no `Set-Cookie`, so a losing, stale, logout, or other revocation
response cannot erase a newer winner that reaches the browser first. Simultaneous
credential requests with no usable predecessor cannot be serialized as one
profile login: they may create separate concurrent records under the
allowed-device policy, and each successful response may set its own cookie. The
browser profile retains only the last processed response cookie. Any
inaccessible record expires normally, and an out-of-order lower-phase cookie
can reduce availability but cannot inherit or grant the authority of another
record.

A delayed successful cookie-less credential response is a new authentication,
not a continuation of the session currently visible in the browser. It may set
an active cookie after a later-created session was ended by current-session
logout; that logout revokes only the record its request presented, not other
independent sessions or pending credential attempts. Revoke-all invalidates
every subject record committed before its epoch increment, so a delayed
response for one of those records installs only an inert cookie; a credential
issuance that commits after revoke-all is a new login and may be active. Future
tests cover both commit and response orders. Topic 05 does not introduce an
otherwise-identifying pre-login browser-profile credential merely to serialize
anonymous requests.

## Revocation triggers and propagation

Authorization reads the central SQL record on every authority-bearing
provisional or authenticated request/transition, so no instance-local positive
cache delays a committed revocation.

| Trigger | Target action | Maximum propagation |
| --- | --- | --- |
| First target-session issuance for a subject | Atomically set irreversible `legacy_authority_disabled_at` before issuing the cookie | Effective at SQL commit; every legacy authorization consults it |
| Current-session logout | Commit `revoked: logout`, then return `204` without mutating the inert cookie | Effective at SQL commit for requests not already authorized |
| Absolute provisional/authenticated expiry | Compare SQL UTC on every authorization | At the exact expiry instant |
| Idle expiry | None | Not applicable; no idle timeout |
| Known access-entitlement expiry | Compare normalized UTC entitlement on every request | At the exact known instant |
| Manual workbook deactivation, earlier entitlement change, or credential-cell change | Synchronous revalidation when the central observation reaches five minutes; fingerprint change revokes all | No more than five minutes of reusable positive eligibility for any provisional/authenticated transition or request |
| `FACEID` changes between exact `Ativo` and `Inativo` | No mutation or revocation of an existing session; capture the new policy only after fresh credential validation | The next fresh credential login reads it synchronously; existing sessions retain their issuance-time policy and ordinary deadlines |
| Backend-owned account deactivation | Revoke all and increment subject epoch in the account transaction | Effective at SQL commit |
| Credential reset | Increment credential version and revoke all before reset success | Effective at SQL commit; manual legacy changes use the five-minute revalidation bound |
| Administrator revoke-all | Increment subject epoch and revoke matching records | Effective at SQL commit |
| Target identifier leakage | Revoke current or all affected subject sessions | Effective at SQL commit |
| Legacy-handle leakage during migration | Set the mapped subject's irreversible legacy cutoff, or disable legacy acceptance globally when scope is uncertain | Effective at the central SQL commit |
| Session-store corruption, loss, or restore incident | Fail session traffic closed; retire every pre-incident verifier-key ID outside SQL; restore/reconcile records; create a new global epoch and verifier key; if legacy compatibility is or was in scope, force its acceptance flag off and retire/rotate the legacy signing key outside the restored store; resume only after every instance acknowledges the new epoch/keys | Failed lookups block immediately; every pre-incident target identifier and legacy handle is permanently invalid before traffic resumes, including credentials represented by a restored backup |
| Legacy signing-key incident during migration | Atomically disable legacy acceptance, increment the global epoch, then rotate the legacy key | Effective at the central control commit for new authorization attempts |
| New verifier-key exposure or authority incident | Fail session traffic closed, retire the affected key ID in external secret control, increment the global epoch, deploy a new key, and require every instance to acknowledge it before resume | New authorization is blocked immediately on incident-control activation; affected identifiers are rejected no later than five minutes after activation and before traffic resumes |

An already authorized request may continue under the current domain behavior
after a concurrent revocation commits. Topic 05 makes no wall-clock or
cancellation bound for that in-flight work; Topic 07 owns timeout and
cancellation. No request that begins authorization after the commit may
succeed. Later idempotency and transaction milestones must decide whether a
specific high-risk write needs a second pre-commit authorization check; this
ADR does not redesign progress, assessment, feedback, or certificate authority.

## Failure contract

The target separates invalid authority from unavailable authority. It does not
reuse a generic communication message as an invalid-session transition.

| Condition | HTTP class and effect |
| --- | --- |
| Missing, malformed, unknown, expired, revoked, or rotated-out cookie on status, phase, completion, revoke-all, or protected APIs | `401`; no domain work and no `Set-Cookie`; browser enters explicit anonymous/invalid-session presentation; current-session logout retains its explicit `204` exception below |
| Validly signed legacy handle with no compatibility binding after ledger enforcement | `401`; no domain work and no row-derived subject; caller must reauthenticate |
| Legacy verifier bound to different subject/metadata, multiple rows despite uniqueness, or corrupt/unreadable binding | `503` integrity/availability failure; fail closed, emit no credential detail, and raise an incident; an identical deterministic repeat is idempotent success, not this failure |
| Valid session in the wrong phase | `403`; no phase change and no domain work |
| Session subject differs from an independently owned resource subject | `403`; no domain work; privacy-safe audit event |
| Invalid credentials | `401`; issue no cookie and leave any existing valid profile session unchanged |
| Missing, blank, unreadable, or unexpected `BD - PLATAFORMA.FACEID` policy during credential login | `503` authority-data/configuration failure; issue no target identifier or legacy handle, emit no `Set-Cookie`, and leave any existing valid profile session unchanged |
| Matched subject newly known to be ineligible | `403`; revoke all sessions for that subject in the eligibility transaction; emit no `Set-Cookie` and never disturb another subject's valid profile session |
| Invalid/missing Origin or session request header | `403` before body/domain work |
| Competing transition, already active challenge, or not-yet-final bound Face result | `409`; current record remains authoritative |
| Session store unavailable or transaction outcome unknown | `503`; no legacy/memory fallback and no `Set-Cookie`; optional bounded `Retry-After` |
| Eligibility source unavailable when revalidation is due | `503`; no provisional transition, authenticated promotion, or protected work; do not convert the condition to invalid credentials |
| Backend/network loss before an authoritative response | Frontend blocks provisional transitions and protected work and shows availability/retry state; it does not forge anonymous, authenticated, or logged-out state |

Invalid cookie classes are intentionally indistinguishable to the caller.
Detailed reason is retained only as privacy-safe server state. Any presence of
the target cookie is decisive during migration—whether it is valid, invalid,
expired, revoked, rotated-out, wrong-phase, or temporarily unverifiable because
the store is unavailable. Legacy `IndexVerificado` is considered only when the
target cookie is absent; target-cookie failure never falls back to it. Cookie
absence is not itself permission to downgrade: legacy middleware also loads the
mapped subject from SQL and rejects every handle when that subject's irreversible
`legacy_authority_disabled_at` is set. Store unavailability is `503` rather
than unchecked legacy acceptance.

Rotation and authenticated elevation are one SQL transaction. On transaction
failure the old record remains in its prior phase and no new cookie is sent. On
commit the old verifier cannot authorize immediately, even if response
delivery later fails.

Logout is effect-idempotent. When its transaction revokes the active presented
verifier, it returns `204` without `Set-Cookie`; the retained browser value is
inert. No cookie, an obviously malformed cookie, an unknown/terminal record, a
stale compare-and-replace loser, and repeated logout also return `204` without
`Set-Cookie` and without revealing prior state. If a well-formed presented
identifier cannot be checked because the store is unavailable, logout returns
`503` without `Set-Cookie` and does not pretend revocation succeeded.

## Target API roles

Topic 05 preserves current domain payloads while adding or adopting these exact
session roles. None exists merely because it is listed here.

| Method and path | Owner and allowed phase | Success/cookie behavior | Idempotency and failure classes |
| --- | --- | --- | --- |
| `POST /plataforma_v2/login-FaceID` | Session authority; public credential validation; backend reads the exact `BD - PLATAFORMA.FACEID` policy | A target-mode request carries `X-Machado-Session-Request: 1`; exact `Inativo` creates `authenticated` directly, while exact `Ativo` creates only the minimum provisional phase required for registration/Face; `200` atomically sets `legacy_authority_disabled_at`, conditionally replaces an active profile session or overwrites an absent/unusable cookie, and issues the selected identifier; it never returns `IndexVerificado`. Before global stop-issuance, a cookie-less legacy-mode request may return the current handle only after its verifier-to-subject binding commits; an identical deterministic binding repeat is `200` | No automatic transport retry; missing/blank/unexpected Face policy is `503` without a target identifier, legacy handle, or `Set-Cookie`; same-active-predecessor conflict is `409` with no `Set-Cookie`; legacy mode with any target cookie or for a target-adopted subject is `409` upgrade-required without a handle; store/eligibility or binding-integrity unavailability is `503`; invalid credentials are `401`; known ineligibility is `403` without cookie mutation |
| `POST /plataforma_v2/sessions/current/registration-enrollment` | Session authority; `credential-verified` with required registration, or already `registration-pending` | `204`; first success rotates to `registration-pending`; no body identifier | Repeated in registration-pending is `204` without another rotation; `401`, `403`, `409`, `503` |
| `POST /plataforma_v2/CadastroFoto_e_FaceID` | Registration domain plus session authority; `registration-pending`; a `face-pending` repeat is conflict only | `200` after registration state is reconciled and one provider challenge is durably bound; returns only provider data required by the SDK and rotates to `face-pending`; application session/provider session identifiers are absent | No blind retry of upload/external creation; an ambiguous external outcome does not promote or rotate, keeps the phase provisional, marks the private flow reconciliation-required, and makes repeat return `409` until the separately authorized registration-reconciliation decision resolves it; session failures use `401`/`403`/`409`/`503`; current non-session domain status/body remains unchanged until its own redesign |
| `POST /plataforma_v2/FaceID` | Face domain plus session authority; `credential-verified` with existing registration; a `face-pending` repeat is conflict only | `200`; creates at most one active bound challenge, returns only provider data required by the SDK, and rotates to `face-pending` | Repeated while creating/active is `409`, never a second untracked challenge; session failures use `401`/`403`/`503`; current non-session domain status/body remains unchanged until its own redesign |
| `POST /plataforma_v2/sessions/current/face-completion` | Session authority plus private Face adapter; `face-pending`; the resulting `authenticated` state supports a current-state repeat | `200` on passing result; accepts no client result or provider session ID, reads the one bound result, atomically creates authenticated session, rotates cookie, and returns current status | Bound result pending is `409` and preserves phase/cookie; definitive factor failure atomically revokes the active verifier without cookie mutation and returns `403`; provider/network/store unavailability is `503` and preserves `face-pending`/cookie; challenge is consumed once; repeat under the resulting authenticated cookie returns `200` current status without creating another session |
| `GET /plataforma_v2/sessions/current` | Session authority; any active phase | `200` no-store JSON with `authenticationPhase`, `serverTime`, `expiresAt`, `eligibilityRevalidateAt`, and allowed next-operation roles; never an identifier | Safe/idempotent; invalid `401` and preserving `503` emit no `Set-Cookie` |
| `DELETE /plataforma_v2/sessions/current` | Session authority; any/none | `204`; revoke the active record when present; every outcome leaves the cookie untouched and an active success makes its retained value inert | Effect-idempotent; `503` only when a presented well-formed record cannot be authoritatively checked; no response emits `Set-Cookie` |
| `DELETE /plataforma_v2/sessions` | Session authority; `authenticated` | `204`; revoke all subject sessions without mutating the caller cookie | Do not automatically retry an ambiguous response; after success the revoked caller cannot invoke it again and a repeat without new authentication is `401`; `403`, `503`; no response emits `Set-Cookie` |
| Existing protected learning routes | Session middleware then existing owning domain; `authenticated` | Authenticate cookie, derive subject, strip session credential from domain input, preserve scoped domain response until its own redesign | Read/write idempotency remains owned by the later API reliability/domain milestones; session failures follow this ADR |

`authenticationPhase` and allowed next-operation roles in a status response are
presentation hints. Middleware reloads the authoritative record and recomputes
permission for every operation; replaying or editing a response grants nothing.

The current public
`GET /plataforma_v2/FaceID_resultado/:Azure_Face_API_LivenessSession_sessionID`
is not the promotion API. It remains a compatibility-only public lookup until
the later **Protect Face-result lookup** milestone. Topic 05 must create and
adopt the session-bound completion role first; the public route cannot promote,
create, refresh, or prove an authenticated session. This sequencing resolves
the dependency on later Face-result security without claiming that this task
protects the current route.

## Current-versus-target endpoint matrix

This table covers the unchanged 14-route inventory. **Target** entries are
future authorization classifications, not implemented behavior.

| Current endpoint | Source-observed current authority | Topic 05 target disposition |
| --- | --- | --- |
| `POST /landingpage/solicitacaoorcamento` | Public | Remains session-free; quote authority is outside Topic 05 |
| `POST /conecta/processa-recomendacao` | Public | Remains outside learning-session authority |
| `POST /clientes/processa-formulario` | Public client intake | Remains public and session-free |
| `POST /clientes/liberacao-acesso-plataforma` | Public operational route | Remains session-free in Topic 05 and is not granted learner authority; exposure is owned by its later perimeter/domain work |
| `POST /plataforma_v2/login-FaceID` | Public; active match may mint signed row handle | Public credential entry; target-mode success atomically disables legacy authority for the subject, issues/replaces the target cookie, and never returns `IndexVerificado`; bounded legacy mode remains only for not-yet-adopted subjects before global stop-issuance |
| `POST /plataforma_v2/CadastroFoto_e_FaceID` | Multipart `IndexVerificado`; Multer runs first | `registration-pending` only; authenticate before resource-heavy parsing; successful challenge becomes `face-pending` |
| `POST /plataforma_v2/FaceID` | JSON `IndexVerificado` | `credential-verified` only for existing-photo Face challenge; successful creation becomes `face-pending` |
| `GET /plataforma_v2/FaceID_resultado/:Azure_Face_API_LivenessSession_sessionID` | Public | Remains public compatibility-only during Topic 05 and never promotes; later Topic 17 protects/retires it |
| `POST /plataforma_v2/refresh` | JSON `IndexVerificado` | `authenticated` only; derives subject and eligibility from session |
| `POST /plataforma_v2/updates` | JSON `IndexVerificado` | `authenticated` only; later progress/assessment authority still required |
| `POST /plataforma_v2/processa-feedback` | JSON `IndexVerificado` | `authenticated` only; later feedback authority still required |
| `GET /ezdrm-playready-authorization-url` | Public | Remains session-free and otherwise unchanged by Topic 05; media/DRM authority remains later work |
| `POST /plataforma_v2/statusreport` | Public | Remains public and session-free; secure report redesign remains Topic 15 |
| `GET /validacaocertificados/:Solicitante_CertificadoID` | Public | Remains outside learning-session authority |

The public client-intake page, status-report page, viewport warning, and
device/browser warning remain session-free. GATE-01 and GATE-02 behavior is
unchanged. This decision does not silently classify any public static page as a
protected API.

## Browser-key disposition

No key is added, renamed, removed, or given new runtime behavior by this ADR.
Later Topic 05 implementation follows this disposition:

| Exact current key | Target authority | Bounded disposition |
| --- | --- | --- |
| `IndexVerificado` | None | During migration accept only through the immutable full-handle-verifier binding and only for a not-yet-adopted subject within the original four-hour lifetime; target-session issuance atomically and irreversibly disables every subject handle; global issuance then stops at authoritative-client cutover; remove the key and all acceptance after final rejection |
| `Usuário_Foto_Cadastrada` | None | Remove the dead stored mirror when registration/next-operation presentation comes from current-session status |
| `Horário-Encerramento-Sessão` | None | During adoption it may display a server-returned expiry only; remove after server-time countdown/status is adopted |
| `Usuário_Logado` | None | During adoption it may mirror presentation; it never gates backend work; remove after direct and restored pages use current-session validation |
| `Usuário_Autorização_Cadastro` | None | Replace with server-side `registration-pending`; remove when registration enrollment is adopted |
| `Origem_Aviso_Dispositivo` | None | Preserve as the current UI/history compatibility marker through this definition task; decide/remove it in **Guard restored protected pages** alongside restored-page and login-history reconciliation |
| `TempoSessão_Segundos` | None | Remove the dead read when the server-time session timer is adopted |

Malformed or forged storage affects at most presentation/navigation. It never
changes a session record, request subject, phase, expiry, eligibility, or
permission.

## Migration, cutover, and rollback

### Prerequisites

Before the dual stack begins:

1. keep `GATE-01` and `GATE-02` ahead of Topic 05 work;
2. provision and qualify the pulled-forward Azure SQL session slice;
3. verify `api.machadogestao.com`, TLS, host binding, cookie behavior,
   credentialed CORS, exact Origin/CSRF rejection, and no-store behavior with
   synthetic accounts;
4. implement the stable subject mapping and five-minute eligibility
   revalidation; and
5. run the durable legacy-binding seeding gate described below for one full
   four-hour issuance horizon; and
6. characterize and decide the separately recorded partial Face-registration
   states with representative nonproduction data; ambiguity must remain
   provisional and blocked until reconciliation is defined; and
7. add privacy-safe measurements that count target versus legacy authorization
   without recording either credential.

### Topic 05 sequence

1. **Implement revocable sessions.** Add the store, legacy compatibility
   binding, target cookie/status/logout, provisional phases, private Face
   binding, promotion, revocation, and dual-stack middleware. Current runtime
   consumers still use the legacy contract. No provisional phase grants
   authenticated operations.
2. **Adopt authoritative sessions.** Change the frontend to credentialed
   requests, target status/phases, server time, and session-bound Face
   completion behind a production adoption gate. Stop minting
   `IndexVerificado` only when the authoritative frontend and authoritative
   logout are ready for the same coordinated production enablement.
3. **Make logout authoritative.** Call current-session logout, wait for its
   authoritative result, then update presentation/navigation. Repeated logout
   remains idempotent; availability failure is not presented as successful
   revocation. This task stays separately reviewable after adoption work, but
   the production adoption gate cannot open before it is ready.
4. **Guard restored protected pages.** Validate before protected startup on
   direct load and `pageshow`, including `event.persisted` BFCache restoration.
   No protected fetch, media initialization, Face initialization, timer, or
   write begins before validation succeeds.

### Bounded dual stack

Before target acceptance begins, a ledger-seeding release must run continuously
for one complete four-hour legacy lifetime. Every newly issued handle is bound
transactionally by its full-handle HMAC verifier to the stable subject before
the response is sent, while the current payload and four-hour behavior remain
otherwise unchanged. If ledger writes or continuity fail, issuance fails closed
and the four-hour horizon restarts after repair. Handles issued before seeding
therefore expire before enforcement begins; this task performs no production
inspection or backfill.

Only after that horizon may the central ledger-enforcement flag turn on and the
dual-stack window start. From then on, a validly signed handle with no unique
compatibility binding is `401`; a duplicate, conflicting, or corrupt binding is
a `503` integrity incident. Neither case may fall back to the signed row index.

The dual-stack window starts when production first accepts the target cookie
and legacy handle together. Its fixed hard maximum is seven calendar days,
stored in `session_authority_control`; it cannot be extended after the window
starts. A failed adoption must abort/roll back before the sunset rather than
continue dual authority.

The two login modes are explicit during that window:

- a legacy client presents no target cookie, omits
  `X-Machado-Session-Request: 1`, and may receive the current four-hour handle
  only while global issuance is enabled and its subject has no
  `legacy_authority_disabled_at`; it receives no target cookie;
- a target client sends that header; successful issuance atomically sets the
  subject's irreversible `legacy_authority_disabled_at`, creates the target
  record, and returns only the target cookie, never `IndexVerificado`.

Every legacy protected request must verify the existing signature/expiry,
compute the full-handle verifier, derive authority only from its immutable
ledger-bound `subject_id`, and consult SQL for the cutoff plus the same bounded
current eligibility decision. The signed row remains compatibility evidence
for the private adapter, never the authority mapping; row insertion, deletion,
or movement cannot change the bound subject. A set subject cutoff rejects all
of that subject's legacy handles with `401`, even if the target cookie was
evicted, manually removed, expired, or never delivered. SQL or due-eligibility
outage is `503`. This per-subject cutover is deliberately cross-device: old
legacy tabs/devices for that subject must load the target-capable client and
reauthenticate, while existing target sessions on other devices remain allowed.
No rollback or administrator action may clear the subject cutoff.

If first target issuance commits but its response is lost, both the subject
cutoff and target record remain committed. The user performs a fresh target-mode
credential login; neither response loss nor rollback can revive a legacy
handle.

At authoritative-client cutover, the backend stops issuing
`IndexVerificado`. It may accept only handles issued before that instant and
only for subjects that have not adopted target authority, and only until their
existing four-hour expiry. Cutover must occur at least four hours before the
seven-day sunset; otherwise the adoption aborts/rolls back. Four hours after
the final legacy issuance, and never later than the sunset, legacy acceptance
is disabled centrally. Production enablement and stop-issuance require:

- all target synthetic and hosted checks passing;
- the unchanged public surfaces passing;
- a reviewed rollback release pair.

Final legacy rejection occurs mechanically when the last issued handle's
four-hour expiry passes, never later than the sunset; no traffic or telemetry
criterion may extend it. Privacy-safe telemetry must then confirm that zero
legacy-only protected request was accepted after that instant and raise an
incident if the invariant is violated. Compatibility-code removal may follow,
but acceptance cannot be restored while investigating residual callers.

Failure to meet those criteria before global cutover aborts that cutover while
at least one full four-hour drain remains; it cannot clear an adopted subject's
cutoff. Once global stop-issuance starts, final rejection is irreversible and
occurs after the last lifetime or at the seven-day hard sunset, whichever is
earlier. No failure, traffic, rollback, or telemetry creates indefinite dual
authority. After legacy acceptance is disabled, `IndexVerificado` is rejected
even if present with an invalid target cookie.

### Rollback boundaries

- Before target issuance, rollback is code/config only and no target-session
  data is authoritative; legacy handles remain the current authority.
- While legacy issuance remains enabled inside the seven-day window, the
  frontend and backend may roll back together to the reviewed dual-stack pair
  only for subjects that have not adopted target authority. Any adopted subject
  remains target-only and requires cookie-capable code. Target sessions may be
  revoked globally; their raw identifiers are never exported or converted to
  legacy handles, and the subject legacy cutoff is never cleared.
- After legacy issuance stops but before its last four-hour lifetime ends,
  rollback may restore only the reviewed cookie-capable dual-stack code with
  the central stop-issuance control preserved, and only before the hard sunset.
  It must not mint a new legacy handle or extend an existing handle lifetime.
- After central legacy rejection, rollback must preserve target authority. It
  may roll forward/fix the target or use the reviewed target release pair; it
  must not re-enable legacy issuance or acceptance.
- Schema rollback never discards revocation or subject-cutoff evidence while
  any corresponding target or legacy identifier could remain valid.

## Threat and failure model

| Threat/failure | Required target control |
| --- | --- |
| Stolen or replayed target identifier | HttpOnly first-party cookie, no URL/body/storage exposure, verifier-only persistence, absolute expiry, per-session/revoke-all, and rotation |
| Stolen or replayed legacy handle during migration | Accept only with an immutable verifier-to-subject binding for a not-yet-adopted subject within the unchanged four-hour lifetime; first target issuance disables every subject handle and global rejection ends the bounded exception |
| Forged browser flags, Face waiver, or malformed storage | Backend ignores Web Storage and client-projected `Usuário_Status_FaceID` for subject, phase, permission, expiry, eligibility, and Face policy; only the backend-read policy at fresh credential validation can waive Face |
| Stale tab | Shared cookie plus current-session validation; invalid response blocks protected work and reconciles presentation |
| Multiple tabs | One profile cookie; rotation/logout affects all tabs; same-predecessor transitions serialize in SQL, while cookie-less logins follow the explicit independent-login policy |
| Multiple devices | Separate records allowed; new login preserves other devices; revoke-all invalidates every subject record |
| Legacy downgrade after target revocation or cookie loss | Irreversible per-subject legacy cutoff is committed with first target issuance and checked in SQL on every legacy authorization; adopted legacy devices must reauthenticate through the target client |
| Browser-clock or stored-deadline tampering | SQL/server UTC owns every deadline and response supplies display time only |
| Refresh/logout race | Central transaction decides; later-started authorization sees revocation; already authorized work has the documented in-flight boundary |
| Malformed/unknown cookie | Uniform `401`, no parser/domain work, no cookie mutation, no detail oracle; a later successful login can replace it |
| BFCache restoration | `pageshow` revalidation before protected work, including `event.persisted` |
| Backend restart or scale-out | Durable shared SQL records and no process-memory authority |
| Session-store outage | `503` fail closed, no legacy/memory fallback, no false invalid-credential or successful-logout response |
| Eligibility-source outage | Fresh observation usable only within five minutes; when due, `503` before provisional transition, promotion, or protected work |
| Store restore or verifier/signing-key compromise | Fail closed, retire keys outside restored SQL, advance epoch, disable legacy when applicable, and resume only after every instance acknowledges the new authority generation |
| Identifier leakage through logs/diagnostics | No raw/verifier logging; separate trace IDs; redaction tests and diagnostic review |
| Cache replay | `no-store`, no validators, `Vary` rules, no identifier in response body |
| Cross-site or sibling-origin request | SameSite Strict, exact Origin, exact allow-origin, credentialed CORS, custom preflighted header |
| Same-origin script compromise/XSS | HttpOnly blocks identifier reads but not credentialed same-origin requests; backend phase/permission/subject checks still apply, while CSP, output safety, and XSS prevention remain later perimeter work |
| Face client assertion or provider-ID replay | Provider challenge stored privately and bound to subject/session; completion accepts neither result nor provider session ID |
| Workbook row movement | Stable `subject_id`; target sessions use it directly, and legacy authority uses the issuance-time verifier binding before the private adapter re-finds/verifies a mutable row hint |
| Account deactivation or entitlement shortening | Server-time expiry plus synchronous eligibility revalidation at the documented five-minute bound |
| Rotation response loss | Old identifier remains invalid; fresh authentication rather than raw-identifier recovery |

## Synthetic future verification matrix

All examples use invented identifiers, fake UTC clocks, synthetic accounts,
inert SQL records, injected Graph/Face adapters, and denied production
networking.

| ID | Future acceptance coverage |
| --- | --- |
| `SESSION-TARGET-01` | Every protected operation accepts a valid `authenticated` session and derives the one expected subject; no provisional state reaches it |
| `SESSION-TARGET-02` | Missing, malformed, unknown, expired, revoked, rotated-out, and wrong-subject sessions fail closed with the selected class, no domain call, and no stale-response cookie mutation |
| `SESSION-TARGET-03` | `credential-verified`, `registration-pending`, and `face-pending` reach only their exact minimum operations |
| `SESSION-TARGET-04` | Authenticated elevation commits challenge consumption, old-record invalidation, new verifier, and phase atomically |
| `SESSION-TARGET-05` | The old provisional identifier fails immediately after elevation, including on another instance |
| `SESSION-TARGET-06` | Forged values for all seven browser keys never grant backend access or alter the subject/phase |
| `SESSION-TARGET-07` | Browser-clock and stored-deadline tampering do not change server expiry; no request extends either clock |
| `SESSION-TARGET-08` | Known entitlement expiry is exact; manual deactivation/shortening or credential-fingerprint change cannot authorize a provisional transition, promotion, or protected request past the five-minute revalidation bound; stale-source failure is `503` |
| `SESSION-TARGET-09` | Current-session logout revokes only the current record; every successful/repeated/invalid `204` leaves the cookie untouched and the retained value is inert; the subject legacy cutoff prevents fallback resurrection; revoke-all invalidates all subject records including other devices without cookie mutation |
| `SESSION-TARGET-10` | Tabs share one profile cookie and observe rotation/logout; concurrent target devices remain active until their own revocation/expiry; new login does not revoke other target devices; delayed cookie-less login responses follow the documented new-authentication policy in both commit/response orders |
| `SESSION-TARGET-11` | Store failure and unknown transaction outcome fail closed as `503`; restore/key-compromise recovery retires target and legacy keys outside restored SQL, advances epoch, invalidates backup-era credentials, preserves no legacy fallback, and waits for every instance to acknowledge before resume |
| `SESSION-TARGET-12` | Cookie flags, host-only scope, target hostname, credentialed CORS, exact Origin, custom header, preflights, `Vary`, cache, issuance-only cookie mutation, and natural-expiry/overwrite behavior match this ADR |
| `SESSION-TARGET-13` | No application session identifier/verifier appears in URLs, bodies, Web Storage, logs, fixtures, snapshots, public diagnostics, or provider challenge fields |
| `SESSION-TARGET-14` | Face promotion uses only the backend-bound provider result; client verdict/provider-ID assertions cannot promote and the public legacy lookup cannot promote |
| `SESSION-TARGET-15` | Direct and BFCache-restored protected pages validate before any protected request, media, Face runtime, timer, or write; network loss shows availability state |
| `SESSION-TARGET-16` | Dual-stack acceptance is bounded; first target issuance irreversibly disables all legacy authority for that subject across cookie loss and devices; legacy issuance stops globally, remaining not-yet-adopted handles age out for four hours, and central rejection persists after sunset |
| `SESSION-TARGET-17` | Public status report, client intake, quote, Conecta, certificate validation, viewport warning, and device/browser warning retain their existing session-free behavior |
| `SESSION-TARGET-18` | Credential, registration, challenge, completion, status, logout, revoke-all, wrong-stage, and store-outage APIs match their methods, phases, status classes, cookie, and idempotency rules |
| `SESSION-TARGET-19` | Restart and multi-instance tests observe committed rotation/revocation and subject legacy cutoffs without a positive authorization cache; same-predecessor races let only the compare-and-replace winner issue a cookie, every non-issuance response has no `Set-Cookie` in either response order, fresh login overwrites unusable cookies, and cookie-less login/logout races follow the documented last-processed-response/new-authentication policy |
| `SESSION-TARGET-20` | Runtime, dependencies, 14-route current inventory, five legacy placements, current seven-key inventory, public Face-result behavior, and artifact identities remain unchanged by this definition task |
| `SESSION-TARGET-21` | A full four-hour ledger-seeding horizon precedes dual-stack enforcement; every accepted legacy handle resolves through one immutable verifier-to-subject binding; an identical same-second deterministic issuance is idempotent success; pre-ledger/missing bindings fail `401`, differing/corrupt bindings fail `503`, and workbook row movement never changes the subject |
| `SESSION-TARGET-22` | Fresh-login tests prove exact backend-read `FACEID = Ativo` requires backend-bound Face, exact `Inativo` creates authenticated authority directly regardless of photo state, and missing/blank/other values fail `503` without a target identifier, legacy handle, or `Set-Cookie`; edits affect only fresh logins, existing provisional/authenticated sessions retain their captured policy and normal lifetime, browser assertions cannot waive Face, and the later account-authority cutover leaves exactly one policy source |

## Later implementation ownership

This ADR is implemented only through the ordered Topic 05 tasks:

- **Implement revocable sessions** owns the pulled-forward store slice,
  subject/session records, cookie APIs, phases, rotation, eligibility,
  revocation, and private Face binding.
- **Adopt authoritative sessions** owns credentialed frontend requests,
  session-status consumption, provisional navigation, server-time presentation,
  Face completion, and bounded legacy cutover.
- **Make logout authoritative** owns logout request, availability handling,
  presentation cleanup, and cross-tab outcome.
- **Guard restored protected pages** owns direct/BFCache validation and the
  later disposition of restored-page/history compatibility state.

Later topics still own HTTP envelope redesign, retry/idempotency beyond session
transitions, progress/assessment/feedback/certificate authority, account and
password migration, workbook/SQL domain migration, public status-report
security, Face-result route protection, and media/DRM authority.

## Definition-task non-effects

This decision task intentionally changes no runtime JavaScript, dependency,
route, method, payload, status, cookie, CORS rule, workflow, deployment
manifest, infrastructure, account, workbook, Face session, live session,
storage key, logout path, timer, refresh behavior, BFCache behavior, public
surface, frontend artifact, or import graph. No production request or
integration is required to validate it.
