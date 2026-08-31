# Machado Backend

Node.js and Express API for Machado's website, client onboarding, Conecta referral program, learning platform, and certificate validation.

The service connects the frontend applications in [`IvyRoom/sistemas`](https://github.com/IvyRoom/sistemas) to Microsoft Graph, Excel workbooks, email delivery, and Azure AI Face.

## Main application areas

| Area | Responsibility |
|---|---|
| Website | Receives quote requests and sends notification emails. |
| Conecta | Validates and records referrals, then sends notification and confirmation emails. |
| Client onboarding | Records company and participant information in Excel workbooks. |
| Learning platform | Handles login, Face liveness sessions, progress updates, feedback, and status reports. |
| Session authority | Provides a durable revocable-session target behind production-disabled controls while legacy clients remain authoritative. |
| Certificate validation | Verifies public certificate IDs against learning-platform records. |
| Access operations | Sends learning-platform access instructions. |
| DRM support | Returns the PlayReady authorization parameters required by the frontend. |

## Technology stack

- Node.js 24
- Express 4
- Microsoft Graph API
- Microsoft Authentication Library (MSAL)
- Azure AI Face REST API
- Azure SQL Database through `mssql` (connected only when the durable-store latch is enabled)
- Azure App Service
- GitHub Actions

The application uses CommonJS. [`app.js`](app.js) is the thin, import-safe
composition root: it exports `createApp(dependencies)`, constructs Express,
installs the global middleware in its exact order, composes integration adapters
from the injected raw clients, passes their capabilities to the domain handler
factories, and registers the explicit ordered route table. [`server.js`](server.js)
is the production entry point and owns environment loading, production client
construction, listener startup, and the Microsoft Graph token lifecycle.

## Repository structure

| Path | Purpose |
|---|---|
| `app.js` | Thin import-safe Express composition root, raw-client adapter composition, global middleware, handler-factory composition, and explicit ordered route registration. |
| `server.js` | Production configuration, Graph and Face client construction, listener startup, and Graph token refresh. |
| `platform-row-authorization.js` | Signed learning-platform row-handle creation and verification. |
| `domains/session-authority/` | Stable subjects, session phases, verifier-only credentials, eligibility, HTTP/cookie policy, and target/legacy authority orchestration. |
| `domains/quote-requests.js` | Quote-request handler factory, notification template, retry boundary, and error mapping. |
| `domains/conecta-recommendations.js` | Conecta validation, workbook payloads, notification templates, business ordering, retry boundaries, and error mapping. |
| `domains/client-onboarding.js` | Client-intake and access-release handler factories, workbook payloads, email templates, scheduling, retry boundaries, and error mapping. |
| `domains/learning-platform.js` | Learning-platform login, Face workflow, refresh, update, feedback, and status-report business behavior, retry boundaries, and error mapping. |
| `domains/drm.js` | Deterministic PlayReady authorization-output handler factory. |
| `domains/certificate-validation.js` | Public certificate-validation handler factory, retry boundary, and score normalization. |
| `integrations/microsoft-graph.js` | Import-safe Microsoft Graph adapter for exact API paths, verbs, request envelopes, and response-shape access. Each SDK-calling operation makes one attempt. |
| `integrations/azure-face.js` | Import-safe Azure Face adapter for exact endpoint paths, multipart mechanics, and result projection. Each SDK-calling operation makes one attempt. |
| `integrations/azure-sql-session-store.js` | Lazy, fail-closed Azure SQL session store with serializable compare-and-replace transitions. |
| `migrations/001-session-authority.sql` | Forward-only schema for the five session-authority records; it is not applied automatically. |
| `shared/retry.js` | Shared application-level retry helper using the injected sleep hook. |
| `shared/escape-html.js` | Shared HTML escaping for dynamic email values. |
| `img/` | Images used in emails and other backend-generated content. |
| `test/` | Isolated Node.js acceptance tests and test support. |
| `docs/runbooks/` | Repeatable operational and maintenance procedures. |
| `.github/workflows/` | Continuous deployment to Azure App Service. |
| `AGENTS.md` | Repository-specific collaboration and engineering guidance. |

## Prerequisites

- Node.js 24
- npm
- Access to the project's Microsoft Entra application credentials
- Access to the project's Azure AI Face resource

## Environment variables

Create an ignored `.env` file in the repository root:

| Variable | Purpose | Required |
|---|---|---|
| `CLIENT_ID` | Microsoft Entra application client ID | Yes |
| `TENANT_ID` | Microsoft Entra tenant ID | Yes |
| `CLIENT_SECRET` | Microsoft Entra application client secret | Yes |
| `PLATFORM_ROW_AUTHORIZATION_KEY_BASE64` | Stable canonical Base64 that decodes to exactly 32 random bytes and signs learning-platform row authorization handles | Yes |
| `AZURE_FACE_API_ENDPOINT` | Azure AI Face resource endpoint | Yes |
| `AZURE_FACE_API_KEY` | Azure AI Face resource key | Yes |
| `PORT` | HTTP port; defaults to `3000` | No |

The durable-store latch and all session-authority rollout controls default to
`false`. With those defaults, the SQL driver is never connected, the target
routes are not registered, and the deployed API retains its 14-route legacy
behavior.

| Session-authority variable | Purpose |
|---|---|
| `SESSION_AUTHORITY_DURABLE_STORE_REQUIRED` | One-way production composition latch. Every rollout control requires it; once production seeding starts it must remain enabled so configuration rollback cannot restore unchecked legacy authority. |
| `SESSION_AUTHORITY_TARGET_ROUTES_ENABLED` | Registers the five target session routes and enables their strict transport boundary. |
| `SESSION_AUTHORITY_TARGET_ISSUANCE_ENABLED` | Permits target identifier issuance only with the coordinated gates below. |
| `SESSION_AUTHORITY_LEGACY_SEEDING_ENABLED` | Enables transactionally bound legacy-ledger issuance. |
| `SESSION_AUTHORITY_LEGACY_ENFORCEMENT_ENABLED` | Requires every accepted legacy handle to have a qualified immutable ledger binding. |
| `SESSION_AUTHORITY_SUBJECT_ADOPTION_ENABLED` | Permits the irreversible per-subject legacy cutoff on first target issuance. |
| `SESSION_AUTHORITY_PROTECTED_ROUTES_ENABLED` | Selects target middleware for the existing protected learning routes. |
| `SESSION_AUTHORITY_FIRST_PARTY_TOPOLOGY_QUALIFIED` | Records the separately reviewed DNS, TLS, App Service hostname, browser, CORS, Origin, CSRF, cache, and response-boundary proof. |
| `SESSION_AUTHORITY_SQL_CONNECTION_STRING` | Private least-privilege Azure SQL connection string. |
| `SESSION_AUTHORITY_EXPECTED_GENERATION` | Positive database authority generation acknowledged by this deployment; mismatched instances fail session traffic closed during recovery. |
| `SESSION_AUTHORITY_LEGACY_SIGNING_KEY_ID` | Non-secret identity for the existing `PLATFORM_ROW_AUTHORIZATION_KEY_BASE64` material. It is required only with the durable-store latch and is fenced independently from the six session-authority keys. |
| `SESSION_AUTHORITY_SQL_CONNECTION_TIMEOUT_MS`, `SESSION_AUTHORITY_SQL_REQUEST_TIMEOUT_MS`, `SESSION_AUTHORITY_SQL_POOL_MAX`, `SESSION_AUTHORITY_SQL_POOL_IDLE_TIMEOUT_MS` | Bounded connection, request, and pool behavior. Startup rejects connection/request timeouts above 30 seconds, a per-instance pool above 10 connections, or an idle timeout above five minutes; defaults are 5 seconds, 5 seconds, 5 connections, and 30 seconds respectively. |

Six separate key descriptors are required whenever the durable-store latch is
enabled. Each uses an `_KEY_ID` and canonical 32-byte `_KEY_BASE64` variable
under these prefixes: `SESSION_AUTHORITY_TARGET_VERIFIER`,
`SESSION_AUTHORITY_LEGACY_COMPATIBILITY`, `SESSION_AUTHORITY_LOGIN_LOOKUP`,
`SESSION_AUTHORITY_CREDENTIAL_FINGERPRINT`,
`SESSION_AUTHORITY_ACCOUNT_MAPPING`, and
`SESSION_AUTHORITY_FACE_CHALLENGE`. Key IDs and key material must be distinct
by purpose. The legacy signing key ID and material must also differ from all
six. Never print, commit, or reuse any of these values.

Before the durable store can serve authority traffic, its singleton control
record must be explicitly bound to every configured key ID and a non-secret,
domain-separated HMAC commitment while every rollout control is dormant and
all authority-data tables are empty. It stores the four rotatable purpose
bindings, independent domain-framed login-lookup and account-mapping bindings,
the canonical aggregate over all six purposes, and an independent commitment
for the existing legacy signed-handle key. Server startup never initializes
these bindings; an absent, internally inconsistent, or mismatched binding fails
session traffic closed before an authority read or write. Login-lookup and
account-mapping keys are immutable in this milestone. Rotatable keys still
require suspended, generation- and epoch-advancing recovery followed by a
separate new-generation resume; changing the legacy signing key permanently
retires legacy issuance and acceptance because its existing handles contain no
key ID.

Never commit `.env` or credential values.

`AZURE_AI_VISION_NPM_TOKEN_BASE64` is not a runtime variable. It is used only while updating the client-side Face Liveness SDK; see the maintenance runbook below.

## Local development

Install dependencies:

```powershell
npm install
```

Start the API:

```powershell
npm start
```

`npm start` runs `server.js`, loads the production-style environment, constructs
the Microsoft Graph and Azure Face clients, and opens the HTTP listener. Do not
use it for isolated verification. Azure App Service's Windows IISNode host
requires `server.js` through an interceptor; the entry-point check therefore
accepts direct execution and IISNode's preserved `process.argv[1]` application
path while ordinary module imports remain side-effect free.

Run the isolated automated tests:

```powershell
npm test
```

The acceptance tests import `createApp` with a fixed synthetic signing key,
recording raw Graph and Face fakes, and deterministic runtime hooks. `app.js`
composes those clients into the same adapters used by production composition.
HTTP cases use ephemeral loopback listeners that are closed after each test.
Importing `app.js`, `server.js`, or any domain, integration, or shared module
does not load environment configuration, construct production clients, acquire
a Graph token, start a listener, call an external service, or schedule a timer,
and the test suite does not call production integrations.

Unless `PORT` is configured, the server listens on `http://localhost:3000`.

> **Safety:** the application has no automatic local-data isolation. With production credentials, requests may modify live Excel workbooks and send real emails. Confirm the target data and intended recipients before exercising side-effecting endpoints.

## Learning-platform row authorization

After a successful active-account login, the backend returns a four-hour signed authorization handle in the legacy `IndexVerificado` response field. The frontend treats this value as opaque and sends it back to row-scoped platform endpoints. The backend verifies the signature and expiration before deriving the workbook row index; callers cannot select a different learner by changing the value.

The approved replacement is documented separately in the
[session-authority target decision](docs/session-authority.md). Its durable
schema, backend authority, target APIs, and compatibility ledger are
implemented but dormant. No production SQL store or first-party topology has
been qualified, no activation control has been approved, and the frontend has
not adopted the target. Consequently, the signed row handle and current
four-hour client behavior remain unchanged after merge. See the
[qualification runbook](docs/runbooks/qualify-session-authority.md) for the
blocked production boundary.

`PLATFORM_ROW_AUTHORIZATION_KEY_BASE64` must be stable canonical Base64 that
decodes to exactly 32 random bytes. Generate it once through an approved
secret-management process, keep it out of output and shell history, and store
it only in secure transient process state and the Azure App Service secret
setting. Never commit it, persist it in an example file, or reuse another
application credential.

Existing unsigned sessions require a fresh login. Rotating the key invalidates
every outstanding signed handle and also requires affected learners to log in
again.

The project does not currently define lint or build scripts.

## Deployment

Pushes to `main` trigger [the GitHub Actions workflow](.github/workflows/main_plataforma-backend-v3.yml). It installs dependencies with Node.js 24 and deploys the repository artifact to the Production slot of the Azure App Service `Plataforma-Backend-v3`.

> **Deployment gate:** configure `PLATFORM_ROW_AUTHORIZATION_KEY_BASE64` in
> the App Service before merging a version that requires signed row handles.
> The application intentionally fails to start when the setting is missing or
> malformed. Changing App Service settings may restart the service, so perform
> this step in an authorized maintenance window and verify service health
> without revealing the value.

The workflow also supports manual execution through GitHub Actions. It runs the automated tests; build commands run only when a corresponding package script exists.

The workflow's path filters exclude Markdown, `docs/**`, and `test/**`, so a
merge limited to those paths does not trigger a production backend deployment.
They do not exclude production JavaScript, including `app.js`, `server.js`,
`platform-row-authorization.js`, `domains/**`, `integrations/**`, and
`shared/**`; a merge that changes any of those files triggers the production
build, test, and deployment workflow. The deployed repository artifact includes
the domain, integration, and shared modules.

## Maintenance and contributor documentation

- [Current API contract inventory](docs/api-contracts.md)
- [Approved session-authority target decision](docs/session-authority.md)
- [Qualify and activate the session authority](docs/runbooks/qualify-session-authority.md)
- [Update the Face Liveness Web SDK](docs/runbooks/update-face-liveness-sdk.md)
- [Repository collaboration guidance](AGENTS.md)

## Current technical constraints

- `app.js` is the thin application composition root. It creates import-safe
  adapters from the injected raw Graph and Face clients and passes narrowly
  named capabilities into cohesive domain handler factories.
- Domain modules own business ordering, validation, authorization, retry and
  error decisions, templates, recipients, payload construction, deduplication,
  and partial-success behavior.
- Integration modules own exact SDK paths and verbs, request envelopes, Face
  multipart mechanics, external response-shape access, and result projection.
  Each SDK-calling adapter operation performs exactly one underlying attempt;
  retries remain at their existing domain boundaries.
- Microsoft Graph integrations depend on fixed workbook, table, and positional-column contracts.
- Production operations reach live external services; automated tests isolate
  them behind injected recording fakes and deterministic runtime hooks.
- The endpoints are application contracts rather than a versioned public API.
