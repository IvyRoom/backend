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
| Certificate validation | Verifies public certificate IDs against learning-platform records. |
| Access operations | Sends learning-platform access instructions. |
| DRM support | Returns the PlayReady authorization parameters required by the frontend. |

## Technology stack

- Node.js 24
- Express 4
- Microsoft Graph API
- Microsoft Authentication Library (MSAL)
- Azure AI Face REST API
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
| `domains/quote-requests.js` | Quote-request handler factory, notification template, retry boundary, and error mapping. |
| `domains/conecta-recommendations.js` | Conecta validation, workbook payloads, notification templates, business ordering, retry boundaries, and error mapping. |
| `domains/client-onboarding.js` | Client-intake and access-release handler factories, workbook payloads, email templates, scheduling, retry boundaries, and error mapping. |
| `domains/learning-platform.js` | Learning-platform login, Face workflow, refresh, update, feedback, and status-report business behavior, retry boundaries, and error mapping. |
| `domains/drm.js` | Deterministic PlayReady authorization-output handler factory. |
| `domains/certificate-validation.js` | Public certificate-validation handler factory, retry boundary, and score normalization. |
| `integrations/microsoft-graph.js` | Import-safe Microsoft Graph adapter for exact API paths, verbs, request envelopes, and response-shape access. Each SDK-calling operation makes one attempt. |
| `integrations/azure-face.js` | Import-safe Azure Face adapter for exact endpoint paths, multipart mechanics, and result projection. Each SDK-calling operation makes one attempt. |
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
[session-authority target decision](docs/session-authority.md). It is not
implemented: the signed row handle, current route behavior, and current risks
remain exactly as described here and in the API inventory.

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
