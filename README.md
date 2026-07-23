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

- Node.js 20
- Express 4
- Microsoft Graph API
- Microsoft Authentication Library (MSAL)
- Azure AI Face REST API
- Azure App Service
- GitHub Actions

The application currently uses CommonJS and runs from a single `app.js` entry point.

## Repository structure

| Path | Purpose |
|---|---|
| `app.js` | Express application, routes, and external-service integrations. |
| `img/` | Images used in emails and other backend-generated content. |
| `docs/runbooks/` | Repeatable operational and maintenance procedures. |
| `.github/workflows/` | Continuous deployment to Azure App Service. |
| `AGENTS.md` | Repository-specific collaboration and engineering guidance. |

## Prerequisites

- Node.js 20
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
| `PLATFORM_ROW_AUTHORIZATION_KEY_BASE64` | Stable 32-byte key, encoded as Base64, for signing learning-platform row authorization handles | Yes |
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

Run the isolated automated tests:

```powershell
npm test
```

Unless `PORT` is configured, the server listens on `http://localhost:3000`.

> **Safety:** the application has no automatic local-data isolation. With production credentials, requests may modify live Excel workbooks and send real emails. Confirm the target data and intended recipients before exercising side-effecting endpoints.

## Learning-platform row authorization

After a successful active-account login, the backend returns a four-hour signed authorization handle in the legacy `IndexVerificado` response field. The frontend treats this value as opaque and sends it back to row-scoped platform endpoints. The backend verifies the signature and expiration before deriving the workbook row index; callers cannot select a different learner by changing the value.

`PLATFORM_ROW_AUTHORIZATION_KEY_BASE64` must decode to exactly 32 bytes and remain stable across application instances and deployments. Generate a value locally with:

```powershell
node -e "console.log(require('node:crypto').randomBytes(32).toString('base64'))"
```

Store the generated value in ignored local configuration and in the Azure App Service application settings. Never commit it. Rotating the key or deploying this change over an existing unsigned browser session requires affected learners to log in again.

The project does not currently define lint or build scripts.

## Deployment

Pushes to `main` trigger [the GitHub Actions workflow](.github/workflows/main_plataforma-backend-v3.yml). It installs dependencies with Node.js 20 and deploys the repository artifact to the Production slot of the Azure App Service `Plataforma-Backend-v3`.

Configure `PLATFORM_ROW_AUTHORIZATION_KEY_BASE64` in the App Service before deploying a version that requires signed row handles. The application fails to start when the setting is missing or malformed.

The workflow also supports manual execution through GitHub Actions. It runs the automated tests; build commands run only when a corresponding package script exists.

The workflow currently has no path filter, so every merge to `main`—including a documentation-only merge—triggers a production backend deployment.

## Maintenance and contributor documentation

- [Update the Face Liveness Web SDK](docs/runbooks/update-face-liveness-sdk.md)
- [Repository collaboration guidance](AGENTS.md)

## Current technical constraints

- The application is implemented in one `app.js` file.
- Microsoft Graph integrations depend on fixed workbook, table, and positional-column contracts.
- Several operations reach live external services, so automated tests will require dependency isolation or test doubles.
- The endpoints are application contracts rather than a versioned public API.
