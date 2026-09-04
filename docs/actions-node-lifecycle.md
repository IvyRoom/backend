# Actions and Node lifecycle

Reviewed 2026-09-04. Owner: Lucas (`IvyRoom`). Review monthly (next:
2026-10-01), on action-update PRs, security advisories, runner warnings, and
Azure lifecycle notices. This record does not change cloud configuration.

## Separate runtime layers

| Layer | Reviewed baseline and evidence |
| --- | --- |
| Application declaration | `package.json` and the root lockfile entry require Node `24.x`; this is not an Azure-host observation. |
| Repository install/test | Deployment, bot validation, and Face-monitor checking explicitly select `24.x`. All setup-node package-manager caching is disabled. |
| JavaScript actions | The six reviewed actions below declare `node24` internally. `node-version` does not select this runtime. |
| GitHub runner | GitHub-hosted `windows-latest` builds/tests; `ubuntu-latest` deploys/checks Face metadata. Read **Set up job** for the actual runner/image. |
| Azure application runtime | Unverified by this work. Existing App Service Production target, Windows IISNode startup, OIDC settings, and integrations are unchanged. |
| Retired Node 22 Function | `Plataforma-Function-v2` and its exclusive supporting resources were deleted with explicit owner approval on 2026-09-04. The verified scope and retained resources are recorded below. |

Node 24 actions require runner **2.327.1+**. Checkout's separate **2.329.0+**
Docker authenticated-Git requirement is not relied on with persisted
credentials disabled. See [checkout requirements](https://github.com/actions/checkout/blob/3d3c42e5aac5ba805825da76410c181273ba90b1/README.md)
and [GitHub-hosted runner policy](https://docs.github.com/en/actions/concepts/runners/github-hosted-runners).
Any future self-hosted move must also satisfy [ongoing runner updates](https://github.blog/changelog/2026-06-12-github-actions-minimum-version-enforcement-timeline-for-self-hosted-runners/),
not just an old compatibility floor. Azure login's [OIDC implementation](https://github.com/Azure/login/blob/7ddb5af1ef8758cf1353cf3b42f940aee27ba21c/README.md)
requires Azure CLI 2.30+; OIDC permission belongs only to the deploy job.

## Reviewed action provenance

Full upstream commit pins and readable release comments live in
[the workflows](../.github/workflows/). Each existing major tag resolved to
the release commit below on the review date; no action major was upgraded.

| Action | Reviewed release | Verified commit |
| --- | --- | --- |
| `actions/checkout` | [v7.0.1](https://github.com/actions/checkout/releases/tag/v7.0.1) | `3d3c42e5aac5ba805825da76410c181273ba90b1` |
| `actions/setup-node` | [v7.0.0](https://github.com/actions/setup-node/releases/tag/v7.0.0) | `820762786026740c76f36085b0efc47a31fe5020` |
| `actions/upload-artifact` | [v7.0.1](https://github.com/actions/upload-artifact/releases/tag/v7.0.1) | `043fb46d1a93c77aae656e7c1c64a875d1fc6a0a` |
| `actions/download-artifact` | [v8.0.1](https://github.com/actions/download-artifact/releases/tag/v8.0.1) | `3e5f45b2cfb9172054b4087a40e8e0b5a5461e7c` |
| `azure/login` | [v3.0.2](https://github.com/Azure/login/releases/tag/v3.0.2) | `7ddb5af1ef8758cf1353cf3b42f940aee27ba21c` |
| `azure/webapps-deploy` | [v3.0.8](https://github.com/Azure/webapps-deploy/releases/tag/v3.0.8) | `02a81bead70021f5284939794bcec79c271ab383` |

Weekly [Dependabot](../.github/dependabot.yml) retains supported
[SHA-pin version updates and release comments](https://docs.github.com/en/code-security/reference/supply-chain-security/supported-ecosystems-and-repositories#github-actions).
GitHub separately documents [Actions vulnerability alerts](https://docs.github.com/en/actions/how-tos/write-workflows/choose-what-workflows-do/find-and-customize-actions#using-release-management-for-your-custom-actions)
as requiring semantic-version references; inspect upstream advisories as well.
Pinned action source does not freeze hosted images, Node `24.x` patches, Azure
CLI, or Azure-side tooling. In the companion Sistemas repository, SWA's pinned
wrapper still uses a mutable [Docker `stable` image](https://github.com/Azure/static-web-apps-deploy/blob/1a947af9992250f3bc2e68ad0754c0b0c11566c9/Dockerfile).

## Lifecycle and upgrade triggers

Dates from the live official [Node release schedule](https://raw.githubusercontent.com/nodejs/Release/main/schedule.json),
verified 2026-09-04; future dates can change.

| Node | Status at review | Active LTS starts | Maintenance starts | End of life |
| --- | --- | --- | --- | --- |
| 22 | Maintenance LTS | 2024-10-29 | 2025-10-21 | 2027-04-30 |
| 24 | Active LTS | 2025-10-28 | 2026-10-20 | 2028-04-30 |
| 26 | Current | 2026-10-28 | 2027-10-20 | 2029-04-30 |

Keep **24.x**. Consider Node 26 only after LTS promotion and explicit
application, locked-dependency, IISNode, CI, and host compatibility qualification.
Start replacement-major qualification by 2027-10-30 (six months before Node 24
EOL), earlier if an advisory, dependency, action, or hosting deadline requires it.

GitHub's updated [Node 20 removal schedule](https://github.blog/changelog/2025-09-19-deprecation-of-node-20-on-github-actions-runners/)
dates the Node 24 default rollout to 2026-06-16 and Node 20 removal to 2026-09-23.
These actions already declare Node 24; no forced-runtime or insecure-runtime
escape-hatch environment variable is needed.

## Workflow controls and acceptance

- Root permissions are empty. Build/validation/check jobs have only contents
  read; deploy alone has OIDC write; the Face notifier alone has issue write.
  Checkout credentials and package-manager caching are disabled.
- Deployment's build and deploy jobs share a workflow-wide non-cancelling lock,
  including manual runs. Bot validation cancels obsolete runs within its PR.
  Every job has a timeout. GitHub's [default queue](https://docs.github.com/en/actions/how-tos/write-workflows/choose-when-workflows-run/control-workflow-concurrency)
  replaces older pending work and does not guarantee dispatch order; do not
  rerun obsolete production events.
- Deployment uses locked `npm ci`, separate optional-build and mandatory-test
  steps. Separate steps prevent a later PowerShell command from masking an
  earlier failure. No package or lockfile dependency content changed.
- Initial qualification: Node 24.19.0 / npm 11.6.2 installed all 118 locked
  packages, then passed 140 tests with the existing production-network denial.
  Lock and installed-tree inspection found no install/preinstall/postinstall
  hooks, local/git dependencies, or native binding files. No production server
  or integration was started.
- Human PRs still skip the bot-only **Validate dependency update** job. That
  required check's skip is not test execution. Attach local locked-install,
  [syntax and network-denied test](../AGENTS.md#verification), Bash syntax, and
  diff-check evidence to every human PR; use `actionlint` when available.
- Triggers and path filters are unchanged. The deployment YAML, Face monitor,
  docs, and tests are ignored, but `dependabot-validation.yml` is **not**.
  This hardening changes that file, so the full PR **deploys on merge**. Do not
  manually dispatch production simply to obtain a PR check. After merge,
  verify resulting `main` build/test/deployment before branch cleanup.
- Face monitor schedules, main-only guards, exact-title deduplication,
  metadata-only checks, notifier-only authority, and bot isolation are unchanged.

## Retired Node 22 Function

Retired 2026-09-04 at 19:30 UTC with Lucas's explicit resource-level approval.
This was cloud retirement only; no Backend or Sistemas application, deployment
configuration, data, route, or integration was changed.

### Identity, ownership, and former responsibility

- Tenant `49342d16-0605-4267-b540-d1fe7756dbac`, subscription
  `1a2f6756-eaa5-4654-bc88-a69e5e588846`, resource group `Plataforma_v2`.
- Function App resource ID:
  `/subscriptions/1a2f6756-eaa5-4654-bc88-a69e5e588846/resourceGroups/Plataforma_v2/providers/Microsoft.Web/sites/Plataforma-Function-v2`.
- The app was Linux in Brazil South, on Y1 Consumption, with `node|22`.
- Lucas confirmed ownership and no current purpose. Only the disabled
  every-ten-minute `function01` timer remained deployed; `function02` was
  absent, with stale metadata from its former daily schedule. No queue, HTTP,
  or event trigger, trigger backlog, or current caller was found in inspected
  cloud configuration and repositories.
- Recovery source remains at `IvyRoom/functions` main
  `75b0abb308f3bd5f8b175b03ba85a6788d17df09`; its last deployment run was
  `23959179347` on 2026-04-03. It has no meaningful automated tests or tested
  reconstruction procedure.
- Backend history ties `function01.js` to the organic-performance endpoint
  removed 2026-05-01 and `function02.js` to the campaign-performance endpoint
  removed 2026-04-03. Neither Backend nor Sistemas has a current caller.

### Evidence and decision

- Aggregate Azure Monitor and Application Insights metrics covered 93 days,
  2026-06-03 through 2026-09-04. Executions, requests, failures, exceptions,
  and dependencies were all zero.
- Storage showed no queues. Eight operations were metadata reads only; the
  observed objects were Function host containers, its content file share, and
  diagnostics tables. No payload logs or business data were read.
- Cost Management returned HTTP 429, so no dollar saving is claimed. Potential
  cost concerned app/plan execution, storage, and telemetry; retained shared
  resources can continue to incur cost.
- Lucas selected direct final deletion instead of a stop-and-observe window
  after the recovery limitation was presented. No stop/start rollback is
  promised; recovery requires explicit recreation, authentication, deployment,
  and end-to-end testing from the retained source.

### Deleted and retained scope

All seven top-level deletions used the subscription/resource-group prefix above:

| Type | Name |
| --- | --- |
| `Microsoft.Web/sites` | `Plataforma-Function-v2` |
| `Microsoft.Web/serverfarms` | `ASP-Plataformav2-8bd1` |
| `Microsoft.Storage/storageAccounts` | `auxiliarfunctionv2` |
| `Microsoft.Insights/components` | `Plataforma-Function-v2` |
| `Microsoft.ManagedIdentity/userAssignedIdentities` | `Plataforma-Funct-id-a8c1` |
| `Microsoft.AlertsManagement/smartDetectorAlertRules` | `Failure Anomalies - Plataforma-Function-v2` |
| `Microsoft.AlertsManagement/smartDetectorAlertRules` | `Failure Anomalies - Plataforma-Function-v1` |

The plan hosted only this app; storage held only its hosting/diagnostics data;
and the identity, child federated credential `erpoxhbb4prlw`, and role assignment
`40496234-37b3-545f-8ff6-800430f93824` served its deployment. The v2 alert
scoped the deleted Function component; the v1 alert was orphaned against an
already absent component. Deleting the identity removed its child federated
credential; the app-scoped role assignment disappeared with the Function App.
The retained legacy workflow now lacks Azure deployment authentication; stale
OIDC trust was not repaired or repurposed.

Subscription inventory changed from 26 to 19: exactly those seven resources
were removed, none was added, and retained sanitized metadata was unchanged.
Final verification found no Function Apps; all seven IDs were absent, the
identity/FIC returned resource/parent-not-found, and its principal had no role
assignments. Backend monitoring resources remained healthy; Backend and
Sistemas app metadata matched their baselines.

The shared Log Analytics workspace
`DefaultWorkspace-1a2f6756-eaa5-4654-bc88-a69e5e588846-CQ` in
`DefaultResourceGroup-CQ`, shared Smart Detection action group, and all Backend,
Sistemas, DNS, Face, video, and DRM resources remain. Historical telemetry
remains under the workspace's 30-day retention; no purge was performed. The
`IvyRoom/functions` source and workflow history also remain.

Backend before/after verification matched: deployment run `33905055294` stayed
successful at `9d616fd42569d7c9ca16f110c3bd916b546195a3`; the signature image
returned `200 image/jpeg` with 235,139 bytes, and an unregistered path retained
the expected `404 text/html`. Sistemas Static Web App state likewise stayed
unchanged. No production integration endpoint was invoked.

Azure's [Functions runtime table](https://learn.microsoft.com/en-us/azure/azure-functions/functions-versions)
lists 2027-04-30 as Node 22's expected end of support and identifies Node 22 as
the final line for Linux Consumption. This Azure policy is distinct from
upstream Node's [independent schedule](https://raw.githubusercontent.com/nodejs/Release/main/schedule.json),
which currently has the same date. [Linux Consumption retirement](https://learn.microsoft.com/en-us/azure/azure-functions/consumption-plan)
is separately scheduled for 2028-09-30.

This documentation-only change is filtered from Backend production deployment
and does not modify Azure. The milestone remains pending until both repository
records merge and required branch and preview cleanup complete.
