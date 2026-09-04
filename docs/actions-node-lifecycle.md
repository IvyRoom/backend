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
| Dormant Node 22 Function | Inventory and retirement belong to the next milestone, **Audit and retire dormant Node 22 Function**. No host inspection or change is claimed here. |

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
