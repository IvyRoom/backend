# AGENTS.md — backend

Node.js service for the website, client onboarding, Conecta, learning platform,
certificate validation, and their external integrations.

<!-- ========================================================= -->
<!-- SHARED WORKING AGREEMENT — KEEP BYTE-IDENTICAL             -->
<!-- ========================================================= -->
## Working agreement — keep byte-identical across repositories

Keep this entire block byte-for-byte identical in both root `AGENTS.md` files.
Change it in both repositories in the same task, but commit and publish each
repository separately.

### Scope and safety
- Ask before large or structural changes; small, obvious fixes may proceed. Keep
  one concern per change, and do not invent fields, endpoints, dependencies,
  copy, or unrelated refactors. When new user-facing copy is required, match its
  surrounding language and tone and flag it for review.
- Match local naming, language, structure, and conventions. Keep reused names
  accurate for every use. Explain conflicts and propose the convention-following
  alternative instead of silently departing from local practice.
- Prefer clear code to commentary. Comment only a non-obvious reason, security
  invariant, external quirk, or contract that naming cannot express, and match
  the file's existing style.
- Never commit secrets. Use ignored configuration or environment variables, and
  stop if a requested change would expose a credential.
- Before running code, identify its effects. Do not write production data, send
  email or other messages, or exercise a side-effecting external integration
  without explicit approval. A repository-specific read-only exception never
  authorizes writes.
- Verify syntax, tests, logic, and any safe local behavior that adds useful
  signal. Stop local servers you start. Keep approvals narrow, agent-specific,
  and limited to the agreed operation; approval never expands scope or permits
  prohibited Git operations or external side effects.

### Git and publication
- The agent owns feature-branch implementation, verification, commits, normal
  push, and a ready-for-review PR; the repository owner alone merges. Never
  commit on `main`, merge, or enable auto-merge.
- Use one lowercase, hyphenated `type/short-desc` branch per feature and
  repository, with the same feature name for cross-repository work. If work
  starts on `main`, create the branch and report it.
- Commit at natural boundaries. Stage named paths, never `git add -A`. Preserve
  pre-existing user edits; when they are in scope for publication, commit them
  separately rather than folding them into agent work.
- Use Conventional Commits (`feat | fix | refactor | style | docs | chore`), an
  imperative summary of about 50 characters or less, and a body when the reason
  is not obvious. End each commit with a matching-provider `Co-Authored-By:`
  trailer after a blank line.
- Before publishing, self-review the complete diff, run relevant checks, and
  require a clean worktree. Push normally, never force-push, open a PR targeting
  `main`, cross-link any companion PR, and report purpose, verification, risk,
  and deployment effect. Use a draft only for intentionally incomplete or
  failing work.
- Before requesting merge, give one concise briefing covering why, what, how,
  verification, deployment risk, and the decision needed. Correct feedback on
  the same branch and PR with new commits; do not rewrite published history. If
  abandoned, close without merging; reserve revert PRs for changes already
  merged.
- Treat a merge to `main` as production-affecting unless the repository's
  authoritative deployment rules prove the scoped change is filtered out. A
  ready PR is complete work, not merge approval.
- After a reported merge, confirm the PR and resulting `main` CI/deployment
  succeeded before cleanup or new work. Use only safe, proportionate smoke
  checks with no production writes or messages. On failure, preserve the branch
  and context and diagnose.
- Successful post-merge cleanup is mandatory and pre-authorized for the merged
  feature branch: require a clean worktree; fetch/prune `origin`; switch to
  `main`; pull with `--ff-only`; verify local `main` equals `origin/main`;
  delete the local branch with `git branch -d`; if it still exists, delete the
  remote branch only after confirming its PR merged or closed; then verify only
  `main` and active branches remain. Stop on a dirty or diverged `main` or any
  failed prerequisite.
- Never use `git branch -D`, amend, rebase, force-push, or `reset --hard`.

<!-- ========================================================= -->
<!-- REPOSITORY-SPECIFIC GUIDANCE — backend                     -->
<!-- ========================================================= -->
## Runtime and deployment

- `server.js` is the production entry point. It owns environment loading,
  production Graph and Face client construction, listener startup, and the
  Graph token lifecycle. `app.js` exports the import-safe
  `createApp(dependencies)` factory and remains the thin composition root for
  Express construction, exact global middleware order, handler-factory
  composition, raw-client adapter construction, and explicit ordered route
  registration.
- Modules under `domains/` are import-safe handler factories that own their
  domain-specific helpers, constants, templates, payload construction, and
  business, retry, and error decisions. Modules under `integrations/` are
  import-safe thin adapters that own exact SDK paths and verbs, request
  envelopes, Face multipart mechanics, external response-shape access, and
  result projection. Each SDK-calling adapter operation makes exactly one
  underlying attempt; keep retry ownership in the domains. Keep genuinely
  shared behavior in `shared/`, match local identifier language and comment
  style, and keep each concern with its narrowest owning domain.
- Production startup must work both through direct Node execution and Azure
  App Service's Windows IISNode interceptor while ordinary imports of every
  production module stay safe.
- Use `README.md` for current setup and integration orientation. Inspect
  `.github/workflows/main_plataforma-backend-v3.yml` before predicting whether
  a scoped change triggers deployment.
- Preserve existing API payloads, status codes, routes, workbook layouts, and
  recipients unless the task explicitly changes them.

## Production integrations and live schema

- The service has no automatic local-data isolation. Before starting it or
  exercising a route, map every Graph, workbook, Azure Face, and email side
  effect. Graph writes, workbook mutations, Face operations, and real email
  require explicit approval.
- Read-only Microsoft Graph inspection of the workbooks is pre-approved only for
  schema verification. It does not authorize a write, email, Face call, or
  broader production read.
- Workbook contracts are positional: drive-item and table identity plus numeric
  array indexes. Before changing dependent code, read the live workbook and
  verify the table GUID/name, column order and count, calculated and manual
  columns, and relevant `AUXILIAR` values. Comments and dated snapshots are not
  schema guarantees.
- Preserve positional widths and `null` placeholders for cells populated by
  formulas or manual workflow. Do not retry a non-idempotent write when an
  ambiguous success could duplicate data; retain the endpoint-specific rules
  below.

## Error registry

This is the backend/frontend `Erro_XXX` contract. Allocate the next free number
for a new backend error and add its Portuguese message in every consumer. Current
consumer locations include `../sistemas/apps/quote-request/main.js`,
`../sistemas/apps/client-intake/main.js`,
`../sistemas/apps/certificate-validation/main.js`,
`../sistemas/apps/conecta/referral-form/main.js`, and
`../sistemas/plataforma_v2/`. `Erro_000` and `Erro_006` are emitted only by
frontends.

- `Erro_000` — frontend fallback: network/unknown failure reaching the backend
- `Erro_001` — read BD Plataforma
- `Erro_002` — upload FotoReferência to OneDrive
- `Erro_003` — flag FotoReferência as registered in BD Plataforma
- `Erro_004` — create Azure Face liveness session (authToken/sessionID)
- `Erro_005` — read FotoReferência from OneDrive
- `Erro_006` — frontend: FaceLivenessDetector failed to run
- `Erro_007` — read Azure Face liveness session results
- `Erro_008` — write BD Plataforma
- `Erro_009` — write BD Feedbacks
- `Erro_010` — write BD Clientes
- `Erro_011` — read BD Clientes
- `Erro_012` — formulario sendMail
- `Erro_013` — invalid formulario payload (400)
- `Erro_014` — invalid conecta payload (400)
- `Erro_015` — read BD Recomendações
- `Erro_016` — recomendante not found in BD Recomendações (404)
- `Erro_017` — write BD Recomendações
- `Erro_018` — conecta sendMail

### plataforma_v2 row authorization invariant
- `IndexVerificado` is a legacy wire-field name. Its value is a signed,
  four-hour authorization handle, never a raw workbook row index.
- Mint handles only after valid credentials for an active account. Treat the
  frontend value as untrusted and derive `platformRowIndex` only through the
  verifier before any row-scoped Graph, photo-path, workbook-row, array, or
  Face access.
- `CadastroFoto_e_FaceID`, `FaceID`, `refresh`, `updates`, and
  `processa-feedback` must all use the verified index. Never interpolate the
  request value directly into a Graph path, file path, or array lookup.
- `PLATFORM_ROW_AUTHORIZATION_KEY_BASE64` is a required, stable 32-byte key
  stored in ignored local configuration and Azure App Service settings. Never
  commit it or substitute another application credential.

### conecta (processa-recomendacao) design notes
- The recommender is identified by matching URL-borne name + company against
  BD - RECOMENDAÇÕES (normalized: trim, collapsed spaces, lowercase). No match
  = `Erro_016`, so a tampered or mistyped link cannot write anything.
- Fill-or-append: a recommender row whose recommendation columns are all `-`
  is a free slot left by the manual invite process—fill it; otherwise append a
  full row copying the recommender columns from the matched row.
- An identical pending recommendation (same recommender + benefited company +
  recommended company + professional + WhatsApp) skips the write but still
  sends the emails, so a retry after a failed sendMail remains safe.
- Update and append writes are deliberately not `retry()`-wrapped; an ambiguous
  failure after a successful mutation could duplicate the recommendation.
- `RECOMENDACOES_COLUMNS` is a 13-position contract. Reverify it against the
  live table before dependent changes. `PRIMEIRO NOME` is calculated, so pass
  `null` on `rows/add` and let the table formula populate it.
- New rows use an Excel datetime serial for São Paulo “now” in `DATA E HORA`,
  `DATA E HORA ATUALIZAÇÃO`, and `DATA E HORA PRÓXIMO CONTATO`; workbook number
  formatting owns their display. Set `ETAPA` to
  `1. REALIZAR CONTATO INICIAL`, `STATUS` to `A INICIAR`, and
  `NÚMERO PARTICIPANTES` to `-`; reverify those values against live `AUXILIAR`
  lists before changing them.
- WhatsApp must match `+XX XX XXXXX-XXXX`; invalid input returns `Erro_014`.

### processa-formulario design notes
- Whole-form retry is safe because new rows are deduplicated against current
  workbook data: BD Plataforma by email and BD Clientes by CPF.
- The two `rows/add` calls are deliberately not `retry()`-wrapped; an
  ambiguous failure after a successful insert could duplicate a row.
- Keep `null` in cells filled by workbook formulas or manual workflow.

## Verification

Do not start the production server for documentation, syntax, or automated
acceptance checks. HTTP tests must use `createApp` with synthetic dependencies
and ephemeral loopback listeners, then close every listener and timer. Run:

```powershell
node --check app.js
node --check server.js
node --check platform-row-authorization.js
node --check domains/quote-requests.js
node --check domains/conecta-recommendations.js
node --check domains/client-onboarding.js
node --check domains/learning-platform.js
node --check domains/drm.js
node --check domains/certificate-validation.js
node --check integrations/microsoft-graph.js
node --check integrations/azure-face.js
node --check shared/retry.js
node --check shared/escape-html.js
npm test
git diff --check
```
