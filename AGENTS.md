# AGENTS.md — backend

Node.js backend serving the frontends (e.g. endpoints under `/plataforma_v2`).

<!-- ========================================================= -->
<!-- SHARED WORKING AGREEMENT — MIRRORED WITH sistemas/AGENTS.md -->
<!-- ========================================================= -->
## Working agreement — KEEP IN SYNC across repos
> This section is mirrored **verbatim** in `sistemas/AGENTS.md` and
> `backend/AGENTS.md`. If it changes in one, make the identical change in the
> other in the same edit. The agent can access both repos, edit both together,
> and commit each repo separately.

### Who you're working with
Lucas Machado is an experienced founder, business/process operator, and
management specialist who is deliberately building formal software-engineering
practice. He graduated first in his mechanical-engineering class, studied at
Cornell, worked in management consulting and automotive process coordination,
and founded Machado | Método Gerencial para Empresas. He taught himself coding,
GitHub, and Azure while building the company and holds the full product and
business context.

Lucas is fluent in English and prefers technical collaboration, source names,
identifiers, and the rare necessary code comment in English; user-facing copy
remains Brazilian Portuguese. Work with him as an expert product owner and a
fast software-engineering learner:

- Lead with the outcome, then explain the smallest useful mental model.
- Define unfamiliar software terms before relying on them. Never confuse an
  unfamiliar term with limited reasoning ability.
- Recommend genuine best practice and explain the trade-offs; don't silently
  preserve an existing convention merely because it already exists.
- Break structural work into small, independently reviewable tasks and confirm
  the shared reasoning at meaningful boundaries.
- Be assertive and concise without skipping fundamentals. Surface mistakes and
  risks directly; brevity is never an excuse to cut correctness.

### Comments — default to none
Working code that leans on clear names; no explanatory or navigational comments.
We trade ideas; you implement, verify, publish the branch, and open the PR; I
review, test when useful, and merge. The old staged "explain / trim" passes are
gone. Commits still flow continuously throughout (don't batch them).
Narrow exception: a single line is fine when it captures what a name can't — a
non-obvious *why*, a security-critical invariant, a browser/API quirk, or a
documented contract (e.g. an HTML↔JS interface). When editing a file that's
already commented (e.g. backend `app.js`), match its existing style.

### Session hygiene
A long session grows slower, costlier, and less sharp — details get buried in a
big context. At a task/repo boundary, or when the thread is clearly long, flag
that a fresh session would help and write a short handoff (state, decisions,
open threads, next steps) so the new one starts oriented.

### Golden rules
- **Ask before large or structural changes.** Propose, wait for my OK. Small,
  obvious fixes: just do them.
- **One concern per change.** No unrelated refactors in passing.
- **Never invent scope.** No fields or endpoints I didn't ask for. When a
  change genuinely requires new user-facing copy (labels, messages), write it
  to fit the surrounding tone and language, and list it in your handoff so I
  can review the wording.
- **Match the surrounding code** of whichever repo/folder you're in — its
  naming, language, and structure win over your defaults. Flag mismatches
  instead of silently "fixing" them.
- **Keep names true to every use.** When reusing a token, helper, or abstraction
  for a new role, verify its name still describes all uses. Prefer one neutral,
  accurate name over a role-specific name used out of context or duplicate
  aliases for the same value.
- **If a request conflicts with a convention, say so** and propose the
  convention-following alternative.
- **Never commit secrets.** Keys, tokens, connection strings, passwords stay out
  of tracked files (use ignored config / env vars). If a change would add one,
  stop and flag it.
- **Verify before handoff.** Check what's mechanical — syntax, tests, logic —
  yourself. When it adds real signal, also exercise the change yourself in a
  local preview (serve the frontend, drive it in a browser). Before running
  anything, map what it touches: never exercise paths that reach production —
  Graph API, live spreadsheets, real e-mail — or anything else with side
  effects beyond this machine, without my explicit OK. Standing exception:
  **read-only** Graph reads of our workbooks are pre-approved — always verify
  a sheet's real schema (columns, table GUID, AUXILIAR-style lists) by reading
  it before writing endpoint code against it; writes and e-mails stay gated.
  When the task wraps, stop any local preview/stub servers you started so
  their ports (e.g. 3000) are free for my own runs. For interaction features,
  verify the human experience, not only DOM state: where the viewport lands
  after a click, what gains focus, and whether content people need to copy can
  actually be copied (through selection or a copy control, including success
  feedback and a usable failure fallback) — at desktop and mobile widths. I
  still own final behavioural and visual testing.
- **Keep permission approvals agent-specific.** When a command prompts and I
  approve it, prefer a reusable, narrowly scoped rule in the active agent's
  own permission system when supported. Never allowlist what the deny floor
  forbids (merge / rebase / amend / force-push / hard reset) or anything with
  side effects beyond this machine. A normal push of an agreed, verified
  feature branch is allowed as part of the publishing workflow below.

### Git — you publish, I merge
- **You own feature-branch implementation and publication.** Make commits
  (`git add` + `git commit`) at natural boundaries throughout the work — don't
  wait for me. Stage deliberately (named paths, never a blanket `git add -A`)
  so secrets and untracked junk can't slip in. No need to surface intermediate
  commits — I review at the Pull Request / merge level.
- **Commit my uncommitted manual edits too.** When I've hand-edited files and
  left them uncommitted, commit them as their own commit, with a summary and
  description you infer from the diff — don't fold them into your own work.
- **Before publishing, self-review the complete diff and run the relevant
  checks.** Once the agreed scope is complete and the worktree is clean, push
  the current feature branch normally (never force-push) and open a Pull
  Request targeting `main`. Summarize what changed and why, verification,
  risks or deployment impact, and any related PR in the other repo.
- **Open a ready-for-review PR when the work is complete and verified.** Use a
  draft only for intentionally incomplete work, early architectural feedback,
  or known failing checks. Never merge or enable auto-merge; I own the final
  merge decision.
- **Before merge, correct the same PR instead of reverting.** If I request
  changes, add correction commits to the same feature branch, push them, let
  checks rerun, and ask me to review the updated diff. Do not rewrite published
  history. If we abandon the approach, close the PR without merging. Reverting
  is for changes already merged to `main`, normally through a new revert PR.
- **After I merge, verify before cleanup.** Confirm the PR is merged and the
  resulting `main` CI/deployment completed successfully; perform only safe,
  proportionate smoke checks with no production writes or messages. If
  verification fails, keep the branch and task context intact and diagnose it.
- **Clean up only after successful verification.** Require a clean worktree;
  fetch/prune `origin`; switch to `main`; pull with `--ff-only`; verify local
  `main` matches `origin/main`; then delete the local branch with
  `git branch -d`. GitHub deleting the remote branch automatically at merge is
  an accepted exception; otherwise delete it manually only after successful
  verification and confirmation that its PR is merged or closed. If `main` is
  dirty or diverged, stop instead of overwriting anything. Never use
  `git branch -D`, amend, rebase, force-push, or `reset --hard`.
- **Stay on the feature branch; never commit to `main`.** One feature = one
  branch per repo, same feature name across repos. Branch names
  `type/short-desc`, lowercase, hyphens. Starting a new feature while on `main`
  with no branch yet: create and name it yourself — no need to ask — then tell me.
- Conventional Commits: `feat | fix | refactor | style | docs | chore`;
  imperative summary ≤ ~50 chars; body explains *why* when non-obvious. End
  every commit with a `Co-Authored-By:` trailer naming the agent/model that
  wrote it, using the matching provider identity — a footer line after a blank
  line, never on the summary.
- Branches are workspaces; merging to `main` deploys. A ready PR means your
  implementation is complete, not that it is approved; only I decide whether
  to merge.

<!-- ========================================================= -->
<!-- REPO SPECIFICS — backend only                             -->
<!-- ========================================================= -->
## Conventions — current state
Single `app.js` Express app; banner comments split it into sections. Data lives
in Excel workbooks reached through the Microsoft Graph API — tables addressed
by drive-item ID + table GUID, columns by numeric index. The column maps exist
only in the sheets, so verify indexes by reading the sheet (read-only Graph
reads are pre-approved — see the working agreement) before writing code
against it. Legacy
sections use Portuguese identifiers; the `formulario` endpoint
(`/clientes/processa-formulario`) uses English camelCase — match the section
you're editing.

### Error codes — user-visible contract with `sistemas` frontends
Canonical registry — moved here from the old dictionary at the top of
`sistemas/plataforma_v2/login/main.js`. New code = next free number, plus a
message in whichever frontend consumes it (e.g. `SUBMIT_ERROR_MESSAGES` in
`formulario/main.js` and `conecta/main.js`). `Erro_000` and `Erro_006` are
emitted by the frontends themselves, never by the backend.

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

### conecta (processa-recomendacao) design notes
- The recommender is identified by matching URL-borne name + company against
  BD - RECOMENDAÇÕES (normalized: trim, collapsed spaces, lowercase). No match
  = `Erro_016`, so a tampered or mistyped link cannot write anything.
- Fill-or-append: a recommender row whose recommendation columns are all `-`
  is a free slot left by the manual invite process — fill it; otherwise append
  a full new row copying the recommender columns from the matched row.
- An identical pending recommendation (same recommender + company +
  professional + WhatsApp) skips the write but still sends the e-mails, so a
  retry after a failed sendMail stays safe.
- Writes are deliberately not `retry()`-wrapped (same double-insert rationale
  as processa-formulario).
- `RECOMENDACOES_COLUMNS` verified against the sheet on 13/Jul/2026 (13
  columns, table `BD`). `PRIMEIRO NOME` is a calculated column — leave it
  `null` on `rows/add` so the table formula fills it.
- New-recommendation defaults: `DATA E HORA` / `DATA E HORA ATUALIZAÇÃO` /
  `DATA E HORA PRÓXIMO CONTATO` = now as an **Excel datetime serial**
  (America/Sao_Paulo); the display is the sheet's number formatting.
  `ETAPA` = `1. REALIZAR CONTATO INICIAL`, `STATUS` = `A INICIAR`,
  `NÚMERO PARTICIPANTES` stays `-`. Stage/status strings must mirror the
  sheet's AUXILIAR tab lists — renaming there requires updating the constants
  in `app.js`.
- WhatsApp payload must match `+XX XX XXXXX-XXXX` (mirrors the frontend mask);
  anything else is `Erro_014`.

### processa-formulario design notes
- Whole-form retry is safe by design: new rows are deduped against existing
  table data (BD Plataforma by e-mail, BD Clientes by CPF), so a
  failed-then-retried submission doesn't duplicate rows.
- The two `rows/add` calls are deliberately not wrapped in `retry()`: an
  ambiguous failure after a successful insert would double-insert.
- Cells left `null` in new rows are columns the sheet itself fills (formulas /
  manual entry) — don't populate them from code.
