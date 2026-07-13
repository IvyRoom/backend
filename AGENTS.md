# AGENTS.md — backend

Node.js backend serving the frontends (e.g. endpoints under `/plataforma_v2`).

<!-- ========================================================= -->
<!-- SHARED WORKING AGREEMENT — MIRRORED WITH sistemas/AGENTS.md -->
<!-- ========================================================= -->
## Working agreement — KEEP IN SYNC across repos
> This section is mirrored **verbatim** in `sistemas/AGENTS.md` and
> `backend/AGENTS.md`. If it changes in one, make the identical change in the
> other in the same edit. The agent can access both repos and edits both
> together, then flags each for its own commit.

### Who you're working with
An experienced engineer who holds the full product context and stays in control
of the code. Be assertive and concise — skip basics, don't pad. Always be as
concise as the task allows while staying thorough, strategic, and robust;
brevity is the default, never an excuse to cut correctness. Surface trade-offs
and push back when you have a real reason; I want a collaborator, not a yes-man.

### Comments — default to none
Working code that leans on clear names; no explanatory or navigational comments.
We trade ideas, you implement, I eyeball and test, and once it's
production-ready I open the PR and merge — so the old staged "explain / trim"
passes are gone. Commits still flow continuously throughout (don't batch them).
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
  their ports (e.g. 3000) are free for my own runs. I still own final
  behavioural and visual testing.

### Git — you commit, I publish
- **You make the commits** (`git add` + `git commit`) on the current feature
  branch, at natural boundaries throughout the work — don't wait for me. Stage
  deliberately (named paths, never a blanket `git add -A`) so secrets and
  untracked junk can't slip in. No need to surface intermediate commits — I
  review at the Pull Request / merge level.
- **Commit my uncommitted manual edits too.** When I've hand-edited files and
  left them uncommitted, commit them as their own commit, with a summary and
  description you infer from the diff — don't fold them into your own work.
- **I handle everything that leaves my machine or rewrites shared history**:
  Publish Branch / Push to Origin / Pull Requests / merge, all in GitHub
  Desktop. Never push, never open or merge PRs, never rewrite history (no
  amend, rebase, force-push, or `reset --hard`).
- **Stay on the feature branch; never commit to `main`.** One feature = one
  branch per repo, same feature name across repos. Branch names
  `type/short-desc`, lowercase, hyphens. Starting a new feature while on `main`
  with no branch yet: create and name it yourself — no need to ask — then tell me.
- Conventional Commits: `feat | fix | refactor | style | docs | chore`;
  imperative summary ≤ ~50 chars; body explains *why* when non-obvious. End
  every commit with a `Co-Authored-By:` trailer naming the model that wrote it
  (e.g. `Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>`) — a footer
  line after a blank line, never on the summary.
- Branches are workspaces; merging to `main` deploys. Nothing's "ready" until I
  say so.

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
- WhatsApp payload must match `+55 XX XXXXX-XXXX` (mirrors the frontend mask,
  pinned to Brazilian numbers); anything else is `Erro_014`.

### processa-formulario design notes
- Whole-form retry is safe by design: new rows are deduped against existing
  table data (BD Plataforma by e-mail, BD Clientes by CPF), so a
  failed-then-retried submission doesn't duplicate rows.
- The two `rows/add` calls are deliberately not wrapped in `retry()`: an
  ambiguous failure after a successful insert would double-insert.
- Cells left `null` in new rows are columns the sheet itself fills (formulas /
  manual entry) — don't populate them from code.
