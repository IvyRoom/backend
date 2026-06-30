# CLAUDE.md — backend

Node.js backend serving the frontends (e.g. endpoints under `/plataforma_v2`).

<!-- ========================================================= -->
<!-- SHARED WORKING AGREEMENT — MIRRORED WITH sistemas/CLAUDE.md -->
<!-- ========================================================= -->
## Working agreement — KEEP IN SYNC across repos
> This section is mirrored **verbatim** in `sistemas/CLAUDE.md` and
> `backend/CLAUDE.md`. If it changes in one, make the identical change in the
> other in the same edit. Claude can access both repos and edits both together,
> then flags each for its own commit.

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
- **Never invent scope.** No fields, endpoints, or copy I didn't ask for.
- **Match the surrounding code** of whichever repo/folder you're in — its
  naming, language, and structure win over your defaults. Flag mismatches
  instead of silently "fixing" them.
- **If a request conflicts with a convention, say so** and propose the
  convention-following alternative.
- **Never commit secrets.** Keys, tokens, connection strings, passwords stay out
  of tracked files (use ignored config / env vars). If a change would add one,
  stop and flag it.
- **Verify before handoff.** Check what's mechanical — syntax, tests, logic —
  yourself; I own behavioural and visual testing.

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
  every commit with a `Co-Authored-By: Claude <noreply@anthropic.com>` trailer —
  a footer line after a blank line, never on the summary.
- Branches are workspaces; merging to `main` deploys. Nothing's "ready" until I
  say so.

<!-- ========================================================= -->
<!-- REPO SPECIFICS — backend only                             -->
<!-- ========================================================= -->
## Conventions — to be defined
We haven't built here together yet. Fill this in when we start, then keep it
honest about what's actually in the repo.

From what I've seen so far (your example endpoint): Express-style routes, a
Microsoft Graph API client (with retry), JSON responses, coded error strings
(e.g. `Erro_008`). **Existing backend code uses Portuguese identifiers** — match
the surrounding code when editing existing files. For brand-new code we write
together, we'll decide the naming convention explicitly before starting.
