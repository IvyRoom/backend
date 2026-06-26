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
of the code. Be assertive and concise — skip basics, don't pad. Surface
trade-offs and push back when you have a real reason; I want a collaborator, not
a yes-man.

### Build rhythm — default loop for new code
Scale it to the change: a one-line fix goes straight to a commit; the full loop
is for non-trivial new logic. Offer each step and **wait** — never run the whole
loop automatically.
1. **Write first, no comments.** Working code, no explanatory comments; lean on
   clear names. I test behaviour, not prose.
2. **When it works, offer to explain.** Once I confirm it behaves, ask if you
   should add temporary explanatory comments so I can walk your logic.
3. **Then offer to trim** to navigation-only signposts.
4. **Then flag the commit** and draft the message + description.

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

### Git — I drive, you assist
- **I handle all git myself** (commit, push, merge) in GitHub Desktop. Never run
  git commands or rewrite history.
- **One feature = one branch per repo**, same feature name across repos. Branch
  names `type/short-desc`, lowercase, hyphens.
- **Point out commit-worthy moments and draft the message + description.**
  Conventional Commits: `feat | fix | refactor | style | docs | chore`;
  imperative summary ≤ ~50 chars; body explains *why* when non-obvious.
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
