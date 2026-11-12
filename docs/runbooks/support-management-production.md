# Support management production

The frontend is built from `sistemas/apps/support-management/` into Portuguese
`dist/atendimentos/` paths. Customer URLs stay on the existing domain:

- `https://machadogestao.com/atendimentos/login/`
- `https://machadogestao.com/atendimentos/agendamento/`

Both scheduler tabs share the second URL. The browser uses the existing shared
backend origin with `/atendimentos/api`. There is no separate booking service,
database or customer domain. Other frontend/preview hosts cannot call the live
support API through this client. No order, charge or tax-document operation is
included; those belong to future Internal Apps tools.

## Production settings

In Azure App Services → Plataforma-Backend-v3 → Settings → Environment variables
→ App settings, add the following values securely. Never paste keys into chat,
Git, frontend files or the generated `dist/` tree.

| Setting | Value |
| --- | --- |
| `PAGARME_API_KEY_PRODUCTION` | The live secret key from the intended Pagar.me account |
| `PAGARME_PUBLIC_KEY_PRODUCTION` | The matching live public key |
| `SUPPORT_MANAGEMENT_ENABLED` | Leave unset/false until the reviewed release is deployed; then set `true` |

The server requires production `sk_` and `pk_` keys, rejecting `sk_test_` and
`pk_test_` keys in production. Some accounts use `sk_live_`/`pk_live_`, which are
also accepted. Local development
continues to use its existing sandbox keys and cannot accept a live pair.
The secret is server-only; the public key is supplied to the card-tokenization
client. Register `https://machadogestao.com` in the live account's permitted
domains, and confirm its card/Zero Dollar verification configuration. If the
account uses an IP allowlist, permit the backend's intended outbound addresses.
Use the same live Pagar.me account for later Internal Apps charging, because
wallet references belong to that account.

Production defaults to the existing Azure persistent content share at
`HOME/data/support-management-production`, outside the deployment directory.
`SUPPORT_DATA_DIRECTORY` is optional: use it only for another verified durable,
private absolute path outside the application tree. A storage-disabled or
local-cache HOME is not accepted as the default. No Azure service is created.

Run one App Service instance and one Node/IISNode worker for this app. The
runtime takes an exclusive metadata lock; a second owner receives an unavailable
response instead of making competing booking writes. Stop the connected local
runner before customers use this same appointments workbook. Continue automated
local development with synthetic dependencies; do not run two booking writers
against the shared workbook.

## Existing data

Hourly prices use each user's workbook reference rate for the first three
available dates, then fall linearly to 80% on the tenth and remain there.
The current R$1,250 reference therefore reaches R$1,000; unavailable dates do not
advance the discount.

BD - USUÁRIOS remains the single company/user table, with the same columns and
password workflow. Login reads SENHA exactly as before; no password migration or
new account-registration process is introduced. Restrict workbook access under
the existing owner workflow and never print passwords in logs or checks.

Card registration still tokenizes directly with Pagar.me and explicitly requests
verification before recording a usable, active, unexpired card. No raw card
number or CVV reaches Excel or disk. A sandbox verification is not a live bank
check; customers must register their cards in the production environment.

Column K records `production` after a verified live registration. Its wallet
references are copied to all matching company rows together. Sandbox references
are not imported into production, and sandbox saves/deletions cannot change
production references. Explicit live registration may replace old sandbox
references; a partial workbook synchronization remains recoverable through the
existing card-sync action.

The private production directory contains only hashed sessions, reusable wallet
references and verification/pending-operation markers, and booking/availability
journals. An environment marker binds it to the live public key. Do not copy the
sandbox directory into production or switch Pagar.me accounts with the same
metadata directory. A live-key rotation also changes this binding: reconcile
that marker against the same account before changing it.

Daily snapshots retain the last seven snapshot dates under the private
`backups/` directory. They are made on startup and before the first metadata
mutation of a new Brasília date. This is a small rollback copy on the same
persistent share, not an independent disaster-recovery service. Excel keeps its
existing OneDrive storage/version workflow. Do not automatically restore an old
snapshot: first reconcile pending reservations/provider operations to avoid
duplicate events or card requests.

## Release and acceptance

1. Keep support disabled while preparing/reviewing the companion backend and
   frontend changes. Complete the required regression suites and generated
   frontend checks. Publication still requires the owner's explicit instruction.
2. Deploy the reviewed backend and frontend through the existing repository
   workflows. Set the live keys and permitted domain, confirm durable storage and
   one worker, then enable support. Missing or invalid configuration fails closed.
3. Verify login at the public URL, the one-hour expiry, company separation and
   administrator-only settings. Use the existing pilot users/rates/permissions
   in BD - USUÁRIOS; do not export credentials for checks.
4. The owner registers a real company card and verifies the masked card, live
   provider wallet and company I:K references. Registration must not charge it.
   Confirm deletion revokes that selected live wallet card before clearing its
   references, then register again if needed.
5. Make one intended booking. Confirm the row, the 09:00–10:00 Brasília Teams
   event, the user's invitation/instructions and the empty-body contato alert.
   API acceptance is not proof of inbox delivery. Automated checks never perform
   these real card, workbook, calendar or email operations.

The sandbox acceptance already performed verifies the workflow; this short
hosted check verifies the different domain, live configuration and delivery.

## Recovery

Stop this support runtime before changing/restoring its private metadata. A
same-instance dead-PID lock is recoverable at the next initialization attempt.
A lock from another Azure instance is deliberately not removed automatically:
verify that instance/worker is stopped before removing only its stale
`production.lock`. Never remove a lock held by an active writer. An initialization
failure can be retried by a later request after its underlying cause is fixed.

Preserve all card and reservation journals. Use the existing sandbox runbook's
stage-by-stage recovery rules with the production runtime and live account;
never resend an ambiguous external write. Turning support off removes its API
mount on restart and leaves legacy application routes unchanged. It does not
cancel existing reservations or revoke stored cards.

References: [Azure persistent storage](https://learn.microsoft.com/en-us/azure/app-service/operating-system-functionality),
[Azure local-cache behavior](https://learn.microsoft.com/en-us/azure/app-service/app-service-local-cache-overview),
[Pagar.me keys](https://docs.pagar.me/docs/chaves-de-acesso),
[wallet verification](https://docs.pagar.me/docs/wallets).
