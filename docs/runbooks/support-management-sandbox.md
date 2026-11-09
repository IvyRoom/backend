# Support management sandbox

This version provides workbook-backed login, weekday date reservations and
administrator-managed company cards. Reservations do not charge a card or set
an appointment duration. Production `app.js` and
`server.js` do not expose these endpoints. Merging the frontend publishes static
pages, but the functional sandbox is local and uses its own same-origin API.

## Daily local cycle

After the setup below, run `npm run support:dev` from `backend` and use its printed
local URL. The app runs locally while reading the real users workbook through
Graph, registering cards with the Pagar.me **test** API, and synchronizing reusable
company-card references to that workbook. Reservar Data updates the appointments
workbook and sends real email/calendar notifications; this is connected local
development, not an offline or simulated application mode.

Save frontend HTML, CSS, or JavaScript and refresh the browser. Sources are served
directly from the sibling `sistemas/apps/support-management/` directory, without a
build. Saving watched support-backend code gracefully stops and restarts that
backend, releasing its storage lock first. Valid sessions survive backend
restarts without extending their original expiry. Stop the development runner
with Ctrl+C when done.

Use `npm run support:test` for focused automated checks during ordinary edits.
Those checks still use synthetic dependencies and deny external traffic; they do
not replace the connected manual acceptance test. Shared changes need relevant
cross-application regression checks. Run the complete required suites, generated
artifact checks, and hosted-origin verification before publication.

Keep checkpoint commits local. Do not push, open or update publication PRs, mark
drafts ready, or deploy until Lucas explicitly requests publication. Existing
support-management PRs remain parked as drafts during iteration; a push can still
trigger a preview deployment even while a PR is a draft.

## Start locally

1. Keep `backend` and `sistemas` as sibling checkouts containing the companion
   support-management changes. Use Node.js 24 and the backend's locked dependencies
   (`npm ci` if they are missing).
2. Keep the existing Graph `CLIENT_ID`, `TENANT_ID`, and `CLIENT_SECRET` in the
   ignored `backend/.env`. Graph reads the verified BD - USUÁRIOS workbook and
   updates only the card-reference and reservation ranges described below.
   The Graph application needs Mail.Send and Calendars.ReadWrite for the contato
   mailbox as well as its existing workbook access.
3. Create `backend/.env.support-sandbox` locally with these two settings. Enter
   actual **test** keys from the same Pagar.me test account, never live keys:

   ```dotenv
   SUPPORT_PAGARME_PUBLIC_KEY=pk_test_REPLACE_LOCALLY
   SUPPORT_PAGARME_SECRET_KEY=sk_test_REPLACE_LOCALLY
   ```

   The existing `PAGARME_API_KEY_SANDBOX` and `PAGARME_PUBLIC_KEY_SANDBOX` in
   `backend/.env` are also accepted for the secret and public test keys. Leave
   the matching `SUPPORT_PAGARME_*` fields blank to reuse them; configured
   `SUPPORT_PAGARME_*` fields take precedence.

   In the same **test** account's dashboard, register the browser origin used for
   tokenization (`http://127.0.0.1:4174`) through its permitted-domain settings.
   Follow the dashboard's required domain format; do not disable an existing
   allowlist. If an IP allowlist is configured, ensure it also permits this
   machine's outbound API requests. Pagar.me requires domain registration before
   browser tokenization. Loopback acceptance is account-dependent and has not
   been verified without test keys; if the dashboard refuses this local origin,
   stop and arrange an approved hosted sandbox origin before card testing.

4. From `backend`, run `npm run support:dev` and open
   <http://127.0.0.1:4174/support-management/login/>. The startup URL uses port
   4174 unless `SUPPORT_SANDBOX_PORT` is configured. Use `127.0.0.1`, not localhost.
   Do not expose this listener through a tunnel or use the production entry point.
   Use `npm run support:sandbox` when automatic backend restarts are unnecessary.
5. Log in using an active workbook account. An account whose `TIPO DE ACESSO` is
   `Administrador` can register or replace a card in the scheduler.
6. Use the Pagar.me simulator card `4000000000000010`, CVV `123`, and a future
   expiry date, with synthetic cardholder/address details. Do not enter the real
   company card into this sandbox. It cannot prove a real issuer authorization.
7. Refresh, sign out and back in, and restart the sandbox to verify persistence.
   A colleague with the same CNPJ sees the masked card and cannot alter it;
   another company sees its own card. Confirm the customer/card in the Pagar.me
   **test** dashboard and the reusable references in BD - USUÁRIOS. Registration does
   not create an order or charge and does not send an email.
8. Stop the server with Ctrl+C when finished.

Pagar.me uses the same API origin for test/live; this entry point accepts only
`pk_test_` and `sk_test_` keys and rejects responses marked live. The secret key
stays server-side. The public key is returned by `/support-management/api/config`.
Browser tokenization goes straight to Pagar.me; only a single-use token and billing
address reach the backend. PAN/CVV are never sent to Graph or saved by the app.

## Identity and persistence

Verified workbook: drive item `01OSXVECRXPEMKXY4F5RH34SKSN35K757F`, table
`{7C4EBF15-124A-4107-9867-F83E9C664B31}` (`BD`), sheet `USUÁRIOS`.
Authentication requires the exact 11-column A:K contract: CNPJ, RAZÃO SOCIAL,
VALOR REFERÊNCIA, NOME COMPLETO, E-MAIL, SENHA, TIPO DE ACESSO, STATUS,
PAGARME CUSTOMER ID, PAGARME CARD ID, PAGARME AMBIENTE. Header comparison trims
whitespace. Any other header, index or range drift fails closed.
Authentication retrieves only A:H, never card or reference
values. Passwords follow the existing workbook contract and never
appear in responses or local metadata. Inactive, missing, or duplicate email
matches fail.

The server explicitly sends `options.verify_card: true` when creating a card.
Only a successful response to that request with an `active`, unexpired card can
create the local `verification: "pagarme_verify_card_v1"` marker. Wallet existence,
an `active` status on a later GET, or `workbookSynced: true` alone never creates
that marker. It records the accepted verification request, not a guarantee of
future payments; sandbox responses simulate this process and do not contact the
real company's issuer.

The official wallet guide supports Zero Dollar Auth for PSP and Gateway. It
specifically requires acquirer enablement for Gateway. The dashboard's active
credit-card status and PSP model do not expose a separate verification setting;
they are not independent proof of a merchant-level verification flag. Follow the
explicit verification request contract and confirm production behavior before
accepting real cards.

Only after this gate passes does the adapter read the current schema and A:H
again, then update I:K for every matching CNPJ row: I (`PAGARME CUSTOMER ID`),
J (`PAGARME CARD ID`), and K (`PAGARME AMBIENTE`). The values are the trusted local
provider IDs and the literal `sandbox`. No card number, holder, expiry, CVV, or
registration-status column remains in this workbook contract. The masked card
display is fetched from Pagar.me. A missing company or changed schema blocks
synchronization instead of adding rows or columns.

On 2026-11-09, read-only verification confirmed USUÁRIOS!A3:K49 with 46 data
rows and the revised CNPJ/reference-rate schema. No workbook migration is needed
for this version. The six legacy raw/masked-card columns were removed previously.

Workbook references mirror the reusable provider wallet IDs for reporting and
backup. They are never supplied by the browser, exposed by the card API, read for
authentication, or imported as verification evidence. For a missing CNPJ-keyed
local record, agreeing sandbox workbook references can recover the wallet IDs.
Only an exact customer/card match in an existing local record preserves its
verification evidence and pending state. Workbook IDs alone remain unverified;
conflicting or ambiguous references fail closed.
The temporary browser `token_` is never stored. Future purchases can use the
provider `card_id` without storing PAN or CVV; the purchase flow and its required
PSP customer information remain a separate milestone.

One-hour opaque sessions survive local backend restarts and keep their original
absolute expiry. Every protected request rechecks the current workbook company,
password, status and role. Logout and invalidation are persisted before they are
acknowledged. The ignored local session store contains only token hashes,
identity, password fingerprints, original expiry and a stable random fingerprint
salt; it never stores raw bearer tokens or passwords. Corrupt session storage
blocks startup rather than silently resetting it. Only the raw session token is
stored in the browser's sessionStorage. Companies are grouped by the 14 CNPJ
digits. Renaming RAZÃO SOCIAL preserves identity; changing the CNPJ invalidates
the session. Sessions minted against the previous company-label identity require
a fresh login after this schema update.

Pagar.me stores the test customer and card. Ignored `.support-sandbox-data/` holds
hashed company keys, Pagar.me customer/card IDs, pending-operation flags,
the explicit-verification marker, and successful-workbook-sync markers. It also
holds reservation journals with date, CNPJ, user email/name, company, quoted
price, request UUID, stage and Graph event references; no card data or passwords.
Keep this directory between restarts and back it up locally. Card references are
specific to the configured Pagar.me account; do not switch accounts with the same
metadata directory. Replacing a card selects the new card for the company; it does
not delete earlier cards from the test provider wallet.

Only one sandbox process may use a metadata directory; startup uses an exclusive
`sandbox.lock`. After an abnormal termination, verify the recorded PID is no longer
running before removing only that stale lock file. Never remove it while the server
is running. Do not run parallel copies of the sandbox against the same test account.

## Date reservations

Verified on 2026-11-09: BD - ATENDIMENTOS drive item
`01OSXVECSFKTNUUJS3TZDKCLDRXFZCMVYB`, sheet ATENDIMENTOS, table BD
`{7C4EBF15-124A-4107-9867-F83E9C664B31}`, range A3:H55. Its exact columns are DATA,
STATUS, OBSERVAÇÃO, CNPJ, NOME COMPLETO, VALOR POR HORA, DURAÇÃO (MIN), VALOR COBRADO.
The last column has the table formula `=[@[DURAÇÃO (MIN)]]/60*[@[VALOR POR HORA]]`.

The calendar reads dates/status only, starting tomorrow in São Paulo, and omits
weekends. A blank status means vacant; any other status means unavailable.
The first vacant weekday uses the authenticated user's VALOR REFERÊNCIA. Each
following vacant weekday reduces that price by 2.5% of the reference price, to
75% at the eleventh vacant weekday; later days stay at 75%. Prices round to cents.
Unavailable days do not advance the discount. Price and availability are checked
again on reservation; stale quotes require reselection. A global claim queue
prevents different dates from accepting incompatible concurrent quotes.

Reservar Data writes only the existing selected row: B=RESERVADO, D=formatted
CNPJ, E=authenticated name, F=hourly price. Null placeholders preserve DATA,
OBSERVAÇÃO, DURAÇÃO (MIN) and VALOR COBRADO, including the existing formula.
The date is rechecked immediately before the single Graph PATCH; no row is added.
This first version reserves a whole day without setting hours/duration or charging
the card. Workbook edits made concurrently outside this local app are not atomic
with its PATCH; avoid another booking writer during this local trial.

After the row is reserved, the app creates an all-day event in the contato
mailbox's calendar, in São Paulo time, and invites the authenticated user. Graph
sends the attendee invitation automatically. A confirmation goes to
contato@machadogestao.com and the authenticated user, with a calendar attachment
using the same event UID. The organizer gets its calendar entry and confirmation;
the user gets the native invitation and confirmation. Templates are intentionally
minimal. Mail API acceptance is not proof of final inbox delivery.

Each request UUID has an atomic ignored `reservation_<UUID>.json` journal:
workbook → event → confirmation → complete. Replaying a completed request does not
reserve again or resend anything. A pending journal blocks its date even after
restart. External writes have SDK automatic retries disabled. An uncertain row
write requires review; an event/mail failure reports the date reserved with
notifications pending. Refreshing never retries or silently claims delivery.

For manual recovery, stop the runner and preserve the journal. Reconcile only the
same UUID, using the exact stored identity/date/price:

1. At workbook stage, inspect the appointment row. If the write succeeded, verify
   its date, CNPJ, name and rate before advancing to event. If no write happened,
   remove only this journal after proving the row is vacant. Never clear an
   uncertain journal merely to make the button work.
2. At event stage, inspect that date in contato's calendar. If the matching event
   exists, recover its Graph event ID and iCalUId into the journal and advance to
   confirmation. If its absence is established, make one explicit event attempt
   with the existing `eventPayload(record)` helper and the same transactionId;
   persist the returned references before proceeding. Never create a second event
   just because the previous response was lost.
3. At confirmation stage, inspect Sent Items for the matching date/user/UID. If
   already sent, mark complete. Otherwise, after establishing it was not sent,
   make one explicit `sendConfirmation(confirmation(record))` attempt and mark
   complete only after acceptance. Both payload helpers are exported from
   `domains/support-bookings.js`; use the same notification adapter and test
   configuration. An ambiguous outcome stays pending for review.

No automatic recovery worker or cancellation/rescheduling UI is included.

## Uncertain payment responses

Provider writes are never retried automatically. A definite provider rejection
clears the pending flag, allowing corrected test details. A timeout, 5xx,
malformed response, or local persistence failure retains it and blocks another
registration. A verification rejection (`412` from Pagar.me) returns the
sanitized `card_verification_failed` error, restores the previous card reference,
and never writes new card references to the workbook. Missing or unusable verification
responses remain blocked for review rather than claiming successful registration.

A card accepted by the verification gate is stored locally with its verification
marker and `pending: "workbook"` before Graph updates begin. If a workbook update
fails, the card stays registered and the API returns `workbookSyncPending: true`;
replacement remains blocked until
synchronization succeeds. The administrator can retry through the card screen.
`POST /support-management/api/card/sync` fetches the existing card from Pagar.me,
patches the same workbook ranges, and records both `workbookSynced: true` and
`workbookReferencesSynced: true`; it does not
create a customer, card, order, or charge. Partial workbook updates are safe to
repeat. Each attempt is explicit and has no automatic external-request retry.

Verified local card references missing either sync marker also report pending
sync. This includes cards from the earlier masked-only workbook sync; their
reusable IDs are backfilled using the trusted local record without creating a
provider customer or card. On the administrator's first card-screen load, the app
attempts this sync
once and offers the retry action if it fails. Unverified legacy references return
`card: null, verificationRequired: true`, even if previously workbook-synced; they
cannot be backfilled through the sync endpoint and require fresh card entry.
Pagar.me may return the same wallet card ID when the card is entered again. ID
reuse is not a failure signal: acceptance still requires a successful explicit
`verify_card` request and an active, unexpired result. The local marker records
that request contract, not independent issuer evidence. Do not add verification
markers manually or promote an old card based on GET alone.

Every registered-card read and workbook sync also checks the current provider
status and expiry. A deleted, expired, or status-less card is not displayed as
registered and cannot update the workbook. `GET /support-management/api/card`
never writes to Graph or requests authorization. Current workbook role and
company are checked for every sync request, and a colleague cannot trigger it.

Stop the sandbox before reconciling a pending company record. Use the record's
filename (the Pagar.me customer `code`) to locate the exact customer in the test
dashboard. For `pending: "customer"`, confirm there is exactly one matching
customer; preserve its ID as `customerId` and remove only the pending flag. If none
exists, clear the record to `{}` only after confirming the request did not succeed.
For `pending: "card"`, inspect the customer's wallet and the verification outcome.
An existing card or `active` status does not prove that the explicit verification
request succeeded. Preserve the previous verified reference after confirming a
failed replacement created no usable card; if the outcome remains uncertain,
leave the record blocked. Any recovered card without established verification
must use the fresh-entry flow. Never copy PAN/CVV or live credentials into these
files or invent a verification marker during reconciliation.

## Verification and production follow-up

Run `npm test` with Node.js 24. Tests deny non-loopback network connections and use
synthetic Graph/payment dependencies. Frontend checks are listed in `sistemas/AGENTS.md`.
Actual Pagar.me registration and workbook updates remain user-operated acceptance
checks; automated tests do not claim live provider or workbook validation.

Before production, agree on hosting, durable shared
storage/coordination, live-key configuration, and the real-card acceptance test.
There is intentionally no live-key switch in this sandbox implementation.

Official references: [authentication](https://docs.pagar.me/reference/autentica%C3%A7%C3%A3o-2),
[tokenization](https://docs.pagar.me/reference/tokeniza%C3%A7%C3%A3o-1),
[token domain registration](https://docs.pagar.me/reference/criar-token-cart%C3%A3o-1),
[allowlists](https://docs.pagar.me/docs/ip-allowlist),
[create card](https://docs.pagar.me/reference/criar-cart%C3%A3o),
[wallet verification](https://docs.pagar.me/docs/wallets),
[verification request model](https://github.com/pagarme/pagarme-nodejs-sdk/blob/main/src/models/createCardOptionsRequest.ts),
[PSP simulator](https://docs.pagar.me/docs/simulador-psp),
[Gateway simulator](https://docs.pagar.me/docs/simulador-de-cart%C3%A3o-de-cr%C3%A9dito).
