# Support management sandbox

This version provides workbook-backed login, 09:00 Brasília-time Microsoft Teams
reservations and administrator-managed company cards. Invitations hold up to
two hours; actual duration remains a manual workbook entry. Reservations and
nonattendance do not automatically charge a card. Production `app.js` and
`server.js` do not expose these endpoints. Merging the frontend publishes static
pages, but the functional sandbox is local and uses its own same-origin API.

## Daily local cycle

After the setup below, run `npm run support:dev` from `backend` and use its printed
local URL. The app runs locally while reading the real users workbook through
Graph, registering cards with the Pagar.me **test** API, and synchronizing reusable
company-card references to that workbook. Calendar loading maintains availability
in BD - ATENDIMENTOS; Reservar Agenda reserves a row and sends real email/calendar
notifications. This is connected local
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
   updates only the card-reference, availability and reservation ranges described below.
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
5. Log in using a listed workbook account. An account whose `TIPO DE ACESSO` is
   `Administrador` can register a card and enable individual employee booking
   checkboxes in Configurações da Conta PJ, then click Salvar. Employees default to
   allowed unless the workbook stores Não. Without a verified card or permission,
   Agendamento shows the company administrator.
6. Use the Pagar.me simulator card `4000000000000010`, CVV `123`, and a future
   expiry date, with synthetic cardholder details. Do not enter the real
   company card into this sandbox. It cannot prove a real issuer authorization.
7. Refresh, sign out and back in, and restart the sandbox to verify persistence.
   A colleague with the same CNPJ sees the masked card and cannot alter it;
   another company sees its own card. Confirm the customer/card in the Pagar.me
   **test** dashboard and the reusable references in BD - USUÁRIOS. Registration does
   not create an order or charge and does not send an email.
8. For booking acceptance, select the intended date in Agendamento, review the
   rate, click Reservar Agenda, then accept the nonattendance term with Confirmar.
   Verify the reserved row, the 09:00–12:00 Teams invitation and confirmation at
   both recipient accounts. This is a real reservation and notification test.
9. Stop the server with Ctrl+C when finished.

Pagar.me uses the same API origin for test/live; this entry point accepts only
`pk_test_` and `sk_test_` keys and rejects responses marked live. The secret key
stays server-side. The public key is returned by `/support-management/api/config`.
Browser tokenization goes straight to Pagar.me; only a single-use token reaches
the backend. Wallet registration omits optional billing-address fields and keeps
`verify_card: true`. PAN/CVV are never sent to Graph or saved by the app. The live
card representation remains visible and locked during registration, then clears
when the request settles or the administrator leaves the form; saved cards remain
masked. CVV validation runs on blur or submission, never while typing. See the
[official card request model](https://github.com/pagarme/pagarme-nodejs-sdk/blob/main/doc/models/create-card-request.md)
for optional address fields; a later order/antifraud flow may require them.

## Identity and persistence

Verified workbook: drive item `01OSXVECRXPEMKXY4F5RH34SKSN35K757F`, table
`{7C4EBF15-124A-4107-9867-F83E9C664B31}` (`BD`), sheet `USUÁRIOS`.
Authentication verifies the exact eleven-column A:K contract: CNPJ, RAZÃO SOCIAL,
VALOR REFERÊNCIA, NOME COMPLETO, E-MAIL, SENHA, TIPO DE ACESSO,
AUTORIZADO A CONTRATAR ATENDIMENTOS, PAGARME CUSTOMER ID, PAGARME CARD ID,
PAGARME AMBIENTE. Header comparison trims whitespace. Other schema drift fails closed.
Credential validation retrieves only A:H, never provider references. H replaced
STATUS: listed, uniquely identified users can log in; Não prevents booking rather
than login. Sim and blank permission values default to allowed. Passwords never
appear in responses or local metadata; missing/duplicate email matches fail.

Checkbox edits remain local until Salvar submits the complete active-company list.
The button and Salvamento pendente. warning appear only while the list has
unsaved changes, remain during saving or a failed attempt, and disappear after
a successful save or when the original selection is restored.
The server validates every target against the fresh authorization directory and
uses its trusted adapter snapshot, avoiding another schema/directory read. Only
changed H cells are patched; multiple cells share Graph batches of up to twenty.
SDK retries are disabled, all subresponses are checked, and an uncertain save
keeps the draft available for an explicit retry. No column is created or renamed.

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

On 2026-11-11, fresh read-only verification confirmed USUÁRIOS!A3:K49 with 46 data
rows, H named AUTORIZADO A CONTRATAR ATENDIMENTOS and all 46 values Sim. There is
no L permission column. The six legacy raw/masked-card columns were removed previously.

Workbook references mirror the reusable provider wallet IDs for reporting and
backup. They are never supplied by the browser, exposed by the card API, read for
authentication, or imported as verification evidence. For a missing CNPJ-keyed
local record, agreeing sandbox workbook references can recover the wallet IDs.
Reference lookup reads contiguous I:K ranges for the matching CNPJ together,
avoiding one Graph request per employee without reading another company's cells.
Without a registered card, Agendamento shows the pending account-settings warning
to every company user and does not load reservable dates.
Only an exact customer/card match in an existing local record preserves its
verification evidence and pending state. Workbook IDs alone remain unverified;
conflicting or ambiguous references fail closed.
The temporary browser `token_` is never stored. Future purchases can use the
provider `card_id` without storing PAN or CVV; the purchase flow and its required
PSP customer information remain a separate milestone.

One-hour opaque sessions survive local backend restarts and keep their original
absolute expiry, returned as `expiresAt` for background expiry enforcement. The
session timer is not displayed. Every protected
request rechecks the current workbook company,
password and role. There is no inactive-login state; permission Não only blocks
booking. Logout and invalidation are persisted before they are
acknowledged. The ignored local session store contains only token hashes,
identity, password fingerprints, original expiry and a stable random fingerprint
salt; it never stores raw bearer tokens or passwords. Corrupt session storage
blocks startup rather than silently resetting it. Browser sessionStorage holds the
token and a public scheduler-context cache bound to that token and expiry. Login
preloads configuration, masked card and calendar, avoiding serial requests during
navigation; reused context refreshes in the background. Logout, HTTP 401 and
expiry clear or invalidate the cached context. Companies are grouped by the 14 CNPJ
digits. Renaming RAZÃO SOCIAL preserves identity; changing the CNPJ invalidates
the session. Sessions minted against the previous company-label identity require
a fresh login after this schema update.

Pagar.me stores the test customer and card. Ignored `.support-sandbox-data/` holds
hashed company keys, Pagar.me customer/card IDs, pending-operation flags,
the explicit-verification marker, and successful-workbook-sync markers. It also
holds reservation journals with date, CNPJ, user email/name, company, quoted
price, request UUID, acknowledgment time, stage and Graph event references, with an
optional Teams URL; no card data or passwords. `availability-sync.json` stores
only the dates of a pending table append.
Keep this directory between restarts and back it up locally. Card references are
specific to the configured Pagar.me account; do not switch accounts with the same
metadata directory. Replacing a card selects the new card for the company; it does
not delete earlier cards from the test provider wallet.

Only one sandbox process may use a metadata directory; startup uses an exclusive
`sandbox.lock`. After an abnormal termination, verify the recorded PID is no longer
running before removing only that stale lock file. Never remove it while the server
is running. Do not run parallel copies of the sandbox against the same test account.

## Date reservations

Verified schema on 2026-11-10: BD - ATENDIMENTOS drive item
`01OSXVECSFKTNUUJS3TZDKCLDRXFZCMVYB`, sheet ATENDIMENTOS, table BD
`{7C4EBF15-124A-4107-9867-F83E9C664B31}`, initially A3:H55. The table now grows as
dates are added. Its exact columns are DATA,
STATUS, OBSERVAÇÃO, CNPJ, NOME COMPLETO, VALOR POR HORA, DURAÇÃO (MIN), VALOR COBRADO.
The last column has the table formula `=[@[DURAÇÃO (MIN)]]/60*[@[VALOR POR HORA]]`.

The scheduler has burgundy-to-dark-burgundy sidebar/header gradients and a white
workspace. Its course-content header shows the full name; the one-hour session
expiry runs in the background. The card tool follows a Ferramentas do Administrador
navigation divider.
It always displays the current month and next month in approximately 500px-wide
blocks with square day tiles, Sunday–Saturday columns and successive week rows;
they shrink on narrow screens. Gray tiles show unavailable dates without an
INDISPONÍVEL caption. Only tomorrow through the last day of next month, in São Paulo time, can be booked. Weekends,
national holidays and Curitiba holidays are unavailable. A blank workbook status
means vacant; any other status blocks booking. Full displayed months include gray
dates outside the booking horizon. The short date-selection instruction is
followed by the calendar and a native expandable Perguntas Frequentes section.
Its confirmation answer displays the authenticated user's email. The FAQ heading
is centered, and questions expand without triangle markers. A calendar legend
identifies the burgundy-to-green available days and gray reserved/unavailable days.

Holiday rules cover the federal holidays plus Curitiba's September 8, Good Friday
and Corpus Christi; government recesses and optional closures are not added.
Sources: [national holidays](https://www2.camara.leg.br/legin/fed/lei/1940-1949/lei-662-6-abril-1949-347136-normaatualizada-pl.html),
[October 12](https://www2.camara.leg.br/legin/fed/lei/1980-1987/lei-6802-30-junho-1980-357065-publicacaooriginal-1-pl.html),
[November 20](https://normas.leg.br/?urn=urn%3Alex%3Abr%3Afederal%3Alei%3A2023-12-21%3B14759),
[Curitiba law](https://legisladocexterno.curitiba.pr.gov.br/VisualizarHTML.aspx?id=20655)
and [official 2026 calendar](https://www.curitiba.pr.gov.br/noticias/prefeitura-de-curitiba-define-calendario-de-feriados-recessos-e-pontos-facultativos-para-2026/80481).
Easter-based dates use the [USNO computus](https://aa.usno.navy.mil/faq/easter).

When a verified card is present, calendar loading ensures all dates from tomorrow
through the last day of next month have table rows. Automatic availability changes touch only B/C on otherwise
vacant rows: B=`INDISPONÍVEL`, C=`APP: INDISPONÍVEL` marks an app-managed block.
The app clears its own marker when a date becomes eligible. It preserves manual
statuses/notes, existing booking data and every RESERVADO row. No BD - USUÁRIOS
columns are added; its A:K contract stays unchanged.

Before `rows/add`, a durable `availability-sync.json` intent records the sorted
missing dates. The append makes one attempt with retries disabled. A fresh table
read must confirm every date and the calculated H formula on every row before
the intent is cleared. Excel must propagate the existing calculated-column
formula; missing or changed formulas stop the calendar instead of silently
accepting uncalculated rows.

After an ambiguous append or formula failure, stop the runner and inspect the
intent and workbook together. If every recorded date exists exactly once and H
has the verified formula, the next read can reconcile the intent. Otherwise,
establish which rows were added and restore the calculated formula or resolve the
missing dates before clearing the intent. Never blindly repeat the append or
delete its intent to bypass the review.

The first three vacant weekdays use 100% of the authenticated user's VALOR
REFERÊNCIA. Prices then decrease linearly from the third through the tenth vacant
weekday, where they reach 75%; later days stay at 75%. There are seven equal
discount steps, each 25/7 percentage points of the reference price. Prices round to cents.
Unavailable days do not advance the discount. Price and availability are checked
again on reservation; stale quotes require reselection. A global claim queue
prevents different dates from accepting incompatible concurrent quotes.

A verified, active and unexpired company card plus an allowed employee booking
permission are required for calendar access and new reservations. Revocation is
rechecked server-side even with an existing session or cached calendar. Initial context/calendar reads cache provider readiness for up
to 60 seconds, while still reading local verification/pending state. Explicit
card reads are fresh; save/sync invalidate that cache. Every new reservation
performs a fresh Pagar.me GET before accepting the booking. No order or charge is
created by this check, and the user-directory authorization is never cached.

Selecting a date dims the other day tiles and opens a native modal with the date,
09:00 start and hourly price. Reservar Agenda reveals the nonattendance warning
and Confirmar/Cancelar actions, with Cancelar focused by default. No reservation
request is sent until Confirmar; it supplies `acknowledged: true`, and the journal
records `acknowledgedAt`. Cancelar, Escape or dismissal before submission closes
the popup without reserving. Dismissal is blocked while submission is pending.
There is no cancellation API after a reservation is confirmed. The displayed
nonattendance term does not implement an automatic charge.

Reservar Agenda writes only the existing selected row: B=RESERVADO, D=formatted
CNPJ, E=authenticated name, F=hourly price. Null placeholders preserve DATA,
OBSERVAÇÃO, DURAÇÃO (MIN) and VALOR COBRADO, including the existing formula.
The date is rechecked immediately before the single Graph PATCH; reservation
itself adds no row. Meetings start at 09:00 Brasília time and may last up to
3h as needed. DURAÇÃO (MIN), column G, remains manual; H calculates the charge
from that actual duration and the agreed rate. No card is charged.
Workbook edits made concurrently outside this local app are not atomic
with its PATCH; avoid another booking writer during this local trial.

After the row is reserved, the app requests a Microsoft Teams event holding
09:00–12:00 in the contato mailbox's calendar, in Brasília time, and invites the
authenticated user. Its returned Teams link is included when provided. Graph
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
   exists, recover its Graph event ID, iCalUId and available Teams URL into the journal and advance to
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

Explicit registered-card reads and workbook sync fetch the current provider
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

## Company-card deletion

The administrator's trash action reveals inline Deletar Cartão and Cancelar
buttons in the saved card heading, with Cancelar focused. Opening, cancelling,
or dismissing these controls does not send a deletion request.
`DELETE /support-management/api/card` accepts only the opaque
`expectedCardRevision` of the displayed registration, never provider identifiers.
It rechecks the workbook role and current revision, derives the selected company
card locally, and persists `pending: "delete"` before one Pagar.me sandbox wallet DELETE.
A matching deleted response or documented 404 establishes removal. It then stores
`pending: "delete_workbook"`, clears I:K for the authenticated CNPJ, and retains
only the reusable customer ID locally. The original registration form returns
after that cleanup succeeds; the next registration reuses the company customer.
Each fresh verified registration gets a new local revision, so a delayed deletion
retry cannot remove its replacement. Existing registrations use a deterministic
reference digest until their next registration; no workbook column is added.

Failed or uncertain deletion does not claim success. Provider retries first GET
the same card: deleted/404 skips another DELETE, while an active card permits one
explicit deletion attempt. Workbook-only retries never repeat provider deletion.
The pending state makes the card unavailable for new bookings and registration;
the administrator can retry the trash action. No automatic external retries run.
Company card mutations and a booking's payment-check/workbook phase share a
company guard; neither can proceed through the other's payment state change.
The existing provider transactions are not modified. Automated checks use
synthetic cards and workbook clear operations only.

Official request contract: [delete card](https://docs.pagar.me/reference/excluir-cart%C3%A3o).

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
