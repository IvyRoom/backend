# Support management sandbox

This first step provides workbook-backed login and administrator-managed company
cards. Scheduling and purchases are not implemented. Production `app.js` and
`server.js` do not expose these endpoints. Merging the frontend publishes static
pages, but the functional sandbox is local and uses its own same-origin API.

## Start locally

1. Keep `backend` and `sistemas` as sibling checkouts containing the companion
   support-management changes. Use Node.js 24 and the backend's locked dependencies
   (`npm ci` if they are missing).
2. Keep the existing Graph `CLIENT_ID`, `TENANT_ID`, and `CLIENT_SECRET` in the
   ignored `backend/.env`. Graph only reads the verified BD - USUÁRIOS workbook.
3. Create `backend/.env.support-sandbox` locally with these two settings. Enter
   actual **test** keys from the same Pagar.me test account, never live keys:

   ```dotenv
   SUPPORT_PAGARME_PUBLIC_KEY=pk_test_REPLACE_LOCALLY
   SUPPORT_PAGARME_SECRET_KEY=sk_test_REPLACE_LOCALLY
   ```

4. From `backend`, run `npm run support:sandbox` and open
   <http://127.0.0.1:4174/support-management/login/>. The startup URL uses port
   4174 unless `SUPPORT_SANDBOX_PORT` is configured. Use `127.0.0.1`, not localhost.
   Do not expose this listener through a tunnel or use the production entry point.
5. Log in using an active workbook account. An account whose `TIPO DE ACESSO` is
   `Administrador` can register or replace a card in the scheduler.
6. Use the Pagar.me simulator card `4000000000000010`, CVV `123`, and a future
   expiry date, with synthetic cardholder/address details. Do not enter the real
   company card into this sandbox. It cannot prove a real issuer authorization.
7. Refresh, sign out and back in, and restart the sandbox to verify persistence.
   A colleague in the same `EMPRESA` sees the masked card and cannot alter it;
   another company sees its own card. Confirm the customer/card in the Pagar.me
   **test** dashboard. No order, charge, email, or workbook mutation occurs.
8. Stop the server with Ctrl+C when finished.

Pagar.me uses the same API origin for test/live; this entry point accepts only
`pk_test_` and `sk_test_` keys and rejects responses marked live. The secret key
stays server-side. The public key is returned by `/support-management/api/config`.
Browser tokenization goes straight to Pagar.me; only a single-use token and billing
address reach the backend. PAN/CVV are never sent to Graph or saved by the app.

## Identity and persistence

Verified workbook: drive item `01OSXVECRXPEMKXY4F5RH34SKSN35K757F`, table
`{7C4EBF15-124A-4107-9867-F83E9C664B31}` (`BD`), sheet `USUÁRIOS`.
Runtime validates all 13 column headers, then retrieves only A:G, never the legacy
card columns H:M. Passwords follow the existing workbook contract and never appear
in responses or local metadata. Inactive, missing, or duplicate email matches fail.

Four-hour opaque sessions expire on restart and are rechecked against the workbook
on each protected request, including current company, password, status and role.
Only the session token is stored in sessionStorage. The app groups companies by
trimmed, whitespace-collapsed, case-folded `EMPRESA`, because the workbook has no
stable company ID. Renaming a company changes that identity; do not rename labels
or reuse a label for a different company during this trial.

Pagar.me stores the test customer and card. Ignored `.support-sandbox-data/` holds
only hashed company keys, Pagar.me customer/card IDs, and pending-operation flags.
Keep this directory between restarts and back it up locally. Card references are
specific to the configured Pagar.me account; do not switch accounts with the same
metadata directory. Replacing a card selects the new card for the company; it does
not delete earlier cards from the test provider wallet.

Only one sandbox process may use a metadata directory; startup uses an exclusive
`sandbox.lock`. After an abnormal termination, verify the recorded PID is no longer
running before removing only that stale lock file. Never remove it while the server
is running. Do not run parallel copies of the sandbox against the same test account.

## Uncertain payment responses

Writes are never retried automatically. A definite provider rejection clears the
pending flag, allowing corrected test details. A timeout, 5xx, malformed response,
or local persistence failure retains it and blocks another registration.

Stop the sandbox before reconciling a pending company record. Use the record's
filename (the Pagar.me customer `code`) to locate the exact customer in the test
dashboard. For `pending: "customer"`, confirm there is exactly one matching
customer; preserve its ID as `customerId` and remove only the pending flag. If none
exists, clear the record to `{}` only after confirming the request did not succeed.
For `pending: "card"`, inspect the customer's wallet: select the successfully
registered card's ID as `cardId`, or retain the previous `cardId` after confirming
the failed request created no card; then remove the flag. If uncertain, leave the
record blocked. Never copy PAN/CVV or live credentials into these files.

## Verification and production follow-up

Run `npm test` with Node.js 24. Tests deny non-loopback network connections and use
synthetic Graph/payment dependencies. Frontend checks are listed in `sistemas/AGENTS.md`.
Actual Pagar.me registration remains a user-operated acceptance check until local
test keys are supplied; automated tests do not claim live provider validation.

Before production, agree on hosting, a stable company identifier, durable shared
storage/coordination, live-key configuration, and the real-card acceptance test.
There is intentionally no live-key switch in this sandbox implementation.

Official references: [authentication](https://docs.pagar.me/reference/autentica%C3%A7%C3%A3o-2),
[tokenization](https://docs.pagar.me/reference/tokeniza%C3%A7%C3%A3o-1),
[create card](https://docs.pagar.me/reference/criar-cart%C3%A3o),
[PSP simulator](https://docs.pagar.me/docs/simulador-psp),
[Gateway simulator](https://docs.pagar.me/docs/simulador-de-cart%C3%A3o-de-cr%C3%A9dito).
