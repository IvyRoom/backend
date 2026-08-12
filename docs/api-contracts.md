# Current API contracts

This document is the compatibility baseline preserved by the Express
application/server seam. It describes wire behavior observed in backend commit
`0f39f020387d9e41f17fc1a4490272ccb14a7745` and in the frozen `sistemas`
consumer snapshot
[`c68f361de054a936b7a6871d82d75a1cdb457c97`](https://github.com/IvyRoom/sistemas/tree/c68f361de054a936b7a6871d82d75a1cdb457c97).
The registered routes remain implemented in the import-safe
[`app.js`](../app.js) factory; production configuration, dependency
construction, listener startup, and Graph token refresh are implemented in
[`server.js`](../server.js).

This is a source characterization, not a redesigned API specification. The
application/server split preserves the compatibility contracts below unless a
separate change explicitly migrates their consumers. The known-risk behavior is
recorded so that later structural work does not accidentally change it.

## Evidence boundaries

- Route, middleware, retry, external-call, response, and positional-workbook
  details are taken from the repository source at the commits above.
- Every workbook path, table identity, width, index, literal, and `null`
  placeholder below is **source-observed only**. No live workbook or `AUXILIAR`
  list was read for this inventory, and this document does not claim live-schema
  verification.
- Consumer absence means only that no caller was found in the frozen `sistemas`
  repository. It does not mean that an endpoint is unused or removable.
- Configuration values, ignored files, production routes, and external services
  were not inspected or exercised.
- The package declares Node.js `24.x`; the lock resolves Express 4.22.1,
  body-parser 1.20.5, CORS 2.8.6, and Multer 1.4.5-lts.2. There is no custom
  404 middleware or error middleware.

## Global startup and request pipeline

### Module startup

Requiring either `app.js` or `server.js` is side-effect free. Import does not
load `.env`, decode production configuration, construct production Graph or
Face clients, open a listener, acquire a Graph token, call an external service,
or schedule a timer. `app.js` exports `createApp(dependencies)`; calling the
factory creates a fresh Express application and registers the middleware and 14
routes, but does not listen.

Explicit production startup proceeds as follows:

1. `startProductionServer()` calls `dotenv.config()` through its environment
   loader.
2. `PLATFORM_ROW_AUTHORIZATION_KEY_BASE64` is decoded and validated. A missing,
   noncanonical, or wrong-length value throws before production dependencies,
   the Express application, or the listener are created. The value itself is
   outside this document.
3. The Microsoft Graph confidential client is configured from environment names
   `CLIENT_ID`, `TENANT_ID`, and `CLIENT_SECRET`, using authority
   `https://login.microsoftonline.com/${TENANT_ID}`. Its Graph client retains an
   access-token variable that is initially unset, and its auth provider supplies
   the current value through the existing callback shape.
4. The Azure Face client is constructed from `AZURE_FACE_API_ENDPOINT` and an
   `AzureKeyCredential` built from `AZURE_FACE_API_KEY`, with no pipeline options
   that would replace SDK defaults. Production row authorization is constructed
   from the decoded signing key.
5. `createApp(...)` creates the application and registers global middleware in
   this exact order: `cors()`, `express.json()`, then
   `app.use('/img', express.static('img'))`, followed by the 14 routes.
6. `app.listen(process.env.PORT || 3000)` is called and its listener is retained.
   Only after that call, the initial Graph acquisition is invoked without being
   awaited, so the listener has no token-readiness gate. The returned production
   lifecycle seam can cancel the refresh timer and close the listener.
7. Production startup occurs when `server.js` is the direct main module or when
   `process.argv[1]` resolves to `server.js`. The latter supports Azure App
   Service's Windows IISNode interceptor, which requires the application module
   but preserves its path in `argv[1]`. Ordinary module imports do not satisfy
   either entry-point check.

Source: [`app.js` lines 32-51](../app.js#L32-L51),
[`app.js` lines 637-641](../app.js#L637-L641),
[`server.js` lines 16-140](../server.js#L16-L140), and
[`platform-row-authorization.js` lines 12-26](../platform-row-authorization.js#L12-L26).

Graph token acquisition requests client credentials for the exact scope
`https://graph.microsoft.com/.default` and has its own startup policy, separate
from the route retry helper:

- On success, the access token is replaced and the next refresh is scheduled
  for five minutes before expiry, but never sooner than 60 seconds. The failure
  delay resets to 2 seconds.
- On failure, this function propagates no startup/readiness signal and logs no
  error; the listener stays up while the token remains unset, so individual
  Graph calls can still fail. A new acquisition attempt is scheduled after 2
  seconds, then 4, 8, 16, 32, and at most 60 seconds for subsequent failures.
- Any existing timer is cleared before the next timer is assigned.

Source: [`server.js` lines 20-59](../server.js#L20-L59) and
[`server.js` lines 81-90](../server.js#L81-L90).

### Request middleware

All requests traverse the following order before route-specific middleware:

1. `cors()` uses its defaults: `Access-Control-Allow-Origin: *`; preflight
   methods `GET,HEAD,PUT,PATCH,POST,DELETE`; reflected requested headers; no
   credential header; and an empty `204` response for handled `OPTIONS`
   preflights.
2. `express.json()` considers `application/json` bodies, inflates supported
   encodings, enforces the default 100 KiB limit and UTF character sets, and
   uses strict JSON parsing (object or array). Empty/no-body requests leave an
   empty object. There is no URL-encoded parser.
3. `/img` static handling uses the `img` directory relative to the process
   working directory. It serves `GET` and `HEAD` requests, including the current
   `/img/LOGO_PAGAR.ME.png` (`image/png`) and
   `/img/ASSINATURA_E-MAIL.jpg` (`image/jpeg`) assets, with file-derived content
   types. Default directory redirects are enabled: `GET /img` returns an HTML
   `301` redirect to `/img/`; because there is no index file, `/img/` then falls
   through to the default 404. Other misses and unsupported methods also fall
   through to later routing.
4. The matching route and its route-specific middleware run.

The canonical method/path spellings in this document are exact. Express case-
sensitive routing and strict routing are not enabled, so current path matching
is case-insensitive and accepts a trailing slash. JavaScript body and query
property lookups remain case-sensitive. Express also supplies `HEAD` handling
for registered `GET` routes, while CORS terminates matching preflights before
routing.

Source: [`app.js` lines 32-51](../app.js#L32-L51).

### Default parser, 404, and error behavior

- Malformed JSON, oversized JSON, an unsupported JSON charset, and Multer
  parser errors are passed to Express's default error path because there is no
  custom error middleware. They are not route-specific JSON contracts; common
  parser statuses include `400`, `413`, and `415`, with the default HTML error
  representation.
- An unmatched request falls through to Express's default `404` HTML response,
  whose message is `Cannot <METHOD> <path>`. CORS headers have already been
  applied.
- A synchronous error passed through Express reaches the default HTML error
  handler. Error detail depends on `NODE_ENV`; it is not a stable JSON shape.
- These routes use Express 4 async handlers. A rejected handler promise outside
  a local `try`/`catch` is not forwarded automatically to Express's error
  handler, so there is no defined HTTP status/body for that failure. It can
  surface as an unhandled process-level rejection and can leave the request
  unresolved. Route sections call this **no explicit unexpected-error
  contract**.

### Retry layers

`retry(fn, retries = 5)` makes five total attempts for every thrown/rejected
error. It waits 500, 1,000, 1,500, and 2,000 milliseconds between attempts. It
has no status/error filtering, jitter, timeout, or cancellation. Whether retry
is safe therefore depends on each operation. Source:
[`app.js` lines 66-74](../app.js#L66-L74).

The dependency clients add a second retry layer:

- Microsoft Graph Client 3.0.7 installs its default retry middleware. A single
  Graph client call can retry a buffered `429`, `503`, or `504` response up to
  three times (four wire attempts), using `Retry-After` when present and its
  default delay/backoff otherwise. A `PUT`, `PATCH`, or `POST` whose content type
  is `application/octet-stream` is not considered buffered by that middleware.
- The Azure REST pipeline used by the Face client has a default maximum of
  three retries per client call. It covers `429`/`503` responses with a usable
  retry-after header, `408`, most `5xx` responses except `501`/`505`, and its
  defined transient system errors, with retry-after or exponential delay.
- Face routes do not inspect the resolved REST response status. Their
  `Erro_004` and `Erro_007` catches run only when the client call rejects after
  its SDK behavior; a resolved non-success response proceeds to body access.

Consequently, "retry" in the route sections means the five-attempt application
helper, and "not application-retried" means only that the helper is absent;
SDK behavior can still create more than one wire attempt. This distinction is
especially important for non-idempotent writes and sends.

### Signed platform-row authorization

The following five routes run `authorizePlatformRow` before their handler:

- `POST /plataforma_v2/CadastroFoto_e_FaceID`
- `POST /plataforma_v2/FaceID`
- `POST /plataforma_v2/refresh`
- `POST /plataforma_v2/updates`
- `POST /plataforma_v2/processa-feedback`

The middleware reads the exact body field `IndexVerificado`, verifies the
canonical signed handle and its four-hour lifetime, and writes the derived
nonnegative workbook row index to `res.locals.platformRowIndex`. Missing,
non-string, numeric, malformed, noncanonical, forged, future-issued, expired,
or otherwise invalid handles fail closed before the handler and its external
calls with status `401`, JSON body `{}`, and content type
`application/json; charset=utf-8`.

The client-visible value remains opaque. Source-observed encoding is a
canonical base64url JSON payload plus an HMAC-SHA256 signature, separated by
`.`. The canonical payload's exact key order is `v`, `purpose`, `rowIndex`,
`iat`, `exp`; `v` is `1`, `purpose` is `platform-row-authorization`, the row is
a nonnegative safe integer, and `exp - iat` is exactly 14,400 seconds.

Verification authenticates the handle itself; it does not read the workbook,
recheck current login status, prove that the row still exists, or rotate the
handle. A valid handle is reusable before, but not at, its exact expiration
time.

Source: [`platform-row-authorization.js` lines 29-90](../platform-row-authorization.js#L29-L90)
and [`platform-row-authorization.js` lines 130-187](../platform-row-authorization.js#L130-L187).
Production authorization wiring is in
[`server.js` lines 96-102](../server.js#L96-L102); the five route placements are
in [`app.js` lines 455-573](../app.js#L455-L573).

`CadastroFoto_e_FaceID` is the one middleware-order exception: its Multer
middleware parses and buffers the multipart request before authorization.

## Source-observed external resources and workbook positions

The common Graph mailbox/user prefix is:

```text
/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be
```

The exact source-observed table bases are:

```text
PLATFORM_TABLE
/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}

CLIENTS_TABLE
/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECQNNRY4S7VCKBF2SOETFSLESSLH/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}

RECOMMENDATIONS_TABLE
/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRAQXJDB7TBYFGKA5YQJXO3YAOS/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}

FEEDBACK_TABLE
/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECXO7I5R6LKLXJD3VWXORUAF7J37/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}

REFERENCE_PHOTO(rowIndex)
/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/root:/2. ENTREGA/1. CONTROLAR PLATAFORMA/PG - FOTOS DE REFERÊNCIA/{rowIndex}.jpg:/content
```

Source: [`app.js` lines 105-113](../app.js#L105-L113),
[`app.js` lines 261-262](../app.js#L261-L262), and
[`app.js` lines 455-568](../app.js#L455-L568).

Source-observed positional layouts used by multiple routes:

| Resource | Width/positions used by source |
|---|---|
| Platform | Update arrays are normally initialized to width 22. Reads use `0` full name, `1` first name, `2` email/login, `3` password, `4` FaceID status, `5` photo-registered flag, `6` access deadline, `7` login status, `8` completed-topic count, `9` unused here, `10..19` module grades, `20` accumulated grade, `21` certificate ID. |
| Recommendations | Width 13: `0` benefited company, `1` recommender full name, `2` calculated first name, `3` recommender email, `4` date/time, `5` recommended company, `6` professional, `7` WhatsApp, `8` stage, `9` status, `10` update date/time, `11` next-contact date/time, `12` participant count. |
| Clients | Writes use width 13; the exact populated positions are described under `processa-formulario`. |
| Feedback | Appends use width 9 in the exact order described under `processa-feedback`. |

`null` cells in update/append arrays are part of the current positional contract
and remain preserved by the application/server seam; they represent cells
intentionally not set by that operation, including calculated or manually
managed cells.

Excel date conversion for login/refresh subtracts 25,569 days, constructs a
JavaScript `Date`, formats it with locale `pt-BR` and two-digit day, short month,
and numeric year, removes `de` and periods, collapses whitespace, and joins the
result with `/`. That locale/runtime-dependent formatted string is the current
wire value.

## Route inventory

Every `res.json(...)` response below has content type
`application/json; charset=utf-8`. An empty JSON object is the two-byte body
`{}`. Unless stated otherwise, a route has no path parameters, query inputs,
multipart inputs, or route-specific authorization.

### `POST /landingpage/solicitacaoorcamento`

Source: [`app.js` lines 86-95](../app.js#L86-L95).

- **Input:** JSON object with exact, case-sensitive keys
  `Solicitante_NomeCompleto`, `Solicitante_Email`,
  `Solicitante_Telefone`, `Solicitante_Cargo`,
  `Solicitante_NomeEmpresa`, `Solicitante_CNPJ`,
  `Solicitante_NúmerodeParticipantes`, and `Solicitante_Observações`. There is
  no validation; absent values are interpolated as `undefined`, and extra keys
  are ignored.
- **Middleware:** global middleware only; public route.
- **Calls/order:** one Graph `POST` to the common user's `/sendMail`, through
  the five-attempt retry helper. It sends HTML to the internal-contact role at
  `contato@machadogestao.com`, with subject
  `Machado - Nova Solicitação de Orçamento`; there is no CC/BCC. Request
  values are interpolated into that HTML without escaping.
- **Responses:** mail success is `200 {}`. Exhausted mail failure is `500 {}`.
  No `Erro_XXX` is returned. Other failures have no explicit unexpected-error
  contract.
- **Partial success/idempotency:** the mail precedes the response. There is no
  dedupe key; repeat requests resend, and an ambiguous failed/retried send can
  duplicate mail.
- **Consumer:** frozen
  [`apps/quote-request/main.js`](https://github.com/IvyRoom/sistemas/blob/c68f361de054a936b7a6871d82d75a1cdb457c97/apps/quote-request/main.js#L188-L249)
  sends all eight keys, treats any 2xx as success, and does not parse either
  response body. Non-2xx/network failure becomes frontend `Erro_000` behavior.

### `POST /conecta/processa-recomendacao`

Source: [`app.js` lines 105-214](../app.js#L105-L214).

- **Input:** exact JSON keys `recommenderFullName`, `benefitedCompany`,
  `recommendedCompany`, `recommendedProfessional`, and
  `recommendedWhatsapp`. Every value must be a nonempty string after trimming.
  Trimmed WhatsApp must match `+XX XX XXXXX-XXXX` exactly
  (`/^\+\d{2} \d{2} \d{5}-\d{4}$/`). Invalid input returns
  `400 {"error":"Erro_014"}` before external calls.
- **Middleware:** global middleware only; public route. Name plus benefited
  company, carried in the public link by the current consumer, are the only
  recommender match gate.
- **Calls/order:** (1) retry `GET RECOMMENDATIONS_TABLE/rows`; exhausted failure
  is `500 Erro_015`. (2) Normalize name/company with string coercion, trim,
  collapsed internal whitespace, and lowercase; no matched row is
  `404 Erro_016`. (3) Detect an existing recommendation by the normalized
  recommended-company/professional/WhatsApp tuple across the matched rows. A
  duplicate skips the workbook write but continues to email. (4) For a new
  recommendation, update the first free matched slot or append a row, without
  application retry. Mutation failure is `500 Erro_017`, before mail. (5) Retry the internal
  notification mail to `contato@machadogestao.com`, then retry the confirmation
  mail to the recommender email at workbook index `3`. Either exhausted mail
  failure is `500 Erro_018`. Full success is `200 {}`.
- **Mail details:** the internal subject is
  `Machado Conecta - Nova Recomendação Recebida`; the recommender subject is
  `Machado Conecta - Recomendação Registrada`. Dynamic values are HTML-escaped.
  The recommender greeting uses trimmed workbook index `2` unless it is empty or
  `'-'`, then falls back to the first whitespace-delimited token of index `1`.
- **Source-observed workbook payload:** a free slot has `'-'` at all indexes
  `4..12`. A 13-cell `null`-initialized array sets `4`, `10`, and `11` to the
  same São Paulo "now" Excel serial; `5`, `6`, `7` to trimmed request values;
  `8` to `1. REALIZAR CONTATO INICIAL`; and `9` to `A INICIAR`. Free-slot update
  leaves `0`, `1`, `2`, `3`, `12` as `null` placeholders. Append additionally
  copies indexes `0`, `1`, `3` from the first matched row and sets `12` to
  `'-'`; calculated first-name index `2` remains `null`.
- **Partial success/idempotency:** a workbook mutation may be committed despite
  an ambiguous `Erro_017`. A committed workbook row remains when mail fails.
  If internal mail succeeds and confirmation fails, the first mail and workbook
  are partial successes. A repeated identical request skips the workbook write
  but sends both mails again; retry ambiguity can also duplicate either mail.
- **Consumer:** frozen
  [`apps/conecta/referral-form/main.js`](https://github.com/IvyRoom/sistemas/blob/c68f361de054a936b7a6871d82d75a1cdb457c97/apps/conecta/referral-form/main.js#L19-L24)
  maps URL query `ncr`/`eb` to the first two body fields, sends all five fields,
  requires JSON on every response, and recognizes `Erro_014` through
  `Erro_018`.
- **Other errors:** returned-row shape/content errors and failures outside the
  local catches have no explicit unexpected-error contract.

### `POST /clientes/processa-formulario`

Source: [`app.js` lines 225-344](../app.js#L225-L344).

- **Input:** exact top-level JSON keys `company`, `shippingAddress`,
  `legalRepresentative`, `adminAssistant`, `participants`. `company` contains
  `legalName`, `cnpj`, and `address`; both addresses expose `postalCode`,
  `street`, `number`, `complement`, `neighborhood`, `city`, `state`, while
  `shippingAddress` also carries ignored key `useCompanyAddress`. Each person
  object exposes `fullName`, `cpf`, `role`, `areaCode`, `whatsapp`, `email`.
- **Validation:** `company.legalName` must be a nonempty string;
  `participants` must be an array with 1 through 25 entries; and every entry
  must have nonempty-string `fullName`, `email`, and `cpf`. Failure is
  `400 {"error":"Erro_013"}`. Other fields and formats are not validated.
- **Middleware:** global middleware only; public route.
- **Calls/order:** (1) retry `GET PLATFORM_TABLE/rows` (`500 Erro_001` on
  exhaustion); (2) retry `GET CLIENTS_TABLE/rows` (`500 Erro_011`); (3) if any
  new platform rows exist, append them without application retry
  (`500 Erro_008`); (4) if any new client rows exist, append them without
  application retry (`500 Erro_010`); (5)
  retry one HTML notification mail to the internal-contact role at
  `contato@machadogestao.com`, with subject
  `Machado: novo Formulário de Informações Iniciais preenchido`
  (`500 Erro_012`). Dynamic values are HTML-escaped. Full success is `200 {}`.
- **Dedupe:** platform rows dedupe by trimmed/lowercased email at index `2`;
  client rows independently dedupe by digits-only CPF at index `4`. The sets
  include earlier participants in the same request, so the two workbooks can
  receive different participant subsets.
- **Source-observed platform payload:** width 22, initially all `null`; `0` full
  name, `2` email, `3` random integer in
  `[100000000000, 1000000000000)`, `4` `Ativo`, `5` `Não`, `6` local-calendar
  today plus 60 days as an Excel serial, `8` `0`, `10..19` `0`, and `21` a new
  certificate ID. Indexes `1`, `7`, `9`, `20` stay `null`. Certificate format
  is `FMG-XXXX-XXXX`, using eight random characters from exact alphabet
  `0123456789ABCDEFGHJKMNPQRSTVWXYZ`, unique against observed index `21` values
  and IDs generated in this request.
- **Source-observed client payload:** width 13, initially all `null`; `0`
  company legal name, `3` participant full name, `4` participant CPF as
  supplied, `5` shipping street, `6` a Number only if shipping number matches
  `^\d+$` (otherwise the original value), `7` shipping complement or `'-'`,
  `8` neighborhood, `9` city, `10` state, `12` postal code. Indexes `1`, `2`,
  `11` remain `null`.
- **Partial success/idempotency:** writes are sequential and nontransactional.
  Platform may persist before a client-write failure; both may persist before a
  mail failure. Later whole-request retries re-read and normally dedupe visible
  rows, but notification mail always runs and can duplicate. Same-email/new-CPF
  and same-CPF/new-email inputs can intentionally split outcomes.
- **Consumer:** frozen
  [`apps/client-intake/main.js`](https://github.com/IvyRoom/sistemas/blob/c68f361de054a936b7a6871d82d75a1cdb457c97/apps/client-intake/main.js#L325-L397)
  sends the shape above, requires JSON on every response, and recognizes
  `Erro_001`, `Erro_008`, `Erro_010`, `Erro_011`, `Erro_012`, and `Erro_013`.
- **Other errors:** processing failures outside the listed catches have no
  explicit unexpected-error contract.

### `POST /clientes/liberacao-acesso-plataforma`

Source: [`app.js` lines 354-424](../app.js#L354-L424).

- **Input:** no body, query, path, or multipart value is read. The global JSON
  parser can still reject malformed JSON before the route.
- **Middleware:** global middleware only; public operational route.
- **Response:** `200` is sent immediately with an empty body by
  `res.status(200).send()`. The route sets no content type. This response occurs
  before every Graph read or mail.
- **Calls/order:** after responding, it performs an uncaught
  `GET PLATFORM_TABLE/rows`. On success it schedules email work after 1,000 ms.
  Neither that read nor the seven sequential `/sendMail` calls is wrapped in
  the application's retry helper. The loop waits 2,000 ms after each mail,
  including the last.
- **Source-observed workbook/recipients:** hard-coded labels 39 through 45 are
  converted to returned-array indexes `35..41` inclusive. Each row supplies
  first name at `1`, recipient email at `2`, and password at `3`. The dynamic
  learner/participant recipient receives HTML access instructions containing
  those credentials and hard-coded legacy customer copy. The subject is
  `Machado | Método Gerencial para Empresas - Instruções de Acesso à Plataforma`.
- **Partial success/idempotency:** the completed HTTP response cannot reflect a
  read or mail failure. A rejected read occurs after response; a rejected timer
  task stops the remaining suffix of emails. Any prefix can be delivered.
  Repeated requests schedule the same seven recipients again, with no dedupe.
- **Consumer:** no in-repository consumer found in the frozen `sistemas`
  snapshot. This is not evidence that the route is removable.
- **Other errors:** there is no explicit error response contract for background
  or unexpected failure.

### `POST /plataforma_v2/login-FaceID`

Source: [`app.js` lines 434-453](../app.js#L434-L453).

- **Input:** exact JSON keys `Usuário_Login` and `Usuário_Senha`; extra keys are
  ignored. There is no backend format validation.
- **Middleware:** global middleware only; public login route.
- **Calls/order:** retry `GET PLATFORM_TABLE/rows`; exhausted failure is
  `500 {"error":"Erro_001"}`. Rows are scanned in returned order. Login must
  strictly equal index `2`; password must strictly equal `index 3 .toString()`.
- **Responses:** no match is
  `401 {"error":"credenciais_inválidas"}`. A match is `200` with exact keys
  `Usuário_Status_FaceID` from index `4`, `Usuário_Foto_Cadastrada` from `5`,
  `Usuário_PrazoAcesso` from converted Excel date `6`, and
  `Usuário_Status_Login` from `7`. Only when index `7` is exactly `Ativo` is a
  newly minted four-hour signed handle included as `IndexVerificado`.
- **Partial success/idempotency:** the workbook call is read-only, but repeated
  active logins mint time-derived handles. Valid inactive credentials
  deliberately return the four-field `200` body without `IndexVerificado`.
- **Consumer:** frozen
  [`plataforma_v2/login/main.js`](https://github.com/IvyRoom/sistemas/blob/c68f361de054a936b7a6871d82d75a1cdb457c97/plataforma_v2/login/main.js#L92-L126)
  parses JSON before checking status, stores the returned fields, branches on
  login status, and uniquely treats `401` as invalid credentials. The inactive
  `200` drives its expired-login UI.
- **Other errors:** malformed workbook rows, including a non-stringifiable
  password cell, have no explicit unexpected-error contract.

### `POST /plataforma_v2/CadastroFoto_e_FaceID`

Source: [`app.js` lines 455-475](../app.js#L455-L475).

- **Input:** `multipart/form-data` with exact text field `IndexVerificado` and
  exact single file field `file`. `multer().single('file')` uses in-memory
  storage and no source-configured size/count limits; it buffers the upload as
  `req.file.buffer`. A missing file reaches an unguarded buffer access after
  authorization.
- **Middleware order:** global CORS -> global JSON parser (skips multipart) ->
  Multer parse/buffer -> signed row authorization -> handler. Invalid handles
  therefore still incur multipart parsing and buffering, then return `401 {}`.
- **Calls/order:** (1) retry `PUT REFERENCE_PHOTO(rowIndex)` with the file bytes
  (`500 Erro_002`); (2) retry update of `PLATFORM_TABLE/rows/itemAt(index=<row>)`
  with a 22-cell array whose only non-`null` value is index `5` = `Sim`
  (`500 Erro_003`); (3) retry Face `POST /detectLivenessWithVerify-sessions`
  (`500 Erro_004`) with multipart parts `VerifyImage` = the uploaded buffer,
  `livenessOperationMode` = `Passive`, and `deviceCorrelationId` = a new UUID.
- **Face status handling:** no resolved REST status is checked. `Erro_004`
  requires a rejected client promise; a resolved non-success response proceeds
  to `body.authToken`/`body.sessionId` extraction. A new UUID is created for
  each application-level Face attempt.
- **Response:** `200` JSON with exact keys
  `Azure_Face_API_LivenessSession_authToken` and
  `Azure_Face_API_LivenessSession_sessionID`, sourced from Face response fields
  `authToken` and `sessionId`.
- **Partial success/idempotency:** the photo can persist before workbook
  failure; photo and workbook flag can persist before Face failure. Photo
  upload and the fixed-value workbook update are repeatable at the same row,
  but an ambiguous Face retry can create multiple sessions and only the
  returned one is observable.
- **Consumer:** frozen
  [`plataforma_v2/cadastro/main.js`](https://github.com/IvyRoom/sistemas/blob/c68f361de054a936b7a6871d82d75a1cdb457c97/plataforma_v2/cadastro/main.js#L73-L117)
  appends `IndexVerificado` then `file` to browser `FormData`, requires JSON,
  and reads both success keys. It recognizes `Erro_002` through `Erro_004`;
  current `401 {}` becomes its generic error path.
- **Other errors:** Multer errors use the default Express error response;
  missing-file and unexpected handler failures have no explicit route JSON
  contract.

### `POST /plataforma_v2/FaceID`

Source: [`app.js` lines 477-494](../app.js#L477-L494).

- **Input/authorization:** JSON body with exact signed field
  `IndexVerificado`; signed authorization runs before the handler and invalid
  input returns `401 {}`.
- **Calls/order:** (1) retry `GET REFERENCE_PHOTO(rowIndex)` (`500 Erro_005`);
  (2) retry Face `POST /detectLivenessWithVerify-sessions` (`500 Erro_004`) with
  multipart parts `VerifyImage` = downloaded photo,
  `livenessOperationMode` = `Passive`, and a new UUID
  `deviceCorrelationId`.
- **Face status handling:** no resolved REST status is checked. `Erro_004`
  requires a rejected client promise, and each application-level attempt uses a
  newly generated UUID.
- **Response:** `200` JSON with exact keys
  `Azure_Face_API_LivenessSession_authToken` and
  `Azure_Face_API_LivenessSession_sessionID`.
- **Partial success/idempotency:** photo retrieval is read-only. Face session
  creation is not idempotent; ambiguous retries or repeated requests can create
  more than one session.
- **Consumer:** frozen
  [`plataforma_v2/login/main.js`](https://github.com/IvyRoom/sistemas/blob/c68f361de054a936b7a6871d82d75a1cdb457c97/plataforma_v2/login/main.js#L152-L195)
  sends the handle, requires JSON, reads both success keys, and recognizes
  `Erro_004`/`Erro_005`; current `401 {}` becomes its generic error path.
- **Other errors:** unexpected downloaded-photo or Face-response shape failures
  have no explicit error contract.

### `GET /plataforma_v2/FaceID_resultado/:Azure_Face_API_LivenessSession_sessionID`

Source: [`app.js` lines 496-510](../app.js#L496-L510).

- **Input:** exact path parameter slot
  `Azure_Face_API_LivenessSession_sessionID`. No body/query value is read.
- **Middleware:** global middleware only; this Face-result route is public.
- **Calls/order:** retry Face
  `GET /detectLivenessWithVerify-sessions/{sessionId}`, passing the decoded path
  parameter as `{sessionId}`. Exhausted failure is `500 Erro_007`.
- **Face status handling:** no resolved REST status is checked. `Erro_007`
  requires a rejected client promise; a resolved non-success response proceeds
  to the nested result extraction.
- **Response:** `200` JSON with exact keys
  `Azure_Face_API_LivenessSession_LivenessDecision` from
  `body.results.attempts[0].result.livenessDecision`,
  `Azure_Face_API_LivenessSession_MatchConfidence` from
  `.verifyResult.matchConfidence`, and
  `Azure_Face_API_LivenessSession_MatchDecision` from
  `.verifyResult.isIdentical`.
- **Partial success/idempotency:** read-only Face result lookup; retry is
  observationally idempotent, subject to the external session state changing.
- **Consumers:** frozen login
  [`main.js`](https://github.com/IvyRoom/sistemas/blob/c68f361de054a936b7a6871d82d75a1cdb457c97/plataforma_v2/login/main.js#L186-L218)
  and registration
  [`main.js`](https://github.com/IvyRoom/sistemas/blob/c68f361de054a936b7a6871d82d75a1cdb457c97/plataforma_v2/cadastro/main.js#L108-L148)
  require JSON and consume the decisions; registration also displays match
  confidence. Both send `Content-Type: application/json` despite having no
  body, and recognize `Erro_007`.
- **Other errors:** an absent/changed Face attempts/result shape has no explicit
  unexpected-error contract.

### `POST /plataforma_v2/refresh`

Source: [`app.js` lines 512-541](../app.js#L512-L541).

- **Input/authorization:** JSON body with exact signed field
  `IndexVerificado`; invalid authorization is `401 {}`.
- **Calls/order:** retry `GET PLATFORM_TABLE/rows`; exhausted failure is
  `500 Erro_001`. The verified row index selects the returned row.
- **Response:** `200` JSON with exact keys `Usuário_NomeCompleto` (`0`),
  `Usuário_PrimeiroNome` (`1`), `Usuário_Email` (`2`),
  `Usuário_PrazoAcesso` (converted `6`), `Usuário_Status_Login` (`7`),
  `Usuário_Formação_NúmeroTópicosConcluídos` (`8`),
  `Usuário_Formação_NotaMódulo1` through
  `Usuário_Formação_NotaMódulo10` (`10..19`),
  `Usuário_Formação_NotaAcumulado` (`20`), and
  `Usuário_Formação_CertificadoID` (`21`).
- **Partial success/idempotency:** read-only and repeatable; the response can
  change with workbook state. The handle is neither refreshed nor returned.
- **Consumer:** frozen
  [`plataforma_v2/estudo/main.js`](https://github.com/IvyRoom/sistemas/blob/c68f361de054a936b7a6871d82d75a1cdb457c97/plataforma_v2/estudo/main.js#L113-L137)
  sends the handle, requires JSON, and reads every success key. It recognizes
  `Erro_001`; current `401 {}` becomes its generic error path.
- **Other errors:** a verified index no longer present in returned workbook data
  or a short row has no explicit unexpected-error contract.

### `POST /plataforma_v2/updates`

Source: [`app.js` lines 543-558](../app.js#L543-L558).

- **Input/authorization:** exact JSON keys `IndexVerificado`,
  `TipoAtualização`, `NúmeroTópicosConcluídos`, `NúmeroMódulo`, `NotaTeste`.
  Only the handle is validated. Invalid authorization is `401 {}`.
- **Source-observed workbook payload:** an update begins as a 22-cell
  `null`-initialized array and
  always sets index `8` to client-supplied `NúmeroTópicosConcluídos`. Only when
  `TipoAtualização` is exactly
  `NúmeroTópicosConcluídos-e-NotaTeste` does it also assign client-supplied
  `NotaTeste` at JavaScript index `NúmeroMódulo + 9`. Current numeric modules
  1 through 10 keep width 22 and target indexes 10 through 19; other types or
  values follow JavaScript addition/property semantics and can expand the
  array or assign a non-index property. No type, range, monotonicity, ownership,
  or grade validation is performed.
- **Calls/order:** retry update of
  `PLATFORM_TABLE/rows/itemAt(index=<verified row>)`. Exhausted failure is
  `500 Erro_008`; success is `200 {}`.
- **Partial success/idempotency:** the same well-formed values are a fixed-value
  update and normally repeatable, but retry ambiguity and unconstrained client
  values remain part of current behavior.
- **Consumer:** frozen
  [`plataforma_v2/estudo/main.js`](https://github.com/IvyRoom/sistemas/blob/c68f361de054a936b7a6871d82d75a1cdb457c97/plataforma_v2/estudo/main.js#L794-L806)
  sends either exact type above with a numeric module/grade or type
  `NúmeroTópicosConcluídos` with `NúmeroMódulo: "n/a"` and
  `NotaTeste: "n/a"`. It requires success JSON but reads no success key,
  recognizes `Erro_008`, and treats current `401 {}` generically.
- **Other errors:** unexpected serialization/index behavior has no explicit
  route error contract.

### `POST /plataforma_v2/processa-feedback`

Source: [`app.js` lines 560-573](../app.js#L560-L573).

- **Input/authorization:** exact JSON keys `IndexVerificado`,
  `NúmeroTópicosConcluídos`, `Usuário_NomeCompleto`, `Usuário_Email`,
  `Feedback_DataPreenchimento`, `NúmeroMódulo`, `Feedback_TamanhoMódulo`,
  `Feedback_QualidadeConteúdo`, `Feedback_QualidadePlataforma`,
  `Feedback_QualidadeMateriaisImpressos`, `Feedback_Comentários`. Only the
  signed handle is validated; invalid authorization is `401 {}`.
- **Calls/order:** (1) retry update of the verified platform row with a 22-cell
  array whose only non-`null` value is client-supplied completed topics at index
  `8` (`500 Erro_008`); (2) retry append to `FEEDBACK_TABLE/rows/add` with the
  nine client-supplied fields in this exact order: full name, email, fill date,
  module number, module-size score, content-quality score, platform-quality
  score, printed-material-quality score, comments (`500 Erro_009`). Success is
  `200 {}`.
- **Partial success/idempotency:** progress can persist before feedback append
  failure. The append itself is retry-wrapped and has no dedupe key, so an
  ambiguous failure or repeated request can duplicate feedback rows. Identity
  and feedback fields are not derived from the verified workbook row.
- **Consumer:** frozen
  [`plataforma_v2/estudo/main.js`](https://github.com/IvyRoom/sistemas/blob/c68f361de054a936b7a6871d82d75a1cdb457c97/plataforma_v2/estudo/main.js#L962-L986)
  supplies all fields from client state/DOM, requires success JSON but reads no
  success key, recognizes `Erro_008`/`Erro_009`, and treats current `401 {}`
  generically.
- **Other errors:** failures outside the two catches have no explicit route
  error contract.

### `GET /ezdrm-playready-authorization-url`

Source: [`app.js` lines 575-583](../app.js#L575-L583).

- **Input:** exact, case-sensitive query keys `token` and `CustomData`. Missing
  or falsy values become empty strings. No body/path/multipart value is read.
- **Middleware:** global middleware only; public route.
- **Calls/order:** no external call.
- **Response:** status `200`, content type `text/html; charset=utf-8`, and exact
  text body, with only the two substituted values URL-encoded:

  ```text
  p1=5&p2=&p3=&p4=1&p5=0&p6=1&p7=0&p8=0&token=<encodeURIComponent(token)>&CustomData=<encodeURIComponent(CustomData)>
  ```

  Query spelling and output parameter order/casing are compatibility contracts.
- **Partial success/idempotency:** deterministic, side-effect free, and
  repeatable for the same parsed query.
- **Consumer:** no in-repository consumer found in the frozen `sistemas`
  snapshot. A separate direct third-party EZDRM URL exists there, but it does
  not call this backend route. Neither fact implies removability.
- **Other errors:** there is no route-specific JSON error contract; an
  unexpected synchronous error would use Express's default handler.

### `POST /plataforma_v2/statusreport`

Source: [`app.js` lines 585-597](../app.js#L585-L597).

- **Input:** exact JSON keys `linha_inicial` and `linha_final`; no validation or
  authorization. JavaScript `slice(linha_inicial, linha_final + 1)` coercion and
  indexing are current behavior.
- **Middleware:** global middleware only; public reporting route.
- **Calls/order:** retry `GET PLATFORM_TABLE/rows`; exhausted failure is
  `500 Erro_001`.
- **Source-observed response projection:** each selected row becomes a 14-value
  array `[source[0], source[8], ...source[10..21]]`: full name, completed-topic
  count, ten module grades, accumulated grade, certificate ID. The code passes
  exact JavaScript expression `linha_final + 1` as the second argument to
  `slice`. A numeric `linha_final` therefore makes the end inclusive; a string
  first concatenates (for example, `"3" + 1` becomes `"31"`) and is then
  coerced by `slice`. The frozen consumer sends parsed numbers.
- **Response:** `200` JSON with exact key
  `Dados_Extraídos_BD_Plataforma` containing the array of projected rows.
- **Partial success/idempotency:** read-only and repeatable against a stable
  workbook snapshot.
- **Consumer:** frozen
  [`plataforma_v2/statusreport/main.js`](https://github.com/IvyRoom/sistemas/blob/c68f361de054a936b7a6871d82d75a1cdb457c97/plataforma_v2/statusreport/main.js#L168-L204)
  derives the two body values from page query keys `li`/`lf`, requires JSON, and
  consumes projected indexes `0..12`; it ignores returned certificate index
  `13`. It recognizes `Erro_001`.
- **Other errors:** malformed returned rows and failures after the successful
  read have no explicit unexpected-error contract.

### `GET /validacaocertificados/:Solicitante_CertificadoID`

Source: [`app.js` lines 609-635](../app.js#L609-L635).

- **Input:** exact path parameter slot `Solicitante_CertificadoID`. Its decoded
  value is string-coerced, outer-trimmed, and uppercased. No body or query input
  is used; the route is public.
- **Calls/order:** retry `GET PLATFORM_TABLE/rows`; exhausted failure is
  `500 Erro_001`. Match the first row whose source-observed index `21`, after
  the same string/trim/uppercase normalization, equals the requested ID.
- **Threshold:** a present path segment that normalizes to empty, or any present
  ID with no match, is `200 {"Certificado_Válido":false}`. A request ending at
  `/validacaocertificados/` has no parameter segment, does not match this route,
  and uses the default 404 instead. The
  accumulated score comes from index `20`: nonfinite becomes `0`; values `<= 1`
  are multiplied by 100; values `> 1` are treated as already-percent. The
  unrounded normalized score must be at least 70. Below 70 is the same `200`
  false response; exactly 70 is valid.
- **Valid response:** `200` JSON with exact keys `Certificado_Válido: true`,
  `Titular_NomeCompleto` from index `0`, `Acumulado_Percentual` rounded with
  `Math.round`, and `Certificado_ID` from normalized index `21`.
- **Partial success/idempotency:** read-only and repeatable against stable data.
- **Consumer:** frozen
  [`apps/certificate-validation/main.js`](https://github.com/IvyRoom/sistemas/blob/c68f361de054a936b7a6871d82d75a1cdb457c97/apps/certificate-validation/main.js#L51-L90)
  trims/uppercases then URL-encodes the ID, requires JSON, reads the valid flag,
  name, and numeric percentage, and recognizes `Erro_001`. It does not
  currently render `Certificado_ID`, which remains part of the backend response.
- **Other errors:** processing failures after a successful read have no
  explicit unexpected-error contract.

## Compatibility contracts versus known-risk legacy behavior

The route inventory above is the compatibility contract. The application/server
split and its acceptance coverage preserve the following known risks; recording
them is not an implied recommendation or authorization to fix them during later
domain extraction:

- Exactly five routes use signed `IndexVerificado` authorization, and every
  invalid-handle class fails closed with `401 {}` before handler external calls.
- `CadastroFoto_e_FaceID` buffers Multer multipart data before it authorizes the
  signed row.
- Valid inactive login credentials return `200` with status/deadline fields and
  no `IndexVerificado`; the current consumer uses that response for its expired
  account UI.
- `FaceID_resultado` and `statusreport` are public. The latter accepts
  client-selected row bounds and returns names, progress, grades, and certificate
  IDs.
- `liberacao-acesso-plataforma` is a public operational route that responds
  before work, uses a fixed row range, sends credentials, and cannot report
  background partial failure.
- `updates` trusts client-controlled progress and grade fields.
  `processa-feedback` trusts client-controlled progress, identity, date, module,
  ratings, and comments even though its row handle is signed.
- The DRM route's `text/html` body, fixed parameter string, lower-case query
  `token`, upper-case `CustomData`, output order, URL encoding, and empty-value
  behavior are exact compatibility details.
- Certificate IDs are outer-trimmed and case-normalized, invalid certificates
  return `200` false, fractional/whole percentages are both accepted, and the
  inclusive validity threshold is 70% before rounding.
- Non-idempotent mail, Face-session creation, and feedback append operations are
  retry-wrapped in several routes. Their ambiguous and partial-success behavior
  is part of this characterization.
- The source has no explicit unexpected-error response contract for async
  failures outside local catches. In particular, workbook/result shape errors,
  missing multipart file data, and post-response access-release failures must
  not be silently reclassified as existing JSON errors during structural work.
- No `sistemas` consumer was found for access release or the backend EZDRM
  helper. Neither endpoint is therefore considered removable.

## Acceptance-test coverage for the application/server seam

The Node.js 24 built-in suite under [`test/`](../test/) exercises this matrix as
behavior, not source text. HTTP cases call `createApp` with a fixed synthetic
32-byte signing key, recording Graph and Face fakes, and deterministic clock,
random, UUID, sleep, and scheduling hooks. They use native `fetch`, `FormData`,
and `Blob` against ephemeral loopback listeners and close every listener and
timer. Import-safety coverage requires `app.js` and `server.js` without
production configuration. A child-process IISNode simulation proves that the
interceptor-style entry reaches signing-key validation while dotenv and the
production SDKs are replaced with inert/forbidden test doubles; it does not
construct real SDK clients, listen, or call external services. Known unhandled
async failure shapes remain documented above rather than being triggered
in-process or assigned new HTTP contracts.

### Bootstrap and middleware acceptance

| Area | Acceptance coverage |
|---|---|
| Configuration/bootstrap | Importing `app.js` and `server.js` performs no startup or external work; direct Node and IISNode-preserved entry paths are recognized; an IISNode-style child process reaches signing-key validation without loading production SDKs; invalid signing-key configuration fails before dependency construction or listening; explicit production startup uses `PORT || 3000`, preserves SDK construction defaults, and starts Graph acquisition only after the listener call. |
| Graph token lifecycle | First acquisition has no readiness gate; success scheduling uses expiry-minus-five-minutes with a 60-second minimum; failures schedule 2, 4, 8, 16, 32, then 60 seconds; success resets the next failure to 2 seconds; each replacement clears the prior timer, cleanup clears the final timer, and an in-flight acquisition cannot schedule after cleanup. |
| Global middleware | Assert CORS headers/preflight, JSON parsing before static/routes, 100 KiB/strict-parser failures, `/img` -> `/img/` redirect, asset GET/HEAD, directory-index miss, and other static misses in the current order. |
| Routing/defaults | Assert canonical 14-route registration, case-insensitive/trailing-slash matching, default 404 HTML, route JSON content types, DRM HTML content type, and access-release empty 200. |
| Authorization | For each of the five protected routes, missing, numeric, malformed, forged, and expired handles return `401 {}` with no handler dependency calls; valid handles expose only their verified row index. Assert Multer remains before auth only on registration. |

### Route-by-route acceptance

| Route | Happy-path assertions | Failure, ordering, partial-success, and repeat assertions |
|---|---|---|
| `POST /landingpage/solicitacaoorcamento` | Exact eight-key values reach the HTML mail; internal recipient; `200 {}`. | Five-attempt mail exhaustion is `500 {}`; absent values remain unvalidated; repeat/ambiguous send can duplicate. |
| `POST /conecta/processa-recomendacao` | Normalized recommender match; exact free-slot update and append 13-cell arrays; duplicate skips write; internal mail precedes confirmation; `200 {}`. | Every invalid-input branch (`Erro_014`), read (`015`), not-found (`016`), write without application retry (`017`), and each mail partial failure (`018`); committed row and prior mail remain; repeat duplicate still mails. |
| `POST /clientes/processa-formulario` | Exact nested payload; independent email/CPF dedupe; exact 22/13-cell arrays and nulls; platform append -> client append -> internal mail; `200 {}`. | Participant/count validation (`013`); read errors (`001`,`011`); write errors (`008`,`010`); mail error (`012`); each prior effect persists; whole-request retry skips visible duplicate rows but can resend mail. |
| `POST /clientes/liberacao-acesso-plataforma` | Empty `200` completes before Graph; exact source indexes `35..41`; first-name/email/password selection; 1-second schedule and 2-second pacing. | Read/mail have no application retry; caller never receives their failure; a mail failure stops the remaining suffix; repeated request resends all. |
| `POST /plataforma_v2/login-FaceID` | Exact credential match and four response fields; active login adds a valid four-hour handle; inactive login omits it. | Graph failure `Erro_001`; invalid credentials `401 credenciais_inválidas`; inactive remains `200`; row-shape failure is not assigned an existing JSON error. |
| `POST /plataforma_v2/CadastroFoto_e_FaceID` | Multipart field/file parsing, verified row, photo PUT -> 22-cell `Sim` update -> exact Face session parts; exact two-key `200`. | Multer-before-auth; `401 {}`; `Erro_002/003/004`; prior photo/flag persist; missing file has no defined JSON error; ambiguous Face retry may create multiple sessions. |
| `POST /plataforma_v2/FaceID` | Verified-row photo GET -> exact Face session parts; exact two-key `200`. | `401 {}`, `Erro_005`, `Erro_004`; Face retry/repeat can create multiple sessions. |
| `GET /plataforma_v2/FaceID_resultado/:Azure_Face_API_LivenessSession_sessionID` | Exact Face path parameter forwarding and three-key projection. | Public access; five-attempt `Erro_007`; missing attempt/result shape has no defined JSON error. |
| `POST /plataforma_v2/refresh` | Verified index selects exact source positions and returns all 18 exact Unicode keys. | `401 {}`; `Erro_001`; missing row/short width has no defined JSON error; no replacement handle is returned. |
| `POST /plataforma_v2/updates` | Exact 22-cell progress-only and progress-plus-grade arrays; verified row; `200 {}`. | `401 {}`; `Erro_008`; preserve lack of value/type/range validation and fixed-value retry behavior. |
| `POST /plataforma_v2/processa-feedback` | Verified-row 22-cell progress update precedes exact nine-value feedback append; `200 {}`. | `401 {}`, `Erro_008`, `Erro_009`; progress persists before append failure; ambiguous append retry/repeat can duplicate; identity remains client-supplied. |
| `GET /ezdrm-playready-authorization-url` | Exact query casing, fixed text/order, URL encoding, empty defaults, `text/html; charset=utf-8`. | No external call; wrong-case keys behave as absent; preserve deterministic `200`. |
| `POST /plataforma_v2/statusreport` | Numeric bounds use an inclusive end and exact 14-value row projection under the exact response key. | Public access; preserve exact `linha_final + 1` then `slice` coercion for nonnumeric types; no validation; `Erro_001`; malformed rows have no defined JSON error. |
| `GET /validacaocertificados/:Solicitante_CertificadoID` | Trim/uppercase match; present normalized-empty/not-found ID returns false; fractions and whole percentages; exactly 70 valid; exact valid response keys. | Missing path segment uses default 404; `Erro_001`; nonfinite/below-70 false; threshold occurs before rounding; a present invalid ID remains `200`, not 404. |

## Frozen consumer reconciliation

The frozen `sistemas` snapshot has consumers for 12 of the 14 routes. Except
for quote submission, every found consumer parses the response as JSON even
when it ignores all success keys. Only login explicitly distinguishes HTTP
`401`; the other protected-route consumers turn the current `401 {}` into their
generic communication-error UI.

No in-repository consumer was found for:

- `POST /clientes/liberacao-acesso-plataforma`
- `GET /ezdrm-playready-authorization-url`

The search result is inventory evidence only and does not authorize removal.

## Deployment effect

The push workflow ignores `**/*.md`, `docs/**`, and `test/**`, but it does not
ignore `app.js`, `server.js`, or `package.json`. Because the application/server
split changes those production files, merging it to `main` triggers the Node.js
24 build and test job and deployment of the resulting repository artifact to
the Production slot. This merge is production-affecting even though its
documentation and test paths are individually ignored. Manual
`workflow_dispatch` also remains available. Source:
[`main_plataforma-backend-v3.yml` lines 6-17](../.github/workflows/main_plataforma-backend-v3.yml#L6-L17).
