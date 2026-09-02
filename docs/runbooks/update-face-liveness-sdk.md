# Update the Face Liveness Web SDK

Current vendored version: **1.5.0**

The browser SDK is vendored in
`sistemas/apps/learning-platform/azure-ai-vision-face-ui`. Both
`apps/learning-platform/login` and `apps/learning-platform/photo-registration`
load that copy through the public `/plataforma/azure-ai-vision-face-ui/` path.
The Backend repository owns only the trusted package-acquisition automation and
this version record; the private UI packages are never application dependencies.

This process intentionally remains separate from ordinary npm Dependabot.

## Official references

- [SDK release notes](https://github.com/Azure-Samples/azure-ai-vision-sdk/releases)
- [What's new in Azure Face](https://learn.microsoft.com/en-us/azure/ai-services/face/whats-new-face)
- [Understand Face Liveness SDK versions](https://learn.microsoft.com/en-us/azure/ai-services/face/sdk/understand-the-liveness-sdk-versions)
- [Web SDK sample and installation guidance](https://github.com/Azure-Samples/azure-ai-vision-sdk/blob/main/samples/web/README.md)
- [Get Client Assets Access Token API](https://learn.microsoft.com/en-us/rest/api/face/liveness-session-operations/get-client-assets-access-token?view=rest-face-v1.3-preview)

## Trusted automation boundary

`.github/workflows/face-sdk-update.yml` checks for a release every Tuesday at
12:23 UTC and also accepts a manual exact-version request. It runs only from the
Backend default branch. The job uses the non-production
`face-sdk-automation` environment with deployment records disabled; it does not
use Azure login, OIDC, either production environment, deployment credentials, or
an application deployment action.

Configure that environment once, restrict it to `main`, and store:

- `AZURE_FACE_API_ENDPOINT` and `AZURE_FACE_API_KEY` as environment secrets;
- `FACE_SDK_APP_PRIVATE_KEY` as an environment secret; and
- `FACE_SDK_APP_CLIENT_ID` as an environment variable.

The private GitHub App must be installed only on `IvyRoom/backend` and
`IvyRoom/sistemas`. Grant it only repository contents read/write and pull request
read/write; metadata read is implicit. Do not grant Actions, deployments,
environments, secrets, workflows, administration, bypass, or Azure permissions.
GitHub Apps do not expose separate approval or merge permissions to withhold from
these required write scopes, so keep the App out of branch-protection bypass
lists and require the ordinary human review and green-check policy. The workflow
creates drafts and contains no approve, ready-for-review, merge, or auto-merge
API call.

Before enabling the schedule, rotate or retire the old repository-level
`AZURE_FACE_API_NPM_TEMPORARY_TOKEN`. A private package credential must exist
only in the main-restricted `face-sdk-automation` environment. Pull-request jobs
must not reference that environment or any of its secrets.

The tracked Backend `.npmrc` contains only the private scope and an environment
variable placeholder. Never commit a package password, Face key, registry
response, GitHub App key, or installation token.

## What the updater does

1. It first requires the Backend version record and the Sistemas vendor manifest
   to agree, then reads the latest public `@azure/ai-vision-face-ui` version or
   validates the exact version supplied by a manual run. Drift fails before any
   network or package work. An equal version exits without requesting a private
   credential; downgrades are rejected.
2. Only when a newer version exists, it exchanges the Face resource key for the
   short-lived `base64AccessToken` required by the private Azure Artifacts feed.
   The response and child-process output are captured and never printed.
3. It installs `@azure/ai-vision-face-ui` under a new
   `$RUNNER_TEMP/face-sdk-*` directory with lifecycle scripts, lockfile writes,
   audit, funding output, and development dependencies disabled. npm's cache and
   temporary `.npmrc` stay inside the same directory.
4. It verifies the exact package identities
   `@azure/ai-vision-face-ui` and `@azure-ai-vision-face/ui-assets`, the requested
   package version, the loader's `clientSDKversion`, required JavaScript and
   WebAssembly engines, images, localization, and the absence of symlinks or
   special files. Windows-reserved, invalid, or case-colliding paths are rejected.
5. It removes Face, registry, GitHub, password, token, and private-key variables
   before invoking the credential-free Sistemas synchronizer. That synchronizer
   first refuses any drift in the existing vendor tree, then builds and validates
   a complete candidate before replacement. It removes files retired upstream
   instead of merging directories and canonicalizes Git-tracked text to LF before
   recording hashes.
6. It restores the login-owned `Brightness.svg` byte-for-byte and these three
   Brazilian Portuguese overrides:

   ```json
   {
     "AZAIF_IncreaseBrightness": "Coloque o brilho da tela no máximo e afaste-se de janelas muito iluminadas.",
     "AZAIF_IncreaseBrightnessHighestSetting": "A tela piscará algumas vezes para processar o FaceID.",
     "AZAIF_IncreaseBrightnessTurnedUp": "Coloquei o brilho no máximo e me afastei de janelas muito iluminadas."
   }
   ```

7. It updates the Backend current-version marker and the Sistemas vendor manifest
   together, verifies that they agree, and proves Backend `package.json`,
   `package-lock.json`, and `.npmrc` are unchanged.
8. It prepares and structurally validates both repositories before obtaining a
   narrowly scoped GitHub App installation token. It verifies that neither
   repository's `main` tip advanced during preparation, then proposes the same
   deterministic branch, such as `chore/update-face-liveness-sdk-1-6-0`, in each
   repository and opens two cross-linked draft pull requests.
9. A `finally` cleanup and a separate always-running workflow cleanup remove the
   install tree, cache, temporary configuration, and downloaded packages after
   success or failure.

The updater stops rather than competing with an existing Face SDK proposal. A
rerun may continue an exact matching branch and draft pull request; it updates
only its delimited description block and preserves maintainer notes outside that
block. A validated matching pair already marked ready for review is a successful
no-op. The updater never force-pushes, closes a pull request or preview, deletes a
branch, approves, merges, enables auto-merge, or changes Azure.

## Pull-request isolation

The generated pull-request workflows use only `contents: read`, check out with
persisted credentials disabled, and run the repository verification suites
without secrets. They have no environment, artifact handoff, OIDC, Azure,
deployment, reusable-workflow, server-start, approval, or merge path.

The Sistemas Static Web Apps workflow rejects same-repository branches beginning
`chore/update-face-liveness-sdk-` in both its build/deploy job and its preview
close job. Therefore an automated proposal cannot deploy, create or close a
preview, enter Production, or receive the Static Web Apps token. Merging a
separately reviewed Sistemas SDK update to `main` retains the ordinary production
deployment path. Backend runbook-only changes remain excluded from its runtime
deployment workflow.

The automation does not rewrite frozen Sistemas release-qualification baselines,
opaque WASM expectations, compatibility prose, or the canonical production
artifact identity. A genuine SDK proposal can therefore make those deliberate
checks fail until a reviewer qualifies the new release and updates each expected
value on the proposal branch. Do not mechanically bless new hashes.

## Review a generated proposal

1. Read the target release and migration notes in the official sources. The Face
   “What's new” page can lag behind the SDK release feed.
2. Confirm both draft pull requests have the same version and branch, are
   cross-linked, and contain only the expected runbook and vendored-SDK changes.
3. Confirm the Backend dependency files and tracked `.npmrc` are unchanged. No
   credential, registry response, temporary package metadata, or package content
   outside the intended Sistemas vendor tree may appear in a diff or log.
4. Review removed and added vendor files. Confirm the complete asset directory
   was replaced, `clientSDKversion` matches the proposal, the custom brightness
   image is byte-identical to
   `sistemas/apps/learning-platform/login/img/Brightness.svg`, and all three
   Portuguese messages are exact.
5. Deliberately update any frozen version, file, WASM, or artifact expectations
   only after reviewing the release. Run the complete Backend and Sistemas suites
   from their `AGENTS.md` files and `git diff --check` in both repositories.
6. In Sistemas, run `node scripts/check-face-sdk-vendor.mjs` and inspect the
   generated vendor manifest. Do not waive structural, browser, runtime, API,
   workbook, retry, signed-handle, artifact, or production-network-denial checks.

Do not combine an SDK proposal with unrelated formatting, dependency, application,
or infrastructure changes.

## Browser qualification

Serve the Sistemas source preview with:

```powershell
node scripts/serve-frontend.mjs
```

Use `/plataforma/login/` for the login flow and
`/plataforma/cadastro-foto/` for enrollment. Do not edit a source file to replace
the Backend URL. The learning-platform runtime reads the production origin from
`apps/shared/backend-origin.js`; automated tests use an injected synthetic origin
and deny the production network. The browser routes moved to `/plataforma/`, but
the Backend API namespace intentionally remains `/plataforma_v2`.

With an approved test account and webcam permission, confirm that the liveness
interface loads, JavaScript/WebAssembly/image/localization requests do not return
404, the Machado brightness guidance appears, success completes, and cancellation
plus one expected failure remain understandable. This smoke test creates live
Face sessions. Exercise photo registration only with explicit approval because
it can write a reference photo and update the live workbook.

## Release and recovery

Keep both pull requests draft until review and every required check are complete.
Merge only during an appropriate maintenance window. Merge and verify the
Sistemas proposal first because it supplies the production browser files; then
merge the Backend version record, which does not require a runtime deployment.
Any version mismatch left between the repositories makes the next automation run
fail loudly before network or package work.

After deployment, repeat the approved production FaceID smoke test. If it fails,
create and review a new revert pull request for the Sistemas SDK change, deploy
the last known-good assets, and keep the failed version out of the current-version
record. Never rewrite history or force-push a recovery.
