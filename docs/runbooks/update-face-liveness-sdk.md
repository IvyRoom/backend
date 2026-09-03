# Update the Face Liveness Web SDK

Current vendored version: **1.5.0**

The browser SDK is vendored in
`sistemas/apps/learning-platform/azure-ai-vision-face-ui`. The login and photo
registration applications load that committed copy through the public
`/plataforma/azure-ai-vision-face-ui/` path. It is not a Backend dependency and
does not update at build time or runtime.

This process intentionally remains separate from ordinary npm Dependabot.

## Automatic version check

Backend's `.github/workflows/face-sdk-version-check.yml` runs every Tuesday at
12:23 UTC and can also be started manually. It compares the version marker above
with the stable `latest` metadata for `@azure/ai-vision-face-ui` on the public npm
registry.

The checker uses no repository or Azure secret, private registry, environment,
OIDC credential, or deployment action. It reads metadata only: it does not
download package contents, modify either repository, or create a pull request.

When the versions differ, it opens one Backend issue for that published version
and assigns the workflow actor. GitHub notifies the assignee; email delivery
depends on that account's GitHub notification settings. An existing open or
closed issue prevents duplicate notices. When the versions match, the workflow
finishes successfully without creating an issue.

The issue is a prompt for Lucas and Codex to review the release and perform the
manual process below.

## Official references

- [SDK release notes](https://github.com/Azure-Samples/azure-ai-vision-sdk/releases)
- [What's new in Azure Face](https://learn.microsoft.com/en-us/azure/ai-services/face/whats-new-face)
- [Understand Face Liveness SDK versions](https://learn.microsoft.com/en-us/azure/ai-services/face/sdk/understand-the-liveness-sdk-versions)
- [Web SDK sample and installation guidance](https://github.com/Azure-Samples/azure-ai-vision-sdk/blob/main/samples/web/README.md)
- [Get Client Assets Access Token API](https://learn.microsoft.com/en-us/rest/api/face/liveness-session-operations/get-client-assets-access-token?view=rest-face-v1.3-preview)

## Before starting an update

1. Review the official release feed, Azure Face notes, and migration guidance.
2. Choose the exact stable version; do not install an unreviewed range.
3. Read both repositories' `AGENTS.md` files and require clean, synchronized
   `main` branches.
4. Create the same update branch in Backend and Sistemas.
5. Confirm Node.js 24, npm, Backend's ignored `npmrc_password.http`, and its
   ignored `.env` are available.
6. Confirm the test browser has webcam permission and use an approved test
   account for live FaceID qualification.

Do not combine an SDK update with unrelated dependency, application, formatting,
or infrastructure changes.

## Manual update

### 1. Obtain the temporary package credential

Run Backend's ignored `npmrc_password.http` request and copy only the
`base64AccessToken` response field. Store it only in the current PowerShell
session:

```powershell
$env:AZURE_AI_VISION_NPM_TOKEN_BASE64 = '<TOKEN_BASE64>'
```

The tracked `.npmrc` reads this value. Never paste it into `.npmrc`, another
file, a commit, a PR, a message, or a log.

### 2. Download the exact package into a temporary directory

From Backend, replace `x.y.z` and run:

```powershell
$targetVersion = 'x.y.z'
$tempSdkRoot = Join-Path ([IO.Path]::GetTempPath()) ("face-sdk-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempSdkRoot | Out-Null
npm install "@azure/ai-vision-face-ui@$targetVersion" `
  --prefix $tempSdkRoot `
  --userconfig .\.npmrc `
  --cache "$tempSdkRoot\npm-cache" `
  --ignore-scripts `
  --package-lock=false `
  --no-save `
  --no-audit `
  --no-fund
(Get-Content "$tempSdkRoot\node_modules\@azure\ai-vision-face-ui\package.json" -Raw | ConvertFrom-Json).version
```

Require the reported version to equal `$targetVersion`. The package, npm cache,
and any temporary metadata must remain under the generated `$tempSdkRoot`.
Backend's `package.json`, `package-lock.json`, and `.npmrc` must not change.

### 3. Validate the downloaded layout

Both checks must return `True`:

```powershell
Test-Path "$tempSdkRoot\node_modules\@azure\ai-vision-face-ui\FaceLivenessDetector.js"
Test-Path "$tempSdkRoot\node_modules\@azure-ai-vision-face\ui-assets\facelivenessdetector-assets"
```

The two source paths intentionally use different scopes:

- `@azure/ai-vision-face-ui` contains `FaceLivenessDetector.js`.
- `@azure-ai-vision-face/ui-assets` contains the asset directory.

Stop and consult Microsoft's current guidance if either identity or path changed.

### 4. Replace the vendored files

Replace, rather than merge:

- `$tempSdkRoot/node_modules/@azure/ai-vision-face-ui/FaceLivenessDetector.js`
  into
  `sistemas/apps/learning-platform/azure-ai-vision-face-ui/FaceLivenessDetector.js`;
- `$tempSdkRoot/node_modules/@azure-ai-vision-face/ui-assets/facelivenessdetector-assets`
  into
  `sistemas/apps/learning-platform/azure-ai-vision-face-ui/facelivenessdetector-assets`.

Replacing the complete asset directory prevents files retired upstream from
remaining in production.

### 5. Restore Machado's overrides

Restore
`sistemas/apps/learning-platform/login/img/Brightness.svg` byte-for-byte at:

`sistemas/apps/learning-platform/azure-ai-vision-face-ui/facelivenessdetector-assets/images/Brightness.svg`

Then restore these exact values in
`sistemas/apps/learning-platform/azure-ai-vision-face-ui/facelivenessdetector-assets/i18n/pt-BR/en.json`:

```json
{
  "AZAIF_IncreaseBrightness": "Coloque o brilho da tela no máximo e afaste-se de janelas muito iluminadas.",
  "AZAIF_IncreaseBrightnessHighestSetting": "A tela piscará algumas vezes para processar o FaceID.",
  "AZAIF_IncreaseBrightnessTurnedUp": "Coloquei o brilho no máximo e me afastei de janelas muito iluminadas."
}
```

These entries are excerpts, not a replacement for the complete localization
file.

### 6. Update and verify the version

Change the version marker at the top of this runbook. Require it to match both
the installed package version and `clientSDKversion` inside the vendored
`FaceLivenessDetector.js`.

Review every added, removed, and changed vendor file. Do not mechanically accept
new hashes, WebAssembly files, browser behavior, or artifact baselines.

### 7. Remove all temporary package material

Whether the update succeeds or fails:

1. Verify `$tempSdkRoot` is the exact generated `face-sdk-*` directory beneath
   the operating-system temporary directory.
2. Remove that directory recursively.
3. Remove `AZURE_AI_VISION_NPM_TOKEN_BASE64` from the PowerShell session or close
   the terminal.
4. Prove Backend's package files and tracked `.npmrc` are unchanged.

Never run cleanup against an unresolved variable, repository root, home
directory, or broad temporary directory.

## Qualification

Run the complete Backend and Sistemas suites from their `AGENTS.md` files,
including both repositories' `git diff --check`. Existing Sistemas tests protect
the vendored file inventory, version, assets, localization, WebAssembly,
application contracts, production-network denial, and generated artifact.

Serve Sistemas source with:

```powershell
node scripts/serve-frontend.mjs
```

Use `/plataforma/login/` for login and `/plataforma/cadastro-foto/` for
enrollment. Do not edit source files to replace the Backend URL. With explicit
approval for live Face operations, verify asset loading, the customized
brightness guidance, success, cancellation, and one expected failure. Exercise
photo registration only with explicit approval because it can write a reference
photo and update the live workbook.

## Release and recovery

Open one reviewable PR in each repository with an intended final change. Merge
and verify Sistemas first because it supplies the production browser files; then
merge the Backend version record, which is deployment-neutral.

After Sistemas deploys, repeat the approved production FaceID smoke test. If it
fails, create a reviewed revert PR that restores the last known-good SDK files
and keep the failed version out of the Backend marker. Never rewrite history or
force-push recovery.
