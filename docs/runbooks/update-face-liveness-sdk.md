# Update the Face Liveness Web SDK

Current vendored version: **1.4.8**

> **Review status (2026-07-15):** the official releases page lists version 1.5.0 as latest, while Machado still uses 1.4.8. Evaluating and installing 1.5.0 is a separate SDK-upgrade task, not part of the documentation update that created this runbook.

This runbook documents the current manual process for updating the Face Liveness Web SDK used by `sistemas/plataforma_v2`. The backend repository is an authenticated download workspace; the browser bundle and assets used in production live in the frontend repository.

Both `plataforma_v2/login` and `plataforma_v2/cadastro` import the vendored bundle.

This procedure changes both repositories:

- `backend`, for the temporary package installation and this version record
- `sistemas`, for the vendored SDK file and assets

## Official references

- [SDK release notes](https://github.com/Azure-Samples/azure-ai-vision-sdk/releases)
- [What's new in Azure Face](https://learn.microsoft.com/en-us/azure/ai-services/face/whats-new-face)
- [Understand Face Liveness SDK versions](https://learn.microsoft.com/en-us/azure/ai-services/face/sdk/understand-the-liveness-sdk-versions)
- [Web SDK sample and installation guidance](https://github.com/Azure-Samples/azure-ai-vision-sdk/blob/main/samples/web/README.md)
- [Get Client Assets Access Token API](https://learn.microsoft.com/en-us/rest/api/face/liveness-session-operations/get-client-assets-access-token?view=rest-face-v1.3-preview)

## Before starting

1. Review both the official release feed and the Azure Face “What's new” page. The latter can lag behind the release feed.
2. Read the target version's release and migration notes. Choose the exact version before installing it.
3. Confirm that `backend` and `sistemas` have clean working trees.
4. Create the same feature branch in both repositories, for example `chore/update-face-liveness-sdk`.
5. Confirm that Node.js 20, npm, the ignored `backend/npmrc_password.http` request file, and the backend's ignored `.env` are available.
6. Confirm that the test browser has webcam permission and that the frontend can be served locally from the `sistemas` repository root.
7. Use an approved test account for the final FaceID test. A locally running backend can still reach live Microsoft Graph and Azure resources.

Do not combine this update with unrelated formatting, comment cleanup, or application changes.

## Update procedure

### 1. Obtain the temporary package token

Run the request in `backend/npmrc_password.http` and copy the response field named `base64AccessToken`. Do not copy the separate `accessToken` field.

In the PowerShell terminal attached to `backend`, store the token only for that terminal session:

```powershell
$env:AZURE_AI_VISION_NPM_TOKEN_BASE64 = '<TOKEN_BASE64>'
```

The tracked `.npmrc` reads this variable for both `_password` entries. Do not paste the token into `.npmrc`, another file, a commit, or a message. Closing the terminal removes the session variable.

### 2. Install the selected package version

In the same terminal, replace `x.y.z` with the version chosen from the release notes and run:

```powershell
$targetVersion = 'x.y.z'
npm install "@azure/ai-vision-face-ui@$targetVersion"
node -p "require('./node_modules/@azure/ai-vision-face-ui/package.json').version"
```

Confirm that the reported installed version matches `$targetVersion`. Using an explicit version keeps the reviewed release and the downloaded release aligned.

### 3. Validate the installed file layout

Run:

```powershell
Test-Path .\node_modules\@azure\ai-vision-face-ui\FaceLivenessDetector.js
Test-Path .\node_modules\@azure-ai-vision-face\ui-assets\facelivenessdetector-assets
```

Both commands must return `True`. If either path is missing, stop and consult the current official Web SDK instructions instead of guessing a replacement path.

The two source paths intentionally use different package scopes:

- `@azure/ai-vision-face-ui` contains `FaceLivenessDetector.js`.
- `@azure-ai-vision-face/ui-assets` contains `facelivenessdetector-assets`.

### 4. Replace the vendored frontend files

Copy:

- `backend/node_modules/@azure/ai-vision-face-ui/FaceLivenessDetector.js`
  to `sistemas/plataforma_v2/azure-ai-vision-face-ui/FaceLivenessDetector.js`
- `backend/node_modules/@azure-ai-vision-face/ui-assets/facelivenessdetector-assets`
  to `sistemas/plataforma_v2/azure-ai-vision-face-ui/facelivenessdetector-assets`

Replace the existing asset directory rather than merging the two versions. This prevents files removed by Microsoft from remaining in the vendored copy.

### 5. Reapply Machado's frontend overrides

Replace:

`sistemas/plataforma_v2/azure-ai-vision-face-ui/facelivenessdetector-assets/images/Brightness.svg`

with:

`sistemas/plataforma_v2/login/img/Brightness.svg`

Then update these values in:

`sistemas/plataforma_v2/azure-ai-vision-face-ui/facelivenessdetector-assets/i18n/pt-BR/en.json`

```json
{
  "AZAIF_IncreaseBrightness": "Coloque o brilho da tela no máximo e afaste-se de janelas muito iluminadas.",
  "AZAIF_IncreaseBrightnessHighestSetting": "A tela piscará algumas vezes para processar o FaceID.",
  "AZAIF_IncreaseBrightnessTurnedUp": "Coloquei o brilho no máximo e me afastei de janelas muito iluminadas."
}
```

These are excerpts to locate and replace in the existing JSON file, not a replacement for the whole file.

### 6. Update the version record

Change **Current vendored version** at the top of this runbook to the installed version.

Also verify that `clientSDKversion` inside the new `FaceLivenessDetector.js` reports the same version.

### 7. Remove the temporary backend dependency

Preserve the existing cleanup sequence:

1. Delete `backend/node_modules`.
2. Delete `backend/package-lock.json`.
3. Remove the `@azure/ai-vision-face-ui` dependency added to `backend/package.json`.
4. Close the terminal containing `AZURE_AI_VISION_NPM_TOKEN_BASE64`.
5. Open a new terminal in `backend` and run:

```powershell
npm install
```

The final backend dependency files must not retain `@azure/ai-vision-face-ui`; it is only a temporary source for the vendored frontend assets.

Run `git diff -- package.json package-lock.json`. It should show no dependency-file changes. Stop and investigate any remaining diff rather than committing regenerated dependency versions accidentally.

Automating this temporary installation, asset synchronization, and cleanup is a separate future improvement.

### 8. Test locally

1. Start the backend with `npm start`.
2. In `sistemas/plataforma_v2/login/main.js`, temporarily change the backend base URL to:

```text
http://localhost:3000/plataforma_v2
```

3. Serve the frontend locally from the `sistemas` repository root so that the existing `/plataforma_v2/` paths remain valid.
4. Complete the login FaceID flow with an approved test account. This flow performs live Microsoft Graph reads and creates live Azure Face sessions.
5. Confirm that:
   - the liveness interface loads;
   - its JavaScript, WebAssembly, image, and localization assets load without `404` responses;
   - the customized brightness instructions appear;
   - the success path completes;
   - cancellation and one expected failure path remain understandable.
6. Test the enrollment flow in `plataforma_v2/cadastro` only when explicitly approved. It can write a reference photo and update a live workbook record.
7. Restore the production base URL before committing:

```text
https://plataforma-backend-v3.azurewebsites.net/plataforma_v2
```

### 9. Review the changes

Before committing, confirm:

- The backend package files contain no temporary Face UI dependency.
- No token or credential appears in either repository's diff.
- The frontend production backend URL has been restored.
- The custom `Brightness.svg` and Brazilian Portuguese strings are present.
- The runbook and vendored JavaScript report the same SDK version.
- Each repository contains only the intended SDK-update changes.
- `git diff --check` passes in both repositories.

### 10. Deploy and verify

1. Commit and open one reviewable pull request in each repository that has an intended final change.
2. Merge during an appropriate platform maintenance window.
3. Deploy `sistemas`; that repository contains the production Web SDK files.
4. Do not deploy backend runtime changes merely because it was used to download the package. Deploy backend code only if the target release explicitly requires compatible backend changes.
5. If the version record in this backend runbook is merged, note that the current backend workflow still performs a production deployment even for documentation-only changes.
6. Repeat the FaceID smoke test in production.

If a pre-merge test fails, do not merge the SDK update. If the production smoke test fails, revert the frontend SDK update commit or pull request and redeploy the previous known-good assets.
