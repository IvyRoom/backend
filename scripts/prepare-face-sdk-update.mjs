import { createHash } from 'node:crypto';
import { lstat, mkdir, readFile, readdir, rm, writeFile, appendFile } from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

export const MAIN_PACKAGE = '@azure/ai-vision-face-ui';
export const ASSET_PACKAGE = '@azure-ai-vision-face/ui-assets';
export const VERSION_PATTERN = /^\d+\.\d+\.\d+$/;
export const UPDATE_BRANCH_PREFIX = 'chore/update-face-liveness-sdk-';

const RUNBOOK_RELATIVE_PATH = 'docs/runbooks/update-face-liveness-sdk.md';
const SISTEMAS_MANIFEST_RELATIVE_PATH = 'scripts/face-sdk-vendor.json';
const PROTECTED_BACKEND_FILES = ['package.json', 'package-lock.json', '.npmrc'];
const VERSION_MARKER = /^(Current vendored version: \*\*)[^*]+(\*\*)$/m;
const REQUIRED_ASSET_FILES = [
    'facelivenessdetector-assets/i18n/pt-BR/en.json',
    'facelivenessdetector-assets/images/Brightness.svg',
    'facelivenessdetector-assets/images/Smile.svg',
    'facelivenessdetector-assets/images/logo.svg',
    'facelivenessdetector-assets/images/FaceId.svg',
    'facelivenessdetector-assets/images/activeMotionVisualHint.png',
    'facelivenessdetector-assets/js/AzureAIVisionFace.js',
    'facelivenessdetector-assets/js/AzureAIVisionFace.wasm',
    'facelivenessdetector-assets/js/AzureAIVisionFace_SIMD.js',
    'facelivenessdetector-assets/js/AzureAIVisionFace_SIMD.wasm',
];

function sha256(buffer) {
    return createHash('sha256').update(buffer).digest('hex');
}

async function fileDigests(root, relativePaths) {
    const entries = await Promise.all(relativePaths.map(async (relativePath) => {
        const contents = await readFile(path.join(root, relativePath));
        return [relativePath, sha256(contents)];
    }));
    return Object.fromEntries(entries);
}

function assertDigestsUnchanged(before, after) {
    for (const [relativePath, digest] of Object.entries(before)) {
        if (after[relativePath] !== digest) {
            throw new Error(`Protected Backend file changed: ${relativePath}`);
        }
    }
}

function parseSemver(value, description = 'version') {
    if (!VERSION_PATTERN.test(value)) {
        throw new Error(`${description} must be an exact stable x.y.z version`);
    }

    const separator = value.indexOf('-');
    const release = separator === -1 ? value : value.slice(0, separator);
    const prerelease = separator === -1 ? null : value.slice(separator + 1);
    return {
        value,
        release: release.split('.').map(Number),
        prerelease: prerelease?.split('.') ?? null,
    };
}

function compareIdentifiers(left, right) {
    const leftNumeric = /^\d+$/.test(left);
    const rightNumeric = /^\d+$/.test(right);
    if (leftNumeric && rightNumeric) return Number(left) - Number(right);
    if (leftNumeric) return -1;
    if (rightNumeric) return 1;
    return left.localeCompare(right, 'en');
}

export function compareSemver(leftValue, rightValue) {
    const left = parseSemver(leftValue, 'Left version');
    const right = parseSemver(rightValue, 'Right version');

    for (let index = 0; index < 3; index += 1) {
        if (left.release[index] !== right.release[index]) {
            return left.release[index] - right.release[index];
        }
    }

    if (left.prerelease === null && right.prerelease === null) return 0;
    if (left.prerelease === null) return 1;
    if (right.prerelease === null) return -1;

    const length = Math.max(left.prerelease.length, right.prerelease.length);
    for (let index = 0; index < length; index += 1) {
        if (left.prerelease[index] === undefined) return -1;
        if (right.prerelease[index] === undefined) return 1;
        const comparison = compareIdentifiers(left.prerelease[index], right.prerelease[index]);
        if (comparison !== 0) return comparison;
    }
    return 0;
}

export function versionBranch(version) {
    parseSemver(version, 'Target version');
    return `${UPDATE_BRANCH_PREFIX}${version.toLowerCase().replace(/[^0-9a-z]+/g, '-')}`;
}

export function sanitizeEnvironment(environment) {
    const forbiddenName = /(?:TOKEN|PASSWORD|SECRET|CREDENTIAL|PRIVATE_KEY|API_KEY|AZURE_FACE|ACTIONS_ID_TOKEN|GH_TOKEN|GITHUB_TOKEN|NPM_CONFIG_(?:USERCONFIG|GLOBALCONFIG|.*AUTH|REGISTRY|CACHE))/i;
    return Object.fromEntries(
        Object.entries(environment).filter(([name]) => !forbiddenName.test(name)),
    );
}

function commandResult(spawnImpl, command, args, options, failureMessage) {
    const result = spawnImpl(command, args, {
        ...options,
        encoding: 'utf8',
        shell: false,
        windowsHide: true,
        maxBuffer: 4 * 1024 * 1024,
    });
    if (result.error || result.status !== 0) {
        throw new Error(failureMessage);
    }
    return String(result.stdout ?? '').trim();
}

function npmCommand(platform = process.platform) {
    return platform === 'win32' ? 'npm.cmd' : 'npm';
}

function parseNpmVersion(stdout) {
    let value;
    try {
        value = JSON.parse(stdout);
    } catch {
        value = stdout.replace(/^['"]|['"]$/g, '');
    }
    if (Array.isArray(value)) value = value.at(-1);
    if (typeof value !== 'string') {
        throw new Error('The public package registry returned an invalid version');
    }
    parseSemver(value, 'Published package version');
    return value;
}

async function readCurrentVersion(backendRoot) {
    const runbookPath = path.join(backendRoot, RUNBOOK_RELATIVE_PATH);
    const runbook = await readFile(runbookPath, 'utf8');
    const matches = [...runbook.matchAll(new RegExp(VERSION_MARKER.source, 'gm'))];
    if (matches.length !== 1) {
        throw new Error('The Face SDK runbook must contain exactly one current-version marker');
    }
    const currentVersion = matches[0][0].match(/\*\*([^*]+)\*\*/)?.[1];
    parseSemver(currentVersion ?? '', 'Current vendored version');
    return { currentVersion, runbook, runbookPath };
}

async function readSistemasVersion(sistemasRoot) {
    const manifestPath = path.join(sistemasRoot, SISTEMAS_MANIFEST_RELATIVE_PATH);
    let manifest;
    try {
        manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
    } catch {
        throw new Error('The Sistemas Face SDK vendor manifest is unavailable or invalid');
    }
    parseSemver(manifest?.version ?? '', 'Sistemas vendored version');
    const expectedPackages = [MAIN_PACKAGE, ASSET_PACKAGE];
    if (
        manifest.schemaVersion !== 1
        || !Array.isArray(manifest.packages)
        || manifest.packages.length !== expectedPackages.length
        || manifest.packages.some((entry, index) => (
            entry?.name !== expectedPackages[index] || entry.version !== manifest.version
        ))
    ) {
        throw new Error('The Sistemas Face SDK vendor manifest is unavailable or invalid');
    }
    return manifest.version;
}

async function writeCurrentVersion({ runbook, runbookPath }, targetVersion) {
    const updated = runbook.replace(VERSION_MARKER, `$1${targetVersion}$2`);
    if (updated === runbook) {
        throw new Error('The Face SDK runbook version marker was not updated');
    }
    await writeFile(runbookPath, updated, 'utf8');
}

function assertWorkspacePath(workspace, tempRoot) {
    const resolvedWorkspace = path.resolve(workspace);
    const resolvedTempRoot = path.resolve(tempRoot);
    const relative = path.relative(resolvedTempRoot, resolvedWorkspace);
    if (
        !relative
        || relative.startsWith('..')
        || path.isAbsolute(relative)
        || path.dirname(resolvedWorkspace) !== resolvedTempRoot
        || !path.basename(resolvedWorkspace).startsWith('face-sdk-')
    ) {
        throw new Error('The temporary workspace must be a new face-sdk-* child of RUNNER_TEMP');
    }
    return resolvedWorkspace;
}

function parseSistemasSummary(stdout, targetVersion) {
    let summary;
    try {
        summary = JSON.parse(stdout);
    } catch {
        throw new Error('The Sistemas Face SDK updater returned an invalid summary');
    }
    if (
        summary?.version !== targetVersion
        || !Number.isSafeInteger(summary.files)
        || summary.files <= 0
        || !Number.isSafeInteger(summary.bytes)
        || summary.bytes <= 0
        || typeof summary.sha256 !== 'string'
        || !/^[0-9a-f]{64}$/.test(summary.sha256)
    ) {
        throw new Error('The Sistemas Face SDK updater returned an invalid summary');
    }
    return summary;
}

async function assertRegularFile(filePath, description) {
    const stats = await lstat(filePath);
    if (!stats.isFile() || stats.isSymbolicLink()) {
        throw new Error(`${description} must be a regular file`);
    }
}

async function assertSafeTree(directory, description) {
    const stats = await lstat(directory);
    if (!stats.isDirectory() || stats.isSymbolicLink()) {
        throw new Error(`${description} must be a real directory`);
    }
    const entries = await readdir(directory, { withFileTypes: true });
    if (entries.length === 0) throw new Error(`${description} must not be empty`);
    for (const entry of entries) {
        const entryPath = path.join(directory, entry.name);
        if (entry.isSymbolicLink()) {
            throw new Error(`${description} must not contain symbolic links`);
        }
        if (entry.isDirectory()) {
            await assertSafeTree(entryPath, description);
        } else if (!entry.isFile()) {
            throw new Error(`${description} may contain only regular files and directories`);
        }
    }
}

async function readPackageManifest(packageRoot, expectedName) {
    await assertRegularFile(path.join(packageRoot, 'package.json'), `${expectedName} package.json`);
    let manifest;
    try {
        manifest = JSON.parse(await readFile(path.join(packageRoot, 'package.json'), 'utf8'));
    } catch {
        throw new Error(`${expectedName} has an invalid package.json`);
    }
    if (manifest.name !== expectedName) {
        throw new Error(`Unexpected private package identity for ${expectedName}`);
    }
    return manifest;
}

export async function validateInstalledPackages(nodeModulesRoot, targetVersion) {
    parseSemver(targetVersion, 'Target version');
    const mainRoot = path.join(nodeModulesRoot, '@azure', 'ai-vision-face-ui');
    const assetsRoot = path.join(nodeModulesRoot, '@azure-ai-vision-face', 'ui-assets');
    await assertSafeTree(mainRoot, MAIN_PACKAGE);
    await assertSafeTree(assetsRoot, ASSET_PACKAGE);
    const mainManifest = await readPackageManifest(mainRoot, MAIN_PACKAGE);
    await readPackageManifest(assetsRoot, ASSET_PACKAGE);
    if (mainManifest.version !== targetVersion) {
        throw new Error('The installed Face SDK version does not match the requested version');
    }

    const loaderPath = path.join(mainRoot, 'FaceLivenessDetector.js');
    await assertRegularFile(loaderPath, 'FaceLivenessDetector.js');
    const loader = await readFile(loaderPath, 'utf8');
    const escapedVersion = targetVersion.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    if (!new RegExp(`clientSDKversion\\s*:\\s*["']${escapedVersion}["']`).test(loader)) {
        throw new Error('FaceLivenessDetector.js does not embed the requested clientSDKversion');
    }

    const assetTree = path.join(assetsRoot, 'facelivenessdetector-assets');
    await assertSafeTree(assetTree, 'Face Liveness asset tree');
    for (const relativePath of REQUIRED_ASSET_FILES) {
        await assertRegularFile(path.join(assetsRoot, relativePath), relativePath);
    }
    return { mainRoot, assetsRoot, loaderPath, assetTree };
}

async function acquireRegistryPassword(endpoint, apiKey, fetchImpl) {
    if (!endpoint || !apiKey) {
        throw new Error('Face SDK automation credentials are not configured');
    }
    let endpointUrl;
    try {
        endpointUrl = new URL(endpoint);
    } catch {
        throw new Error('The Face API endpoint is invalid');
    }
    const azureFaceHostname = /^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.cognitiveservices\.azure\.com$/i;
    if (
        endpointUrl.protocol !== 'https:'
        || !azureFaceHostname.test(endpointUrl.hostname)
        || endpointUrl.username
        || endpointUrl.password
        || endpointUrl.port
        || (endpointUrl.pathname !== '/' && endpointUrl.pathname !== '')
        || endpointUrl.search
        || endpointUrl.hash
    ) {
        throw new Error('The Face API endpoint is invalid');
    }
    const url = new URL('/face/v1.3-preview.1/settings/getClientAssetsAccessToken', endpointUrl);

    let response;
    try {
        response = await fetchImpl(url, {
            method: 'GET',
            headers: { 'Ocp-Apim-Subscription-Key': apiKey },
            redirect: 'error',
            signal: AbortSignal.timeout(30_000),
        });
    } catch {
        throw new Error('Could not obtain the short-lived Face SDK package credential');
    }
    if (!response.ok) {
        throw new Error('Could not obtain the short-lived Face SDK package credential');
    }

    let body;
    try {
        body = await response.json();
    } catch {
        throw new Error('The Face SDK package credential response was invalid');
    }
    const password = body?.base64AccessToken;
    if (typeof password !== 'string' || password.length < 16 || !/^[A-Za-z0-9+/]+={0,2}$/.test(password)) {
        throw new Error('The Face SDK package credential response was invalid');
    }
    return password;
}

async function appendOutputs(outputFile, values) {
    if (!outputFile) return;
    const lines = Object.entries(values).map(([key, value]) => `${key}=${value}`).join('\n');
    await appendFile(outputFile, `${lines}\n`, 'utf8');
}

export async function prepareFaceSdkUpdate({
    backendRoot,
    sistemasRoot,
    workspace,
    tempRoot,
    requestedVersion = 'latest',
    outputFile,
    environment = process.env,
    fetchImpl = globalThis.fetch,
    spawnImpl = spawnSync,
    platform = process.platform,
} = {}) {
    if (!backendRoot || !sistemasRoot || !workspace || !tempRoot) {
        throw new Error('Backend root, Sistemas root, workspace, and temp root are required');
    }
    const resolvedBackendRoot = path.resolve(backendRoot);
    const resolvedSistemasRoot = path.resolve(sistemasRoot);
    const resolvedWorkspace = assertWorkspacePath(workspace, tempRoot);
    const npm = npmCommand(platform);
    const versionRecord = await readCurrentVersion(resolvedBackendRoot);
    const sistemasVersion = await readSistemasVersion(resolvedSistemasRoot);
    if (versionRecord.currentVersion !== sistemasVersion) {
        throw new Error(
            'Backend and Sistemas record different Face SDK versions; reconcile the companion proposals before continuing',
        );
    }
    const protectedBefore = await fileDigests(resolvedBackendRoot, PROTECTED_BACKEND_FILES);

    let targetVersion = requestedVersion;
    if (requestedVersion === 'latest') {
        const stdout = commandResult(
            spawnImpl,
            npm,
            ['view', `${MAIN_PACKAGE}@latest`, 'version', '--json', '--registry=https://registry.npmjs.org/', '--ignore-scripts'],
            { cwd: path.resolve(tempRoot), env: sanitizeEnvironment(environment) },
            'Could not query the public Face SDK version',
        );
        targetVersion = parseNpmVersion(stdout);
    } else {
        parseSemver(requestedVersion, 'Requested version');
    }

    const comparison = compareSemver(targetVersion, versionRecord.currentVersion);
    if (comparison < 0) throw new Error('Face SDK automation does not propose version downgrades');
    if (comparison === 0) {
        const result = {
            'update-available': 'false',
            'current-version': versionRecord.currentVersion,
            'target-version': targetVersion,
            branch: versionBranch(targetVersion),
        };
        await appendOutputs(outputFile, result);
        return result;
    }

    let createdWorkspace = false;
    try {
        await mkdir(resolvedWorkspace, { recursive: false });
        createdWorkspace = true;
        const npmConfig = path.join(resolvedWorkspace, '.npmrc');
        await writeFile(npmConfig, await readFile(path.join(resolvedBackendRoot, '.npmrc')));
        const registryPassword = await acquireRegistryPassword(
            environment.AZURE_FACE_API_ENDPOINT,
            environment.AZURE_FACE_API_KEY,
            fetchImpl,
        );
        const installRoot = path.join(resolvedWorkspace, 'install');
        const npmCache = path.join(resolvedWorkspace, 'npm-cache');
        const installEnvironment = {
            ...sanitizeEnvironment(environment),
            AZURE_AI_VISION_NPM_TOKEN_BASE64: registryPassword,
            NPM_CONFIG_USERCONFIG: npmConfig,
            npm_config_cache: npmCache,
        };
        commandResult(
            spawnImpl,
            npm,
            [
                'install',
                `${MAIN_PACKAGE}@${targetVersion}`,
                '--prefix', installRoot,
                '--no-save',
                '--package-lock=false',
                '--ignore-scripts',
                '--no-audit',
                '--no-fund',
                '--omit=dev',
                '--loglevel=silent',
                '--registry=https://registry.npmjs.org/',
            ],
            { cwd: resolvedWorkspace, env: installEnvironment },
            'The temporary Face SDK package installation failed',
        );

        const nodeModulesRoot = path.join(installRoot, 'node_modules');
        await validateInstalledPackages(nodeModulesRoot, targetVersion);
        const sistemasUpdater = path.join(resolvedSistemasRoot, 'scripts', 'update-face-sdk-vendor.mjs');
        await assertRegularFile(sistemasUpdater, 'Sistemas Face SDK updater');
        const sistemasStdout = commandResult(
            spawnImpl,
            process.execPath,
            [
                sistemasUpdater,
                '--package-root', installRoot,
                '--target-version', targetVersion,
                '--repository-root', resolvedSistemasRoot,
            ],
            { cwd: resolvedSistemasRoot, env: sanitizeEnvironment(environment) },
            'The credential-free Sistemas Face SDK synchronization failed',
        );
        const sistemasSummary = parseSistemasSummary(sistemasStdout, targetVersion);

        await writeCurrentVersion(versionRecord, targetVersion);
        const protectedAfter = await fileDigests(resolvedBackendRoot, PROTECTED_BACKEND_FILES);
        assertDigestsUnchanged(protectedBefore, protectedAfter);
        const result = {
            'update-available': 'true',
            'current-version': versionRecord.currentVersion,
            'target-version': targetVersion,
            branch: versionBranch(targetVersion),
            'vendor-files': String(sistemasSummary.files),
            'vendor-bytes': String(sistemasSummary.bytes),
            'vendor-sha256': sistemasSummary.sha256,
        };
        await appendOutputs(outputFile, result);
        return result;
    } finally {
        if (createdWorkspace) {
            await rm(resolvedWorkspace, { recursive: true, force: true, maxRetries: 3 });
        }
    }
}

function parseArguments(argv) {
    const options = {};
    for (let index = 0; index < argv.length; index += 2) {
        const flag = argv[index];
        const value = argv[index + 1];
        if (!flag?.startsWith('--') || value === undefined) {
            throw new Error('Every command-line option requires a value');
        }
        options[flag.slice(2)] = value;
    }
    return options;
}

async function main() {
    const args = parseArguments(process.argv.slice(2));
    const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
    const backendRoot = path.resolve(args['backend-root'] ?? path.join(scriptDirectory, '..'));
    const tempRoot = path.resolve(args['temp-root'] ?? process.env.RUNNER_TEMP ?? '');
    const result = await prepareFaceSdkUpdate({
        backendRoot,
        sistemasRoot: args['sistemas-root'],
        workspace: args.workspace,
        tempRoot,
        requestedVersion: args['requested-version'] ?? 'latest',
        outputFile: args['github-output'] ?? process.env.GITHUB_OUTPUT,
    });
    process.stdout.write(result['update-available'] === 'true' ? 'Face SDK proposal prepared.\n' : 'No Face SDK update is available.\n');
}

const invokedDirectly = process.argv[1]
    && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (invokedDirectly) {
    main().catch((error) => {
        process.stderr.write(`Face SDK updater failed: ${error.message}\n`);
        process.exitCode = 1;
    });
}
