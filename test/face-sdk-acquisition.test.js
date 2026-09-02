'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const { pathToFileURL } = require('node:url');

const repositoryRoot = path.resolve(__dirname, '..');
const moduleUrl = pathToFileURL(
    path.join(repositoryRoot, 'scripts', 'prepare-face-sdk-update.mjs'),
).href;
const SYNTHETIC_REGISTRY_PASSWORD = 'A'.repeat(16);

const REQUIRED_ASSET_FILES = [
    'i18n/pt-BR/en.json',
    'images/Brightness.svg',
    'images/Smile.svg',
    'images/logo.svg',
    'images/FaceId.svg',
    'images/activeMotionVisualHint.png',
    'js/AzureAIVisionFace.js',
    'js/AzureAIVisionFace.wasm',
    'js/AzureAIVisionFace_SIMD.js',
    'js/AzureAIVisionFace_SIMD.wasm',
];

async function makeFixture(t, currentVersion = '1.5.0') {
    const root = await fsp.mkdtemp(path.join(os.tmpdir(), 'backend-face-sdk-test-'));
    t.after(() => fsp.rm(root, { recursive: true, force: true }));
    const backendRoot = path.join(root, 'backend');
    const sistemasRoot = path.join(root, 'sistemas');
    const tempRoot = path.join(root, 'runner-temp');
    await Promise.all([
        fsp.mkdir(path.join(backendRoot, 'docs', 'runbooks'), { recursive: true }),
        fsp.mkdir(path.join(sistemasRoot, 'scripts'), { recursive: true }),
        fsp.mkdir(tempRoot, { recursive: true }),
    ]);
    await Promise.all([
        fsp.writeFile(path.join(backendRoot, 'package.json'), '{"private":true}\n'),
        fsp.writeFile(path.join(backendRoot, 'package-lock.json'), '{"lockfileVersion":3}\n'),
        fsp.writeFile(path.join(backendRoot, '.npmrc'), '@azure-ai-vision-face:registry=https://example.invalid/\n_password=${AZURE_AI_VISION_NPM_TOKEN_BASE64}\n'),
        fsp.writeFile(
            path.join(backendRoot, 'docs', 'runbooks', 'update-face-liveness-sdk.md'),
            `# Update\n\nCurrent vendored version: **${currentVersion}**\n`,
        ),
        fsp.writeFile(
            path.join(sistemasRoot, 'scripts', 'update-face-sdk-vendor.mjs'),
            '// fixture updater\n',
        ),
        fsp.writeFile(
            path.join(sistemasRoot, 'scripts', 'face-sdk-vendor.json'),
            `${JSON.stringify({
                schemaVersion: 1,
                version: currentVersion,
                packages: [
                    { name: '@azure/ai-vision-face-ui', version: currentVersion },
                    { name: '@azure-ai-vision-face/ui-assets', version: currentVersion },
                ],
            })}\n`,
        ),
    ]);
    return {
        root,
        backendRoot,
        sistemasRoot,
        tempRoot,
        workspace: path.join(tempRoot, 'face-sdk-fixture'),
        outputFile: path.join(root, 'github-output'),
    };
}

function successFetch(secret = SYNTHETIC_REGISTRY_PASSWORD) {
    return async (url, options) => {
        assert.equal(url.protocol, 'https:');
        assert.equal(url.hostname, 'face-fixture.cognitiveservices.azure.com');
        assert.equal(url.pathname, '/face/v1.3-preview.1/settings/getClientAssetsAccessToken');
        assert.equal(options.method, 'GET');
        assert.equal(options.redirect, 'error');
        assert.equal(options.headers['Ocp-Apim-Subscription-Key'], 'face-api-key-sentinel');
        return {
            ok: true,
            json: async () => ({ base64AccessToken: secret }),
        };
    };
}

// Write the install fixture synchronously because spawnSync test doubles cannot await.
function synchronousSuccessfulSpawn(calls, options = {}) {
    return (command, args, spawnOptions) => {
        calls.push({ command, args: [...args], options: spawnOptions });
        if (args[0] === 'view') {
            return { status: 0, stdout: JSON.stringify(options.latestVersion ?? '1.6.0') };
        }
        if (args[0] === 'install') {
            if (options.installFailure) {
                return { status: 1, stdout: SYNTHETIC_REGISTRY_PASSWORD, stderr: 'registry-response-sentinel' };
            }
            const installRoot = args[args.indexOf('--prefix') + 1];
            const mainRoot = path.join(installRoot, 'node_modules', '@azure', 'ai-vision-face-ui');
            const assetsRoot = path.join(installRoot, 'node_modules', '@azure-ai-vision-face', 'ui-assets');
            fs.mkdirSync(mainRoot, { recursive: true });
            for (const relativePath of REQUIRED_ASSET_FILES) {
                const target = path.join(assetsRoot, 'facelivenessdetector-assets', relativePath);
                fs.mkdirSync(path.dirname(target), { recursive: true });
                fs.writeFileSync(target, `fixture:${relativePath}\n`);
            }
            const version = options.installedVersion ?? '1.6.0';
            fs.writeFileSync(path.join(mainRoot, 'package.json'), JSON.stringify({
                name: options.packageOverrides?.mainName ?? '@azure/ai-vision-face-ui',
                version,
            }));
            fs.writeFileSync(
                path.join(mainRoot, 'FaceLivenessDetector.js'),
                `const sdk={clientSDKversion:"${options.packageOverrides?.embeddedVersion ?? version}"};\n`,
            );
            fs.writeFileSync(path.join(assetsRoot, 'package.json'), JSON.stringify({
                name: options.packageOverrides?.assetsName ?? '@azure-ai-vision-face/ui-assets',
                version,
            }));
            return { status: 0, stdout: 'private-package-output-must-not-be-logged' };
        }
        if (options.sistemasFailure) {
            return { status: 1, stdout: SYNTHETIC_REGISTRY_PASSWORD, stderr: 'private-contents-sentinel' };
        }
        return {
            status: 0,
            stdout: JSON.stringify({
                version: '1.6.0',
                files: 90,
                bytes: 10_000_000,
                sha256: 'a'.repeat(64),
            }),
        };
    };
}

test('version comparison and deterministic branch names reject ambiguous versions', async () => {
    const { compareSemver, versionBranch } = await import(moduleUrl);
    assert.equal(compareSemver('1.6.0', '1.5.9'), 1);
    assert.equal(compareSemver('1.6.0', '1.6.0'), 0);
    assert.equal(versionBranch('1.6.0'), 'chore/update-face-liveness-sdk-1-6-0');
    assert.throws(() => versionBranch('1.6.0-beta.1'), /exact stable x\.y\.z version/);
    assert.throws(() => versionBranch('latest'), /exact stable x\.y\.z version/);
});

test('successful acquisition is isolated, secret-sanitized, and always cleaned', async (t) => {
    const { prepareFaceSdkUpdate } = await import(moduleUrl);
    const fixture = await makeFixture(t);
    const calls = [];
    const result = await prepareFaceSdkUpdate({
        ...fixture,
        requestedVersion: 'latest',
        outputFile: fixture.outputFile,
        environment: {
            PATH: process.env.PATH,
            AZURE_FACE_API_ENDPOINT: 'https://face-fixture.cognitiveservices.azure.com/',
            AZURE_FACE_API_KEY: 'face-api-key-sentinel',
            GH_TOKEN: 'github-token-sentinel',
            FACE_SDK_APP_PRIVATE_KEY: 'private-key-sentinel',
            SAFE_VALUE: 'retained',
        },
        fetchImpl: successFetch(),
        spawnImpl: synchronousSuccessfulSpawn(calls),
        platform: 'linux',
    });

    assert.deepEqual(result, {
        'update-available': 'true',
        'current-version': '1.5.0',
        'target-version': '1.6.0',
        branch: 'chore/update-face-liveness-sdk-1-6-0',
        'vendor-files': '90',
        'vendor-bytes': '10000000',
        'vendor-sha256': 'a'.repeat(64),
    });
    assert.equal(fs.existsSync(fixture.workspace), false);
    assert.match(
        await fsp.readFile(
            path.join(fixture.backendRoot, 'docs', 'runbooks', 'update-face-liveness-sdk.md'),
            'utf8',
        ),
        /Current vendored version: \*\*1\.6\.0\*\*/,
    );
    assert.match(await fsp.readFile(fixture.outputFile, 'utf8'), /vendor-sha256=a{64}/);

    const installCall = calls.find(({ args }) => args[0] === 'install');
    for (const flag of [
        '--no-save',
        '--package-lock=false',
        '--ignore-scripts',
        '--no-audit',
        '--no-fund',
        '--omit=dev',
        '--loglevel=silent',
    ]) assert.ok(installCall.args.includes(flag), `${flag} must isolate the temporary install`);
    assert.equal(installCall.options.shell, false);
    assert.equal(installCall.options.env.AZURE_AI_VISION_NPM_TOKEN_BASE64, SYNTHETIC_REGISTRY_PASSWORD);
    assert.match(installCall.options.env.NPM_CONFIG_USERCONFIG, /face-sdk-fixture[\\/]\.npmrc$/);
    assert.match(installCall.options.env.npm_config_cache, /face-sdk-fixture[\\/]npm-cache$/);

    const sistemasCall = calls.find(({ args }) => args.includes('--target-version'));
    assert.ok(sistemasCall);
    assert.equal(sistemasCall.options.env.SAFE_VALUE, 'retained');
    for (const name of Object.keys(sistemasCall.options.env)) {
        assert.doesNotMatch(name, /TOKEN|PASSWORD|SECRET|PRIVATE_KEY|AZURE_FACE/i);
    }
    assert.equal(sistemasCall.args[sistemasCall.args.indexOf('--package-root') + 1],
        path.join(fixture.workspace, 'install'));

    assert.equal(await fsp.readFile(path.join(fixture.backendRoot, 'package.json'), 'utf8'), '{"private":true}\n');
    assert.equal(await fsp.readFile(path.join(fixture.backendRoot, 'package-lock.json'), 'utf8'), '{"lockfileVersion":3}\n');
});

test('an unchanged version avoids credentials, installation, and workspace creation', async (t) => {
    const { prepareFaceSdkUpdate } = await import(moduleUrl);
    const fixture = await makeFixture(t);
    let fetched = false;
    const calls = [];
    const result = await prepareFaceSdkUpdate({
        ...fixture,
        requestedVersion: '1.5.0',
        environment: {},
        fetchImpl: async () => { fetched = true; throw new Error('must not fetch'); },
        spawnImpl: synchronousSuccessfulSpawn(calls),
    });
    assert.equal(result['update-available'], 'false');
    assert.equal(fetched, false);
    assert.equal(calls.length, 0);
    assert.equal(fs.existsSync(fixture.workspace), false);
});

test('Backend and Sistemas version drift fails loudly before network or package work', async (t) => {
    const { prepareFaceSdkUpdate } = await import(moduleUrl);
    const fixture = await makeFixture(t);
    await fsp.writeFile(
        path.join(fixture.sistemasRoot, 'scripts', 'face-sdk-vendor.json'),
        `${JSON.stringify({
            schemaVersion: 1,
            version: '1.4.0',
            packages: [
                { name: '@azure/ai-vision-face-ui', version: '1.4.0' },
                { name: '@azure-ai-vision-face/ui-assets', version: '1.4.0' },
            ],
        })}\n`,
    );
    let fetched = false;
    let spawned = false;
    await assert.rejects(
        prepareFaceSdkUpdate({
            ...fixture,
            requestedVersion: 'latest',
            environment: {},
            fetchImpl: async () => { fetched = true; throw new Error('must not fetch'); },
            spawnImpl: () => { spawned = true; throw new Error('must not spawn'); },
        }),
        /Backend and Sistemas record different Face SDK versions/,
    );
    assert.equal(fetched, false);
    assert.equal(spawned, false);
    assert.equal(fs.existsSync(fixture.workspace), false);
});

test('the Face resource key is sent only to the approved HTTPS Azure hostname shape', async (t) => {
    const { prepareFaceSdkUpdate } = await import(moduleUrl);
    const fixture = await makeFixture(t);
    for (const endpoint of [
        'http://face-fixture.cognitiveservices.azure.com/',
        'https://attacker.example/',
        'https://user:password@face-fixture.cognitiveservices.azure.com/',
        'https://face-fixture.cognitiveservices.azure.com/unexpected-path',
    ]) {
        let fetched = false;
        await assert.rejects(
            prepareFaceSdkUpdate({
                ...fixture,
                requestedVersion: '1.6.0',
                environment: {
                    AZURE_FACE_API_ENDPOINT: endpoint,
                    AZURE_FACE_API_KEY: 'face-api-key-sentinel',
                },
                fetchImpl: async () => { fetched = true; throw new Error('must not fetch'); },
                spawnImpl: synchronousSuccessfulSpawn([]),
            }),
            /The Face API endpoint is invalid/,
        );
        assert.equal(fetched, false);
        assert.equal(fs.existsSync(fixture.workspace), false);
    }
});

for (const failure of ['install', 'sistemas']) {
    test(`${failure} failure reveals no captured output and removes all temporary files`, async (t) => {
        const { prepareFaceSdkUpdate } = await import(moduleUrl);
        const fixture = await makeFixture(t);
        const calls = [];
        const options = failure === 'install' ? { installFailure: true } : { sistemasFailure: true };
        await assert.rejects(
            prepareFaceSdkUpdate({
                ...fixture,
                requestedVersion: '1.6.0',
                environment: {
                    AZURE_FACE_API_ENDPOINT: 'https://face-fixture.cognitiveservices.azure.com/',
                    AZURE_FACE_API_KEY: 'face-api-key-sentinel',
                },
                fetchImpl: successFetch(),
                spawnImpl: synchronousSuccessfulSpawn(calls, options),
            }),
            (error) => {
                assert.equal(error.message.includes(SYNTHETIC_REGISTRY_PASSWORD), false);
                assert.doesNotMatch(error.message, /registry-response|private-contents/);
                return true;
            },
        );
        assert.equal(fs.existsSync(fixture.workspace), false);
        assert.match(
            await fsp.readFile(
                path.join(fixture.backendRoot, 'docs', 'runbooks', 'update-face-liveness-sdk.md'),
                'utf8',
            ),
            /Current vendored version: \*\*1\.5\.0\*\*/,
        );
    });
}

test('invalid private package identity fails closed and is cleaned', async (t) => {
    const { prepareFaceSdkUpdate } = await import(moduleUrl);
    const fixture = await makeFixture(t);
    await assert.rejects(
        prepareFaceSdkUpdate({
            ...fixture,
            requestedVersion: '1.6.0',
            environment: {
                AZURE_FACE_API_ENDPOINT: 'https://face-fixture.cognitiveservices.azure.com/',
                AZURE_FACE_API_KEY: 'face-api-key-sentinel',
            },
            fetchImpl: successFetch(),
            spawnImpl: synchronousSuccessfulSpawn([], {
                packageOverrides: { assetsName: '@attacker/replacement' },
            }),
        }),
        /Unexpected private package identity/,
    );
    assert.equal(fs.existsSync(fixture.workspace), false);
});

test('a pre-existing workspace is rejected without deleting it', async (t) => {
    const { prepareFaceSdkUpdate } = await import(moduleUrl);
    const fixture = await makeFixture(t);
    await fsp.mkdir(fixture.workspace);
    const sentinel = path.join(fixture.workspace, 'owned-by-another-run');
    await fsp.writeFile(sentinel, 'preserve');
    await assert.rejects(
        prepareFaceSdkUpdate({
            ...fixture,
            requestedVersion: '1.6.0',
            environment: {
                AZURE_FACE_API_ENDPOINT: 'https://face-fixture.cognitiveservices.azure.com/',
                AZURE_FACE_API_KEY: 'face-api-key-sentinel',
            },
            fetchImpl: successFetch(),
            spawnImpl: synchronousSuccessfulSpawn([]),
        }),
        /EEXIST/,
    );
    assert.equal(await fsp.readFile(sentinel, 'utf8'), 'preserve');
});

test('credential endpoint failures are generic and leave no temporary package', async (t) => {
    const { prepareFaceSdkUpdate } = await import(moduleUrl);
    const fixture = await makeFixture(t);
    await assert.rejects(
        prepareFaceSdkUpdate({
            ...fixture,
            requestedVersion: '1.6.0',
            environment: {
                AZURE_FACE_API_ENDPOINT: 'https://face-fixture.cognitiveservices.azure.com/',
                AZURE_FACE_API_KEY: 'face-api-key-sentinel',
            },
            fetchImpl: async () => ({
                ok: false,
                status: 401,
                text: async () => 'private-registry-response-sentinel',
            }),
            spawnImpl: synchronousSuccessfulSpawn([]),
        }),
        (error) => {
            assert.equal(error.message, 'Could not obtain the short-lived Face SDK package credential');
            assert.doesNotMatch(error.message, /sentinel|401/);
            return true;
        },
    );
    assert.equal(fs.existsSync(fixture.workspace), false);
});
