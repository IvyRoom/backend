'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const {
    createGraphTokenLifecycle,
    createProductionDependencies,
    startProductionServer,
} = require('../server');

const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const NOW_MS = 1_800_000_000_000;
const SIGNING_KEY = Buffer.from(Array.from({ length: 32 }, (_, index) => index));
const SIGNING_KEY_BASE64 = SIGNING_KEY.toString('base64');

function createFakeScheduler() {
    let nextTimerId = 1;
    const scheduled = [];
    const cancelled = [];

    function schedule(callback, delay) {
        const timer = { id: nextTimerId++, callback, delay };
        scheduled.push(timer);
        return timer;
    }

    function cancel(timer) {
        cancelled.push(timer);
    }

    return { scheduled, cancelled, schedule, cancel };
}

function createQueuedAcquireToken(outcomes) {
    const calls = [];

    async function acquireToken(request) {
        calls.push(request);
        const outcome = outcomes.shift();

        if (outcome instanceof Error) throw outcome;
        return outcome;
    }

    return { calls, acquireToken };
}

function runSyntheticServer({ port } = {}) {
    const events = [];
    const runtimeHooks = { synthetic: true };
    const appDependencies = { synthetic: true };
    const closeResult = { closed: true };
    const listener = {
        close(callback) {
            events.push('close');
            if (callback) callback();
            return closeResult;
        },
    };
    const graphTokenLifecycle = {
        start() {
            events.push('acquire');
            return new Promise(() => {});
        },
        stop() {
            events.push('stop-token-lifecycle');
        },
    };
    const app = {
        listen(selectedPort) {
            events.push(`listen:${selectedPort}`);
            return listener;
        },
    };
    const environment = {
        PLATFORM_ROW_AUTHORIZATION_KEY_BASE64: SIGNING_KEY_BASE64,
    };
    if (port !== undefined) environment.PORT = port;

    const result = startProductionServer({
        environment,
        loadEnvironment() {
            events.push('load-environment');
        },
        createDependencies(receivedEnvironment, receivedKey, receivedRuntimeHooks) {
            events.push('create-dependencies');
            assert.equal(receivedEnvironment, environment);
            assert.deepEqual(receivedKey, SIGNING_KEY);
            assert.equal(receivedRuntimeHooks, runtimeHooks);
            return { appDependencies, graphTokenLifecycle };
        },
        createApplication(receivedDependencies) {
            events.push('create-application');
            assert.equal(receivedDependencies, appDependencies);
            return app;
        },
        runtimeHooks,
    });

    return {
        app,
        closeResult,
        events,
        graphTokenLifecycle,
        listener,
        result,
    };
}

test('importing server is configuration-, SDK-, listener-, and timer-safe', () => {
    const serverPath = require.resolve('../server');
    const script = `
        const Module = require('node:module');
        const forbidden = new Set([
            'dotenv',
            '@microsoft/microsoft-graph-client',
            '@azure/msal-node',
            '@azure/core-auth',
            '@azure-rest/ai-vision-face',
        ]);
        const originalLoad = Module._load;
        Module._load = function (request, parent, isMain) {
            if (forbidden.has(request)) throw new Error('production module loaded during import: ' + request);
            return originalLoad.call(this, request, parent, isMain);
        };
        require(${JSON.stringify(serverPath)});
    `;
    const environment = { ...process.env };

    delete environment.PLATFORM_ROW_AUTHORIZATION_KEY_BASE64;
    delete environment.CLIENT_ID;
    delete environment.TENANT_ID;
    delete environment.CLIENT_SECRET;
    delete environment.AZURE_FACE_API_KEY;
    delete environment.AZURE_FACE_API_ENDPOINT;

    const child = spawnSync(process.execPath, ['-e', script], {
        cwd: path.resolve(__dirname, '..'),
        encoding: 'utf8',
        env: environment,
        timeout: 5_000,
    });

    assert.ifError(child.error);
    assert.equal(child.signal, null);
    assert.equal(child.status, 0, child.stderr);
    assert.equal(child.stdout, '');
});

test('invalid signing-key configuration fails before dependency construction or listening', () => {
    const events = [];

    assert.throws(
        () => startProductionServer({
            environment: { PLATFORM_ROW_AUTHORIZATION_KEY_BASE64: 'not-canonical-base64' },
            loadEnvironment() {
                events.push('load-environment');
            },
            createDependencies() {
                events.push('create-dependencies');
                throw new Error('must not construct dependencies');
            },
            createApplication() {
                events.push('create-application');
                throw new Error('must not create application');
            },
        }),
        /canonical base64 for exactly 32 bytes/,
    );

    assert.deepEqual(events, ['load-environment']);
});

test('explicit startup loads configuration before reading it', () => {
    const environment = {};
    const events = [];
    const listener = { close() {} };

    startProductionServer({
        environment,
        loadEnvironment() {
            events.push('load-environment');
            environment.PLATFORM_ROW_AUTHORIZATION_KEY_BASE64 = SIGNING_KEY_BASE64;
        },
        createDependencies(receivedEnvironment, receivedKey) {
            events.push('create-dependencies');
            assert.equal(receivedEnvironment, environment);
            assert.deepEqual(receivedKey, SIGNING_KEY);
            return {
                appDependencies: {},
                graphTokenLifecycle: {
                    start() { events.push('acquire'); },
                    stop() {},
                },
            };
        },
        createApplication() {
            return {
                listen(port) {
                    events.push(`listen:${port}`);
                    return listener;
                },
            };
        },
    });

    assert.deepEqual(events, [
        'load-environment',
        'create-dependencies',
        'listen:3000',
        'acquire',
    ]);
});

test('production dependency construction preserves SDK defaults and token provider behavior', async () => {
    const scheduler = createFakeScheduler();
    const graphClient = { name: 'synthetic-graph-client' };
    const faceClient = { name: 'synthetic-face-client' };
    const graphInitOptions = [];
    const msalConfigurations = [];
    const tokenRequests = [];
    const faceCredentials = [];
    const faceClientCalls = [];

    class SyntheticConfidentialClientApplication {
        constructor(configuration) {
            msalConfigurations.push(configuration);
        }

        async acquireTokenByClientCredential(request) {
            tokenRequests.push(request);
            return {
                accessToken: 'synthetic-access-token',
                expiresOn: new Date(NOW_MS + 10 * 60 * 1000),
            };
        }
    }

    class SyntheticAzureKeyCredential {
        constructor(key) {
            this.key = key;
            faceCredentials.push({ key, argumentCount: arguments.length, credential: this });
        }
    }

    const sdk = {
        Client: {
            init(options) {
                graphInitOptions.push(options);
                return graphClient;
            },
        },
        ConfidentialClientApplication: SyntheticConfidentialClientApplication,
        AzureKeyCredential: SyntheticAzureKeyCredential,
        FaceClient(endpoint, credential) {
            faceClientCalls.push({ endpoint, credential, argumentCount: arguments.length });
            return faceClient;
        },
    };
    const environment = {
        CLIENT_ID: 'synthetic-client-id',
        TENANT_ID: 'synthetic-tenant-id',
        CLIENT_SECRET: 'synthetic-client-secret',
        AZURE_FACE_API_KEY: 'synthetic-face-key',
        AZURE_FACE_API_ENDPOINT: 'https://synthetic-face.invalid',
    };

    const { appDependencies, graphTokenLifecycle } = createProductionDependencies(
        environment,
        SIGNING_KEY,
        {
            now: () => NOW_MS,
            schedule: scheduler.schedule,
            cancel: scheduler.cancel,
        },
        sdk,
    );

    assert.deepEqual(msalConfigurations, [{
        auth: {
            clientId: 'synthetic-client-id',
            authority: 'https://login.microsoftonline.com/synthetic-tenant-id',
            clientSecret: 'synthetic-client-secret',
        },
    }]);
    assert.equal(graphInitOptions.length, 1);
    assert.deepEqual(Object.keys(graphInitOptions[0]), ['authProvider']);

    function readGraphAccessToken() {
        let providerResult;
        graphInitOptions[0].authProvider((...args) => { providerResult = args; });
        return providerResult;
    }

    assert.deepEqual(readGraphAccessToken(), [null, undefined]);
    assert.deepEqual(faceCredentials.map(({ key, argumentCount }) => ({ key, argumentCount })), [{
        key: 'synthetic-face-key',
        argumentCount: 1,
    }]);
    assert.deepEqual(faceClientCalls, [{
        endpoint: 'https://synthetic-face.invalid',
        credential: faceCredentials[0].credential,
        argumentCount: 2,
    }]);
    assert.equal(appDependencies.graphClient, graphClient);
    assert.equal(appDependencies.faceClient, faceClient);
    assert.equal(typeof appDependencies.platformRowAuthorization.authorize, 'function');
    assert.equal(typeof appDependencies.platformRowAuthorization.createHandle, 'function');

    await graphTokenLifecycle.start();

    assert.deepEqual(tokenRequests, [{ scopes: [GRAPH_SCOPE] }]);
    assert.deepEqual(readGraphAccessToken(), [null, 'synthetic-access-token']);
    assert.deepEqual(scheduler.scheduled.map(({ delay }) => delay), [5 * 60 * 1000]);

    graphTokenLifecycle.stop();
    assert.deepEqual(scheduler.cancelled, [scheduler.scheduled[0]]);
});

test('startup selects configured PORT, listens before acquisition, and does not await readiness', () => {
    const { app, events, listener, result } = runSyntheticServer({ port: '4567' });

    assert.deepEqual(events, [
        'load-environment',
        'create-dependencies',
        'create-application',
        'listen:4567',
        'acquire',
    ]);
    assert.equal(result.app, app);
    assert.equal(result.listener, listener);
});

test('startup defaults to port 3000 and stops token lifecycle before closing listener', () => {
    const { closeResult, events, result } = runSyntheticServer();
    let closeCallbackCalled = false;

    const returned = result.stop(() => { closeCallbackCalled = true; });

    assert.equal(returned, closeResult);
    assert.equal(closeCallbackCalled, true);
    assert.deepEqual(events, [
        'load-environment',
        'create-dependencies',
        'create-application',
        'listen:3000',
        'acquire',
        'stop-token-lifecycle',
        'close',
    ]);
});

test('token success schedules expiry minus five minutes with a 60-second floor', async () => {
    const scheduler = createFakeScheduler();
    const accessTokens = [];
    const acquisition = createQueuedAcquireToken([
        {
            accessToken: 'first-token',
            expiresOn: new Date(NOW_MS + 15 * 60 * 1000),
        },
        {
            accessToken: 'second-token',
            expiresOn: new Date(NOW_MS + 5 * 60 * 1000 + 30_000),
        },
    ]);
    const lifecycle = createGraphTokenLifecycle({
        acquireToken: acquisition.acquireToken,
        setAccessToken: (token) => { accessTokens.push(token); },
        now: () => NOW_MS,
        schedule: scheduler.schedule,
        cancel: scheduler.cancel,
    });

    await lifecycle.start();
    const firstTimer = scheduler.scheduled[0];
    await firstTimer.callback();

    assert.deepEqual(acquisition.calls, [
        { scopes: [GRAPH_SCOPE] },
        { scopes: [GRAPH_SCOPE] },
    ]);
    assert.deepEqual(accessTokens, ['first-token', 'second-token']);
    assert.deepEqual(scheduler.scheduled.map(({ delay }) => delay), [
        10 * 60 * 1000,
        60_000,
    ]);
    assert.deepEqual(scheduler.cancelled, [firstTimer]);

    lifecycle.stop();
    assert.deepEqual(scheduler.cancelled, [firstTimer, scheduler.scheduled[1]]);
});

test('token failures back off to 60 seconds, success resets delay, and timers are replaced', async () => {
    const scheduler = createFakeScheduler();
    const accessTokens = [];
    const acquisition = createQueuedAcquireToken([
        new Error('failure 1'),
        new Error('failure 2'),
        new Error('failure 3'),
        new Error('failure 4'),
        new Error('failure 5'),
        new Error('failure 6'),
        new Error('failure 7'),
        {
            accessToken: 'recovered-token',
            expiresOn: new Date(NOW_MS + 10 * 60 * 1000),
        },
        new Error('failure after recovery'),
    ]);
    const lifecycle = createGraphTokenLifecycle({
        acquireToken: acquisition.acquireToken,
        setAccessToken: (token) => { accessTokens.push(token); },
        now: () => NOW_MS,
        schedule: scheduler.schedule,
        cancel: scheduler.cancel,
    });

    await lifecycle.start();

    for (let index = 0; index < 8; index++) {
        await scheduler.scheduled[index].callback();
    }

    assert.deepEqual(scheduler.scheduled.map(({ delay }) => delay), [
        2_000,
        4_000,
        8_000,
        16_000,
        32_000,
        60_000,
        60_000,
        5 * 60 * 1000,
        2_000,
    ]);
    assert.deepEqual(accessTokens, ['recovered-token']);
    assert.equal(acquisition.calls.length, 9);
    assert.ok(acquisition.calls.every((request) => (
        request.scopes.length === 1 && request.scopes[0] === GRAPH_SCOPE
    )));
    assert.deepEqual(
        scheduler.cancelled,
        scheduler.scheduled.slice(0, -1),
    );

    lifecycle.stop();
    lifecycle.stop();

    assert.deepEqual(scheduler.cancelled, scheduler.scheduled);
});

test('stopping during token acquisition prevents a later timer from being scheduled', async () => {
    const scheduler = createFakeScheduler();
    const accessTokens = [];
    let resolveAcquisition;
    const acquisition = new Promise((resolve) => { resolveAcquisition = resolve; });
    const lifecycle = createGraphTokenLifecycle({
        acquireToken: () => acquisition,
        setAccessToken: (token) => { accessTokens.push(token); },
        now: () => NOW_MS,
        schedule: scheduler.schedule,
        cancel: scheduler.cancel,
    });

    const start = lifecycle.start();
    lifecycle.stop();
    resolveAcquisition({
        accessToken: 'late-token',
        expiresOn: new Date(NOW_MS + 10 * 60 * 1000),
    });
    await start;

    assert.deepEqual(accessTokens, []);
    assert.deepEqual(scheduler.scheduled, []);
    assert.deepEqual(scheduler.cancelled, []);
});
