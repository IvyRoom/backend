'use strict';

const { randomUUID } = require('node:crypto');
const path = require('node:path');
const { createApp } = require('./app');
const {
    createPlatformRowAuthorizationHandle,
    createPlatformRowAuthorizationInspector,
    createPlatformRowAuthorizer,
    decodePlatformRowAuthorizationKey,
} = require('./platform-row-authorization');
const {
    assertSeparateLegacySigningKey,
    readSessionAuthorityConfiguration,
} = require('./domains/session-authority/configuration');
const {
    LEGACY_SEEDING_HEARTBEAT_INTERVAL_MS,
} = require('./domains/session-authority/constants');
const { createAzureSqlSessionStore } = require('./integrations/azure-sql-session-store');

const MICROSOFT_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const MICROSOFT_GRAPH_INITIAL_DELAY = 2000;
const MICROSOFT_GRAPH_MAXIMUM_DELAY = 60000;
const MICROSOFT_GRAPH_REFRESH_MARGIN = 5 * 60 * 1000;

function isIisnodeEntryPoint() {
    return typeof process.argv[1] === 'string' && path.resolve(process.argv[1]) === path.resolve(__filename);
}

function createGraphTokenLifecycle({
    acquireToken,
    setAccessToken,
    now = () => Date.now(),
    schedule = (callback, delay) => setTimeout(callback, delay),
    cancel = (timer) => clearTimeout(timer),
}) {
    let refreshTimer;
    let retryDelay = MICROSOFT_GRAPH_INITIAL_DELAY;
    let stopped = false;

    function replaceTimer(delay) {
        if (refreshTimer !== undefined) cancel(refreshTimer);
        refreshTimer = schedule(refresh, delay);
    }

    async function refresh() {
        if (stopped) return;

        try {
            const response = await acquireToken({ scopes: [MICROSOFT_GRAPH_SCOPE] });
            if (stopped) return;
            setAccessToken(response.accessToken);
            replaceTimer(Math.max(new Date(response.expiresOn).getTime() - now() - MICROSOFT_GRAPH_REFRESH_MARGIN, MICROSOFT_GRAPH_MAXIMUM_DELAY));
            retryDelay = MICROSOFT_GRAPH_INITIAL_DELAY;
        } catch {
            if (stopped) return;
            replaceTimer(retryDelay);
            retryDelay = Math.min(retryDelay * 2, MICROSOFT_GRAPH_MAXIMUM_DELAY);
        }
    }

    function stop() {
        stopped = true;
        if (refreshTimer !== undefined) cancel(refreshTimer);
        refreshTimer = undefined;
    }

    return { start: refresh, stop };
}

function createSessionAuthorityContinuityLifecycle({
    store,
    enabled,
    ownerId = randomUUID(),
    schedule = (callback, delay) => setTimeout(callback, delay),
    cancel = (timer) => clearTimeout(timer),
} = {}) {
    if (!store || typeof store.close !== 'function') {
        throw new TypeError('Session authority store lifecycle is required');
    }
    if (enabled && typeof store.heartbeatLegacySeedingContinuity !== 'function') {
        throw new TypeError('Session authority continuity heartbeat is required');
    }
    let timer;
    let inFlight;
    let started = false;
    let stopped = false;

    function replaceTimer() {
        if (stopped || !enabled) return;
        timer = schedule(
            () => (inFlight = run()),
            LEGACY_SEEDING_HEARTBEAT_INTERVAL_MS,
        );
    }

    async function run() {
        if (stopped || !enabled) return;
        try {
            await store.heartbeatLegacySeedingContinuity({ ownerId });
        } catch {
            // The store privately latches failures; its next successful heartbeat resets continuity.
        } finally {
            if (!stopped) replaceTimer();
        }
    }

    function start() {
        if (started) return inFlight;
        if (stopped || !enabled) return undefined;
        started = true;
        inFlight = run();
        return inFlight;
    }

    async function stop() {
        stopped = true;
        if (timer !== undefined) cancel(timer);
        timer = undefined;
        if (inFlight) await inFlight;
        await store.close();
    }

    return { start, stop };
}

function loadProductionSdk() {
    return {
        Client: require('@microsoft/microsoft-graph-client').Client,
        ConfidentialClientApplication: require('@azure/msal-node').ConfidentialClientApplication,
        AzureKeyCredential: require('@azure/core-auth').AzureKeyCredential,
        FaceClient: require('@azure-rest/ai-vision-face').default,
        sql: require('mssql'),
    };
}

function createProductionDependencies(environment, platformRowAuthorizationKey, runtimeHooks = {}, sdk = loadProductionSdk()) {
    const { Client, ConfidentialClientApplication, AzureKeyCredential, FaceClient } = sdk;

    const confidentialClient = new ConfidentialClientApplication({
        auth: {
            clientId: environment.CLIENT_ID,
            authority: `https://login.microsoftonline.com/${environment.TENANT_ID}`,
            clientSecret: environment.CLIENT_SECRET,
        },
    });

    let Microsoft_Graph_API_AccessToken;
    const graphClient = Client.init({
        authProvider: (done) => { done(null, Microsoft_Graph_API_AccessToken); },
    });

    const graphTokenLifecycle = createGraphTokenLifecycle({
        acquireToken: (request) => confidentialClient.acquireTokenByClientCredential(request),
        setAccessToken: (accessToken) => { Microsoft_Graph_API_AccessToken = accessToken; },
        ...runtimeHooks,
    });

    const faceCredential = new AzureKeyCredential(environment.AZURE_FACE_API_KEY);
    const faceClient = FaceClient(environment.AZURE_FACE_API_ENDPOINT, faceCredential);
    const sessionConfiguration = readSessionAuthorityConfiguration(environment);
    let sessionAuthority;
    let sessionAuthorityLifecycle;

    if (sessionConfiguration.enabled) {
        if (!sdk.sql) throw new TypeError('The Azure SQL driver is required when session authority is enabled');
        assertSeparateLegacySigningKey(sessionConfiguration.keys, platformRowAuthorizationKey);
        const sessionStore = createAzureSqlSessionStore({
            sql: sdk.sql,
            connectionString: sessionConfiguration.connectionString,
            expectedAuthorityGeneration: sessionConfiguration.expectedAuthorityGeneration,
            loginLookupKeyId: sessionConfiguration.loginLookupKeyBinding.keyId,
            loginLookupKeyCommitment: sessionConfiguration.loginLookupKeyBinding.commitment,
            accountMappingKeyBinding: sessionConfiguration.accountMappingKeyBinding,
            authorityKeysetBinding: sessionConfiguration.authorityKeysetBinding,
            legacySigningKeyBinding: sessionConfiguration.legacySigningKeyBinding,
            options: sessionConfiguration.sqlOptions,
        });
        sessionAuthority = {
            store: sessionStore,
            keys: sessionConfiguration.keys,
            runtimeControls: sessionConfiguration.runtimeControls,
            createSubjectId: randomUUID,
            createSessionId: randomUUID,
            createFlowId: randomUUID,
            createCorrelationId: randomUUID,
        };
        sessionAuthorityLifecycle = createSessionAuthorityContinuityLifecycle({
            store: sessionStore,
            enabled: sessionConfiguration.runtimeControls.legacyLedgerSeedingEnabled,
            schedule: runtimeHooks.schedule,
            cancel: runtimeHooks.cancel,
        });
    }

    return {
        appDependencies: {
            graphClient,
            faceClient,
            platformRowAuthorization: {
                authorize: createPlatformRowAuthorizer(platformRowAuthorizationKey),
                createHandle: (rowIndex, nowMs) => createPlatformRowAuthorizationHandle(
                    rowIndex,
                    platformRowAuthorizationKey,
                    nowMs,
                ),
                inspectHandle: createPlatformRowAuthorizationInspector(platformRowAuthorizationKey),
            },
            ...(sessionAuthority ? { sessionAuthority } : {}),
        },
        graphTokenLifecycle,
        ...(sessionAuthorityLifecycle ? { sessionAuthorityLifecycle } : {}),
    };
}

function startProductionServer({
    environment = process.env,
    loadEnvironment = () => require('dotenv').config(),
    createDependencies = createProductionDependencies,
    createApplication = createApp,
    runtimeHooks,
} = {}) {
    loadEnvironment();

    const platformRowAuthorizationKey = decodePlatformRowAuthorizationKey(environment.PLATFORM_ROW_AUTHORIZATION_KEY_BASE64);
    const {
        appDependencies,
        graphTokenLifecycle,
        sessionAuthorityLifecycle,
    } = createDependencies(environment, platformRowAuthorizationKey, runtimeHooks);
    const app = createApplication(appDependencies);
    const listener = app.listen(environment.PORT || 3000);

    graphTokenLifecycle.start();
    if (sessionAuthorityLifecycle) sessionAuthorityLifecycle.start();

    return {
        app,
        listener,
        stop(callback) {
            graphTokenLifecycle.stop();
            return listener.close((listenerError) => {
                if (!sessionAuthorityLifecycle) {
                    if (callback) callback(listenerError);
                    return;
                }
                Promise.resolve(sessionAuthorityLifecycle.stop()).then(
                    () => { if (callback) callback(listenerError); },
                    (storeError) => { if (callback) callback(storeError); },
                );
            });
        },
    };
}

// IISNode requires this file through its interceptor but preserves the application path in argv[1].
if (require.main === module || isIisnodeEntryPoint()) startProductionServer();

module.exports = {
    createGraphTokenLifecycle,
    createProductionDependencies,
    createSessionAuthorityContinuityLifecycle,
    startProductionServer,
};
