'use strict';

const path = require('node:path');
const { createApp } = require('./app');
const {
    createPlatformRowAuthorizationHandle,
    createPlatformRowAuthorizer,
    decodePlatformRowAuthorizationKey,
} = require('./platform-row-authorization');

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

function loadProductionSdk() {
    return {
        Client: require('@microsoft/microsoft-graph-client').Client,
        ConfidentialClientApplication: require('@azure/msal-node').ConfidentialClientApplication,
        AzureKeyCredential: require('@azure/core-auth').AzureKeyCredential,
        FaceClient: require('@azure-rest/ai-vision-face').default,
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
            },
        },
        graphTokenLifecycle,
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
    } = createDependencies(environment, platformRowAuthorizationKey, runtimeHooks);
    const app = createApplication(appDependencies);
    const listener = app.listen(environment.PORT || 3000);

    graphTokenLifecycle.start();

    return {
        app,
        listener,
        stop(callback) {
            graphTokenLifecycle.stop();
            return listener.close(callback);
        },
    };
}

// IISNode requires this file through its interceptor but preserves the application path in argv[1].
if (require.main === module || isIisnodeEntryPoint()) startProductionServer();

module.exports = {
    createGraphTokenLifecycle,
    createProductionDependencies,
    startProductionServer,
};
