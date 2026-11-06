'use strict';

const path = require('node:path');
const { createSupportApp } = require('./support-app');
const { createSupportUsersAdapter } = require('./integrations/support-users');
const { createPagarmeAdapter } = require('./integrations/pagarme');
const { createSupportCardStore, acquireSupportSandboxLock } = require('./integrations/support-card-store');

async function startSupportSandbox({ environment = process.env } = {}) {
    const dotenv = require('dotenv');
    dotenv.config({ path: path.join(__dirname, '.env'), quiet: true });
    dotenv.config({ path: path.join(__dirname, '.env.support-sandbox'), quiet: true });
    const secretKey = environment.SUPPORT_PAGARME_SECRET_KEY?.trim() || environment.PAGARME_API_KEY_SANDBOX?.trim();
    const publicKey = environment.SUPPORT_PAGARME_PUBLIC_KEY?.trim() || environment.PAGARME_PUBLIC_KEY_SANDBOX?.trim();
    const missing = ['CLIENT_ID', 'TENANT_ID', 'CLIENT_SECRET']
        .filter((name) => !environment[name]?.trim());
    if (!publicKey) missing.push('SUPPORT_PAGARME_PUBLIC_KEY (or PAGARME_PUBLIC_KEY_SANDBOX)');
    if (!secretKey) missing.push('SUPPORT_PAGARME_SECRET_KEY (or PAGARME_API_KEY_SANDBOX)');
    if (missing.length) {
        const error = new Error(`Missing local settings: ${missing.join(', ')}. Add them to the ignored local environment files.`);
        error.code = 'SUPPORT_SETUP';
        throw error;
    }
    if (!/^pk_test_[A-Za-z0-9_-]+$/u.test(publicKey)
        || !/^sk_test_[A-Za-z0-9_-]+$/u.test(secretKey)) {
        const error = new Error('SUPPORT_PAGARME_PUBLIC_KEY and the configured Pagar.me secret must be Pagar.me test keys.');
        error.code = 'SUPPORT_SETUP';
        throw error;
    }
    const port = Number(environment.SUPPORT_SANDBOX_PORT || 4174);
    if (!Number.isInteger(port) || port < 1024 || port > 65535) throw new Error('Invalid sandbox port');
    // Validate test keys before initializing any external client.
    const pagarme = createPagarmeAdapter({ secretKey, publicKey });
    const { ConfidentialClientApplication } = require('@azure/msal-node');
    const { Client } = require('@microsoft/microsoft-graph-client');
    const confidentialClient = new ConfidentialClientApplication({ auth: {
        clientId: environment.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${environment.TENANT_ID}`,
        clientSecret: environment.CLIENT_SECRET,
    } });
    const graphClient = Client.initWithMiddleware({ authProvider: {
        async getAccessToken() {
            const response = await confidentialClient.acquireTokenByClientCredential({ scopes: ['https://graph.microsoft.com/.default'] });
            return response.accessToken;
        },
    } });
    const origin = `http://127.0.0.1:${port}`;
    const app = createSupportApp({
        origin,
        frontendDirectory: path.resolve(__dirname, '../sistemas/apps/support-management'),
        users: createSupportUsersAdapter({ graphClient }),
        pagarme,
        cardStore: createSupportCardStore(path.join(__dirname, '.support-sandbox-data')),
        publicKey,
    });
    const directory = path.join(__dirname, '.support-sandbox-data');
    const releaseLock = await acquireSupportSandboxLock(directory);
    async function releaseLockSafely() {
        try {
            await releaseLock();
        } catch {
            console.error('Sandbox lock could not be released. Review the local data directory before restarting.');
            process.exitCode = 1;
        }
    }
    let listener;
    try {
        listener = app.listen(port, '127.0.0.1');
        await new Promise((resolve, reject) => { listener.once('listening', resolve); listener.once('error', reject); });
    } catch (error) {
        await releaseLockSafely();
        throw error;
    }
    const stopped = new Promise((resolve) => {
        listener.once('close', () => { void releaseLockSafely().then(resolve); });
    });
    return {
        app,
        listener,
        origin,
        async stop() {
            if (listener.listening) listener.close();
            await stopped;
        },
    };
}

if (require.main === module) {
    let sandbox;
    let shutdownRequested = false;
    let shutdownPromise;
    const disconnect = () => { if (process.connected) process.disconnect(); };
    const shutdown = () => {
        shutdownRequested = true;
        if (!sandbox) return;
        shutdownPromise ??= sandbox.stop().then(disconnect).catch(() => {
            console.error('Sandbox shutdown failed. Review the local lock before restarting.');
            process.exitCode = 1;
            disconnect();
        });
    };
    process.once('SIGINT', shutdown);
    process.once('SIGTERM', shutdown);
    process.once('disconnect', shutdown);
    process.on('message', (message) => { if (message?.type === 'support:stop') shutdown(); });
    startSupportSandbox().then((instance) => {
        sandbox = instance;
        const { origin } = instance;
        console.log(`Support sandbox: ${origin}/support-management/login/`);
        if (shutdownRequested) shutdown();
        else if (process.connected) process.send({ type: 'support:ready', origin }, () => {});
    }).catch((error) => {
        console.error(error.code === 'SUPPORT_SETUP' ? error.message : 'Sandbox could not start. Check local Graph configuration, Pagar.me test keys, port availability, and the local lock.');
        process.exitCode = 1;
        disconnect();
    });
}

module.exports = { startSupportSandbox };
