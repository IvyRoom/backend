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
    const port = Number(environment.SUPPORT_SANDBOX_PORT || 4174);
    if (!Number.isInteger(port) || port < 1024 || port > 65535) throw new Error('Invalid sandbox port');
    // Validate test keys before initializing any external client.
    const pagarme = createPagarmeAdapter({ secretKey: environment.SUPPORT_PAGARME_SECRET_KEY, publicKey: environment.SUPPORT_PAGARME_PUBLIC_KEY });
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
        publicKey: environment.SUPPORT_PAGARME_PUBLIC_KEY,
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
    startSupportSandbox().then(({ stop, origin }) => {
        console.log(`Support sandbox: ${origin}/support-management/login/`);
        const shutdown = () => { void stop(); };
        process.once('SIGINT', shutdown);
        process.once('SIGTERM', shutdown);
    }).catch(() => {
        console.error('Sandbox could not start. Check local Graph configuration, Pagar.me test keys, and port availability.');
        process.exitCode = 1;
    });
}

module.exports = { startSupportSandbox };
