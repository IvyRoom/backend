'use strict';

require('./deny-production-network');
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const fs = require('node:fs/promises');
const os = require('node:os');
const express = require('express');
const { createSupportApp } = require('../support-app');
const { createTestApp, startLoopback } = require('./app-test-support');

const ORIGIN = 'https://machadogestao.com';
const user = { companyId: '12345678000195', company: 'Empresa Exemplo', name: 'Pessoa Exemplo',
    email: 'person@example.invalid', password: 'synthetic-password', role: 'Administrador', bookingAllowed: true, referenceRate: 1000 };

function support(overrides = {}) {
    return createSupportApp({ environment: 'production', apiOnly: true, origin: ORIGIN,
        publicKey: 'pk_productionSynthetic', preload: true,
        users: { readUsers: async () => [user] },
        cardStore: { read: async () => ({}) },
        pagarme: { getCard: async () => { throw new Error('No card lookup expected'); } },
        ...overrides });
}

test('production support API mounts before broad legacy middleware and only permits the canonical frontend origin', async (t) => {
    let reads = 0;
    const { app, ledger } = createTestApp({ dependencies: { supportApp: support({ users: { readUsers: async () => { reads++; return [user]; } } }) } });
    const { origin } = await startLoopback(app, t);
    const base = `${origin}/atendimentos/api`;
    for (const invalid of ['https://attacker.invalid', 'https://www.machadogestao.com', 'http://machadogestao.com', 'null']) {
        const response = await fetch(`${base}/config`, { headers: { Origin: invalid } });
        assert.equal(response.status, 403);
        assert.equal(response.headers.get('access-control-allow-origin'), null);
    }
    const preflight = await fetch(`${base}/login`, { method: 'OPTIONS', headers: { Origin: ORIGIN,
        'Access-Control-Request-Method': 'POST', 'Access-Control-Request-Headers': 'content-type,authorization' } });
    assert.equal(preflight.status, 204);
    assert.equal(preflight.headers.get('access-control-allow-origin'), ORIGIN);
    assert.match(preflight.headers.get('access-control-allow-headers'), /Authorization/u);
    assert.equal(reads, 0);
    const config = await fetch(`${base}/config`, { headers: { Origin: ORIGIN } });
    assert.equal(config.status, 200);
    assert.deepEqual(await config.json(), { environment: 'production', publicKey: 'pk_productionSynthetic' });
    assert.equal(config.headers.get('access-control-allow-origin'), ORIGIN);
    assert.equal(config.headers.get('cache-control'), 'no-store');
    const login = await fetch(`${base}/login`, { method: 'POST', headers: { Origin: ORIGIN, 'Content-Type': 'application/json' },
        body: JSON.stringify({ email: user.email, password: user.password }) });
    const loggedIn = await login.json();
    assert.equal(login.status, 200);
    assert.deepEqual(loggedIn.config, { environment: 'production', publicKey: 'pk_productionSynthetic' });
    assert.equal(loggedIn.calendar.paymentMethodRegistered, false);
    assert.equal(loggedIn.user.name, user.name);
    assert.equal(loggedIn.user.password, undefined);
    assert.ok(loggedIn.expiresAt > Date.now() && loggedIn.expiresAt <= Date.now() + 3600_000);
    const unauthorized = await fetch(`${base}/calendar`, { headers: { Origin: ORIGIN } });
    assert.equal(unauthorized.status, 401);
    const oversized = await fetch(`${base}/login`, { method: 'POST', headers: { Origin: ORIGIN, 'Content-Type': 'application/json' },
        body: JSON.stringify({ data: 'x'.repeat(9000) }) });
    assert.equal(oversized.status, 413);
    const legacy = await fetch(`${origin}/unknown-legacy-route`);
    assert.equal(legacy.status, 404);
    assert.equal(legacy.headers.get('access-control-allow-origin'), '*');
    assert.deepEqual(ledger, []);
});

test('local source server serves the Portuguese pages and shared origin module without exposing English source routes', async (t) => {
    const root = await fs.mkdtemp(path.join(os.tmpdir(), 'support-source-routes-'));
    t.after(async () => {
        const resolved = path.resolve(root);
        assert.ok(resolved.startsWith(`${path.resolve(os.tmpdir())}${path.sep}`));
        await fs.rm(resolved, { recursive: true, force: true });
    });
    const directory = path.join(root, 'apps', 'support-management');
    for (const [name, source] of Object.entries({ 'login/index.html': '<h1>Login</h1>', 'scheduler/index.html': '<h1>Agendamento</h1>',
        'modules/client.js': 'export const synthetic = true;', 'style.css': 'body { color: black; }',
        '../shared/backend-origin.js': "export const BACKEND_ORIGIN = 'https://backend.example.invalid';" })) {
        const file = path.resolve(directory, name);
        await fs.mkdir(path.dirname(file), { recursive: true });
        await fs.writeFile(file, source);
    }
    const app = createSupportApp({ origin: 'http://127.0.0.1:4174', frontendDirectory: directory });
    const wrapper = express();
    wrapper.use((req, res) => {
        req.headers.host = '127.0.0.1:4174';
        return app(req, res);
    });
    const { origin } = await startLoopback(wrapper, t);
    for (const route of ['/atendimentos/login/', '/atendimentos/agendamento/', '/atendimentos/style.css', '/atendimentos/modules/client.js', '/shared/backend-origin.js']) {
        assert.equal((await fetch(`${origin}${route}`)).status, 200, route);
    }
    for (const route of ['/atendimentos/scheduler/', '/support-management/login/', '/support-production.js', '/.env']) {
        assert.equal((await fetch(`${origin}${route}`)).status, 404, route);
    }
});

test('production support factory rejects noncanonical and sandbox-origin combinations', () => {
    for (const options of [{ environment: 'production', apiOnly: true, origin: 'https://www.machadogestao.com' },
        { environment: 'production', apiOnly: false, origin: ORIGIN },
        { environment: 'sandbox', apiOnly: true, origin: 'http://127.0.0.1:4174' },
        { environment: 'unknown', apiOnly: true, origin: ORIGIN }]) {
        assert.throws(() => createSupportApp(options));
    }
});
