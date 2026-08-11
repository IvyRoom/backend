'use strict';

process.env.NODE_ENV = 'test';

const test = require('node:test');
const assert = require('node:assert/strict');
const { spawnSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {
    createTestApp,
    startLoopback,
} = require('./app-test-support');

const REPOSITORY_ROOT = path.resolve(__dirname, '..');
const SEND_MAIL_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail';

async function assertDefaultHtml(response, expectedStatus) {
    assert.equal(response.status, expectedStatus);
    assert.match(response.headers.get('content-type') || '', /^text\/html;/i);
    assert.equal(response.headers.get('access-control-allow-origin'), '*');
    const body = await response.text();
    assert.match(body, /<!DOCTYPE html>|<html>/i);
    return body;
}

test('importing app and server without production configuration is side-effect free', (t) => {
    const temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'backend-import-safety-'));
    t.after(() => fs.rmSync(temporaryDirectory, { recursive: true, force: true }));

    const environment = { ...process.env, NODE_ENV: 'test' };
    for (const name of [
        'PLATFORM_ROW_AUTHORIZATION_KEY_BASE64',
        'CLIENT_ID',
        'TENANT_ID',
        'CLIENT_SECRET',
        'AZURE_FACE_API_KEY',
        'AZURE_FACE_API_ENDPOINT',
        'PORT',
    ]) {
        delete environment[name];
    }

    const script = [
        "'use strict';",
        'require(process.argv[1]);',
        'require(process.argv[2]);',
        "process.stdout.write('imports-ok');",
    ].join('\n');
    const result = spawnSync(
        process.execPath,
        [
            '-e',
            script,
            path.join(REPOSITORY_ROOT, 'app.js'),
            path.join(REPOSITORY_ROOT, 'server.js'),
        ],
        {
            cwd: temporaryDirectory,
            env: environment,
            encoding: 'utf8',
            timeout: 5_000,
            windowsHide: true,
        },
    );

    assert.equal(result.error, undefined, result.error && result.error.message);
    assert.equal(result.signal, null);
    assert.equal(result.status, 0, result.stderr);
    assert.equal(result.stdout, 'imports-ok');
});

test('CORS defaults and preflight terminate before route dependencies', async (t) => {
    const { app, graphClient, faceClient } = createTestApp();
    const { origin } = await startLoopback(app, t);

    const ordinaryResponse = await fetch(`${origin}/ezdrm-playready-authorization-url`);
    assert.equal(ordinaryResponse.status, 200);
    assert.equal(ordinaryResponse.headers.get('access-control-allow-origin'), '*');
    assert.equal(ordinaryResponse.headers.get('access-control-allow-credentials'), null);
    await ordinaryResponse.text();

    const preflightResponse = await fetch(`${origin}/landingpage/solicitacaoorcamento`, {
        method: 'OPTIONS',
        headers: {
            Origin: 'https://example.test',
            'Access-Control-Request-Method': 'POST',
            'Access-Control-Request-Headers': 'X-Contract-Test, Content-Type',
        },
    });

    assert.equal(preflightResponse.status, 204);
    assert.equal(await preflightResponse.text(), '');
    assert.equal(preflightResponse.headers.get('access-control-allow-origin'), '*');
    assert.equal(
        preflightResponse.headers.get('access-control-allow-methods'),
        'GET,HEAD,PUT,PATCH,POST,DELETE',
    );
    assert.equal(
        preflightResponse.headers.get('access-control-allow-headers'),
        'X-Contract-Test, Content-Type',
    );
    assert.equal(preflightResponse.headers.get('access-control-allow-credentials'), null);
    assert.equal(graphClient.calls.length, 0);
    assert.equal(faceClient.calls.length, 0);
});

test('JSON middleware accepts objects, arrays, and empty bodies', async (t) => {
    const { app, graphClient, faceClient } = createTestApp();
    graphClient.enqueue('POST', SEND_MAIL_PATH, {}, {}, {});
    const { origin } = await startLoopback(app, t);

    for (const request of [
        { headers: { 'Content-Type': 'application/json' }, body: '{}' },
        { headers: { 'Content-Type': 'application/json' }, body: '[]' },
        {},
    ]) {
        const response = await fetch(`${origin}/landingpage/solicitacaoorcamento`, {
            method: 'POST',
            ...request,
        });
        assert.equal(response.status, 200);
        assert.equal(response.headers.get('content-type'), 'application/json; charset=utf-8');
        assert.deepEqual(await response.json(), {});
    }

    assert.equal(graphClient.calls.length, 3);
    assert.equal(faceClient.calls.length, 0);
    graphClient.assertExhausted();
});

test('URL-encoded bodies remain unparsed', async (t) => {
    const { app, graphClient, faceClient } = createTestApp();
    const { origin } = await startLoopback(app, t);
    const form = new URLSearchParams({
        recommenderFullName: 'Pessoa Recomendante',
        benefitedCompany: 'Empresa Beneficiada',
        recommendedCompany: 'Empresa Recomendada',
        recommendedProfessional: 'Pessoa Profissional',
        recommendedWhatsapp: '+55 41 99999-9999',
    });

    const response = await fetch(`${origin}/conecta/processa-recomendacao`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: form.toString(),
    });

    assert.equal(response.status, 400);
    assert.deepEqual(await response.json(), { error: 'Erro_014' });
    assert.equal(graphClient.calls.length, 0);
    assert.equal(faceClient.calls.length, 0);
});

test('strict, malformed, oversized, and unsupported-charset JSON use default HTML errors', async (t) => {
    const { app, graphClient, faceClient } = createTestApp();
    const { origin } = await startLoopback(app, t);

    const cases = [
        {
            name: 'strict primitive',
            path: '/landingpage/solicitacaoorcamento',
            body: 'true',
            contentType: 'application/json',
            status: 400,
        },
        {
            name: 'malformed before static handling',
            path: '/img/LOGO_PAGAR.ME.png',
            body: '{',
            contentType: 'application/json',
            status: 400,
        },
        {
            name: 'over 100 KiB',
            path: '/landingpage/solicitacaoorcamento',
            body: JSON.stringify({ value: 'x'.repeat(100 * 1024) }),
            contentType: 'application/json',
            status: 413,
        },
        {
            name: 'unsupported charset',
            path: '/landingpage/solicitacaoorcamento',
            body: '{}',
            contentType: 'application/json; charset=iso-8859-1',
            status: 415,
        },
    ];

    for (const contractCase of cases) {
        await t.test(contractCase.name, async () => {
            const response = await fetch(`${origin}${contractCase.path}`, {
                method: 'POST',
                headers: { 'Content-Type': contractCase.contentType },
                body: contractCase.body,
            });
            await assertDefaultHtml(response, contractCase.status);
        });
    }

    assert.equal(graphClient.calls.length, 0);
    assert.equal(faceClient.calls.length, 0);
});

test('/img preserves redirect, asset GET and HEAD, and fallthrough behavior', async (t) => {
    const { app, graphClient, faceClient } = createTestApp();
    const { origin } = await startLoopback(app, t);

    const redirectResponse = await fetch(`${origin}/img`, { redirect: 'manual' });
    assert.equal(redirectResponse.status, 301);
    assert.equal(redirectResponse.headers.get('location'), '/img/');
    assert.match(redirectResponse.headers.get('content-type') || '', /^text\/html;/i);
    assert.match(await redirectResponse.text(), /Redirecting to \/img\//);

    for (const asset of [
        { name: 'LOGO_PAGAR.ME.png', contentType: 'image/png' },
        { name: 'ASSINATURA_E-MAIL.jpg', contentType: 'image/jpeg' },
    ]) {
        const expectedBytes = fs.readFileSync(path.join(REPOSITORY_ROOT, 'img', asset.name));
        const assetUrl = `${origin}/img/${asset.name}`;

        const getResponse = await fetch(assetUrl);
        assert.equal(getResponse.status, 200);
        assert.equal(getResponse.headers.get('content-type'), asset.contentType);
        assert.equal(Number(getResponse.headers.get('content-length')), expectedBytes.length);
        assert.deepEqual(Buffer.from(await getResponse.arrayBuffer()), expectedBytes);

        const headResponse = await fetch(assetUrl, { method: 'HEAD' });
        assert.equal(headResponse.status, 200);
        assert.equal(headResponse.headers.get('content-type'), asset.contentType);
        assert.equal(Number(headResponse.headers.get('content-length')), expectedBytes.length);
        assert.equal((await headResponse.arrayBuffer()).byteLength, 0);
    }

    const directoryMiss = await fetch(`${origin}/img/`);
    const directoryBody = await assertDefaultHtml(directoryMiss, 404);
    assert.match(directoryBody, /Cannot GET \/img\//);

    const assetMiss = await fetch(`${origin}/img/not-present.png`);
    const assetMissBody = await assertDefaultHtml(assetMiss, 404);
    assert.match(assetMissBody, /Cannot GET \/img\/not-present\.png/);

    const unsupportedMethod = await fetch(`${origin}/img/LOGO_PAGAR.ME.png`, { method: 'POST' });
    const unsupportedBody = await assertDefaultHtml(unsupportedMethod, 404);
    assert.match(unsupportedBody, /Cannot POST \/img\/LOGO_PAGAR\.ME\.png/);

    assert.equal(graphClient.calls.length, 0);
    assert.equal(faceClient.calls.length, 0);
});

test('unmatched requests retain Express default CORS-bearing HTML 404', async (t) => {
    const { app, graphClient, faceClient } = createTestApp();
    const { origin } = await startLoopback(app, t);

    const response = await fetch(`${origin}/not-registered`);
    const body = await assertDefaultHtml(response, 404);
    assert.match(body, /Cannot GET \/not-registered/);
    assert.equal(graphClient.calls.length, 0);
    assert.equal(faceClient.calls.length, 0);
});

test('DRM output preserves exact HTML text, query casing, and empty defaults', async (t) => {
    const { app, graphClient, faceClient } = createTestApp();
    const { origin } = await startLoopback(app, t);

    const token = 'a b&c/+?';
    const customData = 'Árvore=1&x y';
    const url = new URL('/ezdrm-playready-authorization-url', origin);
    url.searchParams.set('token', token);
    url.searchParams.set('CustomData', customData);

    const response = await fetch(url);
    assert.equal(response.status, 200);
    assert.equal(response.headers.get('content-type'), 'text/html; charset=utf-8');
    assert.equal(
        await response.text(),
        `p1=5&p2=&p3=&p4=1&p5=0&p6=1&p7=0&p8=0&token=${encodeURIComponent(token)}&CustomData=${encodeURIComponent(customData)}`,
    );

    const wrongCaseResponse = await fetch(
        `${origin}/ezdrm-playready-authorization-url?Token=ignored&customdata=ignored`,
    );
    assert.equal(wrongCaseResponse.status, 200);
    assert.equal(
        await wrongCaseResponse.text(),
        'p1=5&p2=&p3=&p4=1&p5=0&p6=1&p7=0&p8=0&token=&CustomData=',
    );

    const emptyResponse = await fetch(`${origin}/ezdrm-playready-authorization-url`);
    assert.equal(emptyResponse.status, 200);
    assert.equal(
        await emptyResponse.text(),
        'p1=5&p2=&p3=&p4=1&p5=0&p6=1&p7=0&p8=0&token=&CustomData=',
    );

    assert.equal(graphClient.calls.length, 0);
    assert.equal(faceClient.calls.length, 0);
});

test('routing remains case-insensitive and accepts a trailing slash', async (t) => {
    const { app, graphClient, faceClient } = createTestApp();
    const { origin } = await startLoopback(app, t);

    const response = await fetch(
        `${origin}/EZDRM-PLAYREADY-AUTHORIZATION-URL/?token=Case&CustomData=Slash`,
    );
    assert.equal(response.status, 200);
    assert.equal(response.headers.get('content-type'), 'text/html; charset=utf-8');
    assert.equal(
        await response.text(),
        'p1=5&p2=&p3=&p4=1&p5=0&p6=1&p7=0&p8=0&token=Case&CustomData=Slash',
    );
    assert.equal(graphClient.calls.length, 0);
    assert.equal(faceClient.calls.length, 0);
});
