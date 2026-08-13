'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createTestApp } = require('./app-test-support');

const EXPECTED_ROUTES = [
    ['post', '/landingpage/solicitacaoorcamento'],
    ['post', '/conecta/processa-recomendacao'],
    ['post', '/clientes/processa-formulario'],
    ['post', '/clientes/liberacao-acesso-plataforma'],
    ['post', '/plataforma_v2/login-FaceID'],
    ['post', '/plataforma_v2/CadastroFoto_e_FaceID'],
    ['post', '/plataforma_v2/FaceID'],
    ['get', '/plataforma_v2/FaceID_resultado/:Azure_Face_API_LivenessSession_sessionID'],
    ['post', '/plataforma_v2/refresh'],
    ['post', '/plataforma_v2/updates'],
    ['post', '/plataforma_v2/processa-feedback'],
    ['get', '/ezdrm-playready-authorization-url'],
    ['post', '/plataforma_v2/statusreport'],
    ['get', '/validacaocertificados/:Solicitante_CertificadoID'],
];

test('composition root preserves the exact route and authorization middleware table', () => {
    const { app, platformRowAuthorization } = createTestApp();
    const routeLayers = app._router.stack.filter((layer) => layer.route);

    assert.deepEqual(
        routeLayers.map(({ route }) => [
            Object.keys(route.methods).filter((method) => route.methods[method]),
            route.path,
        ]),
        EXPECTED_ROUTES.map(([method, path]) => [[method], path]),
    );

    const routeByPath = new Map(routeLayers.map((layer) => [layer.route.path, layer.route]));
    const protectedMiddlewareIndexes = new Map([
        ['/plataforma_v2/CadastroFoto_e_FaceID', 1],
        ['/plataforma_v2/FaceID', 0],
        ['/plataforma_v2/refresh', 0],
        ['/plataforma_v2/updates', 0],
        ['/plataforma_v2/processa-feedback', 0],
    ]);

    for (const [path, authorizationIndex] of protectedMiddlewareIndexes) {
        const route = routeByPath.get(path);
        assert.equal(route.stack[authorizationIndex].handle, platformRowAuthorization.authorize);
        assert.equal(route.stack.length, path === '/plataforma_v2/CadastroFoto_e_FaceID' ? 3 : 2);
    }

    assert.equal(
        routeLayers.flatMap(({ route }) => route.stack)
            .filter(({ handle }) => handle === platformRowAuthorization.authorize)
            .length,
        5,
    );

    const registrationRoute = routeByPath.get('/plataforma_v2/CadastroFoto_e_FaceID');
    assert.equal(registrationRoute.stack[0].name, 'multerMiddleware');

    for (const [method, path] of EXPECTED_ROUTES) {
        if (protectedMiddlewareIndexes.has(path)) continue;
        assert.equal(routeByPath.get(path).stack.length, 1, `${method.toUpperCase()} ${path}`);
    }
});
