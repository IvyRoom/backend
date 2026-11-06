'use strict';

const path = require('node:path');
const express = require('express');
const { createSupportManagementHandlers } = require('./domains/support-management');

function createSupportApp({ frontendDirectory, origin, ...dependencies }) {
    const url = new URL(origin);
    if (url.hostname !== '127.0.0.1' || url.protocol !== 'http:') throw new Error('Sandbox requires a loopback origin');
    const app = express();
    app.disable('x-powered-by');
    app.use((req, res, next) => {
        res.set({
            'Cache-Control': 'no-store',
            'X-Content-Type-Options': 'nosniff',
            'Referrer-Policy': 'no-referrer',
            'Content-Security-Policy': "default-src 'self'; connect-src 'self' https://api.pagar.me; img-src 'self'; style-src 'self'; script-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'",
        });
        if (req.get('host') !== url.host || (req.get('origin') && req.get('origin') !== url.origin)) return res.status(403).end();
        return next();
    });
    app.use(express.json({ limit: '8kb' }));
    const handlers = createSupportManagementHandlers(dependencies);
    const base = '/support-management/api';
    app.get(`${base}/config`, handlers.config);
    app.post(`${base}/login`, handlers.login);
    app.get(`${base}/session`, handlers.authorize, handlers.session);
    app.post(`${base}/logout`, handlers.authorize, handlers.logout);
    app.get(`${base}/card`, handlers.authorize, handlers.getCard);
    app.put(`${base}/card`, handlers.authorize, handlers.saveCard);
    app.post(`${base}/card/sync`, handlers.authorize, handlers.syncCard);
    app.get('/', (req, res) => res.redirect('/support-management/login/'));
    if (frontendDirectory) app.use('/support-management', express.static(path.resolve(frontendDirectory), { index: 'index.html' }));
    app.use((error, req, res, next) => {
        res.status(error.type === 'entity.too.large' ? 413 : 400).json({ error: 'support_management.invalid_request' });
    });
    return app;
}

module.exports = { createSupportApp };
