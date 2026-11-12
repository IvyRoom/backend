'use strict';

const path = require('node:path');
const express = require('express');
const { createSupportManagementHandlers } = require('./domains/support-management');
const { createSupportBookingHandlers } = require('./domains/support-bookings');

function createSupportApp({ frontendDirectory, origin, preload = false, environment = 'sandbox', apiOnly = false, ...dependencies }) {
    const url = new URL(origin);
    if (environment === 'sandbox') {
        if (url.hostname !== '127.0.0.1' || url.protocol !== 'http:' || apiOnly) throw new Error('Sandbox requires a loopback origin');
    } else if (environment !== 'production' || url.origin !== 'https://machadogestao.com' || !apiOnly) {
        throw new Error('Production requires the canonical Machado origin');
    }
    const app = express();
    app.disable('x-powered-by');
    if (environment === 'production') app.set('trust proxy', 1);
    app.use((req, res, next) => {
        res.set({
            'Cache-Control': 'no-store',
            'X-Content-Type-Options': 'nosniff',
            'Referrer-Policy': 'no-referrer',
            'Content-Security-Policy': "default-src 'self'; connect-src 'self' https://api.pagar.me; img-src 'self'; style-src 'self'; script-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'",
        });
        if ((environment === 'sandbox' && req.get('host') !== url.host)
            || (req.get('origin') && req.get('origin') !== url.origin)) return res.status(403).end();
        if (environment === 'production') {
            res.vary('Origin');
            res.set('Access-Control-Allow-Origin', url.origin);
            if (req.method === 'OPTIONS') {
                res.set({ 'Access-Control-Allow-Methods': 'GET,POST,PUT,PATCH,DELETE,OPTIONS',
                    'Access-Control-Allow-Headers': 'Authorization,Content-Type' });
                return res.status(204).end();
            }
        }
        return next();
    });
    app.use(express.json({ limit: '8kb' }));
    let bookingHandlers;
    const handlers = createSupportManagementHandlers({ ...dependencies, environment,
        ...(preload ? { bootstrap: async (user, directory) => ({ config: { environment, publicKey: dependencies.publicKey },
            ...await bookingHandlers.getCalendarContext(user, directory) }) } : {}) });
    bookingHandlers = createSupportBookingHandlers({ ...dependencies, companyCards: handlers.companyPayment, claimCompany: handlers.claimCompany });
    const base = apiOnly ? '' : '/atendimentos/api';
    app.get(`${base}/config`, handlers.config);
    app.post(`${base}/login`, handlers.login);
    app.get(`${base}/session`, handlers.authorize, handlers.session);
    app.post(`${base}/logout`, handlers.logout);
    app.get(`${base}/card`, handlers.authorize, handlers.getCard);
    app.get(`${base}/permissions`, handlers.authorize, handlers.permissions);
    app.patch(`${base}/permissions`, handlers.authorize, handlers.updatePermission);
    app.put(`${base}/card`, handlers.authorize, handlers.saveCard);
    app.delete(`${base}/card`, handlers.authorize, handlers.deleteCard);
    app.post(`${base}/card/sync`, handlers.authorize, handlers.syncCard);
    app.get(`${base}/calendar`, handlers.authorize, bookingHandlers.calendar);
    app.post(`${base}/reservations`, handlers.authorize, bookingHandlers.reserve);
    if (!apiOnly) {
        app.get('/', (req, res) => res.redirect('/atendimentos/login/'));
        if (frontendDirectory) {
            const directory = path.resolve(frontendDirectory);
            app.use('/atendimentos/agendamento', express.static(path.join(directory, 'scheduler'), { index: 'index.html' }));
            app.use('/atendimentos/login', express.static(path.join(directory, 'login'), { index: 'index.html' }));
            app.use('/atendimentos/modules', express.static(path.join(directory, 'modules')));
            app.use('/atendimentos/img', express.static(path.join(directory, 'img')));
            app.get('/atendimentos/style.css', (req, res) => res.sendFile(path.join(directory, 'style.css')));
            app.get('/shared/backend-origin.js', (req, res) => res.sendFile(path.resolve(directory, '../shared/backend-origin.js')));
        }
    }
    app.use((error, req, res, next) => {
        res.status(error.type === 'entity.too.large' ? 413 : 400).json({ error: 'support_management.invalid_request' });
    });
    return app;
}

module.exports = { createSupportApp };
