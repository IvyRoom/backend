'use strict';

const net = require('node:net');

const GUARD_INSTALLED = Symbol.for('machado.session-authority.test-network-guard');
const LOOPBACK_HOSTS = new Set([
    'localhost',
    '127.0.0.1',
    '::1',
    '::ffff:127.0.0.1',
]);

propagatePreloadToNodeChildren();

if (!globalThis[GUARD_INSTALLED]) {
    const originalConnect = net.Socket.prototype.connect;

    net.Socket.prototype.connect = function guardedTestConnect(...args) {
        const target = readConnectTarget(args);
        if (target.network && !isLoopbackHost(target.host)) {
            const error = new Error('Automated tests deny non-loopback network connections');
            error.code = 'ERR_TEST_NETWORK_DENIED';
            throw error;
        }
        return originalConnect.apply(this, args);
    };

    Object.defineProperty(globalThis, GUARD_INSTALLED, { value: true });
}

function readConnectTarget(args) {
    let normalized = args;
    if (Array.isArray(args[0])) normalized = args[0];
    const first = normalized[0];

    if (first && typeof first === 'object') {
        if (typeof first.path === 'string') return { network: false };
        return { network: true, host: first.host || first.hostname || 'localhost' };
    }
    if (typeof first === 'number') {
        return {
            network: true,
            host: typeof normalized[1] === 'string' ? normalized[1] : 'localhost',
        };
    }
    if (typeof first === 'string') return { network: false };
    return { network: true, host: '' };
}

function isLoopbackHost(value) {
    if (typeof value !== 'string') return false;
    const canonical = value.toLowerCase().replace(/^\[|\]$/gu, '').replace(/\.$/u, '');
    return LOOPBACK_HOSTS.has(canonical);
}

function propagatePreloadToNodeChildren() {
    const preloadOption = `--require=${JSON.stringify(__filename)}`;
    const existingOptions = process.env.NODE_OPTIONS || '';
    if (existingOptions.includes(preloadOption)) return;

    process.env.NODE_OPTIONS = existingOptions
        ? `${existingOptions} ${preloadOption}`
        : preloadOption;
}
