'use strict';

const assert = require('node:assert/strict');
const { spawnSync } = require('node:child_process');
const net = require('node:net');
const test = require('node:test');

test('repository test command denies every non-loopback TCP connection', () => {
    assert.throws(
        () => net.connect({ host: 'production-network.invalid', port: 443 }),
        (error) => error && error.code === 'ERR_TEST_NETWORK_DENIED',
    );
});

test('network denial preload propagates to Node children while preserving loopback', () => {
    const script = `
        'use strict';
        const assert = require('node:assert/strict');
        const net = require('node:net');

        assert.throws(
            () => net.connect({ host: 'production-network.invalid', port: 443 }),
            (error) => error && error.code === 'ERR_TEST_NETWORK_DENIED',
        );

        const server = net.createServer((socket) => socket.end());
        server.listen(0, '127.0.0.1', () => {
            const socket = net.connect({
                host: '127.0.0.1',
                port: server.address().port,
            });
            socket.setTimeout(2_000, () => socket.destroy(new Error('loopback connection timed out')));
            socket.once('error', (error) => {
                server.close();
                throw error;
            });
            socket.once('connect', () => {
                socket.end();
                server.close(() => process.stdout.write('child-network-guard-ok'));
            });
        });
    `;
    const child = spawnSync(process.execPath, ['-e', script], {
        cwd: __dirname,
        encoding: 'utf8',
        env: { ...process.env },
        timeout: 5_000,
        windowsHide: true,
    });

    assert.equal(child.error, undefined, child.error && child.error.message);
    assert.equal(child.signal, null);
    assert.equal(child.status, 0, child.stderr);
    assert.equal(child.stderr, '');
    assert.equal(child.stdout, 'child-network-guard-ok');
});
