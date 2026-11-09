'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { spawn } = require('node:child_process');

const WATCHED_FILES = new Set([
    'support-app.js', 'support-sandbox-server.js', '.env', '.env.support-sandbox',
    'domains/support-management.js', 'integrations/support-users.js',
    'integrations/pagarme.js', 'integrations/support-card-store.js',
    'integrations/support-session-store.js',
    'domains/support-bookings.js', 'integrations/support-appointments.js',
    'integrations/support-booking-store.js', 'integrations/support-booking-notifications.js',
]);

function isSupportDevelopmentFile(filename) {
    return WATCHED_FILES.has(String(filename).replace(/\\/gu, '/'));
}

function startSupportDevelopment({
    root = path.resolve(__dirname, '..'),
    spawnChild = () => spawn(process.execPath, ['support-sandbox-server.js'], {
        cwd: root, stdio: ['inherit', 'inherit', 'inherit', 'ipc'], windowsHide: true,
    }),
    watchDirectory = (directory, listener) => fs.watch(directory, listener),
    log = console.log,
    report = console.error,
    debounceMs = 150,
    stopTimeoutMs = 30_000,
} = {}) {
    let current;
    let stopped = false;
    let timer;
    let restarting;
    let stopping;
    const watchers = [];

    function launch() {
        const child = spawnChild();
        const state = { child, exited: false };
        state.completion = new Promise((resolve) => {
            function finish(result) {
                if (state.exited) return;
                state.exited = true;
                resolve(result);
                if (!stopped && !state.stopRequested) log('Support server stopped. Save a watched backend file or local configuration to retry.');
            }
            child.once('error', () => finish({ error: true }));
            child.once('exit', (code, signal) => finish({ code, signal }));
        });
        current = state;
    }

    function stopChild() {
        const state = current;
        if (!state || state.exited) return Promise.resolve();
        if (state.stopping) return state.stopping;
        state.stopRequested = true;
        state.stopping = new Promise((resolve, reject) => {
            const deadline = setTimeout(() => reject(new Error('Support server did not stop. It was left running; inspect it before restarting.')), stopTimeoutMs);
            state.completion.then((result) => {
                clearTimeout(deadline);
                if (result.error || result.code !== 0) reject(new Error('Support server shutdown failed. Check its local lock before restarting.'));
                else resolve();
            });
            // Windows signals can terminate a child before its lock is released; use IPC.
            if (state.child.connected) state.child.send({ type: 'support:stop' }, () => {});
            // Ctrl+C may reach both Windows processes; a disconnected child can already be stopping.
        });
        return state.stopping;
    }

    function restart() {
        if (stopped) return Promise.resolve();
        if (restarting) return restarting;
        log('Restarting support backend...');
        restarting = stopChild().then(() => { if (!stopped) launch(); })
            .finally(() => { restarting = undefined; });
        return restarting;
    }

    function changed(directory, filename) {
        if (!filename || stopped) return;
        const relative = path.relative(root, path.join(directory, String(filename)));
        if (!isSupportDevelopmentFile(relative)) return;
        clearTimeout(timer);
        timer = setTimeout(() => { void restart().catch((error) => report(error.message)); }, debounceMs);
    }

    function stop() {
        if (stopping) return stopping;
        stopped = true;
        clearTimeout(timer);
        watchers.forEach((watcher) => watcher.close());
        stopping = Promise.resolve(restarting).catch(() => {}).then(stopChild);
        return stopping;
    }

    try {
        for (const directory of [root, path.join(root, 'domains'), path.join(root, 'integrations')]) {
            const watcher = watchDirectory(directory, (event, filename) => changed(directory, filename));
            watcher.on('error', () => report('Support file watcher failed. Restart the development command.'));
            watchers.push(watcher);
        }
        launch();
        log('Support development: frontend files update on browser refresh; backend changes restart automatically.');
        log('Press Ctrl+C to stop this local session. No build, push, PR, or deployment runs here.');
        return { restart, stop };
    } catch (error) {
        watchers.forEach((watcher) => watcher.close());
        throw error;
    }
}

if (require.main === module) {
    try {
        const development = startSupportDevelopment();
        const shutdown = () => {
            void development.stop().catch((error) => {
                console.error(error.message);
                process.exitCode = 1;
            });
        };
        process.once('SIGINT', shutdown);
        process.once('SIGTERM', shutdown);
    } catch {
        console.error('Support development could not start. Check the local checkout and file watcher.');
        process.exitCode = 1;
    }
}

module.exports = { startSupportDevelopment, isSupportDevelopmentFile };
