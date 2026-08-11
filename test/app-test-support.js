'use strict';

const { once } = require('node:events');
const {
    PLATFORM_ROW_AUTHORIZATION_DURATION_SECONDS,
    createPlatformRowAuthorizationHandle,
    readPlatformRowIndex,
} = require('../platform-row-authorization');
const { createApp } = require('../app');

const TEST_NOW_MS = 1_800_000_000_000;
const TEST_SIGNING_KEY = Buffer.from(Array.from({ length: 32 }, (_, index) => index));
const TEST_WRONG_SIGNING_KEY = Buffer.from(Array.from({ length: 32 }, (_, index) => index + 1));
const TEST_UUID = '00000000-0000-4000-8000-000000000000';

function createTestPlatformRowAuthorization({
    nowMs = TEST_NOW_MS,
    signingKey = TEST_SIGNING_KEY,
} = {}) {
    return {
        authorize(req, res, next) {
            try {
                res.locals.platformRowIndex = readPlatformRowIndex(
                    req.body && req.body.IndexVerificado,
                    signingKey,
                    nowMs,
                );
                return next();
            } catch {
                return res.status(401).json({});
            }
        },
        createHandle(rowIndex) {
            return createPlatformRowAuthorizationHandle(rowIndex, signingKey, nowMs);
        },
        nowMs,
        signingKey,
    };
}

function changeCharacter(value, index = 0) {
    const replacement = value[index] === 'A' ? 'B' : 'A';
    return `${value.slice(0, index)}${replacement}${value.slice(index + 1)}`;
}

function createInvalidPlatformRowHandles({
    rowIndex = 42,
    nowMs = TEST_NOW_MS,
} = {}) {
    const validHandle = createPlatformRowAuthorizationHandle(rowIndex, TEST_SIGNING_KEY, nowMs);
    const [payloadSegment, signatureSegment] = validHandle.split('.');
    const expiredIssuedAt = nowMs - PLATFORM_ROW_AUTHORIZATION_DURATION_SECONDS * 1000;

    return {
        missing: undefined,
        numeric: rowIndex,
        malformed: 'not-a-platform-row-handle',
        forged: `${payloadSegment}.${changeCharacter(signatureSegment)}`,
        expired: createPlatformRowAuthorizationHandle(
            rowIndex,
            TEST_SIGNING_KEY,
            expiredIssuedAt,
        ),
    };
}

function outcomeQueue(label, ledger) {
    const queues = new Map();

    function queueKey(method, path) {
        return `${String(method).toUpperCase()} ${path}`;
    }

    function enqueue(method, path, ...outcomes) {
        const key = queueKey(method, path);
        const queue = queues.get(key) || [];
        queue.push(...outcomes);
        queues.set(key, queue);
        return api;
    }

    async function dispatch(method, path, body, details = {}) {
        const normalizedMethod = String(method).toUpperCase();
        const call = {
            type: 'external-call',
            client: label,
            method: normalizedMethod,
            path,
            body,
            payload: body,
            ...details,
        };
        api.calls.push(call);
        ledger.push(call);

        const key = queueKey(normalizedMethod, path);
        const queue = queues.get(key);
        if (!queue || queue.length === 0) {
            throw new Error(`Unexpected ${label} call: ${key}`);
        }

        const outcome = queue.shift();
        if (outcome instanceof Error) throw outcome;
        if (typeof outcome === 'function') return outcome(call);
        return outcome;
    }

    function pendingOutcomes() {
        return Array.from(queues.entries())
            .filter(([, outcomes]) => outcomes.length > 0)
            .map(([key, outcomes]) => ({ key, count: outcomes.length }));
    }

    function assertExhausted() {
        const pending = pendingOutcomes();
        if (pending.length > 0) {
            throw new Error(`Unconsumed ${label} outcomes: ${JSON.stringify(pending)}`);
        }
    }

    const api = {
        calls: [],
        enqueue,
        queue: enqueue,
        pendingOutcomes,
        assertExhausted,
        dispatch,
    };

    return api;
}

function createRecordingGraphClient({ ledger = [] } = {}) {
    const recording = outcomeQueue('graph', ledger);

    return Object.assign(recording, {
        api(path) {
            return {
                get: () => recording.dispatch('GET', path),
                post: (body) => recording.dispatch('POST', path, body),
                put: (body) => recording.dispatch('PUT', path, body),
                update: (body) => recording.dispatch('UPDATE', path, body),
            };
        },
    });
}

function createRecordingFaceClient({ ledger = [] } = {}) {
    const recording = outcomeQueue('face', ledger);

    return Object.assign(recording, {
        path(...pathArguments) {
            const [path, ...parameters] = pathArguments;
            const details = { pathArguments: [...pathArguments], parameters };
            return {
                get: () => recording.dispatch('GET', path, undefined, details),
                post: (body) => recording.dispatch('POST', path, body, details),
            };
        },
    });
}

function createManualScheduler({ ledger = [] } = {}) {
    let nextTimerId = 1;
    const tasks = new Map();

    function schedule(callback, delay) {
        const id = nextTimerId;
        nextTimerId += 1;
        tasks.set(id, { id, callback, delay });
        ledger.push({ type: 'timer-scheduled', id, delay });
        return id;
    }

    function cancel(id) {
        const existed = tasks.delete(id);
        ledger.push({ type: 'timer-cancelled', id, existed });
    }

    async function run(id) {
        const task = tasks.get(id);
        if (!task) throw new Error(`Unknown or cancelled timer: ${id}`);
        tasks.delete(id);
        ledger.push({ type: 'timer-run', id, delay: task.delay });
        return task.callback();
    }

    async function runNext() {
        const [nextTask] = tasks.values();
        if (!nextTask) throw new Error('No pending timer');
        return run(nextTask.id);
    }

    return {
        schedule,
        cancel,
        run,
        runNext,
        get pending() {
            return Array.from(tasks.values(), ({ id, delay }) => ({ id, delay }));
        },
        ledger,
    };
}

function createDeferred() {
    let resolve;
    let reject;
    const promise = new Promise((resolvePromise, rejectPromise) => {
        resolve = resolvePromise;
        reject = rejectPromise;
    });
    return { promise, resolve, reject };
}

function createValueSequence(values) {
    const sequence = [...values];
    let index = 0;
    return (...args) => {
        if (index >= sequence.length) throw new Error('Deterministic value sequence exhausted');
        const value = sequence[index];
        index += 1;
        return typeof value === 'function' ? value(...args) : value;
    };
}

function deterministicRandomInt(min, max) {
    return max === undefined ? 0 : min;
}

function createTestApp(overrides = {}) {
    const ledger = overrides.ledger || [];
    const dependencyOverrides = overrides.dependencies || {};
    const graphClient = dependencyOverrides.graphClient
        || overrides.graphClient
        || createRecordingGraphClient({ ledger });
    const faceClient = dependencyOverrides.faceClient
        || overrides.faceClient
        || createRecordingFaceClient({ ledger });
    const platformRowAuthorization = dependencyOverrides.platformRowAuthorization
        || overrides.platformRowAuthorization
        || createTestPlatformRowAuthorization();
    const scheduler = overrides.scheduler || createManualScheduler({ ledger });
    const sleep = dependencyOverrides.sleep || overrides.sleep || (async (delay) => {
        ledger.push({ type: 'sleep', delay });
    });

    const dependencies = {
        graphClient,
        faceClient,
        platformRowAuthorization,
        now: () => new Date(TEST_NOW_MS),
        randomInt: deterministicRandomInt,
        uuid: () => TEST_UUID,
        sleep,
        schedule: scheduler.schedule,
        ...dependencyOverrides,
    };

    for (const name of ['now', 'randomInt', 'uuid', 'schedule']) {
        if (overrides[name] !== undefined) dependencies[name] = overrides[name];
    }

    const app = createApp(dependencies);

    return {
        app,
        dependencies,
        ledger,
        graphClient,
        faceClient,
        platformRowAuthorization,
        scheduler,
    };
}

async function startLoopback(app, testContext) {
    const server = app.listen(0, '127.0.0.1');
    await once(server, 'listening');
    let closed = false;

    async function close() {
        if (closed) return;
        closed = true;
        await new Promise((resolve, reject) => {
            server.close((error) => {
                if (error) reject(error);
                else resolve();
            });
            if (typeof server.closeAllConnections === 'function') server.closeAllConnections();
        });
    }

    if (testContext && typeof testContext.after === 'function') testContext.after(close);

    const address = server.address();
    return {
        server,
        origin: `http://127.0.0.1:${address.port}`,
        close,
    };
}

module.exports = {
    TEST_NOW_MS,
    TEST_SIGNING_KEY,
    TEST_WRONG_SIGNING_KEY,
    createTestPlatformRowAuthorization,
    createInvalidPlatformRowHandles,
    createRecordingGraphClient,
    createRecordingFaceClient,
    createManualScheduler,
    createDeferred,
    createValueSequence,
    createTestApp,
    startLoopback,
};
