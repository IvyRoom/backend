'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createMicrosoftGraphAdapter } = require('../integrations/microsoft-graph');
const { createAzureFaceAdapter } = require('../integrations/azure-face');

test('integration adapter factories do not access raw clients during construction', () => {
    const forbiddenClient = new Proxy({}, {
        get() {
            throw new Error('raw client accessed during adapter construction');
        },
    });

    const microsoftGraph = createMicrosoftGraphAdapter({ graphClient: forbiddenClient });
    const azureFace = createAzureFaceAdapter({ faceClient: forbiddenClient });

    assert.equal(typeof microsoftGraph.readPlatformRows, 'function');
    assert.equal(typeof microsoftGraph.sendMail, 'function');
    assert.equal(typeof azureFace.createLivenessSession, 'function');
    assert.equal(typeof azureFace.readLivenessSessionResult, 'function');
});

test('integration adapters keep response-shape projection separate from SDK attempts', async (t) => {
    await t.test('Microsoft Graph', async () => {
        let valueAccesses = 0;
        const response = Object.defineProperty({}, 'value', {
            get() {
                valueAccesses += 1;
                return [{ values: [['row']] }];
            },
        });
        const microsoftGraph = createMicrosoftGraphAdapter({
            graphClient: {
                api() {
                    return { get: async () => response };
                },
            },
        });

        const rawResponse = await microsoftGraph.readPlatformRows();
        assert.equal(rawResponse, response);
        assert.equal(valueAccesses, 0);
        const rows = microsoftGraph.extractRows(rawResponse);
        assert.equal(valueAccesses, 1);
        assert.deepEqual(microsoftGraph.extractRowCells(rows[0]), ['row']);
    });

    await t.test('Azure Face', async () => {
        let bodyAccesses = 0;
        const response = Object.defineProperty({}, 'body', {
            get() {
                bodyAccesses += 1;
                return { authToken: 'token', sessionId: 'session' };
            },
        });
        const azureFace = createAzureFaceAdapter({
            faceClient: {
                path() {
                    return { post: async () => response };
                },
            },
        });

        const rawResponse = await azureFace.createLivenessSession(Buffer.from('photo'), 'uuid');
        assert.equal(rawResponse, response);
        assert.equal(bodyAccesses, 0);
        assert.deepEqual(azureFace.extractLivenessSession(rawResponse), {
            authToken: 'token',
            sessionId: 'session',
        });
        assert.equal(bodyAccesses, 2);
    });

    await t.test('Azure Face result', async () => {
        let bodyAccesses = 0;
        const response = Object.defineProperty({}, 'body', {
            get() {
                bodyAccesses += 1;
                return {
                    results: {
                        attempts: [{
                            result: {
                                livenessDecision: 'realface',
                                verifyResult: { matchConfidence: 0.99, isIdentical: true },
                            },
                        }],
                    },
                };
            },
        });
        const azureFace = createAzureFaceAdapter({
            faceClient: {
                path() {
                    return { get: async () => response };
                },
            },
        });

        const rawResponse = await azureFace.readLivenessSessionResult('session');
        assert.equal(rawResponse, response);
        assert.equal(bodyAccesses, 0);
        assert.deepEqual(azureFace.extractLivenessSessionResult(rawResponse), {
            livenessDecision: 'realface',
            matchConfidence: 0.99,
            matchDecision: true,
        });
        assert.equal(bodyAccesses, 3);
    });
});
