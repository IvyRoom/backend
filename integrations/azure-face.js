'use strict';

const LIVENESS_SESSIONS_PATH = '/detectLivenessWithVerify-sessions';
const LIVENESS_SESSION_RESULT_PATH = '/detectLivenessWithVerify-sessions/{sessionId}';

function createAzureFaceAdapter({ faceClient }) {
    return {
        createLivenessSession(verifyImage, deviceCorrelationId) {
            return faceClient.path(LIVENESS_SESSIONS_PATH).post({
                contentType: 'multipart/form-data',
                body: [
                    { name: 'VerifyImage', body: verifyImage },
                    { name: 'livenessOperationMode', body: 'Passive' },
                    { name: 'deviceCorrelationId', body: deviceCorrelationId },
                ],
            });
        },

        extractLivenessSession(response) {
            return {
                authToken: response.body.authToken,
                sessionId: response.body.sessionId,
            };
        },

        readLivenessSessionResult(sessionId) {
            return faceClient.path(LIVENESS_SESSION_RESULT_PATH, sessionId).get();
        },

        extractLivenessSessionResult(response) {
            return {
                livenessDecision: response.body.results.attempts[0].result.livenessDecision,
                matchConfidence: response.body.results.attempts[0].result.verifyResult.matchConfidence,
                matchDecision: response.body.results.attempts[0].result.verifyResult.isIdentical,
            };
        },

        extractBoundLivenessSessionResult(response) {
            const body = response && response.body;
            if (!body || typeof body !== 'object') {
                throw new TypeError('Malformed Face liveness result');
            }
            if (body.status === 'NotStarted' || body.status === 'Running') {
                return { providerState: 'pending' };
            }
            if (body.status === 'Failed' || body.status === 'Canceled') {
                throw new TypeError('Unavailable Face liveness result');
            }
            if (body.status !== 'Succeeded') {
                throw new TypeError('Malformed Face liveness status');
            }

            const attempts = body.results && body.results.attempts;
            if (!Array.isArray(attempts) || attempts.length === 0) {
                throw new TypeError('Malformed Face liveness attempts');
            }
            if (attempts.some((attempt) => (
                !attempt
                || !Number.isSafeInteger(attempt.attemptId)
                || attempt.attemptId < 1
            ))) {
                throw new TypeError('Malformed Face liveness attempt identifier');
            }
            if (new Set(attempts.map(({ attemptId }) => attemptId)).size !== attempts.length) {
                throw new TypeError('Malformed Face liveness attempt identifier');
            }

            const attempt = attempts.reduce((latest, candidate) => (
                candidate.attemptId > latest.attemptId
                    ? candidate
                    : latest
            ), attempts[0]);
            if (attempt.attemptStatus === 'Failed' || attempt.attemptStatus === 'Canceled') {
                throw new TypeError('Unavailable Face liveness attempt');
            }
            if (attempt.attemptStatus !== 'Succeeded') {
                throw new TypeError('Malformed Face liveness attempt status');
            }

            const result = attempt.result;
            const verifyResult = result && result.verifyResult;
            if (
                !result
                || !['realface', 'spoofface', 'uncertain'].includes(result.livenessDecision)
                || !verifyResult
                || typeof verifyResult.matchConfidence !== 'number'
                || !Number.isFinite(verifyResult.matchConfidence)
                || verifyResult.matchConfidence < 0
                || verifyResult.matchConfidence > 1
                || typeof verifyResult.isIdentical !== 'boolean'
            ) {
                throw new TypeError('Malformed definitive Face liveness result');
            }
            return {
                livenessDecision: result.livenessDecision,
                matchConfidence: verifyResult.matchConfidence,
                matchDecision: verifyResult.isIdentical,
            };
        },
    };
}

module.exports = { createAzureFaceAdapter };
