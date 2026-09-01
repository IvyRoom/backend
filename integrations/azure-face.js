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
    };
}

module.exports = { createAzureFaceAdapter };
