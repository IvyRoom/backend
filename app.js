'use strict';

const crypto = require('node:crypto');
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const { v4: uuidv4 } = require('uuid');
const { createQuoteRequestHandler } = require('./domains/quote-requests');
const { createConectaRecommendationHandler } = require('./domains/conecta-recommendations');
const { createClientOnboardingHandlers } = require('./domains/client-onboarding');
const { ConverteData, createLearningPlatformHandlers } = require('./domains/learning-platform');
const { createDrmHandler } = require('./domains/drm');
const { createCertificateValidationHandler } = require('./domains/certificate-validation');
const { createSessionAuthority } = require('./domains/session-authority/authority');
const { createSessionAuthorityHandlers } = require('./domains/session-authority/handlers');
const { createTargetHttpBoundary } = require('./domains/session-authority/http');
const { createMicrosoftGraphAdapter } = require('./integrations/microsoft-graph');
const { createAzureFaceAdapter } = require('./integrations/azure-face');
const { createRetry } = require('./shared/retry');

function createApp({
    graphClient,
    faceClient,
    platformRowAuthorization = {},
    now = () => new Date(),
    randomInt = (...args) => crypto.randomInt(...args),
    uuid = uuidv4,
    sleep = (delay) => new Promise((resolve) => setTimeout(resolve, delay)),
    schedule = (callback, delay) => setTimeout(callback, delay),
    sessionAuthority,
} = {}) {
    const {
        authorize: authorizePlatformRow,
        createHandle: createPlatformRowAuthorizationHandle,
        inspectHandle: inspectPlatformRowAuthorizationHandle,
    } = platformRowAuthorization;

    const retry = createRetry({ sleep });
    const microsoftGraph = createMicrosoftGraphAdapter({ graphClient });
    const azureFace = createAzureFaceAdapter({ faceClient });
    const authority = sessionAuthority && (sessionAuthority.authority || createSessionAuthority({
        store: sessionAuthority.store,
        keys: sessionAuthority.keys,
        runtimeControls: sessionAuthority.runtimeControls,
        randomBytes: sessionAuthority.randomBytes,
        createSubjectId: sessionAuthority.createSubjectId,
        createSessionId: sessionAuthority.createSessionId,
        createFlowId: sessionAuthority.createFlowId,
        createCorrelationId: sessionAuthority.createCorrelationId,
        formatLegacyAccessDate: ConverteData,
        legacyHandleAuthority: {
            createHandle: createPlatformRowAuthorizationHandle,
            inspectHandle: inspectPlatformRowAuthorizationHandle,
        },
        accountSource: {
            async readRows() {
                const response = await microsoftGraph.readPlatformRows();
                return microsoftGraph.extractRows(response).map((row) => (
                    microsoftGraph.extractRowCells(row)
                ));
            },
            async readRowsLegacy() {
                const response = await retry(() => microsoftGraph.readPlatformRows());
                return microsoftGraph.extractRows(response).map((row) => (
                    microsoftGraph.extractRowCells(row)
                ));
            },
            uploadReferencePhoto: (rowIndex, photo) => (
                microsoftGraph.uploadReferencePhoto(rowIndex, photo)
            ),
            markPhotoRegistered: (rowIndex) => microsoftGraph.updatePlatformRow(
                rowIndex,
                [null, null, null, null, null, 'Sim', null, null, null, null, null,
                    null, null, null, null, null, null, null, null, null, null, null],
            ),
            downloadReferencePhoto: (rowIndex) => microsoftGraph.downloadReferencePhoto(rowIndex),
        },
        faceSource: {
            async createLivenessSession(referenceImage, correlationId) {
                const response = await azureFace.createLivenessSession(referenceImage, correlationId);
                const result = azureFace.extractLivenessSession(response);
                return {
                    authToken: result.authToken,
                    privateChallengeId: result.sessionId,
                };
            },
            async readLivenessSessionResult(privateChallengeId) {
                const response = await azureFace.readLivenessSessionResult(privateChallengeId);
                return azureFace.extractBoundLivenessSessionResult(response);
            },
        },
    }));
    const sessionContext = sessionAuthority && { ...sessionAuthority, authority };

    const app = express();
    if (sessionContext) {
        app.use(createTargetHttpBoundary({
            ...sessionContext.http,
            targetRoutesEnabled: authority.runtimeControls.targetRoutesEnabled,
            protectedRoutesEnabled: authority.runtimeControls.protectedRoutesEnabled,
        }));
        const legacyCors = cors();
        app.use((req, res, next) => (
            res.locals.sessionAuthorityTransport ? next() : legacyCors(req, res, next)
        ));
    } else {
        app.use(cors());
    }
    app.use(express.json());
    app.use('/img', express.static('img'));

    const processQuoteRequest = createQuoteRequestHandler({ microsoftGraph, retry });
    const processConectaRecommendation = createConectaRecommendationHandler({
        microsoftGraph,
        retry,
        now,
    });
    const {
        processClientIntake,
        releasePlatformAccess,
    } = createClientOnboardingHandlers({
        microsoftGraph,
        retry,
        now,
        randomInt,
        sleep,
        schedule,
    });
    const {
        loginWithFaceId,
        sendPlatformDataReadFailure,
        registerPhotoAndFaceId,
        createFaceIdSession,
        getFaceIdResult,
        refresh,
        updateProgress,
        processFeedback,
        getStatusReport,
    } = createLearningPlatformHandlers({
        microsoftGraph,
        azureFace,
        retry,
        uuid,
        createPlatformRowAuthorizationHandle,
    });
    const getPlayReadyAuthorizationUrl = createDrmHandler();
    const validateCertificate = createCertificateValidationHandler({ microsoftGraph, retry });
    const referencePhotoUpload = multer().single('file');
    const sessionHandlers = sessionContext && createSessionAuthorityHandlers({
        authority,
        legacyHandlers: {
            loginWithFaceId,
            sendPlatformDataReadFailure,
            registerPhotoAndFaceId,
            createFaceIdSession,
        },
        legacyAuthorizePlatformRow: authorizePlatformRow,
        uploadReferencePhoto: referencePhotoUpload,
    });

    app.post('/landingpage/solicitacaoorcamento', processQuoteRequest);
    app.post('/conecta/processa-recomendacao', processConectaRecommendation);
    app.post('/clientes/processa-formulario', processClientIntake);
    app.post('/clientes/liberacao-acesso-plataforma', releasePlatformAccess);
    app.post('/plataforma_v2/login-FaceID', sessionHandlers ? sessionHandlers.loginWithFaceId : loginWithFaceId);
    if (sessionHandlers && authority.runtimeControls.targetRoutesEnabled) {
        app.post(
            '/plataforma_v2/sessions/current/registration-enrollment',
            sessionHandlers.registrationEnrollment,
        );
    }
    if (sessionHandlers) {
        app.post(
            '/plataforma_v2/CadastroFoto_e_FaceID',
            sessionHandlers.targetRegistrationPreauthorization,
            sessionHandlers.parseReferencePhoto,
            sessionHandlers.authorizePlatformRequest,
            sessionHandlers.registerPhotoAndFaceId,
        );
        app.post(
            '/plataforma_v2/FaceID',
            sessionHandlers.authorizePlatformRequest,
            sessionHandlers.createFaceIdSession,
        );
    } else {
        app.post('/plataforma_v2/CadastroFoto_e_FaceID', referencePhotoUpload, authorizePlatformRow, registerPhotoAndFaceId);
        app.post('/plataforma_v2/FaceID', authorizePlatformRow, createFaceIdSession);
    }
    if (sessionHandlers && authority.runtimeControls.targetRoutesEnabled) {
        app.post(
            '/plataforma_v2/sessions/current/face-completion',
            sessionHandlers.faceCompletion,
        );
        app.get('/plataforma_v2/sessions/current', sessionHandlers.getCurrentSession);
        app.delete('/plataforma_v2/sessions/current', sessionHandlers.logoutCurrentSession);
        app.delete('/plataforma_v2/sessions', sessionHandlers.revokeAllSessions);
    }
    app.get('/plataforma_v2/FaceID_resultado/:Azure_Face_API_LivenessSession_sessionID', getFaceIdResult);
    app.post('/plataforma_v2/refresh', sessionHandlers ? sessionHandlers.authorizeProtectedLearning : authorizePlatformRow, refresh);
    app.post('/plataforma_v2/updates', sessionHandlers ? sessionHandlers.authorizeProtectedLearning : authorizePlatformRow, updateProgress);
    app.post('/plataforma_v2/processa-feedback', sessionHandlers ? sessionHandlers.authorizeProtectedLearning : authorizePlatformRow, processFeedback);
    app.get('/ezdrm-playready-authorization-url', getPlayReadyAuthorizationUrl);
    app.post('/plataforma_v2/statusreport', getStatusReport);
    app.get('/validacaocertificados/:Solicitante_CertificadoID', validateCertificate);

    return app;
}

module.exports = { createApp };
