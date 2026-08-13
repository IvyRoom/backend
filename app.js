'use strict';

const crypto = require('node:crypto');
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const { v4: uuidv4 } = require('uuid');
const { createQuoteRequestHandler } = require('./domains/quote-requests');
const { createConectaRecommendationHandler } = require('./domains/conecta-recommendations');
const { createClientOnboardingHandlers } = require('./domains/client-onboarding');
const { createLearningPlatformHandlers } = require('./domains/learning-platform');
const { createDrmHandler } = require('./domains/drm');
const { createCertificateValidationHandler } = require('./domains/certificate-validation');
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
} = {}) {
    const {
        authorize: authorizePlatformRow,
        createHandle: createPlatformRowAuthorizationHandle,
    } = platformRowAuthorization;

    const app = express();
    app.use(cors());
    app.use(express.json());
    app.use('/img', express.static('img'));

    const retry = createRetry({ sleep });
    const processQuoteRequest = createQuoteRequestHandler({ graphClient, retry });
    const processConectaRecommendation = createConectaRecommendationHandler({
        graphClient,
        retry,
        now,
    });
    const {
        processClientIntake,
        releasePlatformAccess,
    } = createClientOnboardingHandlers({
        graphClient,
        retry,
        now,
        randomInt,
        sleep,
        schedule,
    });
    const {
        loginWithFaceId,
        registerPhotoAndFaceId,
        createFaceIdSession,
        getFaceIdResult,
        refresh,
        updateProgress,
        processFeedback,
        getStatusReport,
    } = createLearningPlatformHandlers({
        graphClient,
        faceClient,
        retry,
        uuid,
        createPlatformRowAuthorizationHandle,
    });
    const getPlayReadyAuthorizationUrl = createDrmHandler();
    const validateCertificate = createCertificateValidationHandler({ graphClient, retry });

    app.post('/landingpage/solicitacaoorcamento', processQuoteRequest);
    app.post('/conecta/processa-recomendacao', processConectaRecommendation);
    app.post('/clientes/processa-formulario', processClientIntake);
    app.post('/clientes/liberacao-acesso-plataforma', releasePlatformAccess);
    app.post('/plataforma_v2/login-FaceID', loginWithFaceId);
    app.post('/plataforma_v2/CadastroFoto_e_FaceID', multer().single('file'), authorizePlatformRow, registerPhotoAndFaceId);
    app.post('/plataforma_v2/FaceID', authorizePlatformRow, createFaceIdSession);
    app.get('/plataforma_v2/FaceID_resultado/:Azure_Face_API_LivenessSession_sessionID', getFaceIdResult);
    app.post('/plataforma_v2/refresh', authorizePlatformRow, refresh);
    app.post('/plataforma_v2/updates', authorizePlatformRow, updateProgress);
    app.post('/plataforma_v2/processa-feedback', authorizePlatformRow, processFeedback);
    app.get('/ezdrm-playready-authorization-url', getPlayReadyAuthorizationUrl);
    app.post('/plataforma_v2/statusreport', getStatusReport);
    app.get('/validacaocertificados/:Solicitante_CertificadoID', validateCertificate);

    return app;
}

module.exports = { createApp };
