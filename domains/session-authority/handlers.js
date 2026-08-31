'use strict';

const { SESSION_PHASES } = require('./constants');
const {
    applySessionResponseHardening,
    formatSessionIssuanceCookie,
    readSessionCookie,
} = require('./http');
const {
    authorityConflict,
    invalidAuthority,
    isSessionAuthorityError,
    statusForSessionAuthorityError,
} = require('./errors');

function createSessionAuthorityHandlers({
    authority,
    legacyHandlers,
    legacyAuthorizePlatformRow,
    uploadReferencePhoto,
} = {}) {
    if (!authority || typeof authority !== 'object') throw new TypeError('Session authority is required');
    if (!legacyHandlers || typeof legacyHandlers !== 'object') throw new TypeError('Legacy handlers are required');
    if (typeof legacyHandlers.sendPlatformDataReadFailure !== 'function') {
        throw new TypeError('Legacy platform-data failure handler is required');
    }
    if (typeof legacyAuthorizePlatformRow !== 'function') {
        throw new TypeError('Legacy platform-row authorization is required');
    }
    if (typeof uploadReferencePhoto !== 'function') {
        throw new TypeError('Reference-photo upload middleware is required');
    }

    const targetRegistrationPreauthorization = targetOnly(async (req, res, next) => {
        try {
            const authorized = await authority.authorizeCurrent(requireIdentifier(req), {
                allowedPhases: [
                    SESSION_PHASES.registrationPending,
                    SESSION_PHASES.facePending,
                ],
                revalidate: true,
            });
            if (authorized.session.phase === SESSION_PHASES.facePending) {
                throw authorityConflict('face-challenge-active');
            }
            return next();
        } catch (error) {
            return sendFailure(res, error);
        }
    });

    const parseReferencePhoto = conditionalMiddleware({
        target: uploadReferencePhoto,
        legacy: uploadReferencePhoto,
    });

    const authorizePlatformRequest = conditionalMiddleware({
        target: (_req, _res, next) => next(),
        legacy: authority.runtimeControls.durableStoreRequired
            ? authorizeLegacyPlatformRow
            : legacyAuthorizePlatformRow,
    });

    async function loginWithFaceId(req, res) {
        if (!isTargetRequest(res)) {
            if (!authority.runtimeControls.durableStoreRequired) {
                return legacyHandlers.loginWithFaceId(req, res);
            }
            return runLegacy(res, () => authority.loginLegacyWithSeeding({
                login: req.body && req.body.Usuário_Login,
                password: req.body && req.body.Usuário_Senha,
            }), legacyHandlers.sendPlatformDataReadFailure);
        }

        const cookie = readSessionCookie(req);
        return run(res, () => authority.loginTarget({
            login: req.body && req.body.Usuário_Login,
            password: req.body && req.body.Usuário_Senha,
            presentedIdentifier: cookie.state === 'present' ? cookie.value : null,
        }));
    }

    async function registerPhotoAndFaceId(req, res) {
        if (!isTargetRequest(res)) return legacyHandlers.registerPhotoAndFaceId(req, res);
        return run(res, () => authority.createRegistrationChallenge(
            requireIdentifier(req),
            req.file && req.file.buffer,
        ));
    }

    async function createFaceIdSession(req, res) {
        if (!isTargetRequest(res)) return legacyHandlers.createFaceIdSession(req, res);
        return run(res, () => authority.createExistingPhotoChallenge(requireIdentifier(req)));
    }

    const registrationEnrollment = targetHandler((req) => (
        authority.registrationEnrollment(requireIdentifier(req))
    ));
    const faceCompletion = targetHandler((req) => authority.completeFace(requireIdentifier(req)));
    const getCurrentSession = targetHandler((req) => authority.current(requireIdentifier(req)));
    const logoutCurrentSession = targetHandler((req) => {
        const cookie = readSessionCookie(req);
        return authority.logout(cookie.state === 'present' ? cookie.value : undefined);
    });
    const revokeAllSessions = targetHandler((req) => authority.revokeAll(requireIdentifier(req)));

    function authorizeProtectedLearning(req, res, next) {
        if (!isTargetRequest(res)) {
            if (authority.runtimeControls.durableStoreRequired) {
                return authorizeLegacyPlatformRow(req, res, next);
            }
            return legacyAuthorizePlatformRow(req, res, next);
        }

        authority.authorizeProtected(requireIdentifier(req)).then((authorized) => {
            res.locals.subjectId = authorized.subjectId;
            res.locals.platformRowIndex = authorized.platformRowIndex;
            if (req.body && typeof req.body === 'object') delete req.body.IndexVerificado;
            next();
        }).catch((error) => sendFailure(res, error));
        return undefined;
    }

    function authorizeLegacyPlatformRow(req, res, next) {
        const cookie = readSessionCookie(req);
        if (cookie.state !== 'missing') return sendFailure(res, invalidAuthority('target-cookie-decisive'));
        const rawHandle = req.body && req.body.IndexVerificado;
        authority.authorizeLegacy(rawHandle).then((authorized) => {
            res.locals.subjectId = authorized.subjectId;
            res.locals.platformRowIndex = authorized.platformRowIndex;
            next();
        }).catch((error) => sendFailure(res, error));
        return undefined;
    }

    return {
        authorizePlatformRequest,
        authorizeProtectedLearning,
        createFaceIdSession,
        faceCompletion,
        getCurrentSession,
        loginWithFaceId,
        logoutCurrentSession,
        parseReferencePhoto,
        registrationEnrollment,
        registerPhotoAndFaceId,
        revokeAllSessions,
        targetRegistrationPreauthorization,
    };
}

function conditionalMiddleware({ target, legacy }) {
    return function modeAwareMiddleware(req, res, next) {
        return (isTargetRequest(res) ? target : legacy)(req, res, next);
    };
}

function targetOnly(handler) {
    return function targetOnlyMiddleware(req, res, next) {
        if (!isTargetRequest(res)) return next();
        return handler(req, res, next);
    };
}

function targetHandler(operation) {
    return function targetRouteHandler(req, res) {
        return run(res, () => operation(req));
    };
}

async function run(res, operation) {
    applySessionResponseHardening(res);
    try {
        return sendResult(res, await operation());
    } catch (error) {
        return sendFailure(res, error);
    }
}

async function runLegacy(res, operation, sendPlatformDataReadFailure) {
    try {
        const result = await operation();
        if (!result || !Number.isInteger(result.status) || result.issuance) {
            return sendLegacyFailure(res, new Error('Invalid legacy compatibility result'));
        }
        if (result.status === 204) return res.status(204).end();
        return res.status(result.status).json(result.body === undefined ? {} : result.body);
    } catch (error) {
        if (
            isSessionAuthorityError(error)
            && error.reason === 'legacy-platform-data-read-failed'
        ) {
            return sendPlatformDataReadFailure(res);
        }
        return sendLegacyFailure(res, error);
    }
}

function sendResult(res, result) {
    if (!result || !Number.isInteger(result.status)) {
        return sendFailure(res, new Error('Invalid session-authority result'));
    }
    if (result.issuance) {
        res.setHeader('Set-Cookie', formatSessionIssuanceCookie({
            identifier: result.issuance.identifier,
            expiresAt: result.issuance.expiresAt,
            now: result.issuance.serverTime,
        }));
    }
    if (result.status === 204) return res.status(204).end();
    return res.status(result.status).json(result.body === undefined ? {} : result.body);
}

function sendFailure(res, error) {
    applySessionResponseHardening(res);
    const status = isSessionAuthorityError(error)
        ? statusForSessionAuthorityError(error)
        : 503;
    if (isSessionAuthorityError(error) && error.retryAfter !== undefined) {
        res.setHeader('Retry-After', String(error.retryAfter));
    }
    if (status === 401 && error && error.reason === 'invalid-credentials') {
        return res.status(status).json({ error: 'credenciais_inválidas' });
    }
    return res.status(status).json({});
}

function sendLegacyFailure(res, error) {
    const status = isSessionAuthorityError(error)
        ? statusForSessionAuthorityError(error)
        : 503;
    if (status === 401 && error && error.reason === 'invalid-credentials') {
        return res.status(status).json({ error: 'credenciais_inválidas' });
    }
    return res.status(status).json({});
}

function requireIdentifier(req) {
    const cookie = readSessionCookie(req);
    if (cookie.state !== 'present') throw invalidAuthority('invalid-session');
    return cookie.value;
}

function isTargetRequest(res) {
    return Boolean(
        res
        && res.locals
        && res.locals.sessionAuthorityTransport
        && res.locals.sessionAuthorityTransport.isTarget,
    );
}

module.exports = {
    createSessionAuthorityHandlers,
    isTargetRequest,
};
