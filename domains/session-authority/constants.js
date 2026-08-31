'use strict';

const SESSION_COOKIE_NAME = '__Host-machado-session';
const SESSION_REQUEST_HEADER = 'x-machado-session-request';
const SESSION_REQUEST_HEADER_VALUE = '1';
const SESSION_FRONTEND_ORIGIN = 'https://machadogestao.com';
const SESSION_API_HOSTNAME = 'api.machadogestao.com';

const PROVISIONAL_LIFETIME_MS = 20 * 60 * 1000;
const AUTHENTICATED_LIFETIME_MS = 4 * 60 * 60 * 1000;
const ELIGIBILITY_REVALIDATION_MS = 5 * 60 * 1000;
const LEGACY_LIFETIME_MS = 4 * 60 * 60 * 1000;
const LEGACY_SUNSET_MAXIMUM_MS = 7 * 24 * 60 * 60 * 1000;
const LEGACY_SEEDING_HEARTBEAT_INTERVAL_MS = 30 * 1000;
const LEGACY_SEEDING_LEASE_MS = 2 * 60 * 1000;

const SESSION_PHASES = Object.freeze({
    credentialVerified: 'credential-verified',
    registrationPending: 'registration-pending',
    facePending: 'face-pending',
    authenticated: 'authenticated',
    expired: 'expired',
    revoked: 'revoked',
    rotatedOut: 'rotated-out',
});

const ACTIVE_SESSION_PHASES = Object.freeze([
    SESSION_PHASES.credentialVerified,
    SESSION_PHASES.registrationPending,
    SESSION_PHASES.facePending,
    SESSION_PHASES.authenticated,
]);

const NEXT_OPERATION_ROLES = Object.freeze({
    registrationEnrollment: 'registration-enrollment',
    registrationChallenge: 'registration-challenge',
    faceChallenge: 'face-challenge',
    faceCompletion: 'face-completion',
    protectedLearning: 'protected-learning',
    revokeAll: 'revoke-all',
});

const TARGET_SESSION_ROUTES = Object.freeze([
    Object.freeze({ method: 'POST', path: '/plataforma_v2/sessions/current/registration-enrollment' }),
    Object.freeze({ method: 'POST', path: '/plataforma_v2/sessions/current/face-completion' }),
    Object.freeze({ method: 'GET', path: '/plataforma_v2/sessions/current' }),
    Object.freeze({ method: 'DELETE', path: '/plataforma_v2/sessions/current' }),
    Object.freeze({ method: 'DELETE', path: '/plataforma_v2/sessions' }),
]);

const TARGET_ROLE_PATHS = Object.freeze({
    login: '/plataforma_v2/login-FaceID',
    registration: '/plataforma_v2/CadastroFoto_e_FaceID',
    faceChallenge: '/plataforma_v2/FaceID',
});

const PROTECTED_LEARNING_PATHS = Object.freeze([
    '/plataforma_v2/refresh',
    '/plataforma_v2/updates',
    '/plataforma_v2/processa-feedback',
]);

module.exports = {
    ACTIVE_SESSION_PHASES,
    AUTHENTICATED_LIFETIME_MS,
    ELIGIBILITY_REVALIDATION_MS,
    LEGACY_LIFETIME_MS,
    LEGACY_SEEDING_HEARTBEAT_INTERVAL_MS,
    LEGACY_SEEDING_LEASE_MS,
    LEGACY_SUNSET_MAXIMUM_MS,
    NEXT_OPERATION_ROLES,
    PROTECTED_LEARNING_PATHS,
    PROVISIONAL_LIFETIME_MS,
    SESSION_API_HOSTNAME,
    SESSION_COOKIE_NAME,
    SESSION_FRONTEND_ORIGIN,
    SESSION_PHASES,
    SESSION_REQUEST_HEADER,
    SESSION_REQUEST_HEADER_VALUE,
    TARGET_ROLE_PATHS,
    TARGET_SESSION_ROUTES,
};
