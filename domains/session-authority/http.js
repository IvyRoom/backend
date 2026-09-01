'use strict';

const {
    PROTECTED_LEARNING_PATHS,
    SESSION_API_HOSTNAME,
    SESSION_COOKIE_NAME,
    SESSION_FRONTEND_ORIGIN,
    SESSION_REQUEST_HEADER,
    SESSION_REQUEST_HEADER_VALUE,
    TARGET_ROLE_PATHS,
    TARGET_SESSION_ROUTES,
} = require('./constants');
const { parseOpaqueIdentifier } = require('./cryptography');

const SESSION_COOKIE_STATES = Object.freeze({
    missing: 'missing',
    malformed: 'malformed',
    present: 'present',
});

const TARGET_REQUEST_KINDS = Object.freeze({
    passThrough: 'pass-through',
    target: 'target',
    blockedTarget: 'blocked-target',
});

const SESSION_REQUEST_HEADER_HTTP_NAME = SESSION_REQUEST_HEADER
    .split('-')
    .map((part) => `${part.slice(0, 1).toUpperCase()}${part.slice(1)}`)
    .join('-');
const ALLOWED_REQUEST_HEADERS = `Content-Type, ${SESSION_REQUEST_HEADER_HTTP_NAME}`;
const ACTUAL_VARY = 'Origin, Cookie';
const PREFLIGHT_VARY = 'Origin, Access-Control-Request-Method, Access-Control-Request-Headers';
const VALIDATOR_HEADERS = new Set(['etag', 'last-modified']);
const RESPONSE_HARDENED = Symbol('session-authority-response-hardened');

const DEFAULT_PROTECTED_ROUTES = Object.freeze(PROTECTED_LEARNING_PATHS.map((path) => Object.freeze({
    method: 'POST',
    path,
})));

function missingCookie() {
    return { state: SESSION_COOKIE_STATES.missing };
}

function malformedCookie() {
    return { state: SESSION_COOKIE_STATES.malformed };
}

function isCanonicalIdentifier(value) {
    try {
        parseOpaqueIdentifier(value);
        return true;
    } catch {
        return false;
    }
}

function parseSessionCookieHeader(cookieHeader) {
    if (cookieHeader === undefined || cookieHeader === null || cookieHeader === '') {
        return missingCookie();
    }
    if (typeof cookieHeader !== 'string') {
        return malformedCookie();
    }

    let identifier;
    let matches = 0;

    for (const rawSegment of cookieHeader.split(';')) {
        const segment = rawSegment.trim();
        const separatorIndex = segment.indexOf('=');
        const name = separatorIndex === -1 ? segment : segment.slice(0, separatorIndex);
        if (name !== SESSION_COOKIE_NAME) {
            if (name.trim() === SESSION_COOKIE_NAME) return malformedCookie();
            continue;
        }

        matches += 1;
        if (matches > 1 || separatorIndex === -1) return malformedCookie();
        identifier = segment.slice(separatorIndex + 1);
    }

    if (matches === 0) return missingCookie();
    if (!isCanonicalIdentifier(identifier)) return malformedCookie();
    return { state: SESSION_COOKIE_STATES.present, value: identifier };
}

function readHeader(req, name) {
    if (!req || !req.headers || typeof req.headers !== 'object') return undefined;
    return req.headers[String(name).toLowerCase()];
}

function readSessionCookie(req) {
    return parseSessionCookieHeader(readHeader(req, 'cookie'));
}

function requireDate(value, name) {
    const date = value instanceof Date ? new Date(value.getTime()) : new Date(value);
    if (!Number.isFinite(date.getTime())) throw new TypeError(`${name} must be a valid date`);
    return date;
}

function formatSessionIssuanceCookie({ identifier, expiresAt, now = new Date() } = {}) {
    if (!isCanonicalIdentifier(identifier)) {
        throw new TypeError('Session identifier must use the canonical transport encoding');
    }

    const authoritativeExpiry = requireDate(expiresAt, 'expiresAt');
    const issuedAt = requireDate(now, 'now');
    const remainingMilliseconds = authoritativeExpiry.getTime() - issuedAt.getTime();
    if (remainingMilliseconds <= 0) {
        throw new RangeError('An issuance cookie requires a future authoritative expiry');
    }

    const maxAgeSeconds = Math.ceil(remainingMilliseconds / 1000);
    return [
        `${SESSION_COOKIE_NAME}=${identifier}`,
        'Path=/',
        'Secure',
        'HttpOnly',
        'SameSite=None',
        'Partitioned',
        `Max-Age=${maxAgeSeconds}`,
        `Expires=${authoritativeExpiry.toUTCString()}`,
    ].join('; ');
}

function isValidatorHeader(name) {
    return VALIDATOR_HEADERS.has(String(name).toLowerCase());
}

function withoutValidatorHeaders(headers) {
    if (Array.isArray(headers)) {
        const filtered = [];
        for (let index = 0; index < headers.length; index += 2) {
            if (!isValidatorHeader(headers[index])) {
                filtered.push(headers[index], headers[index + 1]);
            }
        }
        return filtered;
    }
    if (!headers || typeof headers !== 'object') return headers;
    return Object.fromEntries(Object.entries(headers).filter(([name]) => !isValidatorHeader(name)));
}

function applySessionResponseHardening(res) {
    if (!res || typeof res.setHeader !== 'function' || typeof res.removeHeader !== 'function') {
        throw new TypeError('A Node-compatible response is required');
    }

    if (!res[RESPONSE_HARDENED]) {
        const setHeader = res.setHeader;
        res.setHeader = function setSessionHeader(name, value) {
            if (isValidatorHeader(name)) {
                this.removeHeader(name);
                return this;
            }
            return setHeader.call(this, name, value);
        };

        if (typeof res.writeHead === 'function') {
            const writeHead = res.writeHead;
            res.writeHead = function writeSessionHead(statusCode, statusMessage, headers) {
                this.removeHeader('ETag');
                this.removeHeader('Last-Modified');

                if (typeof statusMessage === 'string') {
                    return writeHead.call(
                        this,
                        statusCode,
                        statusMessage,
                        withoutValidatorHeaders(headers),
                    );
                }

                return writeHead.call(this, statusCode, withoutValidatorHeaders(statusMessage));
            };
        }

        Object.defineProperty(res, RESPONSE_HARDENED, { value: true });
    }

    res.removeHeader('ETag');
    res.removeHeader('Last-Modified');
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
    res.setHeader('Referrer-Policy', 'no-referrer');
    return res;
}

function normalizeMethod(method) {
    return typeof method === 'string' ? method.trim().toUpperCase() : '';
}

function normalizePath(path) {
    if (typeof path !== 'string' || path === '') return '';
    const queryIndex = path.indexOf('?');
    const pathname = queryIndex === -1 ? path : path.slice(0, queryIndex);
    return pathname.length > 1 ? pathname.replace(/\/+$/u, '').toLowerCase() : pathname;
}

function requestPath(req) {
    if (req && typeof req.path === 'string') return normalizePath(req.path);
    if (req && typeof req.url === 'string') return normalizePath(req.url);
    return '';
}

function routeIndex(routes, routeType) {
    const indexed = new Map();
    for (const route of routes) {
        if (!route || typeof route !== 'object') {
            throw new TypeError(`${routeType} routes must contain method/path objects`);
        }
        const method = normalizeMethod(route.method);
        const path = normalizePath(route.path);
        if (!method || !path.startsWith('/')) {
            throw new TypeError(`${routeType} routes must contain valid methods and paths`);
        }

        const entry = indexed.get(path) || { methods: [], routeType };
        if (!entry.methods.includes(method)) entry.methods.push(method);
        indexed.set(path, entry);
    }
    return indexed;
}

function parseRequestedHeaders(value) {
    if (value === undefined) return [];
    if (typeof value !== 'string') return null;
    return value.split(',').map((name) => name.trim().toLowerCase()).filter(Boolean);
}

function hasTargetPreflightHeader(req) {
    const requestedHeaders = parseRequestedHeaders(readHeader(req, 'access-control-request-headers'));
    return Array.isArray(requestedHeaders) && requestedHeaders.includes(SESSION_REQUEST_HEADER);
}

function passThroughClassification() {
    return {
        kind: TARGET_REQUEST_KINDS.passThrough,
        isTarget: false,
        blocked: false,
    };
}

function targetClassification({
    entry,
    preflight,
    effectiveMethod,
    cookie,
}) {
    return {
        kind: TARGET_REQUEST_KINDS.target,
        isTarget: true,
        blocked: false,
        preflight,
        methodAllowed: entry.methods.includes(effectiveMethod),
        allowedMethods: [...entry.methods],
        routeType: entry.routeType,
        cookieState: cookie.state,
    };
}

function blockedTargetClassification({ preflight, cookie }) {
    return {
        kind: TARGET_REQUEST_KINDS.blockedTarget,
        isTarget: false,
        blocked: true,
        preflight,
        reason: 'target-routes-disabled',
        cookieState: cookie.state,
    };
}

function createTargetRequestClassifier({
    targetRoutesEnabled = false,
    protectedRoutesEnabled = false,
    protectedRoutes = DEFAULT_PROTECTED_ROUTES,
} = {}) {
    if (typeof targetRoutesEnabled !== 'boolean') {
        throw new TypeError('targetRoutesEnabled must be a boolean');
    }
    if (typeof protectedRoutesEnabled !== 'boolean') {
        throw new TypeError('protectedRoutesEnabled must be a boolean');
    }

    const targetRouteIndex = routeIndex(TARGET_SESSION_ROUTES, 'session');
    const dualModeRouteIndex = routeIndex([
        { method: 'POST', path: TARGET_ROLE_PATHS.login },
        { method: 'POST', path: TARGET_ROLE_PATHS.registration },
        { method: 'POST', path: TARGET_ROLE_PATHS.faceChallenge },
    ], 'dual-mode');
    const protectedRouteIndex = routeIndex(protectedRoutes, 'protected');

    return function classifyTargetRequest(req) {
        const path = requestPath(req);
        const requestMethod = normalizeMethod(req && req.method);
        const preflight = requestMethod === 'OPTIONS';
        const effectiveMethod = preflight
            ? normalizeMethod(readHeader(req, 'access-control-request-method'))
            : requestMethod;
        const cookie = readSessionCookie(req);
        const exactTargetHeader = readHeader(req, SESSION_REQUEST_HEADER)
            === SESSION_REQUEST_HEADER_VALUE;
        const targetPreflightHeader = preflight && hasTargetPreflightHeader(req);
        const hasTargetCookie = cookie.state !== SESSION_COOKIE_STATES.missing;
        const dualModeTargetSignal = exactTargetHeader || targetPreflightHeader || hasTargetCookie;

        const sessionEntry = targetRouteIndex.get(path);
        if (sessionEntry && targetRoutesEnabled) {
            return targetClassification({ entry: sessionEntry, preflight, effectiveMethod, cookie });
        }

        const dualModeEntry = dualModeRouteIndex.get(path);
        if (dualModeEntry && dualModeTargetSignal) {
            if (!targetRoutesEnabled) return blockedTargetClassification({ preflight, cookie });
            return targetClassification({
                entry: dualModeEntry,
                preflight,
                effectiveMethod,
                cookie,
            });
        }

        const protectedEntry = protectedRouteIndex.get(path);
        if (protectedEntry && protectedRoutesEnabled && dualModeTargetSignal) {
            return targetClassification({
                entry: protectedEntry,
                preflight,
                effectiveMethod,
                cookie,
            });
        }

        return passThroughClassification();
    };
}

function setTargetCorsHeaders(res, classification, targetOrigin) {
    res.setHeader('Access-Control-Allow-Origin', targetOrigin);
    res.setHeader('Access-Control-Allow-Credentials', 'true');
    res.setHeader('Vary', classification.preflight ? PREFLIGHT_VARY : ACTUAL_VARY);

    if (classification.preflight) {
        res.setHeader('Access-Control-Allow-Methods', classification.allowedMethods.join(', '));
        res.setHeader('Access-Control-Allow-Headers', ALLOWED_REQUEST_HEADERS);
    }
}

function respondEmptyJson(res, statusCode) {
    return res.status(statusCode).json({});
}

function createTargetHttpBoundary({
    classifyRequest,
    targetRoutesEnabled = false,
    protectedRoutesEnabled = false,
    protectedRoutes = DEFAULT_PROTECTED_ROUTES,
    targetHostname = SESSION_API_HOSTNAME,
    targetOrigin = SESSION_FRONTEND_ORIGIN,
} = {}) {
    if (typeof targetHostname !== 'string' || targetHostname === '') {
        throw new TypeError('targetHostname must be a non-empty string');
    }
    if (targetHostname !== SESSION_API_HOSTNAME && process.env.NODE_ENV !== 'test') {
        throw new Error('A synthetic target hostname may be injected only in tests');
    }
    if (typeof targetOrigin !== 'string' || targetOrigin === '') {
        throw new TypeError('targetOrigin must be a non-empty string');
    }
    if (targetOrigin !== SESSION_FRONTEND_ORIGIN && process.env.NODE_ENV !== 'test') {
        throw new Error('A synthetic target origin may be injected only in tests');
    }

    const classify = classifyRequest || createTargetRequestClassifier({
        targetRoutesEnabled,
        protectedRoutesEnabled,
        protectedRoutes,
    });

    if (typeof classify !== 'function') throw new TypeError('classifyRequest must be a function');

    return function targetHttpBoundary(req, res, next) {
        const classification = classify(req);
        if (classification.blocked) {
            applySessionResponseHardening(res);
            return respondEmptyJson(res, 503);
        }
        if (!classification.isTarget) return next();

        applySessionResponseHardening(res);
        setTargetCorsHeaders(res, classification, targetOrigin);

        if (!res.locals) res.locals = {};
        res.locals.sessionAuthorityTransport = classification;

        if (readHeader(req, 'host') !== targetHostname) {
            return respondEmptyJson(res, 403);
        }

        if (classification.preflight) {
            const origin = readHeader(req, 'origin');
            const requestedHeaders = parseRequestedHeaders(
                readHeader(req, 'access-control-request-headers'),
            );
            const headersAllowed = Array.isArray(requestedHeaders)
                && requestedHeaders.includes(SESSION_REQUEST_HEADER)
                && requestedHeaders.every((name) => (
                    name === 'content-type' || name === SESSION_REQUEST_HEADER
                ));

            if (
                origin !== targetOrigin
                || !classification.methodAllowed
                || !headersAllowed
            ) {
                return respondEmptyJson(res, 403);
            }

            return res.status(204).end();
        }

        if (!classification.methodAllowed) return respondEmptyJson(res, 403);

        const method = normalizeMethod(req && req.method);
        const unsafe = !['GET', 'HEAD', 'OPTIONS'].includes(method);
        if (unsafe && readHeader(req, 'origin') !== targetOrigin) {
            return respondEmptyJson(res, 403);
        }
        if (
            classification.cookieState !== SESSION_COOKIE_STATES.missing
            && readHeader(req, SESSION_REQUEST_HEADER) !== SESSION_REQUEST_HEADER_VALUE
        ) {
            return respondEmptyJson(res, 403);
        }

        return next();
    };
}

module.exports = {
    ACTUAL_VARY,
    ALLOWED_REQUEST_HEADERS,
    DEFAULT_PROTECTED_ROUTES,
    PREFLIGHT_VARY,
    SESSION_COOKIE_NAME,
    SESSION_COOKIE_STATES,
    SESSION_REQUEST_HEADER,
    SESSION_REQUEST_HEADER_HTTP_NAME,
    SESSION_REQUEST_HEADER_VALUE,
    TARGET_ORIGIN: SESSION_FRONTEND_ORIGIN,
    TARGET_REQUEST_KINDS,
    applySessionResponseHardening,
    createTargetHttpBoundary,
    createTargetRequestClassifier,
    formatSessionIssuanceCookie,
    isCanonicalIdentifier,
    parseSessionCookieHeader,
    readSessionCookie,
};
