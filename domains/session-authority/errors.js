'use strict';

const ERROR_CLASSES = Object.freeze({
    invalid: 'invalid-authority',
    forbidden: 'forbidden-authority',
    conflict: 'authority-conflict',
    unavailable: 'authority-unavailable',
});

class SessionAuthorityError extends Error {
    constructor(errorClass, message, options = {}) {
        super(message, options);
        this.name = 'SessionAuthorityError';
        this.errorClass = errorClass;
        this.reason = options.reason;
        this.retryAfter = options.retryAfter;
    }
}

function invalidAuthority(reason = 'invalid') {
    return new SessionAuthorityError(ERROR_CLASSES.invalid, 'Session authority is invalid', { reason });
}

function forbiddenAuthority(reason = 'forbidden') {
    return new SessionAuthorityError(ERROR_CLASSES.forbidden, 'Session operation is forbidden', { reason });
}

function authorityConflict(reason = 'conflict') {
    return new SessionAuthorityError(ERROR_CLASSES.conflict, 'Session transition conflicted', { reason });
}

function authorityUnavailable(reason = 'unavailable', options = {}) {
    return new SessionAuthorityError(ERROR_CLASSES.unavailable, 'Session authority is unavailable', {
        ...options,
        reason,
    });
}

function isSessionAuthorityError(error) {
    return error instanceof SessionAuthorityError;
}

function statusForSessionAuthorityError(error) {
    switch (error && error.errorClass) {
    case ERROR_CLASSES.invalid: return 401;
    case ERROR_CLASSES.forbidden: return 403;
    case ERROR_CLASSES.conflict: return 409;
    case ERROR_CLASSES.unavailable: return 503;
    default: return 503;
    }
}

module.exports = {
    ERROR_CLASSES,
    SessionAuthorityError,
    authorityConflict,
    authorityUnavailable,
    forbiddenAuthority,
    invalidAuthority,
    isSessionAuthorityError,
    statusForSessionAuthorityError,
};
