'use strict';

const { ELIGIBILITY_REVALIDATION_MS } = require('./constants');
const {
    authorityUnavailable,
    forbiddenAuthority,
    invalidAuthority,
} = require('./errors');
const { createCredentialFingerprint } = require('./cryptography');

const ACCESS_TIME_ZONE = 'America/Sao_Paulo';
const EXCEL_EPOCH_UTC_MS = Date.UTC(1899, 11, 31);
const EXCEL_FAKE_LEAP_DAY_SERIAL = 60;
const MAX_EXCEL_ACCESS_DATE_SERIAL = 2_958_464;
const CIVIL_DATE_SEARCH_WINDOW_MS = 36 * 60 * 60 * 1000;

const civilDateFormatter = new Intl.DateTimeFormat('en-CA-u-ca-iso8601-nu-latn', {
    day: '2-digit',
    month: '2-digit',
    timeZone: ACCESS_TIME_ZONE,
    year: 'numeric',
});

function defaultExtractCells(row) {
    if (Array.isArray(row)) return row;
    if (!row || typeof row !== 'object' || !Array.isArray(row.values)) {
        throw new TypeError('Malformed workbook row');
    }
    if (row.values.length === 1 && Array.isArray(row.values[0])) return row.values[0];
    return row.values;
}

function exactCredentialCellValue(value) {
    if (
        value === null
        || value === undefined
        || !['string', 'number', 'boolean', 'bigint'].includes(typeof value)
    ) {
        throw new TypeError('Malformed workbook credential cell');
    }

    return value.toString();
}

function readRows(rows, extractCells) {
    if (!Array.isArray(rows) || typeof extractCells !== 'function') {
        throw authorityUnavailable('malformed-account-rows');
    }

    try {
        return rows.map((row, rowIndex) => {
            const cells = extractCells(row);
            if (
                !Array.isArray(cells)
                || cells.length < 8
                || typeof cells[2] !== 'string'
                || cells[2].length === 0
            ) {
                throw new TypeError('Malformed workbook row');
            }

            return {
                cells,
                exactCredential: exactCredentialCellValue(cells[3]),
                rowIndex,
            };
        });
    } catch (_error) {
        throw authorityUnavailable('malformed-account-row');
    }
}

function requireFingerprintKey(credentialFingerprintKey) {
    if (!credentialFingerprintKey) {
        throw new TypeError('A credential fingerprint key is required');
    }

    return credentialFingerprintKey;
}

function projectAccount(row, credentialFingerprintKey, includeFreshLoginPolicy) {
    const { cells, exactCredential, rowIndex } = row;
    const account = {
        accountStatus: cells[7],
        accessDateSerial: cells[6],
        credentialFingerprint: createCredentialFingerprint(
            exactCredential,
            credentialFingerprintKey,
        ),
        rowIndex,
    };
    Object.defineProperty(account, 'exactLogin', {
        enumerable: false,
        value: cells[2],
    });

    if (includeFreshLoginPolicy) {
        account.faceAuthRequired = readFaceRequirement(cells[4]);
        account.photoRegistrationStatus = cells[5];
    }

    return Object.freeze(account);
}

function readCredentialAccount(rows, {
    login,
    password,
    extractCells = defaultExtractCells,
    credentialFingerprintKey,
} = {}) {
    if (
        typeof login !== 'string'
        || login.length === 0
        || typeof password !== 'string'
        || password.length === 0
    ) {
        throw invalidAuthority('invalid-credentials');
    }

    const readableRows = readRows(rows, extractCells);
    const matchingLoginRows = readableRows.filter(({ cells }) => cells[2] === login);
    if (matchingLoginRows.length === 0) throw invalidAuthority('invalid-credentials');
    if (matchingLoginRows.length > 1) {
        throw authorityUnavailable('duplicate-credential-match');
    }

    const [matchingRow] = matchingLoginRows;
    if (matchingRow.exactCredential !== password) {
        throw invalidAuthority('invalid-credentials');
    }

    return projectAccount(
        matchingRow,
        requireFingerprintKey(credentialFingerprintKey),
        true,
    );
}

function readMappedAccount(rows, exactLogin, {
    extractCells = defaultExtractCells,
    credentialFingerprintKey,
} = {}) {
    if (typeof exactLogin !== 'string' || exactLogin.length === 0) {
        throw authorityUnavailable('malformed-account-mapping');
    }

    const readableRows = readRows(rows, extractCells);
    const matchingRows = readableRows.filter(({ cells }) => cells[2] === exactLogin);
    if (matchingRows.length === 0) throw authorityUnavailable('missing-account-mapping');
    if (matchingRows.length > 1) throw authorityUnavailable('duplicate-account-mapping');

    return projectAccount(
        matchingRows[0],
        requireFingerprintKey(credentialFingerprintKey),
        false,
    );
}

function readFaceRequirement(value) {
    if (value === 'Ativo') return true;
    if (value === 'Inativo') return false;
    throw authorityUnavailable('invalid-face-policy');
}

function excelSerialToCivilDate(value) {
    if (typeof value !== 'number' || !Number.isFinite(value)) {
        throw new TypeError('The Excel access date is invalid');
    }

    const serial = Math.floor(value);
    if (
        serial < 1
        || serial > MAX_EXCEL_ACCESS_DATE_SERIAL
        || serial === EXCEL_FAKE_LEAP_DAY_SERIAL
    ) {
        throw new TypeError('The Excel access date is invalid');
    }

    const realDaysAfterEpoch = serial < EXCEL_FAKE_LEAP_DAY_SERIAL ? serial : serial - 1;
    const date = new Date(EXCEL_EPOCH_UTC_MS + realDaysAfterEpoch * 24 * 60 * 60 * 1000);

    return Object.freeze({
        day: date.getUTCDate(),
        month: date.getUTCMonth() + 1,
        year: date.getUTCFullYear(),
    });
}

function addCivilDay({ year, month, day }) {
    const nextDate = new Date(Date.UTC(year, month - 1, day + 1));
    return {
        day: nextDate.getUTCDate(),
        month: nextDate.getUTCMonth() + 1,
        year: nextDate.getUTCFullYear(),
    };
}

function civilDateAt(timestampMs) {
    const parts = {};
    for (const part of civilDateFormatter.formatToParts(new Date(timestampMs))) {
        if (part.type === 'year' || part.type === 'month' || part.type === 'day') {
            parts[part.type] = Number(part.value);
        }
    }

    return parts;
}

function compareCivilDates(left, right) {
    if (left.year !== right.year) return left.year - right.year;
    if (left.month !== right.month) return left.month - right.month;
    return left.day - right.day;
}

function civilDayStartUtc(civilDate) {
    const nominalUtc = Date.UTC(civilDate.year, civilDate.month - 1, civilDate.day);
    let low = nominalUtc - CIVIL_DATE_SEARCH_WINDOW_MS;
    let high = nominalUtc + CIVIL_DATE_SEARCH_WINDOW_MS;

    while (low < high) {
        const midpoint = low + Math.floor((high - low) / 2);
        if (compareCivilDates(civilDateAt(midpoint), civilDate) < 0) {
            low = midpoint + 1;
        } else {
            high = midpoint;
        }
    }

    if (compareCivilDates(civilDateAt(low), civilDate) !== 0) {
        throw new TypeError('The Excel access date is invalid');
    }

    return new Date(low);
}

function excelSerialToEntitlementExpiry(value) {
    return civilDayStartUtc(addCivilDay(excelSerialToCivilDate(value)));
}

function readNow(now) {
    const value = now instanceof Date ? now.getTime() : now;
    if (typeof value !== 'number' || !Number.isFinite(value)) {
        throw new TypeError('The eligibility observation time is invalid');
    }

    const observedAt = new Date(value);
    if (!Number.isFinite(observedAt.getTime())) {
        throw new TypeError('The eligibility observation time is invalid');
    }

    return observedAt;
}

function normalizeEligibility(account, now) {
    if (!account || typeof account !== 'object') {
        throw authorityUnavailable('malformed-account-authority');
    }
    if (account.accountStatus !== 'Ativo') {
        throw forbiddenAuthority('account-inactive');
    }

    let entitlementExpiresAt;
    try {
        entitlementExpiresAt = excelSerialToEntitlementExpiry(account.accessDateSerial);
    } catch (_error) {
        throw forbiddenAuthority('invalid-entitlement');
    }

    const eligibilityObservedAt = readNow(now);
    if (eligibilityObservedAt.getTime() >= entitlementExpiresAt.getTime()) {
        const error = forbiddenAuthority('entitlement-expired');
        error.entitlementExpiresAt = entitlementExpiresAt;
        throw error;
    }

    const eligibilityRevalidateAt = new Date(Math.min(
        eligibilityObservedAt.getTime() + ELIGIBILITY_REVALIDATION_MS,
        entitlementExpiresAt.getTime(),
    ));

    return Object.freeze({
        eligibilityObservedAt,
        eligibilityRevalidateAt,
        eligibilityState: 'eligible',
        entitlementExpiresAt,
    });
}

module.exports = {
    ACCESS_TIME_ZONE,
    excelSerialToCivilDate,
    excelSerialToEntitlementExpiry,
    normalizeEligibility,
    readCredentialAccount,
    readFaceRequirement,
    readMappedAccount,
};
