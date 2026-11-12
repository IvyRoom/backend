'use strict';

const API_ORIGIN = 'https://api.pagar.me/core/v5';
const PAGE_SIZE = 100;
const MAX_PAGES = 100;
const DEFINITE_REJECTIONS = new Set([400, 401, 403, 404, 412, 422]);

function createPagarmeAdapter({ secretKey, publicKey, environment = 'sandbox', fetchImpl = fetch, now = Date.now }) {
    const secretPattern = environment === 'production' ? /^sk_(?!test_)[A-Za-z0-9_-]+$/u : /^sk_test_[A-Za-z0-9_-]+$/u;
    const publicPattern = environment === 'production' ? /^pk_(?!test_)[A-Za-z0-9_-]+$/u : /^pk_test_[A-Za-z0-9_-]+$/u;
    if (!['sandbox', 'production'].includes(environment) || !secretPattern.test(secretKey || '') ||
        !publicPattern.test(publicKey || '') ||
        typeof fetchImpl !== 'function' || typeof now !== 'function') {
        throw failure('invalid_configuration');
    }

    const authorization = `Basic ${Buffer.from(`${secretKey}:`).toString('base64')}`;

    async function request(path, { method = 'GET', body, cardVerification = false, cardLookup = false } = {}) {
        let response;
        try {
            response = await fetchImpl(`${API_ORIGIN}${path}`, {
                method,
                redirect: 'error',
                headers: {
                    Authorization: authorization,
                    Accept: 'application/json',
                    ...(body ? { 'Content-Type': 'application/json' } : {}),
                },
                ...(body ? { body: JSON.stringify(body) } : {}),
                signal: AbortSignal.timeout(15000),
            });
        } catch {
            // Provider errors can contain request data; only fixed errors cross this boundary.
            throw failure('request_failed');
        }
        if (!response?.ok) {
            const definiteRejection = DEFINITE_REJECTIONS.has(response?.status);
            const code = cardLookup && response?.status === 404 ? 'card_not_found'
                : cardVerification && response?.status === 412 ? 'card_verification_failed'
                : definiteRejection ? 'request_rejected' : 'request_failed';
            const error = failure(code);
            if (definiteRejection) error.definiteRejection = true;
            throw error;
        }
        try {
            const result = await response.json();
            requireEnvironment(result, environment);
            return result;
        } catch {
            throw failure('request_failed');
        }
    }

    async function list(path, query = '') {
        const values = [];
        for (let page = 1; page <= MAX_PAGES; page += 1) {
            const result = await request(`${path}?${query}size=${PAGE_SIZE}&page=${page}`);
            if (!Array.isArray(result.data)) throw failure('invalid_response');
            result.data.forEach((value) => requireEnvironment(value, environment));
            values.push(...result.data);
            const hasNext = Boolean(result.paging?.next) || result.data.length === PAGE_SIZE;
            if (!hasNext) return values;
        }
        throw failure('invalid_response');
    }

    return {
        async findCompanyCustomer(code) {
            requireCode(code);
            const customers = await list('/customers', `code=${encodeURIComponent(code)}&`);
            const matches = customers.filter(customer => customer.code === code);
            if (matches.length > 1) throw failure('ambiguous_customer');
            return matches.length ? projectCustomer(matches[0], code, environment) : null;
        },

        async createCompanyCustomer({ code, name }) {
            requireCode(code);
            if (typeof name !== 'string' || !name.trim() || name.length > 64 || /[\u0000-\u001f]/u.test(name)) {
                throw failure('invalid_input');
            }
            return projectCustomer(await request('/customers', {
                method: 'POST',
                body: { code, name },
            }), code, environment);
        },

        async listCards(customerId) {
            requireId(customerId, 'cus');
            const cards = await list(`/customers/${customerId}/cards`);
            return cards.map((value) => projectCard(value, environment));
        },

        async getCard(customerId, cardId) {
            requireId(customerId, 'cus');
            requireId(cardId, 'card');
            const card = await request(`/customers/${customerId}/cards/${cardId}`, { cardLookup: true });
            if (card.id !== cardId) throw failure('invalid_response');
            return projectCard(card, environment);
        },

        async deleteCard(customerId, cardId) {
            requireId(customerId, 'cus');
            requireId(cardId, 'card');
            try {
                const card = await request(`/customers/${customerId}/cards/${cardId}`, { method: 'DELETE', cardLookup: true });
                if (card.id !== cardId || card.status !== 'deleted') throw failure('invalid_response');
                return { deleted: true };
            } catch (error) {
                if (error.code === 'pagarme.card_not_found') return { deleted: true };
                throw error;
            }
        },

        async createCard(customerId, { token, billingAddress }) {
            requireId(customerId, 'cus');
            requireId(token, 'token');
            const card = projectCard(await request(`/customers/${customerId}/cards`, {
                method: 'POST',
                cardVerification: true,
                body: {
                    token,
                    ...(billingAddress === undefined ? {} : { billing_address: projectAddress(billingAddress) }),
                    // The v5 SDK documents authorization before saving with options.verify_card.
                    options: { verify_card: true },
                },
            }), environment);
            const current = new Date(now());
            const year = card.expYear < 100 ? 2000 + card.expYear : card.expYear;
            if (card.status !== 'active' || !Number.isFinite(current.getTime()) ||
                year < current.getUTCFullYear() ||
                (year === current.getUTCFullYear() && card.expMonth < current.getUTCMonth() + 1)) {
                throw failure('invalid_response');
            }
            // Only a successful explicit verification request can create this evidence; GET cannot.
            return { ...card, verificationRequested: true };
        },
    };
}

function requireId(value, prefix) {
    if (typeof value !== 'string' || !new RegExp(`^${prefix}_[A-Za-z0-9]+$`, 'u').test(value)) {
        throw failure('invalid_input');
    }
}

function requireCode(code) {
    if (typeof code !== 'string' || !/^[A-Za-z0-9_-]{1,52}$/u.test(code)) throw failure('invalid_input');
}

function requireEnvironment(value, environment) {
    if (!value || typeof value !== 'object' || Array.isArray(value) ||
        ['live', 'livemode'].some((key) => Object.hasOwn(value, key) && value[key] !== (environment === 'production'))) {
        throw failure('invalid_response');
    }
}

function projectCustomer(value, code, environment) {
    requireEnvironment(value, environment);
    requireId(value.id, 'cus');
    if (value.code !== code) throw failure('invalid_response');
    return { id: value.id, code: value.code };
}

function projectCard(value, environment) {
    requireEnvironment(value, environment);
    requireId(value.id, 'card');
    if (typeof value.brand !== 'string' || !/^[A-Za-z0-9 -]{1,32}$/u.test(value.brand) ||
        !/^\d{4}$/u.test(value.last_four_digits || '') ||
        !Number.isInteger(value.exp_month) || value.exp_month < 1 || value.exp_month > 12 ||
        !Number.isInteger(value.exp_year) || value.exp_year < 0 || value.exp_year > 9999 ||
        (value.holder_name !== undefined && (typeof value.holder_name !== 'string'
            || value.holder_name.length > 64 || /[\u0000-\u001f\u007f]/u.test(value.holder_name))) ||
        (value.status !== undefined && !['active', 'deleted', 'expired'].includes(value.status)) ||
        (value.created_at !== undefined && (typeof value.created_at !== 'string' ||
            !Number.isFinite(Date.parse(value.created_at))))) {
        throw failure('invalid_response');
    }
    return {
        id: value.id,
        brand: value.brand,
        lastFour: value.last_four_digits,
        expMonth: value.exp_month,
        expYear: value.exp_year,
        holderName: value.holder_name,
        createdAt: value.created_at,
        status: value.status,
    };
}

function projectAddress(value) {
    if (!value || typeof value !== 'object' || Array.isArray(value)) throw failure('invalid_input');
    const address = {};
    for (const field of ['line_1', 'line_2', 'zip_code', 'city', 'state', 'country']) {
        if (value[field] !== undefined) {
            if (typeof value[field] !== 'string') throw failure('invalid_input');
            address[field] = value[field];
        }
    }
    return address;
}

function failure(code) {
    const error = new Error('Pagar.me operation failed');
    error.code = `pagarme.${code}`;
    return error;
}

module.exports = { createPagarmeAdapter };
