'use strict';

const crypto = require('node:crypto');

const SESSION_DURATION = 4 * 60 * 60 * 1000;
const normalize = (value) => String(value || '').trim().replace(/\s+/gu, ' ').toLowerCase();
const digest = (value) => crypto.createHash('sha256').update(value).digest('hex');

function companyCode(company) {
    return `support_${digest(normalize(company)).slice(0, 44)}`;
}

function publicUser(user) {
    return { name: user.name, company: user.company, role: user.role };
}

function publicCard(card) {
    return { brand: card.brand, lastFour: card.lastFour, expMonth: card.expMonth, expYear: card.expYear };
}

function validBillingAddress(address) {
    if (!address || typeof address !== 'object' || Array.isArray(address)) return false;
    const allowed = ['line_1', 'line_2', 'zip_code', 'city', 'state', 'country'];
    if (Object.keys(address).some((key) => !allowed.includes(key))) return false;
    return ['line_1', 'city'].every((key) => typeof address[key] === 'string' && address[key].trim().length > 0 && address[key].length <= 256)
        && (address.line_2 === undefined || (typeof address.line_2 === 'string' && address.line_2.length <= 256))
        && /^\d{8}$/.test(address.zip_code) && /^[A-Z]{2}$/.test(address.state) && address.country === 'BR';
}

function createSupportManagementHandlers({ users, pagarme, cardStore, publicKey, now = Date.now }) {
    const sessions = new Map();
    const attempts = new Map();
    const busyCompanies = new Set();
    const sessionSalt = crypto.randomBytes(32).toString('hex');
    const fingerprint = (user) => digest(`${sessionSalt}:${user.password}`);
    const fail = (res, status, reason) => res.status(status).json({ error: `support_management.${reason}` });

    async function readUser(email) {
        const matches = (await users.readUsers()).filter((user) => normalize(user.email) === normalize(email));
        return matches.length === 1 && matches[0].status === 'Ativo' && normalize(matches[0].company) ? matches[0] : null;
    }

    function sweep() {
        for (const [key, session] of sessions) if (session.expiresAt <= now()) sessions.delete(key);
        for (const [key, attempt] of attempts) if (attempt.until <= now()) attempts.delete(key);
    }

    return {
        config(req, res) {
            return res.json({ environment: 'sandbox', publicKey });
        },
        async login(req, res) {
            sweep();
            const { email, password } = req.body || {};
            if (typeof email !== 'string' || email.length > 254 || typeof password !== 'string' || !password || password.length > 256) return fail(res, 401, 'invalid_credentials');
            // Bound both maps; failed login attempts must not exhaust sandbox memory.
            const key = req.ip;
            const attempt = attempts.get(key) || { count: 0, until: now() + 60_000 };
            if (attempt.count >= 10 || (attempts.size >= 1000 && !attempts.has(key))) return fail(res, 429, 'try_later');
            attempt.count += 1;
            attempts.set(key, attempt);
            try {
                const user = await readUser(email);
                const candidate = Buffer.from(digest(password));
                const expected = Buffer.from(digest(user ? String(user.password) : crypto.randomUUID()));
                if (!user || !crypto.timingSafeEqual(candidate, expected)) return fail(res, 401, 'invalid_credentials');
                if (sessions.size >= 500) return fail(res, 429, 'try_later');
                const token = crypto.randomBytes(32).toString('base64url');
                sessions.set(token, { email: normalize(user.email), company: normalize(user.company), fingerprint: fingerprint(user), expiresAt: now() + SESSION_DURATION });
                return res.json({ token, user: publicUser(user) });
            } catch {
                return fail(res, 503, 'users_unavailable');
            }
        },
        async authorize(req, res, next) {
            sweep();
            const token = (req.get('authorization') || '').match(/^Bearer ([A-Za-z0-9_-]{43})$/u)?.[1];
            const session = sessions.get(token);
            if (!session) return fail(res, 401, 'session_expired');
            try {
                const user = await readUser(session.email);
                if (!user || normalize(user.company) !== session.company || fingerprint(user) !== session.fingerprint) {
                    sessions.delete(token);
                    return fail(res, 401, 'session_expired');
                }
                res.locals.supportUser = user;
                res.locals.supportToken = token;
                return next();
            } catch {
                return fail(res, 503, 'users_unavailable');
            }
        },
        session(req, res) {
            return res.json({ user: publicUser(res.locals.supportUser) });
        },
        logout(req, res) {
            sessions.delete(res.locals.supportToken);
            return res.status(204).end();
        },
        async getCard(req, res) {
            const code = companyCode(res.locals.supportUser.company);
            try {
                const record = await cardStore.read(code);
                if (record.pending) return fail(res, 409, 'card_review_required');
                if (!record.cardId) return res.json({ card: null });
                return res.json({ card: publicCard(await pagarme.getCard(record.customerId, record.cardId)) });
            } catch {
                return fail(res, 503, 'card_unavailable');
            }
        },
        async saveCard(req, res) {
            const user = res.locals.supportUser;
            if (user.role !== 'Administrador') return fail(res, 403, 'administrator_required');
            const { token, billingAddress } = req.body || {};
            if (Object.keys(req.body || {}).some((key) => !['token', 'billingAddress'].includes(key)) || typeof token !== 'string' || !/^token_[A-Za-z0-9]{1,200}$/u.test(token) || !validBillingAddress(billingAddress)) return fail(res, 400, 'invalid_card');
            const code = companyCode(user.company);
            if (busyCompanies.has(code)) return fail(res, 409, 'card_busy');
            busyCompanies.add(code);
            try {
                let record = await cardStore.read(code);
                if (record.pending) return fail(res, 409, 'card_review_required');
                if (!record.customerId) {
                    let customer = await pagarme.findCompanyCustomer(code);
                    if (!customer) {
                        // Persist intent before a non-idempotent call so an ambiguous result cannot be replayed.
                        await cardStore.write(code, { pending: 'customer' });
                        try {
                            customer = await pagarme.createCompanyCustomer({ code, name: user.company.slice(0, 64) });
                        } catch (error) {
                            if (error.definiteRejection) await cardStore.write(code, {});
                            throw error;
                        }
                    }
                    record = { customerId: customer.id };
                    await cardStore.write(code, record);
                }
                await cardStore.write(code, { ...record, pending: 'card' });
                let card;
                try {
                    card = await pagarme.createCard(record.customerId, { token, billingAddress });
                } catch (error) {
                    if (error.definiteRejection) await cardStore.write(code, record);
                    throw error;
                }
                await cardStore.write(code, { customerId: record.customerId, cardId: card.id });
                return res.json({ card: publicCard(card) });
            } catch {
                return fail(res, 503, 'card_save_failed');
            } finally {
                busyCompanies.delete(code);
            }
        },
    };
}

module.exports = { createSupportManagementHandlers, companyCode };
