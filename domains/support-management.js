'use strict';

const crypto = require('node:crypto');
const { createMemorySupportSessionStore } = require('../integrations/support-session-store');

const SESSION_DURATION = 60 * 60 * 1000;
const CARD_VERIFICATION = 'pagarme_verify_card_v1';
const normalize = (value) => String(value || '').trim().replace(/\s+/gu, ' ').toLowerCase();
const digest = (value) => crypto.createHash('sha256').update(value).digest('hex');

function companyCode(company) {
    return `support_${digest(normalize(company)).slice(0, 44)}`;
}

function publicUser(user) {
    return { name: user.name, company: user.company, role: user.role, email: user.email, bookingAllowed: user.bookingAllowed === true };
}

function publicCard(card) {
    return { brand: card.brand, lastFour: card.lastFour, expMonth: card.expMonth, expYear: card.expYear };
}

function cardRevision(record) {
    return record.cardRevision || digest(`${record.customerId}:${record.cardId}`);
}

function validBillingAddress(address) {
    if (!address || typeof address !== 'object' || Array.isArray(address)) return false;
    const allowed = ['line_1', 'line_2', 'zip_code', 'city', 'state', 'country'];
    if (Object.keys(address).some((key) => !allowed.includes(key))) return false;
    return ['line_1', 'city'].every((key) => typeof address[key] === 'string' && address[key].trim().length > 0 && address[key].length <= 256)
        && (address.line_2 === undefined || (typeof address.line_2 === 'string' && address.line_2.length <= 256))
        && /^\d{8}$/.test(address.zip_code) && /^[A-Z]{2}$/.test(address.state) && address.country === 'BR';
}

function createSupportManagementHandlers({ users, pagarme, cardStore, publicKey, now = Date.now, sessionStore = createMemorySupportSessionStore(), bootstrap }) {
    const attempts = new Map();
    const busyCompanies = new Set();
    const bookingCompanies = new Map();
    const cardReads = new Map();
    const paymentCache = new Map();
    const fingerprint = (user) => digest(`${sessionStore.salt}:${user.password}`);
    const companyIdentity = (user) => normalize(user.companyId || user.company);
    const fail = (res, status, reason) => res.status(status).json({ error: `support_management.${reason}` });

    function companyPermissions(user, directory) {
        return { employees: directory.filter((person) => companyIdentity(person) === companyIdentity(user)).map((person) => ({
            email: normalize(person.email), name: person.name,
            allowed: person.bookingAllowed === true,
            active: /^[^\s@]+@[^\s@]+\.[^\s@]+$/u.test(person.email)
                && directory.filter((entry) => normalize(entry.email) === normalize(person.email)).length === 1,
        })) };
    }

    // Keep wallet mutations outside each booking's payment check and workbook write.
    function claimCompany(user) {
        const code = companyCode(companyIdentity(user));
        if (busyCompanies.has(code) && !bookingCompanies.has(code)) return null;
        busyCompanies.add(code);
        bookingCompanies.set(code, (bookingCompanies.get(code) || 0) + 1);
        let released = false;
        return () => {
            if (released) return;
            released = true;
            const remaining = bookingCompanies.get(code) - 1;
            if (remaining) bookingCompanies.set(code, remaining);
            else {
                bookingCompanies.delete(code);
                busyCompanies.delete(code);
            }
        };
    }

    function readUser(email, directory) {
        const matches = directory.filter((user) => normalize(user.email) === normalize(email));
        return matches.length === 1 && normalize(matches[0].company) ? matches[0] : null;
    }

    async function sweep() {
        await sessionStore.prune(now());
        for (const [key, attempt] of attempts) if (attempt.until <= now()) attempts.delete(key);
    }

    function usableCard(card) {
        const year = card?.expYear < 100 ? 2000 + card.expYear : card?.expYear;
        return card?.status === 'active'
            && Number.isInteger(card.expMonth) && card.expMonth >= 1 && card.expMonth <= 12
            && Number.isInteger(year) && year >= 2000 && year <= 9999
            && Date.UTC(year, card.expMonth, 1) > now();
    }

    const hasVerification = (record) => record.verification === CARD_VERIFICATION;

    async function readCompanyCard(user) {
        const code = companyCode(companyIdentity(user));
        if (cardReads.has(code)) return cardReads.get(code);
        const reading = readCompanyCardRecord(user, code);
        cardReads.set(code, reading);
        try { return await reading; }
        finally { cardReads.delete(code); }
    }

    async function readCompanyCardRecord(user, code) {
        let record = await cardStore.read(code);
        if (!Object.keys(record).length && user.companyId && users.readCompanyCardReferences) {
            const references = await users.readCompanyCardReferences(user.companyId);
            if (references) {
                const previous = cardStore.findByReferences ? await cardStore.findByReferences(references) : null;
                record = previous ? { ...previous } : { customerId: references.customerId, cardId: references.cardId };
                // Preserve local verification evidence across the CNPJ migration; workbook IDs alone cannot create it.
                delete record.workbookSynced;
                delete record.workbookReferencesSynced;
                await cardStore.write(code, record);
            }
        }
        return { code, record };
    }

    async function synchronizeWorkbook(code, record, company, card) {
        if (!hasVerification(record) || !usableCard(card)) throw new Error('Card verification required');
        await users.writeCompanyCard(company, card, {
            customerId: record.customerId, cardId: record.cardId, environment: 'sandbox',
        });
        await cardStore.write(code, { customerId: record.customerId, cardId: record.cardId, cardRevision: cardRevision(record),
            verification: CARD_VERIFICATION, workbookSynced: true, workbookReferencesSynced: true });
        return { card: publicCard(card), cardRevision: cardRevision(record) };
    }

    async function companyPayment(user, { cached = false } = {}) {
        const { code, record } = await readCompanyCard(user);
        if (['delete', 'delete_workbook'].includes(record.pending)) return { card: null, cardDeletionPending: true, cardRevision: cardRevision(record) };
        if (record.pending && record.pending !== 'workbook') throw Object.assign(new Error('Card review required'), { code: 'card_review_required' });
        if (!record.cardId) return { card: null };
        if (!hasVerification(record)) return { card: null, verificationRequired: true, cardRevision: cardRevision(record) };
        const previous = paymentCache.get(code);
        const key = `${record.customerId}:${record.cardId}:${cardRevision(record)}`;
        if (cached && previous?.key === key && previous.until > now()) return { ...previous.result };
        const card = await pagarme.getCard(record.customerId, record.cardId);
        const result = usableCard(card) ? { card: publicCard(card), cardRevision: cardRevision(record),
            ...(record.workbookSynced !== true || record.workbookReferencesSynced !== true ? { workbookSyncPending: true } : {}) }
            : { card: null, verificationRequired: true, cardRevision: cardRevision(record) };
        paymentCache.set(code, { key, result, until: now() + 60_000 });
        return result;
    }

    async function initialContext(user, directory) {
        if (!bootstrap) return {};
        const permissions = user.role === 'Administrador' ? { permissions: companyPermissions(user, directory) } : {};
        try { return { ...await bootstrap(user, directory), ...permissions }; }
        catch { return { config: { environment: 'sandbox', publicKey }, ...permissions }; }
    }

    async function finishCardDeletion(code, record, company) {
        await users.clearCompanyCard(company);
        await cardStore.write(code, record.customerId ? { customerId: record.customerId } : {});
        return { card: null };
    }

    return {
        claimCompany,
        companyPayment,
        config(req, res) {
            return res.json({ environment: 'sandbox', publicKey });
        },
        async login(req, res) {
            try {
                await sweep();
                const { email, password } = req.body || {};
                if (typeof email !== 'string' || email.length > 254 || typeof password !== 'string' || !password || password.length > 256) return fail(res, 401, 'invalid_credentials');
                // Bound both maps; failed login attempts must not exhaust sandbox memory.
                const key = req.ip;
                const attempt = attempts.get(key) || { count: 0, until: now() + 60_000 };
                if (attempt.count >= 10 || (attempts.size >= 1000 && !attempts.has(key))) return fail(res, 429, 'try_later');
                attempt.count += 1;
                attempts.set(key, attempt);
                const directory = await users.readUsers();
                const user = readUser(email, directory);
                const candidate = Buffer.from(digest(password));
                const expected = Buffer.from(digest(user ? String(user.password) : crypto.randomUUID()));
                if (!user || !crypto.timingSafeEqual(candidate, expected)) return fail(res, 401, 'invalid_credentials');
                if (sessionStore.size >= 500) return fail(res, 429, 'try_later');
                const token = crypto.randomBytes(32).toString('base64url');
                const expiresAt = now() + SESSION_DURATION;
                await sessionStore.set(digest(token), { email: normalize(user.email), company: companyIdentity(user), fingerprint: fingerprint(user), expiresAt });
                return res.json({ token, user: publicUser(user), expiresAt, ...await initialContext(user, directory) });
            } catch {
                return fail(res, 503, 'users_unavailable');
            }
        },
        async authorize(req, res, next) {
            try {
                await sweep();
                const token = (req.get('authorization') || '').match(/^Bearer ([A-Za-z0-9_-]{43})$/u)?.[1];
                if (!token) return fail(res, 401, 'session_expired');
                const tokenHash = digest(token);
                const session = await sessionStore.get(tokenHash);
                if (!session) return fail(res, 401, 'session_expired');
                const directory = await users.readUsers();
                const user = readUser(session.email, directory);
                if (!user || companyIdentity(user) !== session.company || fingerprint(user) !== session.fingerprint) {
                    await sessionStore.delete(tokenHash);
                    return fail(res, 401, 'session_expired');
                }
                res.locals.supportUser = user;
                res.locals.supportTokenHash = tokenHash;
                res.locals.supportUsers = directory;
                res.locals.supportExpiresAt = session.expiresAt;
                return next();
            } catch {
                return fail(res, 503, 'users_unavailable');
            }
        },
        async session(req, res) {
            return res.json({ user: publicUser(res.locals.supportUser), expiresAt: res.locals.supportExpiresAt,
                ...await initialContext(res.locals.supportUser, res.locals.supportUsers) });
        },
        async logout(req, res) {
            try {
                await sessionStore.delete(res.locals.supportTokenHash);
                return res.status(204).end();
            } catch {
                return fail(res, 503, 'session_unavailable');
            }
        },
        async getCard(req, res) {
            try {
                return res.json(await companyPayment(res.locals.supportUser));
            } catch (error) {
                if (error.code === 'card_review_required') return fail(res, 409, 'card_review_required');
                return fail(res, 503, 'card_unavailable');
            }
        },
        permissions(req, res) {
            const user = res.locals.supportUser;
            if (user.role !== 'Administrador') return fail(res, 403, 'administrator_required');
            return res.json(companyPermissions(user, res.locals.supportUsers));
        },
        async updatePermission(req, res) {
            const user = res.locals.supportUser;
            if (user.role !== 'Administrador') return fail(res, 403, 'administrator_required');
            const body = req.body;
            if (!body || typeof body !== 'object' || Array.isArray(body)
                || Object.keys(body).some((key) => key !== 'employees') || !Array.isArray(body.employees) || body.employees.length > 1000
                || body.employees.some((person) => !person || typeof person !== 'object' || Array.isArray(person)
                    || Object.keys(person).some((key) => !['email', 'allowed'].includes(key))
                    || typeof person.email !== 'string' || person.email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/u.test(person.email)
                    || typeof person.allowed !== 'boolean')) return fail(res, 400, 'invalid_request');
            const updates = new Map(body.employees.map((person) => [normalize(person.email), person.allowed]));
            if (updates.size !== body.employees.length) return fail(res, 400, 'invalid_request');
            const directory = res.locals.supportUsers;
            const targets = companyPermissions(user, directory).employees.filter((person) => person.active);
            if (targets.length !== updates.size || targets.some((person) => !updates.has(person.email))) return fail(res, 404, 'employee_not_found');
            const code = companyCode(companyIdentity(user));
            if (busyCompanies.has(code)) return fail(res, 409, 'permission_busy');
            busyCompanies.add(code);
            try {
                await users.writeBookingPermissions(user.companyId, [...updates].map(([email, allowed]) => ({ email, allowed })), directory);
                return res.json(companyPermissions(user, directory.map((person) => companyIdentity(person) === companyIdentity(user) && updates.has(normalize(person.email))
                    ? { ...person, bookingAllowed: updates.get(normalize(person.email)) } : person)));
            } catch {
                return fail(res, 503, 'permission_save_failed');
            } finally { busyCompanies.delete(code); }
        },
        async deleteCard(req, res) {
            const user = res.locals.supportUser;
            if (user.role !== 'Administrador') return fail(res, 403, 'administrator_required');
            if (!req.body || typeof req.body !== 'object' || Array.isArray(req.body)
                || Object.keys(req.body).some((key) => key !== 'expectedCardRevision')
                || typeof req.body.expectedCardRevision !== 'string' || !/^[a-f0-9]{64}$/u.test(req.body.expectedCardRevision)) return fail(res, 400, 'invalid_request');
            const code = companyCode(companyIdentity(user));
            if (busyCompanies.has(code)) return fail(res, 409, 'card_busy');
            busyCompanies.add(code);
            paymentCache.delete(code);
            try {
                let { record } = await readCompanyCard(user);
                if (record.cardId && req.body.expectedCardRevision !== cardRevision(record)) return fail(res, 409, 'card_changed');
                if (record.pending && !['workbook', 'delete', 'delete_workbook'].includes(record.pending)) return fail(res, 409, 'card_review_required');
                if (!record.cardId || record.pending === 'delete_workbook') return res.json(await finishCardDeletion(code, record, user.companyId || user.company));
                let deleted = false;
                if (record.pending === 'delete') {
                    try { deleted = (await pagarme.getCard(record.customerId, record.cardId)).status === 'deleted'; }
                    catch (error) { if (error.code === 'pagarme.card_not_found') deleted = true; else throw error; }
                }
                if (!deleted) {
                    const previous = record;
                    await cardStore.write(code, { ...record, pending: 'delete' });
                    try {
                        if ((await pagarme.deleteCard(record.customerId, record.cardId)).deleted !== true) throw new Error('Card deletion not confirmed');
                    } catch (error) {
                        if (error.definiteRejection) await cardStore.write(code, previous);
                        throw error;
                    }
                }
                record = { customerId: record.customerId, cardId: record.cardId, cardRevision: cardRevision(record), pending: 'delete_workbook' };
                await cardStore.write(code, record);
                return res.json(await finishCardDeletion(code, record, user.companyId || user.company));
            } catch {
                return fail(res, 503, 'card_delete_failed');
            } finally {
                busyCompanies.delete(code);
            }
        },
        async saveCard(req, res) {
            const user = res.locals.supportUser;
            if (user.role !== 'Administrador') return fail(res, 403, 'administrator_required');
            const { token, billingAddress } = req.body || {};
            if (Object.keys(req.body || {}).some((key) => !['token', 'billingAddress'].includes(key)) || typeof token !== 'string' || !/^token_[A-Za-z0-9]{1,200}$/u.test(token)
                || (billingAddress !== undefined && !validBillingAddress(billingAddress))) return fail(res, 400, 'invalid_card');
            const code = companyCode(companyIdentity(user));
            if (busyCompanies.has(code)) return fail(res, 409, 'card_busy');
            busyCompanies.add(code);
            paymentCache.delete(code);
            try {
                let { record } = await readCompanyCard(user);
                if (record.pending === 'workbook' && hasVerification(record)
                    && usableCard(await pagarme.getCard(record.customerId, record.cardId))) return fail(res, 409, 'workbook_sync_required');
                if (record.pending && record.pending !== 'workbook') return fail(res, 409, 'card_review_required');
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
                    if (error.code === 'pagarme.card_verification_failed') return fail(res, 422, 'card_verification_failed');
                    throw error;
                }
                // A successful explicit verification request may return the existing wallet card ID.
                if (card.verificationRequested !== true || !usableCard(card)) throw new Error('Card verification not confirmed');
                const confirmed = { customerId: record.customerId, cardId: card.id, cardRevision: crypto.randomBytes(32).toString('hex'),
                    verification: CARD_VERIFICATION, pending: 'workbook' };
                // Persist the confirmed provider card before Excel so a retry cannot register it twice.
                await cardStore.write(code, confirmed);
                try {
                    return res.json(await synchronizeWorkbook(code, confirmed, user.companyId || user.company, card));
                } catch {
                    return res.json({ card: publicCard(card), workbookSyncPending: true, cardRevision: cardRevision(confirmed) });
                }
            } catch {
                return fail(res, 503, 'card_save_failed');
            } finally {
                busyCompanies.delete(code);
            }
        },
        async syncCard(req, res) {
            const user = res.locals.supportUser;
            if (user.role !== 'Administrador') return fail(res, 403, 'administrator_required');
            if (Object.keys(req.body || {}).length) return fail(res, 400, 'invalid_request');
            const code = companyCode(companyIdentity(user));
            if (busyCompanies.has(code)) return fail(res, 409, 'card_busy');
            busyCompanies.add(code);
            paymentCache.delete(code);
            try {
                const { record } = await readCompanyCard(user);
                if (record.pending && record.pending !== 'workbook') return fail(res, 409, 'card_review_required');
                if (!record.cardId) return fail(res, 404, 'card_unavailable');
                if (!hasVerification(record)) return fail(res, 409, 'card_verification_required');
                const card = await pagarme.getCard(record.customerId, record.cardId);
                if (!usableCard(card)) return fail(res, 409, 'card_verification_required');
                await cardStore.write(code, { customerId: record.customerId, cardId: record.cardId, cardRevision: cardRevision(record),
                    verification: CARD_VERIFICATION, pending: 'workbook' });
                return res.json(await synchronizeWorkbook(code, record, user.companyId || user.company, card));
            } catch {
                return fail(res, 503, 'workbook_sync_failed');
            } finally {
                busyCompanies.delete(code);
            }
        },
    };
}

module.exports = { createSupportManagementHandlers, companyCode };
