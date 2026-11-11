'use strict';

const { escapeHtml } = require('../shared/escape-html');
const { appointmentAvailability, bookingCalendarRange, isCalendarDate } = require('./support-availability');

const CONTACT = 'contato@machadogestao.com';
const TIME_ZONE = 'E. South America Standard Time';
const CONFIRMATION_SUBJECT = 'Atendimento Consultivo - Machado | Método Gerencial para Empresas';
const UUID = /^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/u;
const normalize = (value) => String(value || '').trim().toLowerCase();

function localDate(nowMs) {
    const parts = Object.fromEntries(new Intl.DateTimeFormat('en', {
        timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit', day: '2-digit',
    }).formatToParts(new Date(nowMs)).map(({ type, value }) => [type, value]));
    return `${parts.year}-${parts.month}-${parts.day}`;
}

function calendarDays(rows, referenceRate, nowMs, pendingDates = new Set()) {
    const baseCents = Math.round(referenceRate * 100);
    if (typeof referenceRate !== 'number' || !Number.isFinite(referenceRate) || !Number.isSafeInteger(baseCents) || baseCents <= 0) throw new Error('Reference rate unavailable');
    const today = localDate(nowMs);
    const range = bookingCalendarRange(today);
    let vacantRank = 0;
    return rows.filter(({ date }) => isCalendarDate(date) && date >= range.first && date <= range.last)
        .sort((a, b) => a.date.localeCompare(b.date)).map(({ date, status }) => {
            const available = appointmentAvailability(date, today) && !String(status || '').trim() && !pendingDates.has(date);
            const discountStep = available ? Math.min(Math.max(vacantRank++ - 2, 0), 7) : null;
            return { date, available, priceCents: available ? Math.round(baseCents * (28 - discountStep) / 28) : null,
                pricePosition: available ? discountStep / 7 : null };
        });
}

function emailAddress(value) {
    return typeof value === 'string' && value.length <= 254
        && /^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$/u.test(value);
}

function appointmentDateTime(booking) {
    return `${booking.date.split('-').reverse().join('/')} às 09:00 (horário de Brasília)`;
}

function confirmation(booking) {
    if (typeof booking.meetingUrl !== 'string' || !booking.meetingUrl) throw new Error('Teams link unavailable');
    const firstName = booking.name.trim().split(/\s+/u)[0];
    const meetingUrl = escapeHtml(booking.meetingUrl);
    return { subject: CONFIRMATION_SUBJECT,
        body: { contentType: 'HTML', content: `<p>Prezado(a) ${escapeHtml(firstName)},</p><p>Seu atendimento consultivo junto a nosso diretor está agendado para ${escapeHtml(appointmentDateTime(booking))}.</p><p>Por favor confirme sua presença no invite do calendário.</p><p>O link de acesso ao encontro é: <a href="${meetingUrl}">${meetingUrl}</a></p><p>Atenciosamente,</p><p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg" width="600" /></p>` },
        toRecipients: [{ emailAddress: { address: normalize(booking.email) } }],
    };
}

function eventPayload(booking) {
    return { subject: `Atendimento Consultivo: ${booking.name}`,
        body: { contentType: 'HTML', content: `<p>${escapeHtml(booking.name)}<br>${escapeHtml(booking.company)}</p><p>${escapeHtml(appointmentDateTime(booking))}.</p>` },
        isAllDay: false,
        start: { dateTime: `${booking.date}T09:00:00`, timeZone: TIME_ZONE },
        end: { dateTime: `${booking.date}T10:00:00`, timeZone: TIME_ZONE },
        isOnlineMeeting: true, onlineMeetingProvider: 'teamsForBusiness',
        attendees: normalize(booking.email) === CONTACT ? [] : [{ emailAddress: { address: booking.email, name: booking.name }, type: 'required' }],
        showAs: 'busy', transactionId: booking.id,
    };
}

function createSupportBookingHandlers({ appointments, bookingStore, bookingNotifications, users, companyCards = async () => ({ card: null }), claimCompany = () => () => {}, now = Date.now }) {
    const busyDates = new Set();
    const busyIds = new Set();
    let claims = Promise.resolve();
    let pendingClaims = 0;
    const fail = (res, status, code) => res.status(status).json({ error: `support_management.${code}` });
    async function readCalendar(user) {
        const pendingDates = await bookingStore.pendingDates();
        const today = localDate(now());
        const rows = appointments.syncAvailability ? await appointments.syncAvailability({ today }) : await appointments.readDates();
        return { days: calendarDays(rows, user.referenceRate, now(), pendingDates), rows };
    }
    async function getCalendarContext(user, directory = []) {
        const card = await companyCards(user, { cached: true });
        const administrators = directory.filter((person) => person.role === 'Administrador'
            && normalize(person.companyId || person.company) === normalize(user.companyId || user.company));
        const canBook = Boolean(card.card) && user.bookingAllowed === true;
        return { card, calendar: { days: canBook ? (await readCalendar(user)).days : [], canBook,
            paymentMethodRegistered: Boolean(card.card),
            administratorName: [...new Set(administrators.map((person) => person.name))].join(', ') || 'não informado',
            userEmail: user.email, today: localDate(now()) } };
    }
    return {
        getCalendarContext,
        async calendar(req, res) {
            try { return res.json((await getCalendarContext(res.locals.supportUser, res.locals.supportUsers)).calendar); }
            catch { return fail(res, 503, 'calendar_unavailable'); }
        },
        async reserve(req, res) {
            const user = res.locals.supportUser;
            const { date, priceCents, requestId, acknowledged } = req.body || {};
            if (Object.keys(req.body || {}).some((key) => !['date', 'priceCents', 'requestId', 'acknowledged'].includes(key)) || acknowledged !== true
                || !isCalendarDate(date) || !Number.isSafeInteger(priceCents) || priceCents <= 0
                || typeof requestId !== 'string' || !UUID.test(requestId)
                || typeof user.companyId !== 'string' || !/^\d{14}$/u.test(user.companyId) || !emailAddress(user.email)) return fail(res, 400, 'invalid_request');
            if (busyDates.has(date) || busyIds.has(requestId)) return fail(res, 409, 'date_unavailable');
            busyDates.add(date);
            busyIds.add(requestId);
            const queued = pendingClaims > 0;
            pendingClaims++;
            const previousClaim = claims;
            let releaseClaim;
            claims = new Promise((resolve) => {
                releaseClaim = () => { pendingClaims--; resolve(); };
            });
            let reserved = false;
            let record;
            let releaseCompany;
            try {
                await previousClaim;
                const previous = await bookingStore.read(requestId);
                if (previous) {
                    if (previous.date !== date || previous.priceCents !== priceCents || previous.companyId !== user.companyId
                        || normalize(previous.email) !== normalize(user.email)) return fail(res, 409, 'reservation_review_required');
                    if (previous.stage === 'complete') return res.json({ date, reserved: true });
                    if (previous.stage !== 'workbook') return res.json({ date, reserved: true, notificationPending: true });
                    return fail(res, 409, 'reservation_review_required');
                }
                releaseCompany = claimCompany(user);
                if (!releaseCompany) return fail(res, 409, 'card_busy');
                // Authorization just read this directory; queued claims must recheck it after waiting.
                const directory = !queued && Array.isArray(res.locals.supportUsers) ? res.locals.supportUsers
                    : users ? await users.readUsers() : [user];
                const current = directory.filter((person) => normalize(person.email) === normalize(user.email));
                if (current.length !== 1 || current[0].companyId !== user.companyId || current[0].password !== user.password || current[0].bookingAllowed !== true) return fail(res, 403, 'booking_permission_required');
                if (!(await companyCards(user)).card) return fail(res, 403, 'payment_method_required');
                const calendar = await readCalendar(current[0]);
                const day = calendar.days.find((entry) => entry.date === date);
                if (!day?.available) return fail(res, 409, 'date_unavailable');
                if (day.priceCents !== priceCents) return fail(res, 409, 'price_changed');
                record = { id: requestId, date, companyId: user.companyId, email: normalize(user.email), name: user.name,
                    company: user.company, priceCents, createdAt: now(), acknowledgedAt: now(), stage: 'workbook' };
                await bookingStore.write(record);
                // Claim quotes in order: occupying an earlier date changes every later vacant price.
                releaseClaim();
                releaseClaim = undefined;
                try {
                    await appointments.reserveDate({ date, companyId: user.companyId, name: user.name, hourlyRate: priceCents / 100 }, calendar.rows);
                } catch (error) {
                    if (error.definiteRejection) {
                        await bookingStore.remove(requestId);
                        record = undefined;
                    }
                    if (error.code === 'support_appointments.date_unavailable') return fail(res, 409, 'date_unavailable');
                    throw error;
                }
                reserved = true;
                releaseCompany();
                releaseCompany = undefined;
                record = { ...record, stage: 'event' };
                await bookingStore.write(record);
                const event = await bookingNotifications.createEvent(eventPayload(record));
                record = { ...record, ...event, stage: 'confirmation' };
                await bookingStore.write(record);
                await bookingNotifications.sendConfirmation(confirmation(record));
                await bookingStore.write({ ...record, stage: 'complete' });
                return res.json({ date, reserved: true });
            } catch {
                if (reserved) return res.json({ date, reserved: true, notificationPending: true });
                return fail(res, record ? 409 : 503, record ? 'reservation_review_required' : 'reservation_failed');
            } finally {
                releaseCompany?.();
                releaseClaim?.();
                busyDates.delete(date);
                busyIds.delete(requestId);
            }
        },
    };
}

module.exports = { createSupportBookingHandlers, calendarDays, eventPayload, confirmation };
