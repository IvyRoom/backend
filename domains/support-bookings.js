'use strict';

const { escapeHtml } = require('../shared/escape-html');
const { appointmentAvailability, bookingCalendarRange, isCalendarDate } = require('./support-availability');

const CONTACT = 'contato@machadogestao.com';
const TIME_ZONE = 'E. South America Standard Time';
const TIME_LABEL = 'Início às 09:00 (horário de Brasília), com duração de até 2h, conforme necessário.';
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

function textCalendar(value) {
    return String(value).replace(/\\/gu, '\\\\').replace(/\r\n|\r|\n/gu, '\\n').replace(/;/gu, '\\;').replace(/,/gu, '\\,');
}

function foldCalendarLine(value) {
    let line = '';
    const lines = [];
    for (const character of value) {
        if (Buffer.byteLength(line + character) > 75) { lines.push(line); line = ' '; }
        line += character;
    }
    lines.push(line);
    return lines.join('\r\n');
}

function calendarTime(date, time) {
    const zone = new Intl.DateTimeFormat('en', { timeZone: 'America/Sao_Paulo', timeZoneName: 'longOffset' })
        .formatToParts(new Date(`${date}T12:00:00Z`)).find(({ type }) => type === 'timeZoneName').value;
    const offset = zone === 'GMT' ? '+00:00' : zone.slice(3);
    return new Date(`${date}T${time}:00${offset}`).toISOString().replace(/[-:]/gu, '').replace(/\.\d{3}/u, '');
}

function confirmation(booking) {
    const displayDate = booking.date.split('-').reverse().join('/');
    const price = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(booking.priceCents / 100);
    const subject = `Atendimento consultivo - ${displayDate}`;
    const description = `${booking.name}\n${booking.company}\n${displayDate}\n${TIME_LABEL}\nPreço por hora: ${price}${booking.meetingUrl ? `\n${booking.meetingUrl}` : ''}`;
    const calendar = [
        'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Machado//Atendimentos Consultivos//PT-BR', 'METHOD:REQUEST',
        'BEGIN:VEVENT', `UID:${textCalendar(booking.iCalUid)}`,
        `DTSTAMP:${new Date(booking.createdAt).toISOString().replace(/[-:]/gu, '').replace(/\.\d{3}/u, '')}`,
        `DTSTART:${calendarTime(booking.date, '09:00')}`,
        `DTEND:${calendarTime(booking.date, '11:00')}`,
        `SUMMARY:${textCalendar(subject)}`, `DESCRIPTION:${textCalendar(description)}`,
        `ORGANIZER:mailto:${CONTACT}`, `ATTENDEE;RSVP=TRUE:mailto:${booking.email}`,
        ...(booking.meetingUrl ? [`URL:${booking.meetingUrl}`] : []),
        'STATUS:CONFIRMED', 'SEQUENCE:0', 'END:VEVENT', 'END:VCALENDAR',
    ].map(foldCalendarLine).join('\r\n') + '\r\n';
    return { subject, body: { contentType: 'HTML', content: `<p>${escapeHtml(booking.name)}<br>${escapeHtml(booking.company)}</p><p>${displayDate}<br>${TIME_LABEL}<br>Preço por hora: ${escapeHtml(price)}</p>${booking.meetingUrl ? `<p><a href="${escapeHtml(booking.meetingUrl)}">Ingressar no Microsoft Teams</a></p>` : ''}` },
        toRecipients: Array.from(new Set([CONTACT, normalize(booking.email)])).map((address) => ({ emailAddress: { address } })),
        attachments: [{ '@odata.type': '#microsoft.graph.fileAttachment', name: 'atendimento.ics', contentType: 'text/calendar; method=REQUEST; charset=UTF-8', contentBytes: Buffer.from(calendar, 'utf8').toString('base64') }],
    };
}

function eventPayload(booking) {
    const message = confirmation({ ...booking, iCalUid: booking.iCalUid || booking.id });
    return { subject: message.subject, body: message.body, isAllDay: false,
        start: { dateTime: `${booking.date}T09:00:00`, timeZone: TIME_ZONE },
        end: { dateTime: `${booking.date}T11:00:00`, timeZone: TIME_ZONE },
        isOnlineMeeting: true, onlineMeetingProvider: 'teamsForBusiness',
        attendees: normalize(booking.email) === CONTACT ? [] : [{ emailAddress: { address: booking.email, name: booking.name }, type: 'required' }],
        showAs: 'busy', transactionId: booking.id,
    };
}

function createSupportBookingHandlers({ appointments, bookingStore, bookingNotifications, companyCards = async () => ({ card: null }), claimCompany = () => () => {}, now = Date.now }) {
    const busyDates = new Set();
    const busyIds = new Set();
    let claims = Promise.resolve();
    const fail = (res, status, code) => res.status(status).json({ error: `support_management.${code}` });
    async function readCalendar(user) {
        const pendingDates = await bookingStore.pendingDates();
        const today = localDate(now());
        const rows = appointments.syncAvailability ? await appointments.syncAvailability({ today }) : await appointments.readDates();
        return calendarDays(rows, user.referenceRate, now(), pendingDates);
    }
    async function getCalendarContext(user, directory = []) {
        const card = await companyCards(user, { cached: true });
        const administrators = directory.filter((person) => person.status === 'Ativo' && person.role === 'Administrador'
            && normalize(person.companyId || person.company) === normalize(user.companyId || user.company));
        return { card, calendar: { days: card.card ? await readCalendar(user) : [],
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
            const previousClaim = claims;
            let releaseClaim;
            claims = new Promise((resolve) => { releaseClaim = resolve; });
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
                if (!(await companyCards(user)).card) return fail(res, 403, 'payment_method_required');
                const day = (await readCalendar(user)).find((entry) => entry.date === date);
                if (!day?.available) return fail(res, 409, 'date_unavailable');
                if (day.priceCents !== priceCents) return fail(res, 409, 'price_changed');
                record = { id: requestId, date, companyId: user.companyId, email: normalize(user.email), name: user.name,
                    company: user.company, priceCents, createdAt: now(), acknowledgedAt: now(), stage: 'workbook' };
                await bookingStore.write(record);
                // Claim quotes in order: occupying an earlier date changes every later vacant price.
                releaseClaim();
                releaseClaim = undefined;
                try {
                    await appointments.reserveDate({ date, companyId: user.companyId, name: user.name, hourlyRate: priceCents / 100 });
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
