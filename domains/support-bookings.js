'use strict';

const { escapeHtml } = require('../shared/escape-html');

const CONTACT = 'contato@machadogestao.com';
const TIME_ZONE = 'E. South America Standard Time';
const UUID = /^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/u;
const normalize = (value) => String(value || '').trim().toLowerCase();

function validDate(value) {
    return typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/u.test(value)
        && Number.isFinite(Date.parse(`${value}T00:00:00Z`))
        && new Date(`${value}T00:00:00Z`).toISOString().slice(0, 10) === value;
}

function nextDate(date) {
    return new Date(Date.parse(`${date}T00:00:00Z`) + 86400000).toISOString().slice(0, 10);
}

function tomorrow(nowMs) {
    const parts = Object.fromEntries(new Intl.DateTimeFormat('en', {
        timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit', day: '2-digit',
    }).formatToParts(new Date(nowMs)).map(({ type, value }) => [type, value]));
    return nextDate(`${parts.year}-${parts.month}-${parts.day}`);
}

function calendarDays(rows, referenceRate, nowMs, pendingDates = new Set()) {
    const baseCents = Math.round(referenceRate * 100);
    if (typeof referenceRate !== 'number' || !Number.isFinite(referenceRate) || !Number.isSafeInteger(baseCents) || baseCents <= 0) throw new Error('Reference rate unavailable');
    const firstDate = tomorrow(nowMs);
    let vacantRank = 0;
    return rows.filter(({ date }) => validDate(date) && date >= firstDate
        && ![0, 6].includes(new Date(`${date}T12:00:00Z`).getUTCDay()))
        .sort((a, b) => a.date.localeCompare(b.date)).map(({ date, status }) => {
            const available = !String(status || '').trim() && !pendingDates.has(date);
            const rank = available ? Math.min(vacantRank++, 10) : null;
            return { date, available, priceCents: available ? Math.round(baseCents * (40 - rank) / 40) : null,
                pricePosition: available ? rank / 10 : null };
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

function confirmation(booking) {
    const displayDate = booking.date.split('-').reverse().join('/');
    const price = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(booking.priceCents / 100);
    const subject = `Atendimento consultivo - ${displayDate}`;
    const description = `${booking.name}\n${booking.company}\n${displayDate}\nPreço por hora: ${price}`;
    const calendar = [
        'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Machado//Atendimentos Consultivos//PT-BR', 'METHOD:REQUEST',
        'BEGIN:VEVENT', `UID:${textCalendar(booking.iCalUid)}`,
        `DTSTAMP:${new Date(booking.createdAt).toISOString().replace(/[-:]/gu, '').replace(/\.\d{3}/u, '')}`,
        `DTSTART;VALUE=DATE:${booking.date.replace(/-/gu, '')}`,
        `DTEND;VALUE=DATE:${nextDate(booking.date).replace(/-/gu, '')}`,
        `SUMMARY:${textCalendar(subject)}`, `DESCRIPTION:${textCalendar(description)}`,
        `ORGANIZER:mailto:${CONTACT}`, `ATTENDEE;RSVP=TRUE:mailto:${booking.email}`,
        'STATUS:CONFIRMED', 'SEQUENCE:0', 'END:VEVENT', 'END:VCALENDAR',
    ].map(foldCalendarLine).join('\r\n') + '\r\n';
    return { subject, body: { contentType: 'HTML', content: `<p>${escapeHtml(booking.name)}<br>${escapeHtml(booking.company)}</p><p>${displayDate}<br>Preço por hora: ${escapeHtml(price)}</p>` },
        toRecipients: Array.from(new Set([CONTACT, normalize(booking.email)])).map((address) => ({ emailAddress: { address } })),
        attachments: [{ '@odata.type': '#microsoft.graph.fileAttachment', name: 'atendimento.ics', contentType: 'text/calendar; method=REQUEST; charset=UTF-8', contentBytes: Buffer.from(calendar, 'utf8').toString('base64') }],
    };
}

function eventPayload(booking) {
    const message = confirmation({ ...booking, iCalUid: booking.iCalUid || booking.id });
    return { subject: message.subject, body: message.body, isAllDay: true,
        start: { dateTime: `${booking.date}T00:00:00`, timeZone: TIME_ZONE },
        end: { dateTime: `${nextDate(booking.date)}T00:00:00`, timeZone: TIME_ZONE },
        attendees: normalize(booking.email) === CONTACT ? [] : [{ emailAddress: { address: booking.email, name: booking.name }, type: 'required' }],
        showAs: 'busy', transactionId: booking.id,
    };
}

function createSupportBookingHandlers({ appointments, bookingStore, bookingNotifications, now = Date.now }) {
    const busyDates = new Set();
    const busyIds = new Set();
    let claims = Promise.resolve();
    const fail = (res, status, code) => res.status(status).json({ error: `support_management.${code}` });
    async function readCalendar(user) {
        const pendingDates = await bookingStore.pendingDates();
        return calendarDays(await appointments.readDates(), user.referenceRate, now(), pendingDates);
    }
    return {
        async calendar(req, res) {
            try { return res.json({ days: await readCalendar(res.locals.supportUser) }); }
            catch { return fail(res, 503, 'calendar_unavailable'); }
        },
        async reserve(req, res) {
            const user = res.locals.supportUser;
            const { date, priceCents, requestId } = req.body || {};
            if (Object.keys(req.body || {}).some((key) => !['date', 'priceCents', 'requestId'].includes(key))
                || !validDate(date) || !Number.isSafeInteger(priceCents) || priceCents <= 0
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
                const day = (await readCalendar(user)).find((entry) => entry.date === date);
                if (!day?.available) return fail(res, 409, 'date_unavailable');
                if (day.priceCents !== priceCents) return fail(res, 409, 'price_changed');
                record = { id: requestId, date, companyId: user.companyId, email: normalize(user.email), name: user.name,
                    company: user.company, priceCents, createdAt: now(), stage: 'workbook' };
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
                releaseClaim?.();
                busyDates.delete(date);
                busyIds.delete(requestId);
            }
        },
    };
}

module.exports = { createSupportBookingHandlers, calendarDays, eventPayload, confirmation };
