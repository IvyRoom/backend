'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');

const UUID = /^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/u;
const STAGES = ['workbook', 'event', 'confirmation', 'complete'];

function validateBooking(value) {
    const keys = ['id', 'date', 'companyId', 'email', 'name', 'company', 'priceCents', 'createdAt', 'stage', 'eventId', 'iCalUid'];
    if (!value || typeof value !== 'object' || Array.isArray(value)
        || Object.keys(value).some((key) => !keys.includes(key))
        || typeof value.id !== 'string' || !UUID.test(value.id)
        || typeof value.date !== 'string' || !/^\d{4}-\d{2}-\d{2}$/u.test(value.date)
        || !Number.isFinite(Date.parse(`${value.date}T00:00:00Z`))
        || new Date(`${value.date}T00:00:00Z`).toISOString().slice(0, 10) !== value.date
        || typeof value.companyId !== 'string' || !/^\d{14}$/u.test(value.companyId)
        || typeof value.email !== 'string' || !value.email || value.email.length > 254 || /[\r\n]/u.test(value.email)
        || typeof value.name !== 'string' || !value.name || value.name.length > 255 || /[\x00-\x1f\x7f]/u.test(value.name)
        || typeof value.company !== 'string' || !value.company || value.company.length > 4096 || /[\x00-\x1f\x7f]/u.test(value.company)
        || !Number.isSafeInteger(value.priceCents) || value.priceCents <= 0
        || !Number.isSafeInteger(value.createdAt) || value.createdAt <= 0
        || !STAGES.includes(value.stage)
        || (value.eventId !== undefined && (typeof value.eventId !== 'string' || !value.eventId || value.eventId.length > 2048 || /[\r\n]/u.test(value.eventId)))
        || (value.iCalUid !== undefined && (typeof value.iCalUid !== 'string' || !value.iCalUid || value.iCalUid.length > 1024 || /[\r\n]/u.test(value.iCalUid)))
        || (['confirmation', 'complete'].includes(value.stage) && (!value.eventId || !value.iCalUid))) {
        throw new Error('Invalid support booking record');
    }
    return { ...value };
}

function createSupportBookingStore(directory) {
    function filename(id) {
        if (typeof id !== 'string' || !UUID.test(id)) throw new Error('Invalid support booking ID');
        return path.join(directory, `reservation_${id}.json`);
    }
    async function read(id) {
        try {
            const record = validateBooking(JSON.parse(await fs.readFile(filename(id), 'utf8')));
            if (record.id !== id) throw new Error('Booking ID mismatch');
            return record;
        } catch (error) {
            if (error.code === 'ENOENT') return null;
            throw new Error('Support booking storage unavailable');
        }
    }
    return {
        read,
        async write(value) {
            const record = validateBooking(value);
            const target = filename(record.id);
            const temporary = `${target}.${crypto.randomUUID()}.tmp`;
            await fs.mkdir(directory, { recursive: true });
            try {
                await fs.writeFile(temporary, JSON.stringify(record), { flag: 'wx', mode: 0o600 });
                await fs.rename(temporary, target);
            } finally {
                await fs.rm(temporary, { force: true });
            }
        },
        async remove(id) { await fs.rm(filename(id), { force: true }); },
        async pendingDates() {
            const names = await fs.readdir(directory).catch((error) => { if (error.code === 'ENOENT') return []; throw error; });
            const dates = new Set();
            for (const name of names) {
                const match = /^reservation_(.+)\.json$/u.exec(name);
                if (!match) continue;
                const record = await read(match[1]);
                if (record && record.stage !== 'complete') dates.add(record.date);
            }
            return dates;
        },
    };
}

module.exports = { createSupportBookingStore };
