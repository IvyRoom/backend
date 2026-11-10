'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');

function validateRecord(record) {
    if (!record || typeof record !== 'object' || Array.isArray(record)) throw new Error('Invalid card reference record');
    const result = {};
    for (const [key, value] of Object.entries(record)) {
        const valid = key === 'workbookReferencesSynced' ? value === true
            : key === 'workbookSynced' ? typeof value === 'boolean' : typeof value === 'string' && (
            (key === 'customerId' && /^cus_[A-Za-z0-9]+$/.test(value))
            || (key === 'cardId' && /^card_[A-Za-z0-9]+$/.test(value))
            || (key === 'cardRevision' && /^[a-f0-9]{64}$/u.test(value))
            || (key === 'verification' && value === 'pagarme_verify_card_v1')
            || (key === 'pending' && ['customer', 'card', 'workbook', 'delete', 'delete_workbook'].includes(value))
        );
        if (!valid) throw new Error('Invalid card reference record');
        result[key] = value;
    }
    if ((result.cardId && !result.customerId)
        || (result.pending === 'card' && !result.customerId)
        || (result.pending === 'workbook' && (!result.customerId || !result.cardId))
        || (['delete', 'delete_workbook'].includes(result.pending) && (!result.customerId || !result.cardId))
        || (result.workbookSynced !== undefined && (!result.customerId || !result.cardId))
        || (result.workbookReferencesSynced !== undefined && (!result.customerId || !result.cardId || result.workbookSynced !== true))
        || (result.verification !== undefined && (!result.customerId || !result.cardId))
        || (result.cardRevision !== undefined && (!result.customerId || !result.cardId))
        || (result.pending === 'workbook' && result.workbookSynced === true)
        || (result.pending === 'customer' && (result.customerId || result.cardId))) {
        throw new Error('Invalid card reference record');
    }
    return result;
}

async function acquireSupportSandboxLock(directory) {
    await fs.mkdir(directory, { recursive: true });
    const lockPath = path.join(directory, 'sandbox.lock');
    const lock = await fs.open(lockPath, 'wx', 0o600);
    try {
        await lock.writeFile(String(process.pid));
    } catch (error) {
        try {
            await lock.close();
        } finally {
            await fs.rm(lockPath, { force: true });
        }
        throw error;
    }
    let released;
    return function release() {
        if (!released) released = lock.close().then(() => fs.rm(lockPath));
        return released;
    };
}

function createSupportCardStore(directory) {
    function filename(companyCode) {
        if (!/^support_[a-f0-9]{44}$/.test(companyCode)) throw new Error('Invalid company code');
        return path.join(directory, `${companyCode}.json`);
    }

    return {
        async findByReferences({ customerId, cardId }) {
            const reference = validateRecord({ customerId, cardId });
            const names = await fs.readdir(directory).catch((error) => { if (error.code === 'ENOENT') return []; throw error; });
            const matches = [];
            for (const name of names) {
                if (!/^support_[a-f0-9]{44}\.json$/u.test(name)) continue;
                const record = validateRecord(JSON.parse(await fs.readFile(path.join(directory, name), 'utf8')));
                if (record.customerId === reference.customerId && record.cardId === reference.cardId) matches.push(record);
            }
            if (matches.length > 1) throw new Error('Ambiguous local card references');
            return matches[0] || null;
        },
        async read(companyCode) {
            try {
                return validateRecord(JSON.parse(await fs.readFile(filename(companyCode), 'utf8')));
            } catch (error) {
                if (error.code === 'ENOENT') return {};
                throw new Error('Card reference storage unavailable');
            }
        },
        async write(companyCode, record) {
            const target = filename(companyCode);
            const safeRecord = validateRecord(record);
            await fs.mkdir(directory, { recursive: true });
            const temporary = `${target}.${crypto.randomUUID()}.tmp`;
            try {
                await fs.writeFile(temporary, JSON.stringify(safeRecord), { flag: 'wx', mode: 0o600 });
                await fs.rename(temporary, target);
            } finally {
                await fs.rm(temporary, { force: true });
            }
        },
    };
}

module.exports = { createSupportCardStore, acquireSupportSandboxLock };
