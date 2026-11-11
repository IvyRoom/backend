'use strict';

const { RetryHandlerOptions } = require('@microsoft/microsoft-graph-client');

const WORKSHEET_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRXPEMKXY4F5RH34SKSN35K757F/workbook/worksheets/{00000000-0001-0000-0000-000000000000}';
const TABLE_PATH = `${WORKSHEET_PATH}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}`;
const COLUMNS = [
    'CNPJ', 'RAZÃO SOCIAL', 'VALOR REFERÊNCIA', 'NOME COMPLETO', 'E-MAIL', 'SENHA',
    'TIPO DE ACESSO', 'AUTORIZADO A CONTRATAR ATENDIMENTOS', 'PAGARME CUSTOMER ID', 'PAGARME CARD ID', 'PAGARME AMBIENTE',
];

function createSupportUsersAdapter({ graphClient }) {
    const snapshots = new WeakMap();
    async function readVerifiedTable() {
        const [columns, range] = await Promise.all([
            graphClient.api(`${TABLE_PATH}/columns?$select=index,name`).get(),
            graphClient.api(`${TABLE_PATH}/range?$select=address,rowCount,columnCount`).get(),
        ]);
        const expected = COLUMNS;
        if (!Array.isArray(columns.value) || columns.value.length !== expected.length
            || columns.value.some((column, index) => column?.index !== index
                || typeof column.name !== 'string' || column.name.trim() !== expected[index])) {
            throw new Error('support_users.schema_mismatch');
        }

        const address = /^USUÁRIOS!A3:K([3-9]|[1-9]\d+)$/.exec(range.address);
        const lastRow = Number(address?.[1]);
        if (!address
            || range.columnCount !== expected.length || range.rowCount !== lastRow - 2) {
            throw new Error('support_users.schema_mismatch');
        }
        return lastRow;
    }

    async function readRows(address, count, width) {
        const response = await graphClient.api(`${WORKSHEET_PATH}/range(address='${address}')?$select=values`).get();
        if (!Array.isArray(response.values) || response.values.length !== count
            || response.values.some((row) => !Array.isArray(row) || row.length !== width)) {
            throw new Error('support_users.schema_mismatch');
        }
        return response.values;
    }

    async function readVerifiedUsers() {
        const lastRow = await readVerifiedTable();
        // H is booking authorization; this workbook no longer has a login-status column.
        const rows = lastRow === 3 ? [] : await readRows(`A4:H${lastRow}`, lastRow - 3, 8);
        const users = rows.map((row) => {
            const companyId = normalizeCompanyId(row[0]);
            if (!companyId || typeof row[2] !== 'number' || !Number.isFinite(row[2]) || row[2] <= 0) {
                throw new Error('support_users.invalid_company_data');
            }
            const permission = String(row[7] ?? '').trim();
            return { companyId, company: String(row[1]), referenceRate: row[2],
                name: String(row[3]), email: String(row[4]), password: String(row[5]),
                role: String(row[6]), bookingAllowed: !permission || permission === 'Sim' };
        });
        snapshots.set(users, rows.map((row, index) => ({ row: index + 4,
            companyId: normalizeCompanyId(row[0]), email: String(row[4]).trim().toLowerCase(),
            permission: String(row[7] ?? '').trim() })));
        return users;
    }

    async function readCompanyRows(companyId) {
        const normalized = normalizeCompanyId(companyId);
        if (!normalized) throw new Error('support_users.invalid_company');
        const lastRow = await readVerifiedTable();
        const values = lastRow === 3 ? [] : await readRows(`A4:A${lastRow}`, lastRow - 3, 1);
        const rows = values.flatMap((row, index) => normalizeCompanyId(row[0]) === normalized ? [index + 4] : []);
        if (!rows.length) throw new Error('support_users.company_not_found');
        return rows;
    }

    async function writeCompanyReferences(companyId, values) {
        const ranges = contiguousRanges(await readCompanyRows(companyId));
        for (const { first, last } of ranges) {
            await graphClient.api(`${WORKSHEET_PATH}/range(address='I${first}:K${last}')`)
                .middlewareOptions([new RetryHandlerOptions(0, 0, () => false)]).update({
                    values: Array.from({ length: last - first + 1 }, () => [...values]),
                });
        }
    }

    return {
        readUsers: () => readVerifiedUsers(),

        async writeBookingPermissions(companyId, employees, directory) {
            const company = normalizeCompanyId(companyId);
            if (!company || !Array.isArray(employees) || employees.length > 1000
                || employees.some((person) => !person || typeof person !== 'object' || Array.isArray(person)
                    || Object.keys(person).some((key) => !['email', 'allowed'].includes(key))
                    || typeof person.email !== 'string' || person.email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/u.test(person.email)
                    || typeof person.allowed !== 'boolean')) throw new Error('support_users.invalid_permission');
            const desired = new Map(employees.map((person) => [person.email.trim().toLowerCase(), person.allowed ? 'Sim' : 'Não']));
            if (desired.size !== employees.length) throw new Error('support_users.invalid_permission');
            // Reuse only the fresh, adapter-owned snapshot from this request's authorization.
            const rows = snapshots.get(directory) || snapshots.get(await readVerifiedUsers());
            const counts = new Map();
            for (const person of rows) counts.set(person.email, (counts.get(person.email) || 0) + 1);
            const eligible = rows.filter((person) => person.companyId === company && counts.get(person.email) === 1
                && /^[^\s@]+@[^\s@]+\.[^\s@]+$/u.test(person.email));
            if (eligible.length !== desired.size || eligible.some((person) => !desired.has(person.email))) throw new Error('support_users.employee_not_found');
            const changed = eligible.filter((person) => person.permission !== desired.get(person.email));
            const ranges = [];
            for (const person of changed) {
                const previous = ranges.at(-1);
                if (previous && previous.last + 1 === person.row) {
                    previous.last = person.row;
                    previous.values.push([desired.get(person.email)]);
                } else ranges.push({ first: person.row, last: person.row, values: [[desired.get(person.email)]] });
            }
            const noRetry = () => [new RetryHandlerOptions(0, 0, () => false)];
            if (ranges.length === 1) {
                await graphClient.api(`${WORKSHEET_PATH}/range(address='H${ranges[0].first}:H${ranges[0].last}')`)
                    .middlewareOptions(noRetry()).update({ values: ranges[0].values });
            } else for (let offset = 0; offset < ranges.length; offset += 20) {
                const requests = ranges.slice(offset, offset + 20).map((range, index) => ({
                    id: String(index + 1), method: 'PATCH',
                    url: `${WORKSHEET_PATH}/range(address='H${range.first}:H${range.last}')`,
                    ...(index ? { dependsOn: [String(index)] } : {}),
                    headers: { 'Content-Type': 'application/json' }, body: { values: range.values },
                }));
                const result = await graphClient.api('/$batch').middlewareOptions(noRetry()).post({ requests });
                if (!Array.isArray(result?.responses) || result.responses.length !== requests.length
                    || new Set(result.responses.map((response) => response.id)).size !== requests.length
                    || result.responses.some((response) => !requests.some((request) => request.id === response.id)
                        || !Number.isInteger(response.status) || response.status < 200 || response.status >= 300)) {
                    throw new Error('support_users.permission_save_failed');
                }
            }
        },

        async writeCompanyCard(company, _card, references) {
            if (!normalizeCompanyId(company)) throw new Error('support_users.invalid_company');
            validateCardReferences(references);
            await writeCompanyReferences(company, [references.customerId, references.cardId, references.environment]);
        },

        async clearCompanyCard(companyId) {
            await writeCompanyReferences(companyId, ['', '', '']);
        },

        async readCompanyCardReferences(companyId) {
            const ranges = contiguousRanges(await readCompanyRows(companyId));
            let references = null;
            for (const { first, last } of ranges) {
                const records = await readRows(`I${first}:K${last}`, last - first + 1, 3);
                for (const row of records) {
                    const values = row.map((value) => String(value ?? '').trim());
                    if (values.every((value) => !value)) continue;
                    const current = { customerId: values[0], cardId: values[1], environment: values[2] };
                    validateCardReferences(current);
                    if (references && (references.customerId !== current.customerId
                        || references.cardId !== current.cardId || references.environment !== current.environment)) {
                        throw new Error('support_users.conflicting_card_references');
                    }
                    references = current;
                }
            }
            return references;
        },
    };
}

function contiguousRanges(rows) {
    const ranges = [];
    for (const row of rows) {
        const previous = ranges.at(-1);
        if (previous && previous.last + 1 === row) previous.last = row;
        else ranges.push({ first: row, last: row });
    }
    return ranges;
}

function normalizeCompanyId(value) {
    if (typeof value !== 'string' || !/^[\d.\-/\s]+$/u.test(value)) return '';
    const digits = value.replace(/\D/gu, '');
    return /^\d{14}$/u.test(digits) ? digits : '';
}

function validateCardReferences(references) {
    if (!references || typeof references !== 'object' || Array.isArray(references)
        || Object.keys(references).some((key) => !['customerId', 'cardId', 'environment'].includes(key))
        || typeof references.customerId !== 'string' || typeof references.cardId !== 'string'
        || !/^cus_[A-Za-z0-9]{1,200}$/u.test(references.customerId)
        || !/^card_[A-Za-z0-9]{1,200}$/u.test(references.cardId)
        || references.environment !== 'sandbox') {
        throw new Error('support_users.invalid_card_references');
    }
}

module.exports = { createSupportUsersAdapter };
