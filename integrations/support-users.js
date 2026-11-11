'use strict';

const { RetryHandlerOptions } = require('@microsoft/microsoft-graph-client');

const WORKSHEET_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRXPEMKXY4F5RH34SKSN35K757F/workbook/worksheets/{00000000-0001-0000-0000-000000000000}';
const TABLE_PATH = `${WORKSHEET_PATH}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}`;
const COLUMNS = [
    'CNPJ', 'RAZÃO SOCIAL', 'VALOR REFERÊNCIA', 'NOME COMPLETO', 'E-MAIL', 'SENHA',
    'TIPO DE ACESSO', 'STATUS', 'PAGARME CUSTOMER ID', 'PAGARME CARD ID', 'PAGARME AMBIENTE',
];
const PERMISSION_COLUMN = 'CONTRATAR ATENDIMENTOS';

function createSupportUsersAdapter({ graphClient }) {
    let addingPermissionColumn;
    async function readVerifiedTable() {
        const [columns, range] = await Promise.all([
            graphClient.api(`${TABLE_PATH}/columns?$select=index,name`).get(),
            graphClient.api(`${TABLE_PATH}/range?$select=address,rowCount,columnCount`).get(),
        ]);
        const expected = columns.value?.length === COLUMNS.length + 1 ? [...COLUMNS, PERMISSION_COLUMN] : COLUMNS;
        if (!Array.isArray(columns.value) || columns.value.length !== expected.length
            || columns.value.some((column, index) => column?.index !== index
                || typeof column.name !== 'string' || column.name.trim() !== expected[index])) {
            throw new Error('support_users.schema_mismatch');
        }

        const address = /^USUÁRIOS!A3:([KL])([3-9]|[1-9]\d+)$/.exec(range.address);
        const lastRow = Number(address?.[2]);
        if (!address || address[1] !== (expected.length === 12 ? 'L' : 'K')
            || range.columnCount !== expected.length || range.rowCount !== lastRow - 2) {
            throw new Error('support_users.schema_mismatch');
        }
        return { lastRow, permissions: expected.length === 12 };
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
        const { lastRow, permissions } = await readVerifiedTable();
        if (lastRow === 3) return [];

        // Authentication never retrieves provider references from I:K.
        const [rows, allowed] = await Promise.all([
            readRows(`A4:H${lastRow}`, lastRow - 3, 8),
            permissions ? readRows(`L4:L${lastRow}`, lastRow - 3, 1) : [],
        ]);
        return rows.map((row, index) => {
            const companyId = normalizeCompanyId(row[0]);
            if (!companyId || typeof row[2] !== 'number' || !Number.isFinite(row[2]) || row[2] <= 0) {
                throw new Error('support_users.invalid_company_data');
            }
            return {
                companyId,
                company: String(row[1]),
                referenceRate: row[2],
                name: String(row[3]),
                email: String(row[4]),
                password: String(row[5]),
                role: String(row[6]),
                status: String(row[7]),
                bookingAllowed: allowed[index]?.[0] === 'Sim',
            };
        });
    }

    async function readCompanyRows(companyId) {
        const normalized = normalizeCompanyId(companyId);
        if (!normalized) throw new Error('support_users.invalid_company');
        const { lastRow } = await readVerifiedTable();
        const values = lastRow === 3 ? [] : await readRows(`A4:A${lastRow}`, lastRow - 3, 1);
        const rows = values.flatMap((row, index) => normalizeCompanyId(row[0]) === normalized ? [index + 4] : []);
        if (!rows.length) throw new Error('support_users.company_not_found');
        return rows;
    }

    async function ensurePermissionColumn() {
        if (!addingPermissionColumn) addingPermissionColumn = (async () => {
            if (!(await readVerifiedTable()).permissions) {
                const used = await graphClient.api(`${WORKSHEET_PATH}/usedRange(valuesOnly=true)?$select=address`).get();
                const address = /^(?:'USUÁRIOS'|USUÁRIOS)!([A-Z]+)([1-9]\d*)(?::([A-Z]+)([1-9]\d*))?$/u.exec(used.address);
                const endColumn = address?.[3] || address?.[1];
                // Preserve manual cells outside the table; an occupied extension needs review.
                if (!endColumn || endColumn.length > 1 || endColumn >= 'L') throw new Error('support_users.permission_column_occupied');
                await graphClient.api(`${TABLE_PATH}/columns`)
                    .middlewareOptions([new RetryHandlerOptions(0, 0, () => false)])
                    .post({ name: PERMISSION_COLUMN });
                if (!(await readVerifiedTable()).permissions) throw new Error('support_users.schema_mismatch');
            }
        })().finally(() => { addingPermissionColumn = undefined; });
        return addingPermissionColumn;
    }

    async function permissionRow(companyId, email) {
        const { lastRow, permissions } = await readVerifiedTable();
        const rows = lastRow === 3 ? [] : await readRows(`A4:H${lastRow}`, lastRow - 3, 8);
        const matching = rows.flatMap((row, index) => String(row[4]).trim().toLowerCase() === email ? [{ row, index: index + 4 }] : []);
        if (matching.length !== 1 || normalizeCompanyId(matching[0].row[0]) !== companyId || matching[0].row[7] !== 'Ativo') throw new Error('support_users.employee_not_found');
        return { row: matching[0].index, permissions };
    }

    return {
        readUsers: () => readVerifiedUsers(),

        async writeBookingPermission(companyId, email, allowed) {
            const company = normalizeCompanyId(companyId);
            if (!company || typeof email !== 'string' || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/u.test(email)
                || typeof allowed !== 'boolean') throw new Error('support_users.invalid_permission');
            const address = email.trim().toLowerCase();
            let target = await permissionRow(company, address);
            if (!target.permissions) {
                await ensurePermissionColumn();
                target = await permissionRow(company, address);
            }
            const { row } = target;
            await graphClient.api(`${WORKSHEET_PATH}/range(address='L${row}:L${row}')`)
                .middlewareOptions([new RetryHandlerOptions(0, 0, () => false)])
                .update({ values: [[allowed ? 'Sim' : 'Não']] });
        },

        async writeCompanyCard(company, _card, references) {
            if (!normalizeCompanyId(company)) throw new Error('support_users.invalid_company');
            validateCardReferences(references);
            const rows = await readCompanyRows(company);
            for (const row of rows) {
                await graphClient.api(`${WORKSHEET_PATH}/range(address='I${row}:K${row}')`).update({
                    values: [[references.customerId, references.cardId, references.environment]],
                });
            }
        },

        async clearCompanyCard(companyId) {
            const rows = await readCompanyRows(companyId);
            for (const row of rows) {
                await graphClient.api(`${WORKSHEET_PATH}/range(address='I${row}:K${row}')`).update({
                    values: [['', '', '']],
                });
            }
        },

        async readCompanyCardReferences(companyId) {
            const rows = await readCompanyRows(companyId);
            let references = null;
            for (const row of rows) {
                const [[customerId, cardId, environment]] = await readRows(`I${row}:K${row}`, 1, 3);
                const values = [customerId, cardId, environment].map((value) => String(value ?? '').trim());
                if (values.every((value) => !value)) continue;
                const current = { customerId: values[0], cardId: values[1], environment: values[2] };
                validateCardReferences(current);
                if (references && (references.customerId !== current.customerId
                    || references.cardId !== current.cardId || references.environment !== current.environment)) {
                    throw new Error('support_users.conflicting_card_references');
                }
                references = current;
            }
            return references;
        },
    };
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
