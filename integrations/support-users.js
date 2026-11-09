'use strict';

const WORKSHEET_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRXPEMKXY4F5RH34SKSN35K757F/workbook/worksheets/{00000000-0001-0000-0000-000000000000}';
const TABLE_PATH = `${WORKSHEET_PATH}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}`;
const COLUMNS = [
    'CNPJ', 'RAZÃO SOCIAL', 'VALOR REFERÊNCIA', 'NOME COMPLETO', 'E-MAIL', 'SENHA',
    'TIPO DE ACESSO', 'STATUS', 'PAGARME CUSTOMER ID', 'PAGARME CARD ID', 'PAGARME AMBIENTE',
];

function createSupportUsersAdapter({ graphClient }) {
    async function readVerifiedTable() {
        const [columns, range] = await Promise.all([
            graphClient.api(`${TABLE_PATH}/columns?$select=index,name`).get(),
            graphClient.api(`${TABLE_PATH}/range?$select=address,rowCount,columnCount`).get(),
        ]);
        if (!Array.isArray(columns.value) || columns.value.length !== COLUMNS.length
            || columns.value.some((column, index) => column?.index !== index
                || typeof column.name !== 'string' || column.name.trim() !== COLUMNS[index])) {
            throw new Error('support_users.schema_mismatch');
        }

        const address = /^USUÁRIOS!A3:K([3-9]|[1-9]\d+)$/.exec(range.address);
        const lastRow = Number(address?.[1]);
        if (!address || range.columnCount !== COLUMNS.length || range.rowCount !== lastRow - 2) {
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
        if (lastRow === 3) return [];

        // Authentication never retrieves provider references from I:K.
        const rows = await readRows(`A4:H${lastRow}`, lastRow - 3, 8);
        return rows.map((row) => {
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
            };
        });
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

    return {
        readUsers: () => readVerifiedUsers(),

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
