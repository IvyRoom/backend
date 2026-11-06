'use strict';

const WORKSHEET_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRXPEMKXY4F5RH34SKSN35K757F/workbook/worksheets/{00000000-0001-0000-0000-000000000000}';
const TABLE_PATH = `${WORKSHEET_PATH}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}`;
const COLUMNS = [
    'EMPRESA', 'NOME COMPLETO', 'PRIMEIRO NOME', 'E-MAIL', 'SENHA',
    'TIPO DE ACESSO', 'STATUS LOGIN', 'PAGARME CUSTOMER ID', 'PAGARME CARD ID', 'AMBIENTE PAGARME',
];

function createSupportUsersAdapter({ graphClient }) {
    async function readVerifiedUsers() {
        const columns = await graphClient.api(`${TABLE_PATH}/columns?$select=index,name`).get();
        if (!Array.isArray(columns.value) || columns.value.length !== COLUMNS.length
            || columns.value.some((column, index) => column.index !== index || column.name !== COLUMNS[index])) {
            throw new Error('support_users.schema_mismatch');
        }

        const range = await graphClient.api(`${TABLE_PATH}/range?$select=address,rowCount,columnCount`).get();
        const address = /^USUÁRIOS!A3:J([3-9]|[1-9]\d+)$/.exec(range.address);
        const lastRow = Number(address?.[1]);
        if (!address || range.columnCount !== COLUMNS.length || range.rowCount !== lastRow - 2) {
            throw new Error('support_users.schema_mismatch');
        }
        if (lastRow === 3) return [];

        // Authentication never retrieves provider references from H:J.
        const response = await graphClient.api(`${WORKSHEET_PATH}/range(address='A4:G${lastRow}')?$select=values`).get();
        if (!Array.isArray(response.values) || response.values.length !== lastRow - 3
            || response.values.some((row) => !Array.isArray(row) || row.length !== 7)) {
            throw new Error('support_users.schema_mismatch');
        }

        return response.values.map((row) => ({
            company: String(row[0]),
            name: String(row[1]),
            email: String(row[3]),
            password: String(row[4]),
            role: String(row[5]),
            status: String(row[6]),
        }));
    }

    return {
        readUsers: () => readVerifiedUsers(),

        async writeCompanyCard(company, _card, references) {
            const companyName = normalizeCompany(company);
            if (!companyName) throw new Error('support_users.invalid_company');
            if (!references || typeof references !== 'object' || Array.isArray(references)
                || Object.keys(references).some((key) => !['customerId', 'cardId', 'environment'].includes(key))
                || typeof references.customerId !== 'string' || typeof references.cardId !== 'string'
                || !/^cus_[A-Za-z0-9]{1,200}$/u.test(references.customerId || '')
                || !/^card_[A-Za-z0-9]{1,200}$/u.test(references.cardId || '')
                || references.environment !== 'sandbox') {
                throw new Error('support_users.invalid_card_references');
            }
            const users = await readVerifiedUsers();
            const rows = users.flatMap((user, index) => normalizeCompany(user.company) === companyName ? [index + 4] : []);
            if (!rows.length) throw new Error('support_users.company_not_found');
            for (const row of rows) {
                await graphClient.api(`${WORKSHEET_PATH}/range(address='H${row}:J${row}')`).update({
                    values: [[references.customerId, references.cardId, references.environment]],
                });
            }
        },
    };
}

function normalizeCompany(value) {
    return typeof value === 'string' ? value.trim().replace(/\s+/gu, ' ').toLowerCase() : '';
}

module.exports = { createSupportUsersAdapter };
