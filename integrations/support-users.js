'use strict';

const WORKSHEET_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRXPEMKXY4F5RH34SKSN35K757F/workbook/worksheets/{00000000-0001-0000-0000-000000000000}';
const TABLE_PATH = `${WORKSHEET_PATH}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}`;
const COLUMNS = [
    'EMPRESA', 'NOME COMPLETO', 'PRIMEIRO NOME', 'E-MAIL', 'SENHA',
    'TIPO DE ACESSO', 'STATUS LOGIN', 'NÚMERO', 'TITULAR', 'MÊS', 'ANO', 'CVV', 'STATUS CARTÃO',
];

function createSupportUsersAdapter({ graphClient }) {
    async function readUsers() {
        const columns = await graphClient.api(`${TABLE_PATH}/columns?$select=index,name`).get();
        if (!Array.isArray(columns.value) || columns.value.length !== COLUMNS.length
            || columns.value.some((column, index) => column.index !== index || column.name !== COLUMNS[index])) {
            throw new Error('support_users.schema_mismatch');
        }

        const range = await graphClient.api(`${TABLE_PATH}/range?$select=address,rowCount,columnCount`).get();
        const address = /^USUÁRIOS!A3:M([3-9]|[1-9]\d+)$/.exec(range.address);
        const lastRow = Number(address?.[1]);
        if (!address || range.columnCount !== COLUMNS.length || range.rowCount !== lastRow - 2) {
            throw new Error('support_users.schema_mismatch');
        }
        if (lastRow === 3) return [];

        // The workbook also has legacy card columns; never retrieve their values.
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
        readUsers,

        async writeCompanyCard(company, card) {
            const companyName = normalizeCompany(company);
            const holderName = typeof card?.holderName === 'string' ? card.holderName.trim() : '';
            const year = card?.expYear < 100 ? 2000 + card.expYear : card?.expYear;
            if (!companyName || !/^\d{4}$/u.test(card?.lastFour || '')
                || !holderName || holderName.length > 64 || /[\u0000-\u001f\u007f]/u.test(holderName)
                || /^[=+\-@]/u.test(holderName)
                || !Number.isInteger(card?.expMonth) || card.expMonth < 1 || card.expMonth > 12
                || !Number.isInteger(card?.expYear) || card.expYear < 0
                || !Number.isInteger(year) || year < 2000 || year > 9999) {
                throw new Error('support_users.invalid_card_metadata');
            }
            const users = await readUsers();
            const rows = users.flatMap((user, index) => normalizeCompany(user.company) === companyName ? [index + 4] : []);
            if (!rows.length) throw new Error('support_users.company_not_found');
            for (const row of rows) {
                await graphClient.api(`${WORKSHEET_PATH}/range(address='H${row}:K${row}')`).update({
                    values: [[`**** **** **** ${card.lastFour}`, holderName, card.expMonth, year]],
                });
                // Keep the legacy CVV column L outside both update ranges.
                await graphClient.api(`${WORKSHEET_PATH}/range(address='M${row}')`).update({ values: [['Cadastrado']] });
            }
        },
    };
}

function normalizeCompany(value) {
    return typeof value === 'string' ? value.trim().replace(/\s+/gu, ' ').toLowerCase() : '';
}

module.exports = { createSupportUsersAdapter };
