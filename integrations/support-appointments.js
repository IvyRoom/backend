'use strict';

const { RetryHandlerOptions } = require('@microsoft/microsoft-graph-client');

const WORKSHEET_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSFKTNUUJS3TZDKCLDRXFZCMVYB/workbook/worksheets/{00000000-0001-0000-0000-000000000000}';
const TABLE_PATH = `${WORKSHEET_PATH}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}`;
const COLUMNS = [
    'DATA', 'STATUS', 'OBSERVAÇÃO', 'CNPJ', 'NOME COMPLETO',
    'VALOR POR HORA', 'DURAÇÃO (MIN)', 'VALOR COBRADO',
];
const DAY_MILLISECONDS = 86400000;
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

function createSupportAppointmentsAdapter({ graphClient }) {
    async function readVerifiedTable() {
        const columns = await graphClient.api(`${TABLE_PATH}/columns?$select=index,name`).get();
        if (!Array.isArray(columns.value) || columns.value.length !== COLUMNS.length
            || columns.value.some((column, index) => column?.index !== index
                || typeof column.name !== 'string' || column.name.trim() !== COLUMNS[index])) {
            throw new Error('support_appointments.schema_mismatch');
        }
        const range = await graphClient.api(`${TABLE_PATH}/range?$select=address,rowCount,columnCount`).get();
        const address = /^ATENDIMENTOS!A3:H([3-9]|[1-9]\d+)$/.exec(range.address);
        const lastRow = Number(address?.[1]);
        if (!address || range.columnCount !== COLUMNS.length || range.rowCount !== lastRow - 2) {
            throw new Error('support_appointments.schema_mismatch');
        }
        return lastRow;
    }

    async function readDateRows(address, count) {
        const response = await graphClient.api(`${WORKSHEET_PATH}/range(address='${address}')?$select=values`).get();
        if (!Array.isArray(response.values) || response.values.length !== count
            || response.values.some((row) => !Array.isArray(row) || row.length !== 2)) {
            throw new Error('support_appointments.schema_mismatch');
        }
        return response.values.map((row) => ({ date: excelDate(row[0]), status: String(row[1] ?? '').trim().toUpperCase() }));
    }

    async function readDates() {
        const lastRow = await readVerifiedTable();
        if (lastRow === 3) return [];
        const dates = await readDateRows(`A4:B${lastRow}`, lastRow - 3);
        if (new Set(dates.map(({ date }) => date)).size !== dates.length) {
            throw new Error('support_appointments.schema_mismatch');
        }
        return dates.map((entry, index) => ({ ...entry, excelRow: index + 4 }));
    }

    return {
        readDates,

        async reserveDate({ date, companyId, name, hourlyRate }) {
            let selected;
            try {
                if (!isDate(date) || typeof companyId !== 'string' || !/^\d{14}$/u.test(companyId)
                    || typeof name !== 'string' || !name.trim() || name.length > 300 || /^[=+\-@]/u.test(name.trim())
                    || typeof hourlyRate !== 'number' || !Number.isFinite(hourlyRate) || hourlyRate <= 0) {
                    throw new Error('support_appointments.invalid_reservation');
                }
                const dates = await readDates();
                selected = dates.find((entry) => entry.date === date);
                if (!selected || selected.status) throw unavailable();
                const [current] = await readDateRows(`A${selected.excelRow}:B${selected.excelRow}`, 1);
                if (current.date !== date || current.status) throw unavailable();
            } catch (error) {
                error.definiteRejection = true;
                throw error;
            }

            const formattedCnpj = companyId.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/u, '$1.$2.$3/$4-$5');
            // Null leaves the existing date, notes, duration and calculated charge untouched.
            await graphClient.api(`${WORKSHEET_PATH}/range(address='A${selected.excelRow}:H${selected.excelRow}')`)
                .middlewareOptions([new RetryHandlerOptions(0, 0, () => false)])
                .update({ values: [[null, 'RESERVADO', null, formattedCnpj, name.trim(), hourlyRate, null, null]] });
            return { date: selected.date, status: 'RESERVADO', excelRow: selected.excelRow };
        },
    };
}

function excelDate(value) {
    if (typeof value !== 'number' || !Number.isInteger(value) || value < 61 || value > 2958465) {
        throw new Error('support_appointments.schema_mismatch');
    }
    return new Date(EXCEL_EPOCH + value * DAY_MILLISECONDS).toISOString().slice(0, 10);
}

function isDate(value) {
    return typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/u.test(value)
        && !Number.isNaN(Date.parse(`${value}T00:00:00Z`))
        && new Date(`${value}T00:00:00Z`).toISOString().slice(0, 10) === value;
}

function unavailable() {
    const error = new Error('support_appointments.date_unavailable');
    error.code = 'support_appointments.date_unavailable';
    error.definiteRejection = true;
    return error;
}

module.exports = { createSupportAppointmentsAdapter };
