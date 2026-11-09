'use strict';

const { RetryHandlerOptions } = require('@microsoft/microsoft-graph-client');
const { appointmentAvailability, isCalendarDate, addCalendarDays } = require('../domains/support-availability');

const WORKSHEET_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSFKTNUUJS3TZDKCLDRXFZCMVYB/workbook/worksheets/{00000000-0001-0000-0000-000000000000}';
const TABLE_PATH = `${WORKSHEET_PATH}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}`;
const COLUMNS = [
    'DATA', 'STATUS', 'OBSERVAÇÃO', 'CNPJ', 'NOME COMPLETO',
    'VALOR POR HORA', 'DURAÇÃO (MIN)', 'VALOR COBRADO',
];
const DAY_MILLISECONDS = 86400000;
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const AVAILABILITY_MARKER = 'APP: INDISPONÍVEL';
const CHARGE_FORMULAS = new Set([
    '=[@[DURAÇÃO (MIN)]]/60*[@[VALOR POR HORA]]',
    '=BD[[#This Row],[DURAÇÃO (MIN)]]/60*BD[[#This Row],[VALOR POR HORA]]',
]);

function createSupportAppointmentsAdapter({ graphClient, availabilityJournal }) {
    let operations = Promise.resolve();
    function exclusive(operation) {
        const result = operations.then(operation, operation);
        operations = result.catch(() => {});
        return result;
    }

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

    async function readSnapshotRows(address, count, firstRow) {
        const response = await graphClient.api(`${WORKSHEET_PATH}/range(address='${address}')?$select=values,formulas`).get();
        return snapshotRows(response, count, firstRow);
    }

    async function readSnapshot() {
        const response = await graphClient.api(`${TABLE_PATH}/range?$select=address,rowCount,columnCount,values,formulas`).get();
        const address = /^ATENDIMENTOS!A3:H([3-9]|[1-9]\d+)$/.exec(response.address);
        const lastRow = Number(address?.[1]);
        if (!address || response.columnCount !== COLUMNS.length || response.rowCount !== lastRow - 2
            || !Array.isArray(response.values) || response.values.length !== response.rowCount
            || !Array.isArray(response.values[0]) || response.values[0].length !== COLUMNS.length
            || response.values[0].some((value, index) => typeof value !== 'string' || value.trim() !== COLUMNS[index])
            || !Array.isArray(response.formulas) || response.formulas.length !== response.rowCount
            || !Array.isArray(response.formulas[0]) || response.formulas[0].length !== COLUMNS.length) {
            throw new Error('support_appointments.schema_mismatch');
        }
        const rows = snapshotRows({ values: response.values.slice(1), formulas: response.formulas.slice(1) }, lastRow - 3, 4);
        if (new Set(rows.map(({ date }) => date)).size !== rows.length) throw new Error('support_appointments.schema_mismatch');
        return rows;
    }

    async function syncAvailability({ today }) {
        if (!isCalendarDate(today)) throw new Error('support_appointments.invalid_date');
        if (!availabilityJournal || !['read', 'write', 'clear'].every((method) => typeof availabilityJournal[method] === 'function')) {
            throw new Error('support_appointments.availability_journal_required');
        }
        let rows = await readSnapshot();
        const intent = await availabilityJournal.read();
        if (intent) {
            if (!validIntent(intent) || !intent.dates.every((date) => rows.some((row) => row.date === date))) {
                throw new Error('support_appointments.availability_review_required');
            }
            await availabilityJournal.clear();
        }
        const present = new Set(rows.map(({ date }) => date));
        const missing = Array.from({ length: 30 }, (_, index) => addCalendarDays(today, index + 1)).filter((date) => !present.has(date));
        if (missing.length) {
            // A durable intent prevents replay when rows/add succeeds but its response is lost.
            await availabilityJournal.write({ dates: missing });
            await graphClient.api(`${TABLE_PATH}/rows/add`)
                .middlewareOptions([new RetryHandlerOptions(0, 0, () => false)])
                .post({ values: missing.map((date) => {
                    const available = appointmentAvailability(date, today);
                    return [(Date.parse(`${date}T00:00:00Z`) - EXCEL_EPOCH) / DAY_MILLISECONDS,
                        available ? '' : 'INDISPONÍVEL', available ? '' : AVAILABILITY_MARKER, '', '', '', null, null];
                }) });
            rows = await readSnapshot();
            if (!missing.every((date) => rows.some((row) => row.date === date))) {
                throw new Error('support_appointments.availability_review_required');
            }
            await availabilityJournal.clear();
        }
        const changed = rows.flatMap((row, index) => managedAvailabilityChange(row, today) ? [index] : []);
        if (changed.length) {
            const firstIndex = changed[0];
            const lastIndex = changed[changed.length - 1];
            const firstRow = rows[firstIndex].excelRow;
            const lastRow = rows[lastIndex].excelRow;
            const current = await readSnapshotRows(`A${firstRow}:H${lastRow}`, lastIndex - firstIndex + 1, firstRow);
            if (current.some((row, index) => row.date !== rows[firstIndex + index].date)) throw new Error('support_appointments.schema_mismatch');
            const values = current.map((row) => managedAvailabilityChange(row, today) || [null, null]);
            if (values.some((cells) => cells[0] !== null)) {
                await graphClient.api(`${WORKSHEET_PATH}/range(address='B${firstRow}:C${lastRow}')`)
                    .middlewareOptions([new RetryHandlerOptions(0, 0, () => false)])
                    .update({ values });
            }
            current.forEach((row, index) => { rows[firstIndex + index] = { ...row, status: values[index][0] ?? row.status }; });
        }
        return rows.map(({ date, status, excelRow }) => ({ date, status, excelRow }));
    }

    async function reserveDate({ date, companyId, name, hourlyRate }) {
        let selected;
        try {
            if (!isCalendarDate(date) || typeof companyId !== 'string' || !/^\d{14}$/u.test(companyId)
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
    }

    return {
        readDates,
        syncAvailability: (input) => exclusive(() => syncAvailability(input)),
        reserveDate: (input) => exclusive(() => reserveDate(input)),
    };
}

function snapshotRows(response, count, firstRow) {
    if (!Array.isArray(response.values) || response.values.length !== count
        || !Array.isArray(response.formulas) || response.formulas.length !== count
        || response.values.some((row) => !Array.isArray(row) || row.length !== 8)
        || response.formulas.some((row) => !Array.isArray(row) || row.length !== 8 || !CHARGE_FORMULAS.has(row[7]))) {
        throw new Error('support_appointments.schema_mismatch');
    }
    return response.values.map((cells, index) => ({
        date: excelDate(cells[0]), status: String(cells[1] ?? '').trim().toUpperCase(),
        excelRow: firstRow + index, cells,
    }));
}

function validIntent(intent) {
    return intent && typeof intent === 'object' && !Array.isArray(intent)
        && Object.keys(intent).length === 1 && Array.isArray(intent.dates) && intent.dates.length > 0 && intent.dates.length <= 30
        && intent.dates.every(isCalendarDate) && new Set(intent.dates).size === intent.dates.length;
}

function managedAvailabilityChange(row, today) {
    const note = String(row.cells[2] ?? '');
    const vacant = row.cells.slice(3, 7).every((value) => value === '' || value === null || value === undefined);
    if (!vacant || (row.status !== '' && row.status !== 'INDISPONÍVEL')
        || (note !== AVAILABILITY_MARKER && (note !== '' || row.status !== ''))) return null;
    const available = appointmentAvailability(row.date, today);
    const desired = available ? ['', ''] : ['INDISPONÍVEL', AVAILABILITY_MARKER];
    return row.status === desired[0] && note === desired[1] ? null : desired;
}

function excelDate(value) {
    if (typeof value !== 'number' || !Number.isInteger(value) || value < 61 || value > 2958465) {
        throw new Error('support_appointments.schema_mismatch');
    }
    return new Date(EXCEL_EPOCH + value * DAY_MILLISECONDS).toISOString().slice(0, 10);
}

function unavailable() {
    const error = new Error('support_appointments.date_unavailable');
    error.code = 'support_appointments.date_unavailable';
    error.definiteRejection = true;
    return error;
}

module.exports = { createSupportAppointmentsAdapter };
