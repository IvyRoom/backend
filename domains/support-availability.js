'use strict';

const DAY_MILLISECONDS = 86400000;
// Federal Laws 662/1949, 6.802/1980 and 14.759/2023; Curitiba Law 3.015/1967.
const FIXED_HOLIDAYS = new Map([
    ['01-01', 'Confraternização Universal'], ['04-21', 'Tiradentes'], ['05-01', 'Dia do Trabalho'],
    ['09-07', 'Independência do Brasil'], ['09-08', 'Nossa Senhora da Luz dos Pinhais'],
    ['10-12', 'Nossa Senhora Aparecida'], ['11-02', 'Finados'], ['11-15', 'Proclamação da República'],
    ['11-20', 'Dia Nacional de Zumbi e da Consciência Negra'], ['12-25', 'Natal'],
]);

function isCalendarDate(value) {
    return typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/u.test(value)
        && Number.isFinite(Date.parse(`${value}T00:00:00Z`))
        && new Date(`${value}T00:00:00Z`).toISOString().slice(0, 10) === value;
}

function addCalendarDays(date, days) {
    if (!isCalendarDate(date) || !Number.isInteger(days)) throw new Error('support_availability.invalid_date');
    const result = new Date(Date.parse(`${date}T00:00:00Z`) + days * DAY_MILLISECONDS).toISOString().slice(0, 10);
    if (!isCalendarDate(result)) throw new Error('support_availability.invalid_date');
    return result;
}

function bookingCalendarRange(today) {
    if (!isCalendarDate(today)) throw new Error('support_availability.invalid_date');
    const horizon = addCalendarDays(today, 30);
    const month = new Date(`${horizon}T00:00:00Z`);
    month.setUTCMonth(month.getUTCMonth() + 1, 0);
    return { first: `${today.slice(0, 7)}-01`, last: month.toISOString().slice(0, 10) };
}

function easter(year) {
    // Gregorian computus published by the U.S. Naval Observatory: https://aa.usno.navy.mil/faq/easter
    const integer = Math.trunc;
    const century = integer(year / 100);
    const golden = year - 19 * integer(year / 19);
    const correction = integer((century - 17) / 25);
    let epact = century - integer(century / 4) - integer((century - correction) / 3) + 19 * golden + 15;
    epact -= 30 * integer(epact / 30);
    epact -= integer(epact / 28) * (1 - integer(epact / 28) * integer(29 / (epact + 1)) * integer((21 - golden) / 11));
    let weekday = year + integer(year / 4) + epact + 2 - century + integer(century / 4);
    weekday -= 7 * integer(weekday / 7);
    const offset = epact - weekday;
    const month = 3 + integer((offset + 40) / 44);
    const day = offset + 28 - 31 * integer(month / 4);
    return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function holidayReason(date) {
    if (!isCalendarDate(date)) return null;
    const fixed = FIXED_HOLIDAYS.get(date.slice(5));
    if (fixed) return fixed;
    const easterDate = easter(Number(date.slice(0, 4)));
    if (date === addCalendarDays(easterDate, -2)) return 'Sexta-feira da Paixão';
    if (date === addCalendarDays(easterDate, 60)) return 'Corpus Christi';
    return null;
}

function appointmentAvailability(date, today) {
    if (!isCalendarDate(today)) throw new Error('support_availability.invalid_date');
    if (!isCalendarDate(date) || date <= today || date > addCalendarDays(today, 30)) return false;
    return ![0, 6].includes(new Date(`${date}T12:00:00Z`).getUTCDay()) && !holidayReason(date);
}

module.exports = { appointmentAvailability, bookingCalendarRange, holidayReason, isCalendarDate, addCalendarDays };
