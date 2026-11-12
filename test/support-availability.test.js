'use strict';

require('./deny-production-network');
const test = require('node:test');
const assert = require('node:assert/strict');
const { appointmentAvailability, bookingCalendarRange, holidayReason, addCalendarDays } = require('../domains/support-availability');

test('availability starts tomorrow and continues through the end of next month, excluding weekends and holidays', () => {
    const today = '2026-11-10';
    for (const date of ['2026-11-09', today, '2026-11-14', '2026-11-15', '2026-11-20', '2026-12-25', '2027-01-04']) {
        assert.equal(appointmentAvailability(date, today), false, date);
    }
    for (const date of ['2026-11-11', '2026-12-09', '2026-12-11', '2026-12-31']) {
        assert.equal(appointmentAvailability(date, today), true, date);
    }
});

test('visible calendar always includes exactly the current and next month across short months and year boundaries', () => {
    assert.deepEqual(bookingCalendarRange('2026-11-09'), { first: '2026-11-01', last: '2026-12-31' });
    assert.deepEqual(bookingCalendarRange('2026-11-01'), { first: '2026-11-01', last: '2026-12-31' });
    assert.deepEqual(bookingCalendarRange('2026-02-01'), { first: '2026-02-01', last: '2026-03-31' });
    assert.deepEqual(bookingCalendarRange('2026-10-01'), { first: '2026-10-01', last: '2026-11-30' });
    assert.deepEqual(bookingCalendarRange('2026-01-31'), { first: '2026-01-01', last: '2026-02-28' });
    assert.deepEqual(bookingCalendarRange('2028-01-31'), { first: '2028-01-01', last: '2028-02-29' });
    assert.deepEqual(bookingCalendarRange('2026-12-20'), { first: '2026-12-01', last: '2027-01-31' });
    assert.deepEqual(bookingCalendarRange('2026-12-31'), { first: '2026-12-01', last: '2027-01-31' });
});

test('month windows allow day thirty-one and leap days while rejecting dates in the following month', () => {
    assert.equal(appointmentAvailability('2026-12-11', '2026-11-10'), true);
    assert.equal(appointmentAvailability('2026-12-31', '2026-11-10'), true);
    assert.equal(appointmentAvailability('2027-01-04', '2026-11-10'), false);
    assert.equal(appointmentAvailability('2026-03-31', '2026-02-01'), true);
    assert.equal(appointmentAvailability('2026-04-01', '2026-02-01'), false);
    assert.equal(appointmentAvailability('2028-02-29', '2028-01-10'), true);
    assert.equal(appointmentAvailability('2028-03-01', '2028-01-10'), false);
    assert.equal(appointmentAvailability('2027-01-29', '2026-12-20'), true);
    assert.equal(appointmentAvailability('2027-02-01', '2026-12-20'), false);
});

test('national and Curitiba legal holidays are unavailable, including moving religious dates', () => {
    for (const date of ['2026-01-01', '2026-04-03', '2026-04-21', '2026-05-01', '2026-06-04', '2026-09-07', '2026-09-08', '2026-10-12', '2026-11-02', '2026-11-15', '2026-11-20', '2026-12-25', '2027-03-26', '2027-05-27', '2028-04-14', '2028-06-15']) {
        assert.equal(typeof holidayReason(date), 'string', date);
        assert.equal(appointmentAvailability(date, addCalendarDays(date, -1)), false, date);
    }
    assert.equal(holidayReason('2026-09-08'), 'Nossa Senhora da Luz dos Pinhais');
    assert.equal(holidayReason('2026-06-04'), 'Corpus Christi');
});

test('optional public-service closure dates do not become private-service holidays', () => {
    for (const date of ['2026-02-16', '2026-02-17', '2026-02-18', '2026-04-02', '2026-06-05', '2026-10-28', '2028-12-19', '2026-12-24', '2026-12-31']) {
        assert.equal(holidayReason(date), null, date);
        assert.equal(appointmentAvailability(date, addCalendarDays(date, -1)), true, date);
    }
});

test('invalid dates fail closed and calendar arithmetic preserves leap days and year boundaries', () => {
    for (const value of ['2026-02-30', '2026-11-09T12:00:00Z', '', null]) {
        assert.equal(appointmentAvailability(value, '2026-11-09'), false);
        assert.equal(holidayReason(value), null);
        assert.throws(() => bookingCalendarRange(value), /invalid_date/u);
    }
    assert.equal(addCalendarDays('2028-02-28', 1), '2028-02-29');
    assert.equal(addCalendarDays('2026-12-31', 1), '2027-01-01');
});
