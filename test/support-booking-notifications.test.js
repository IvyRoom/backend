'use strict';

require('./deny-production-network');

const test = require('node:test');
const assert = require('node:assert/strict');
const { RetryHandlerOptions } = require('@microsoft/microsoft-graph-client');
const { createSupportBookingNotifications } = require('../integrations/support-booking-notifications');

const USER_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be';

function fakeClient(response, rejection) {
    const requests = [];
    return {
        requests,
        api(path) {
            const request = { path, options: null, body: null, attempts: 0 };
            requests.push(request);
            return {
                middlewareOptions(options) {
                    request.options = options;
                    return this;
                },
                async post(body) {
                    request.body = body;
                    request.attempts += 1;
                    if (rejection) throw rejection;
                    return response;
                },
            };
        },
    };
}

function assertSingleAttempt(request) {
    assert.equal(request.attempts, 1);
    assert.equal(request.options.length, 1);
    const [options] = request.options;
    assert.ok(options instanceof RetryHandlerOptions);
    assert.equal(options.maxRetries, 0);
    assert.equal(options.delay, 0);
    assert.equal(options.shouldRetry(0, 0, {}, {}, { status: 503 }), false);
}

function assertFixedFailure(code) {
    return (error) => {
        assert.equal(error.message, 'Support booking notification failed');
        assert.equal(error.code, `support_booking_notifications.${code}`);
        assert.deepEqual(Object.keys(error), ['code']);
        assert.equal(error.cause, undefined);
        return true;
    };
}

test('booking notification factory does not access its client during construction', () => {
    const graphClient = new Proxy({}, {
        get() { throw new Error('client accessed during construction'); },
    });
    const adapter = createSupportBookingNotifications({ graphClient });
    assert.equal(typeof adapter.createEvent, 'function');
    assert.equal(typeof adapter.sendConfirmation, 'function');
});

test('calendar creation preserves the domain payload and projects only bounded provider references', async () => {
    const graphClient = fakeClient({
        id: 'AAMkAGFhZC8-AAAB=',
        iCalUId: '040000008200E00074C5B7101A82E00800000000@outlook.com',
        body: { content: 'private provider content' },
        attendees: [{ emailAddress: { address: 'private@example.invalid' } }],
    });
    const adapter = createSupportBookingNotifications({ graphClient });
    const payload = {
        subject: 'Atendimento consultivo',
        transactionId: 'd0b07b46-1311-41fb-984f-b338fd6b8a16',
        isAllDay: true,
        start: { dateTime: '2026-11-12T00:00:00', timeZone: 'E. South America Standard Time' },
        end: { dateTime: '2026-11-13T00:00:00', timeZone: 'E. South America Standard Time' },
        attendees: [{ emailAddress: { address: 'user@example.invalid' }, type: 'required' }],
    };

    assert.deepEqual(await adapter.createEvent(payload), {
        eventId: 'AAMkAGFhZC8-AAAB=',
        iCalUid: '040000008200E00074C5B7101A82E00800000000@outlook.com',
    });
    assert.equal(graphClient.requests.length, 1);
    const [request] = graphClient.requests;
    assert.equal(request.path, `${USER_PATH}/events`);
    assert.equal(request.body, payload);
    assert.equal(request.body.transactionId, payload.transactionId);
    assertSingleAttempt(request);
});

test('confirmation mail preserves the message and exact sendMail envelope without exposing its response', async () => {
    const graphClient = fakeClient({ private: 'unexpected provider response' });
    const adapter = createSupportBookingNotifications({ graphClient });
    const message = {
        subject: 'Data reservada',
        body: { contentType: 'HTML', content: '<p>Reserva confirmada.</p>' },
        toRecipients: [
            { emailAddress: { address: 'contato@machadogestao.com' } },
            { emailAddress: { address: 'user@example.invalid' } },
        ],
        attachments: [{
            '@odata.type': '#microsoft.graph.fileAttachment',
            name: 'atendimento.ics',
            contentType: 'text/calendar',
            contentBytes: Buffer.from('BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n').toString('base64'),
        }],
    };

    assert.equal(await adapter.sendConfirmation(message), undefined);
    assert.equal(graphClient.requests.length, 1);
    const [request] = graphClient.requests;
    assert.equal(request.path, `${USER_PATH}/sendMail`);
    assert.deepEqual(request.body, { message, saveToSentItems: true });
    assert.equal(request.body.message, message);
    assertSingleAttempt(request);
});

test('notification operations sanitize SDK errors and make no application retry', async (t) => {
    for (const operation of ['createEvent', 'sendConfirmation']) {
        await t.test(operation, async () => {
            const rejection = Object.assign(new Error('private body and credentials'), {
                statusCode: 503,
                body: { private: 'sensitive provider data' },
            });
            const graphClient = fakeClient(null, rejection);
            const adapter = createSupportBookingNotifications({ graphClient });
            await assert.rejects(adapter[operation]({ transactionId: 'booking-id' }), assertFixedFailure('request_failed'));
            assert.equal(graphClient.requests.length, 1);
            assertSingleAttempt(graphClient.requests[0]);
        });
    }
});

test('request-construction failures are sanitized without attempting a different request', async () => {
    let attempts = 0;
    const adapter = createSupportBookingNotifications({ graphClient: {
        api() {
            attempts += 1;
            throw new Error('private SDK configuration');
        },
    } });
    await assert.rejects(adapter.createEvent({ transactionId: 'booking-id' }), assertFixedFailure('request_failed'));
    assert.equal(attempts, 1);
});

test('calendar projection rejects absent, oversized, and unsafe identifiers', async (t) => {
    const invalid = [undefined, null, false, 42, '', 'with space', 'with\nnewline', 'with\u007fdelete',
        'with\u0085control', 'with\u200bformat', 'with\u2028separator'];
    for (const field of ['id', 'iCalUId']) {
        for (const value of [...invalid, 'x'.repeat(field === 'id' ? 2049 : 1025)]) {
            await t.test(`${field}: ${JSON.stringify(value)?.slice(0, 32)}`, async () => {
                const graphClient = fakeClient({ id: 'event-id', iCalUId: 'calendar-uid', [field]: value });
                const adapter = createSupportBookingNotifications({ graphClient });
                await assert.rejects(adapter.createEvent({ transactionId: 'booking-id' }), assertFixedFailure('invalid_response'));
                assert.equal(graphClient.requests.length, 1);
                assertSingleAttempt(graphClient.requests[0]);
            });
        }
    }
    for (const response of [undefined, null, []]) {
        const adapter = createSupportBookingNotifications({ graphClient: fakeClient(response) });
        await assert.rejects(adapter.createEvent({ transactionId: 'booking-id' }), assertFixedFailure('invalid_response'));
    }
});

test('calendar response property errors are sanitized after one SDK attempt', async () => {
    const response = Object.defineProperty({}, 'id', {
        get() { throw new Error('private provider response'); },
    });
    const graphClient = fakeClient(response);
    const adapter = createSupportBookingNotifications({ graphClient });
    await assert.rejects(adapter.createEvent({ transactionId: 'booking-id' }), assertFixedFailure('invalid_response'));
    assertSingleAttempt(graphClient.requests[0]);
});

test('invalid payload containers fail before accessing the SDK', async (t) => {
    const graphClient = fakeClient({ id: 'event-id', iCalUId: 'calendar-uid' });
    const adapter = createSupportBookingNotifications({ graphClient });
    for (const operation of ['createEvent', 'sendConfirmation']) {
        await t.test(operation, async () => {
            for (const value of [undefined, null, [], '', 42, true]) {
                await assert.rejects(adapter[operation](value), assertFixedFailure('invalid_request'));
            }
        });
    }
    assert.equal(graphClient.requests.length, 0);
});
