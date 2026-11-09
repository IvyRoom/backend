'use strict';

const { RetryHandlerOptions } = require('@microsoft/microsoft-graph-client');

const USER_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be';

function createSupportBookingNotifications({ graphClient }) {
    async function postOnce(path, body) {
        try {
            // A timed-out write may already have succeeded; only the booking domain may recover it.
            return await graphClient.api(path)
                .middlewareOptions([new RetryHandlerOptions(0, 0, () => false)])
                .post(body);
        } catch {
            throw failure('request_failed');
        }
    }

    return {
        async createEvent(eventPayload) {
            requirePayload(eventPayload);
            const response = await postOnce(`${USER_PATH}/events`, eventPayload);
            try {
                const eventId = response?.id;
                const iCalUid = response?.iCalUId;
                if (!validId(eventId, 2048) || !validId(iCalUid, 1024)) {
                    throw failure('invalid_response');
                }
                const meetingUrl = response?.onlineMeeting?.joinUrl;
                if (meetingUrl !== undefined && meetingUrl !== null) {
                    if (typeof meetingUrl !== 'string' || meetingUrl.length > 4096 || /[\p{C}\s]/u.test(meetingUrl)) throw failure('invalid_response');
                    const url = new URL(meetingUrl);
                    if (url.protocol !== 'https:' || url.username || url.password
                        || !['teams.microsoft.com', 'teams.live.com', 'teams.cloud.microsoft'].includes(url.hostname)) throw failure('invalid_response');
                    return { eventId, iCalUid, meetingUrl };
                }
                return { eventId, iCalUid };
            } catch {
                throw failure('invalid_response');
            }
        },

        async sendConfirmation(message) {
            requirePayload(message);
            await postOnce(`${USER_PATH}/sendMail`, { message, saveToSentItems: true });
        },
    };
}

function requirePayload(value) {
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
        throw failure('invalid_request');
    }
}

function validId(value, maximumLength) {
    return typeof value === 'string' && value.length > 0 && value.length <= maximumLength &&
        !/[\p{C}\s]/u.test(value);
}

function failure(code) {
    const error = new Error('Support booking notification failed');
    error.code = `support_booking_notifications.${code}`;
    return error;
}

module.exports = { createSupportBookingNotifications };
