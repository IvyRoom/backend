'use strict';

function createRetry({ sleep }) {
    return async function retry(fn, retries = 5) {
        for (let i = 0; i < retries; i++) {
            try { return await fn(); }
            catch (err) {
                if (i === retries - 1) throw err;
                await sleep(500 * (i + 1));
            }
        }
    };
}

module.exports = { createRetry };
