/**
 * Retry operasi async dengan jeda tetap antar percobaan.
 */

const DEFAULT_MAX_ATTEMPTS = 4;
const DEFAULT_DELAY_MS = 1000;

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Error transient Google Drive / jaringan yang layak di-retry.
 * @param {unknown} err
 */
function isRetryableDriveError(err) {
    const msg = String(err?.message ?? err ?? "").toLowerCase();
    const code = String(err?.code ?? "").toUpperCase();
    const status = err?.response?.status ?? err?.status;

    if (
        ["ETIMEDOUT", "ESOCKETTIMEDOUT", "ECONNRESET", "EAI_AGAIN", "ENOTFOUND"].includes(
            code
        )
    ) {
        return true;
    }
    if (/time exceeded|timed out|timeout|socket hang up|network/i.test(msg)) {
        return true;
    }
    if ([429, 500, 502, 503, 504].includes(status)) {
        return true;
    }
    const errors =
        err?.response?.data?.error?.errors ??
        err?.errors ??
        [];
    if (errors.some(e => e.reason === "userRateLimitExceeded")) {
        return true;
    }
    if (/user rate limit exceeded|ratelimit/i.test(msg)) {
        return true;
    }
    return false;
}

/**
 * @template T
 * @param {() => Promise<T>} fn
 * @param {{ maxAttempts?: number, delayMs?: number, label?: string, retryIf?: (err: unknown) => boolean }} [options]
 * @returns {Promise<T>}
 */
async function withRetry(fn, options = {}) {
    const maxAttempts = options.maxAttempts ?? DEFAULT_MAX_ATTEMPTS;
    const delayMs = options.delayMs ?? DEFAULT_DELAY_MS;
    const label = options.label ?? "operation";
    const retryIf = options.retryIf ?? (() => true);

    let lastError;
    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
        try {
            return await fn();
        } catch (err) {
            lastError = err;
            if (!retryIf(err)) {
                throw err;
            }
            if (attempt >= maxAttempts) break;
            console.warn(
                `  [retry] ${label} gagal (percobaan ${attempt}/${maxAttempts}), ulang dalam ${delayMs}ms...`
            );
            await sleep(delayMs);
        }
    }
    throw lastError;
}

module.exports = {
    withRetry,
    isRetryableDriveError,
    sleep,
    DEFAULT_MAX_ATTEMPTS,
    DEFAULT_DELAY_MS
};
