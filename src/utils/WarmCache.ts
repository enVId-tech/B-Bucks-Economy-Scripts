/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran. All rights reserved.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 *
 * Optional background cache warming. CacheService entries expire (SERVER_SIDE_CACHE_AGE), so the first dialog opened after a
 * quiet period would otherwise pay the full sheet-read cost. A time-driven trigger calls warmCaches() every few minutes so a
 * cold cache is refilled before anyone is waiting for it.
 *
 * warmCaches() is idempotent and cheap: it calls the SAME cached fetchers the dialogs use, WITHOUT forceRefresh, so anything
 * that is still warm is returned from cache and only missing pieces are read from the sheets. It never changes sheet data.
 */

/** Name of the function the trigger runs. Must match the function below. */
const WARM_CACHE_HANDLER: string = "warmCaches";
const WARM_CACHE_INTERVAL_MINUTES: number = 10;

/** Refills any cache that has expired. Safe to run at any time. */
function warmCaches(): boolean { return runTrusted_(warmCachesEntry_); }

function warmCachesEntry_(): boolean {
    const startedAt = Date.now();
    try {
        // Skip while another run (or a real user action) holds the script lock: never compete with the admin.
        const lock = LockService.getScriptLock();
        if (!lock.tryLock(500)) {
            log("warmCaches skipped: another operation is running.", false);
            return false;
        }
        try {
            fetchSettingsDataCached(JSON.stringify({ forceRefresh: false }));
            fetchServicesDataCached(JSON.stringify({ forceRefresh: false }));
            fetchInvestmentsDataCached(JSON.stringify({ forceRefresh: false }));
            // First page of the log; the rolling cache only reads blocks that are missing
            fetchTransactionsPage(JSON.stringify({ page: 1, pageSize: 25, direction: "jump", includeStats: true }));
        } finally {
            lock.releaseLock();
        }
        log(`warmCaches finished in ${Date.now() - startedAt}ms.`, false);
        return true;
    } catch (error: any) {
        log(`Error in warmCaches: ${error.message}`, false);
        return false;
    }
}

/** Triggers currently installed for warmCaches. */
function getWarmCacheTriggers(): GoogleAppsScript.Script.Trigger[] {
    requireEditor_();
    return ScriptApp.getProjectTriggers().filter(trigger => trigger.getHandlerFunction() === WARM_CACHE_HANDLER);
}

/** True when background warming is installed. */
function isCacheWarmingEnabled(): boolean {
    requireEditor_();
    return getWarmCacheTriggers().length > 0;
}

/**
 * Installs the background trigger (replacing any previous copy so there is never more than one).
 * @returns A short human-readable status message.
 */
function enableCacheWarming(): string {
    requireEditor_();
    try {
        getWarmCacheTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
        ScriptApp.newTrigger(WARM_CACHE_HANDLER).timeBased().everyMinutes(WARM_CACHE_INTERVAL_MINUTES).create();
        warmCaches();
        return `Background cache warming is ON (every ${WARM_CACHE_INTERVAL_MINUTES} minutes).`;
    } catch (error: any) {
        log(`Error in enableCacheWarming: ${error.message}`, true);
        return `Could not enable cache warming: ${error.message}`;
    }
}

/**
 * Removes the background trigger.
 * @returns A short human-readable status message.
 */
function disableCacheWarming(): string {
    requireEditor_();
    try {
        const existing = getWarmCacheTriggers();
        existing.forEach(trigger => ScriptApp.deleteTrigger(trigger));
        return existing.length > 0 ? "Background cache warming is OFF." : "Background cache warming was already off.";
    } catch (error: any) {
        log(`Error in disableCacheWarming: ${error.message}`, true);
        return `Could not disable cache warming: ${error.message}`;
    }
}
