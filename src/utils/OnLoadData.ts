// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. All rights reserved.
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains utility functions for data caching and dialog management in the B-Bucks Economy Scripts project. It provides functions to retrieve cached data, launch modeless dialogs with initial data payloads, and centralizes the logic for opening various dialogs across the application.

/**
 * Retrieves cached data for a given key. It first checks the in-memory cache, then the script properties for permanent storage. If found in properties, it updates the cache for future quick access. If no data is found, it returns an error message as a JSON string.
 * @param cacheKey The key for the cached data to retrieve. This should correspond to a specific dataset like "cachedIndividuals", "cachedServices", etc.
 * @returns A JSON string containing the cached data or an error message if no data is found. The expected structure of the returned data depends on the cacheKey used, but it generally includes relevant information for the application's operations, such as lists of individuals, services, settings, or transactions.
 */
function getCachedData(cacheKey: string): string | void {
    try {
        const cache = CacheService.getScriptCache();

        const cached = cache.get(cacheKey);

        if (cached) return cached;

        const permanentData = PropertiesService.getScriptProperties().getProperty(cacheKey);

        if (permanentData) {
            cache.put(cacheKey, permanentData, SERVER_SIDE_CACHE_AGE);
            return permanentData;
        }

        return
    } catch (error: any) {
        log(`Error in getCachedData for key ${cacheKey}: ${error.message}`, true);
        return
    }
}

/**
 * Centralized setter
 * Handles mutex locking to prevent concurrent write collisions.
 * @param cacheKey The key for the cached data to update. This should correspond to a specific dataset like "cachedIndividuals", "cachedServices", etc.
 * @param data The data to cache, which will be stringified and stored in both the in-memory cache and the script properties for persistence. The structure of this data should align with what is expected for the given cacheKey, such as arrays of individuals, services, settings objects, or transaction records.
 * @returns {boolean} True if the cache was successfully updated, false otherwise.
 */
function setCachedData(cacheKey: string, data: any): boolean {
    try {
        const serializedData = JSON.stringify(data);
        const lock = LockService.getScriptLock();

        try {
            // Attempt to acquire the lock with a timeout to prevent indefinite waiting in case of issues
            lock.waitLock(WAIT_LOCK_TIME);

            // Update the cache and properties within the lock simultaneously to ensure consistency
            CacheService.getScriptCache().put(cacheKey, serializedData, SERVER_SIDE_CACHE_AGE);
            PropertiesService.getScriptProperties().setProperty(cacheKey, serializedData);
        } catch (lockError: any) {
            log(`Failed to acquire lock for cache update on key ${cacheKey}: ${lockError.message}`, true);
            return false;
        } finally {
            lock.releaseLock();
        }
        return true;
    } catch (error: any) {
        log(`Error in setCachedData for key ${cacheKey}: ${error.message}`, true);
        return false;
    }
}
/**
 * Launches a modeless dialog in the Google Sheets UI using a specified HTML template. The dialog is populated with initial data fetched from the cache, which includes individuals, services, settings, and transactions. The function takes parameters for the template name, dialog title, and dimensions (width and height) to customize the appearance of the dialog. This utility function centralizes the logic for opening various dialogs across the application, ensuring consistency in how data is passed and how dialogs are displayed.
 * @param templateName The name of the HTML template file (without the .html extension) to use for the dialog's content. This should correspond to a file in the project's HTML directory, such as "BalanceManager", "InvestmentsManager", etc.
 * @param title The title to display on the dialog window. This should be descriptive of the dialog's purpose, such as "Bank of Banderas - Manual Balance Manager".
 * @param width The width of the dialog in pixels. This should be set based on the expected content and layout of the dialog to ensure a good user experience without excessive scrolling or wasted space.
 * @param height The height of the dialog in pixels. Similar to width, this should be chosen to accommodate the content of the dialog while maintaining usability and aesthetics.
 */
function launchModelessDialog(templateName: string, title: string, width: number, height: number): void {
    const template = HtmlService.createTemplateFromFile(templateName);

    const cache = CacheService.getScriptCache();

    const globalData = {
        [SERVICES_CACHED_KEY]: cache.get(SERVICES_CACHED_KEY) || "{}",
        [SETTINGS_CACHED_KEY]: cache.get(SETTINGS_CACHED_KEY) || "{}",
        [TRANSACTIONS_CACHED_KEY]: cache.get(TRANSACTIONS_CACHED_KEY) || "[]",
        [INVESTMENTS_LEDGER_CACHED_KEY]: cache.get(INVESTMENTS_LEDGER_CACHED_KEY) || "{}"
    };

    template.initialServerPayload = `{
        "${SERVICES_CACHED_KEY}": ${globalData[SERVICES_CACHED_KEY]},
        "${SETTINGS_CACHED_KEY}": ${globalData[SETTINGS_CACHED_KEY]},
        "${TRANSACTIONS_CACHED_KEY}": ${globalData[TRANSACTIONS_CACHED_KEY]},
        "${INVESTMENTS_LEDGER_CACHED_KEY}": ${globalData[INVESTMENTS_LEDGER_CACHED_KEY]}
    }`;

    const html = template.evaluate()
        .setTitle(title)
        .setWidth(width)
        .setHeight(height);

    SpreadsheetApp.getUi().showModelessDialog(html, title);
}

function preloadCacheForAllDialogs(): void {
    // Preload all relevant data into the cache to ensure fast access when dialogs are opened. This can be called onOpen or at strategic points in the application to keep the cache warm.
    fetchServicesDataCached(JSON.stringify({ forceRefresh: true }));
    fetchSettingsDataCached(JSON.stringify({ forceRefresh: true }));
    fetchInvestmentsDataCached(JSON.stringify({ forceRefresh: true }));
    fetchTransactionsDataCached(JSON.stringify({ forceRefresh: true }));
}

/**
 * Exposes a clean, cross-dialog server-side setter for client states.
 * Keeps data alive in Google's high-speed RAM cache for up to 6 hours (21600 seconds).
 * This function can be called from any dialog to update the cache with new data, ensuring that all dialogs have access to the most recent information without needing to refresh or re-fetch from the sheet until necessary.
 * @param key The key for the cached data to update. This should correspond to a specific dataset like "cachedIndividuals", "cachedServices", etc.
 * @param data The data to cache, which will be stringified and stored in the in-memory cache for quick access. The structure of this data should align with what is expected for the given key, such as arrays of individuals, services, settings objects, or transaction records.
 */
function setServerCacheValue(data: string): boolean {
    const { key, value } = JSON.parse(data);

    try {
        const cache = CacheService.getScriptCache();
        // Cache strings up to 100KB per key. 
        cache.put(key, value, SERVER_SIDE_CACHE_AGE);
        return true;
    } catch (error: any) {
        log(`Failed to write to server cache layer: ${error.message} for key: ${key}`, true);
        return false;
    }
}

/**
 * Universally clears specified keys or can wipe out everything.
 * This is useful for ensuring that stale data doesn't persist in the cache, especially after significant updates or when debugging. By providing an array of keys, you can target specific datasets for clearing, or if you want to reset everything, you can call this function with all relevant keys.
 * @param keys An array of keys corresponding to the cached data that should be cleared. This should include all relevant cache keys used in the application, such as "cachedIndividuals", "cachedServices", "cachedSettings", "cachedTransactions", and "cachedInvestmentsLedger". If you want to clear all cached data, you can pass an array containing all these keys.
 * @returns {boolean} True if the cache was successfully cleared for the specified keys, false otherwise.
 */
function clearGlobalCache(keys: string[]): boolean {
    try {
        const cache = CacheService.getScriptCache();
        const props = PropertiesService.getScriptProperties();
        cache.removeAll(keys);

        keys.forEach(key => props.deleteProperty(key));

        log(`Global Cache Pipeline cleared for keys: ${keys.join(", ")}`, false);
        return true;
    } catch (error: any) {
        log(`Failed to purge global cache layer: ${error.message}`, true);
        return false;
    }
}

function clearServerCacheValue(key: string): boolean {
    try {
        const cache = CacheService.getScriptCache();
        const props = PropertiesService.getScriptProperties();
        cache.remove(key);
        props.deleteProperty(key);
        log(`Cleared duplicate server cache value for key: ${key}`, false);
        return true;
    } catch (error: any) {
        log(`Failed to clear duplicate server cache value for key ${key}: ${error.message}`, true);
        return false;
    }
}