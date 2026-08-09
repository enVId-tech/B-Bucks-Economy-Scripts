/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran. All rights reserved.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 * 
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 * 
 * This file contains utility functions for data caching and dialog management in the B-Bucks Economy Scripts project. It provides functions to retrieve cached data, launch modeless dialogs with initial data payloads, and centralizes the logic for opening various dialogs across the application.
 */

interface PeriodConfig {
    name: string;
    rowIndex: number;
    formulas: (string | number | boolean)[];
}

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

function updateTimestamps(): void {
    try {

    } catch (error: any) {
        log(`Error in updateTimestamps: ${error.message}`, true);
    }
}

/**
 * Dynamic Historical Logger
 * This function records the current state of the economy into a historical records sheet. It evaluates formulas from various period sheets, captures their calculated values, and appends them to the historical records sheet with a timestamp. The function handles dynamic periods, formula evaluation, and ensures that the historical records are kept up-to-date for analysis and reporting.
 * The function is designed to be efficient by using batch operations for reading and writing data, and it includes error handling to ensure that any issues during the process are logged for review. It also manages the creation of a temporary calculation sheet to evaluate formulas without affecting the main sheets.
 * Note: This function assumes that the historical records sheet and the period sheets are structured correctly, and that the formulas in the period sheets are valid and return numerical values. It is important to ensure that the sheet names and ranges used in this function match the actual structure of the Google Sheets document.
 * @returns {void} This function does not return a value. It performs operations on the Google Sheets document to update the historical records.
 */
function recordDailyData(): void {
    // Fetch settings data
    const settings: SettingsData | { error: string } = fetchSettingsDataCached(JSON.stringify({ forceRefresh: true }));

    if (!settings || typeof settings === 'string' || 'error' in settings) {
        log(`Failed to fetch settings data for daily economics recording: ${typeof settings === 'string' ? settings : (settings as { error: string }).error}`, true);
        return;
    }

    const logDaily = settings && typeof settings !== 'string' && settings.ledgersAndRecords && settings.ledgersAndRecords.logBankingDaily;

    // Guard against paused executions if global flag exists
    if (logDaily === false) {
        log(`Daily economics logging is currently paused. Skipping recordDailyEconomics execution.`, false);
        return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheetName = typeof DEFAULT_HISTORICAL_RECORDS_SHEET !== 'undefined'
        ? DEFAULT_HISTORICAL_RECORDS_SHEET
        : "Economic Records";

    const historySheet = ss.getSheetByName(sheetName);
    if (!historySheet) {
        log(`Sheet "${sheetName}" not found.`, true);
        return;
    }

    const timestamp = new Date();

    // Settings block setup (B3:Z10)
    const settingsRange = historySheet.getRange("B3:Z10");
    const settingsBlock = settingsRange.getValues();
    const numCols = settingsBlock[0].length;

    const validPeriods: PeriodConfig[] = [];

    // Parse valid periods and formulas
    settingsBlock.forEach((row, rowIndex) => {
        const periodName = row[0]?.toString().trim();
        if (!periodName || periodName === "" || periodName === "Sheet") return;

        validPeriods.push({
            name: periodName,
            rowIndex: rowIndex,
            formulas: row.slice(1)
        });
    });

    if (validPeriods.length === 0) return;

    const tempSheet = ss.insertSheet(`calc_ws_${Date.now()}`);

    try {
        const formulaMatrix: string[][] = [];

        validPeriods.forEach((period) => {
            const rowFormulas = period.formulas.map((cellValue) => {
                let formulaTemplate = cellValue?.toString().trim();
                if (!formulaTemplate) return "";

                // Strip leading '=' if present in the raw settings cell
                if (formulaTemplate.startsWith('=')) {
                    formulaTemplate = formulaTemplate.substring(1);
                }

                // Scope ranges to the target period sheet
                const rangeRegex = /([A-Za-z]+[0-9]+(?::[A-Za-z]+[0-9]+)?)/g;
                const scopedFormula = formulaTemplate.replace(rangeRegex, `'${period.name}'!$1`);

                // Prepend '=' so Google Sheets evaluates it properly
                return `=${scopedFormula}`;
            });
            formulaMatrix.push(rowFormulas);
        });

        // Batch apply formulas to temp sheet
        const calcRange = tempSheet.getRange(1, 1, formulaMatrix.length, numCols - 1);
        calcRange.setFormulas(formulaMatrix);
        SpreadsheetApp.flush();

        // Retrieve calculated outputs
        const calculatedValues = calcRange.getValues();

        // Build historical row matrix
        const outputRows: (string | number | Date | null)[][] = validPeriods.map((period, pIdx) => {
            const rowValues = calculatedValues[pIdx] || [];
            const formattedRow: (string | number | Date | null)[] = [timestamp, period.name];

            for (let colIndex = 0; colIndex < numCols - 1; colIndex++) {
                const val = rowValues[colIndex];

                if (val === undefined || val === "" || (typeof val === 'string' && val.includes("#"))) {
                    formattedRow.push(null);
                } else if (typeof val === 'number') {
                    formattedRow.push(Number(val.toFixed(2)));
                } else {
                    formattedRow.push(val);
                }
            }
            return formattedRow;
        });

        const historicalStartRow = typeof HISTORICAL_RECORDS_ROW_START !== 'undefined'
            ? HISTORICAL_RECORDS_ROW_START
            : 11;

        const lastRow = historySheet.getLastRow();
        const startRow = lastRow < historicalStartRow ? historicalStartRow : lastRow + 1;

        // Auto-expand sheet rows if needed
        const rowsNeeded = outputRows.length;
        const currentMax = historySheet.getMaxRows();
        if ((startRow + rowsNeeded - 1) > currentMax) {
            historySheet.insertRowsAfter(currentMax, (startRow + rowsNeeded - 1) - currentMax);
        }

        // Batch write data
        const targetRange = historySheet.getRange(startRow, 1, outputRows.length, outputRows[0].length);
        targetRange.setValues(outputRows);

        // Batch apply formatting
        if (startRow > historicalStartRow) {
            const templateRange = historySheet.getRange(historicalStartRow, 1, 1, outputRows[0].length);
            templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        }

        // Sort complete data range from row 11 downwards by Column A descending (most recent first)
        const finalLastRow = historySheet.getLastRow();
        const totalColumns = historySheet.getLastColumn();

        if (finalLastRow >= historicalStartRow) {
            const dataRangeToSort = historySheet.getRange(
                historicalStartRow,
                1,
                finalLastRow - historicalStartRow + 1,
                totalColumns
            );

            dataRangeToSort.sort({ column: 1, ascending: false });
        }

    } catch (error: any) {
        log(`Error in recordDailyData: ${error.message}`, true);
    } finally {
        ss.deleteSheet(tempSheet);
    }
}

function saveHistoricalRecords(): boolean {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DEFAULT_HISTORICAL_RECORDS_SHEET);
        if (!sheet) {
            log(`Historical Records sheet "${DEFAULT_HISTORICAL_RECORDS_SHEET}" not found.`, true);
            return false;
        }

        // Save sheet to new tab with timestamped name
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const newSheetName = `Records_${timestamp}`;

        const newSheet = sheet.copyTo(SpreadsheetApp.getActiveSpreadsheet());

        newSheet.setName(newSheetName);

        return true;
    } catch (error: any) {
        log(`Error in saveHistoricalRecords: ${error.message}`, true);
        return false;
    }
}

function resetHistoricalRecords(): boolean {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DEFAULT_HISTORICAL_RECORDS_SHEET);
        if (!sheet) {
            log(`Historical Records sheet "${DEFAULT_HISTORICAL_RECORDS_SHEET}" not found.`, true);
            return false;
        }

        // Clear all rows below the header row (assuming the header is in row 10)
        const lastRow = sheet.getLastRow();
        if (lastRow > HISTORICAL_RECORDS_ROW_START) {
            sheet.getRange(HISTORICAL_RECORDS_ROW_START + 1, 1, lastRow - HISTORICAL_RECORDS_ROW_START, sheet.getLastColumn()).clearContent();
        }

        log(`Historical Records reset successfully.`, false);

        return true;
    } catch (error: any) {
        log(`Error in resetHistoricalRecords: ${error.message}`, true);
        return false;
    }
}