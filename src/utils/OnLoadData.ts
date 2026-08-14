/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran. All rights reserved.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 */

interface PeriodConfig {
    name: string;
    rowIndex: number;
    formulas: (string | number | boolean)[];
}

/**
 * Retrieves cached data for a given key. Checks memory cache, then script properties.
 * @param cacheKey The key for the cached data to retrieve.
 * @returns The cached string payload, or null if no data is found.
 */
function getCachedData(cacheKey: string): string | null {
    try {
        const cache = CacheService.getScriptCache();
        const cached = cache.get(cacheKey);

        if (cached) return cached;

        const permanentData = PropertiesService.getScriptProperties().getProperty(cacheKey);

        if (permanentData) {
            cache.put(cacheKey, permanentData, SERVER_SIDE_CACHE_AGE);
            return permanentData;
        }

        return null;
    } catch (error: any) {
        log(`Error in getCachedData for key ${cacheKey}: ${error.message}`, true);
        return null;
    }
}

/**
 * Retrieves the current ordered list of cached keys (oldest to newest).
 */
function getCacheIndex(): string[] {
    try {
        const raw = PropertiesService.getScriptProperties().getProperty(CACHE_INDEX_KEY);
        return raw ? JSON.parse(raw) : [];
    } catch (e) {
        return [];
    }
}

/**
 * Updates the cache key index, maintaining LRU order and enforcing MAX_CACHE_ENTRIES.
 */
function touchCacheIndex(key: string): void {
    try {
        const props = PropertiesService.getScriptProperties();
        let index = getCacheIndex();

        // Remove existing instance of key to push it to the end (most recent)
        index = index.filter(k => k !== key && k !== CACHE_INDEX_KEY);
        index.push(key);

        // Evict oldest entries if total keys exceed MAX_CACHE_ENTRIES
        while (index.length > MAX_CACHE_ENTRIES) {
            const evictedKey = index.shift();
            if (evictedKey) {
                CacheService.getScriptCache().remove(evictedKey);
                props.deleteProperty(evictedKey);
                log(`Evicted oldest cache entry to enforce max capacity: ${evictedKey}`, false);
            }
        }

        props.setProperty(CACHE_INDEX_KEY, JSON.stringify(index));
    } catch (err: any) {
        log(`Failed to execute touchCacheIndex: ${err.message}.`, false);
    }
}

/**
 * Safely caps payload size by truncating arrays to keep only the most recent items.
 */
function enforcePayloadSizeLimit(data: any, maxBytes: number = MAX_PROPERTY_BYTES): string {
    let serialized = typeof data === 'string' ? data : JSON.stringify(data);

    // If string fits under quota, return as-is
    if (Utilities.newBlob(serialized).getBytes().length <= maxBytes) {
        return serialized;
    }

    // If payload is an array (or parsed object containing an array), keep only most recent entries
    try {
        let parsed = typeof data === 'string' ? JSON.parse(data) : data;

        if (Array.isArray(parsed)) {
            while (parsed.length > 1 && Utilities.newBlob(JSON.stringify(parsed)).getBytes().length > maxBytes) {
                // Remove oldest item from beginning of array
                parsed.shift();
            }
            return JSON.stringify(parsed);
        } else if (typeof parsed === 'object' && parsed !== null) {
            // Optional: If object contains array fields (like a ledger), prune array properties
            for (const key of Object.keys(parsed)) {
                if (Array.isArray(parsed[key])) {
                    while (parsed[key].length > 1 && Utilities.newBlob(JSON.stringify(parsed)).getBytes().length > maxBytes) {
                        parsed[key].shift();
                    }
                }
            }
            return JSON.stringify(parsed);
        }
    } catch (error: any) {
        log(`Failed to parse and truncate payload: ${error.message}`, true);
    }

    return serialized;
}

/**
 * Centralized setter with mutex locking, byte-size bounds, and max entry eviction.
 */
function setCachedData(cacheKey: string, data: any): boolean {
    try {
        // Enforce payload size limits (prune to keep most recent entries)
        const safePayload = enforcePayloadSizeLimit(data, MAX_PROPERTY_BYTES);

        const lock = LockService.getScriptLock();
        const hasLock = lock.tryLock(WAIT_LOCK_TIME);

        if (!hasLock) {
            log(`Failed to acquire lock for cache update on key ${cacheKey}`, true);
            return false;
        }

        try {
            // Save payload to CacheService and PropertiesService
            CacheService.getScriptCache().put(cacheKey, safePayload, SERVER_SIDE_CACHE_AGE);
            PropertiesService.getScriptProperties().setProperty(cacheKey, safePayload);

            // Maintain key index and evict oldest keys if entry limit is exceeded
            touchCacheIndex(cacheKey);
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
 * Launches a modeless dialog in the Google Sheets UI using a specified HTML template.
 */
function launchModelessDialog(templateName: string, title: string, width: number, height: number): void {
    const template = HtmlService.createTemplateFromFile(templateName);
    const cache = CacheService.getScriptCache();

    // Safely structure payload to avoid raw script injection bugs
    const payloadObject = {
        [SETTINGS_CACHED_KEY]: cache.get(SETTINGS_CACHED_KEY) || "{}",
        [SERVICES_CACHED_KEY]: cache.get(SERVICES_CACHED_KEY) || "{}",
        [TRANSACTIONS_CACHED_KEY]: cache.get(TRANSACTIONS_CACHED_KEY) || "[]",
        [INVESTMENTS_LEDGER_CACHED_KEY]: cache.get(INVESTMENTS_LEDGER_CACHED_KEY) || "{}"
    };

    template.initialServerPayload = JSON.stringify(payloadObject);

    const html = template.evaluate()
        .setTitle(title)
        .setWidth(width)
        .setHeight(height);

    SpreadsheetApp.getUi().showModelessDialog(html, title);
}

function preloadCacheForAllDialogs(): void {
    fetchSettingsDataCached(JSON.stringify({ forceRefresh: true }));
    fetchServicesDataCached(JSON.stringify({ forceRefresh: true }));
    fetchInvestmentsDataCached(JSON.stringify({ forceRefresh: true }));
    fetchTransactionsDataCached(JSON.stringify({ forceRefresh: true }));
}

function setServerCacheValue(data: string): boolean {
    try {
        const { key, value } = JSON.parse(data);
        const cache = CacheService.getScriptCache();

        const stringValue = typeof value === 'string' ? value : JSON.stringify(value);
        cache.put(key, stringValue, SERVER_SIDE_CACHE_AGE);
        return true;
    } catch (error: any) {
        log(`Failed to write to server cache layer: ${error.message}`, true);
        return false;
    }
}

/**
 * Clears specific keys and cleans up the key index.
 */
function clearGlobalCache(keys: string[]): boolean {
    try {
        const cache = CacheService.getScriptCache();
        const props = PropertiesService.getScriptProperties();

        cache.removeAll(keys);
        keys.forEach(key => props.deleteProperty(key));

        // Update tracking index
        let index = getCacheIndex();
        index = index.filter(k => !keys.includes(k));
        props.setProperty(CACHE_INDEX_KEY, JSON.stringify(index));

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
        log(`Cleared server cache value for key: ${key}`, false);
        return true;
    } catch (error: any) {
        log(`Failed to clear server cache value for key ${key}: ${error.message}`, true);
        return false;
    }
}

function bulkUpdateTimestamps(): void {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const sheetsToUpdate = spreadsheet.getSheets().filter(sheet => TIMESTAMP_LIST.includes(sheet.getName()));

        const currentUser = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
        const lastModifiedBy = currentUser ? ` by ${currentUser}` : '';
        const timestampText = `Last updated: ${new Date().toLocaleString()}${lastModifiedBy}`;

        sheetsToUpdate.forEach(sheet => {
            sheet.getRange(TIMESTAMP_CELL).setValue(timestampText);
        });
    } catch (error: any) {
        log(`Error in bulkUpdateTimestamps: ${error.message}`, true);
    }
}

function updateTimestampForSheet(): void {
    try {
        const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        if (TIMESTAMP_LIST.includes(activeSheet.getName())) {
            const currentUser = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
            const lastModifiedBy = currentUser ? ` by ${currentUser}` : '';
            activeSheet.getRange(TIMESTAMP_CELL).setValue(`Last updated: ${new Date().toLocaleString()}${lastModifiedBy}`);
        }
    } catch (error: any) {
        log(`Error in updateTimestampForSheet: ${error.message}`, true);
    }
}

function updateTimestamps(): void {
    try {
        bulkUpdateTimestamps();
        if (typeof refreshAllInvestments === 'function') {
            refreshAllInvestments();
        }
    } catch (error: any) {
        log(`Error in updateTimestamps: ${error.message}`, true);
    }
}

function recordDailyData(): void {
    const logDaily = fetchProperty("logBankingDaily", "ledgersAndRecords");

    if (logDaily === false) {
        log(`Daily economics logging is currently paused. Skipping recordDailyEconomics execution.`, false);
        return;
    }

    const today = new Date();
    const dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, 2 = Tuesday, ...

    // Early exit if today is Sunday (0) or Monday (1)
    if (dayOfWeek === 0 || dayOfWeek === 1) {
        log("Execution skipped: No weekend sheet modifications scheduled.", false);
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
    const settingsRange = historySheet.getRange("B3:Z10");
    const settingsBlock = settingsRange.getValues();
    const numCols = settingsBlock[0].length;

    const validPeriods: PeriodConfig[] = [];

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

                if (formulaTemplate.startsWith('=')) {
                    formulaTemplate = formulaTemplate.substring(1);
                }

                // Precision Regex: match cell coordinates without modifying function names
                const cellRangeRegex = /\b(\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?)\b/g;
                const scopedFormula = formulaTemplate.replace(cellRangeRegex, `'${period.name}'!$1`);

                return `=${scopedFormula}`;
            });
            formulaMatrix.push(rowFormulas);
        });

        const calcRange = tempSheet.getRange(1, 1, formulaMatrix.length, numCols - 1);
        calcRange.setFormulas(formulaMatrix);
        SpreadsheetApp.flush();

        const calculatedValues = calcRange.getValues();

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

        const numNewRows = outputRows.length;
        const numNewCols = outputRows[0].length;

        // 1. Insert space directly at the top of the history list
        historySheet.insertRowsBefore(historicalStartRow, numNewRows);

        // 2. Target the newly created empty space
        const targetRange = historySheet.getRange(historicalStartRow, 1, numNewRows, numNewCols);
        targetRange.setValues(outputRows);

        // 3. Copy formatting from existing row below if available
        const existingDataRow = historicalStartRow + numNewRows;
        if (historySheet.getLastRow() >= existingDataRow) {
            const templateRange = historySheet.getRange(existingDataRow, 1, 1, numNewCols);
            templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        }

        // 4. Sort the entire block down to the last row by Column A (Newest to Oldest)
        const finalLastRow = historySheet.getLastRow();
        if (finalLastRow >= historicalStartRow) {
            const dataRangeToSort = historySheet.getRange(
                historicalStartRow,
                1,
                finalLastRow - historicalStartRow + 1,
                numNewCols // Fixed to explicitly match data width
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
        const sheetName = typeof DEFAULT_HISTORICAL_RECORDS_SHEET !== 'undefined'
            ? DEFAULT_HISTORICAL_RECORDS_SHEET
            : "Economic Records";

        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (!sheet) {
            log(`Historical Records sheet "${sheetName}" not found.`, true);
            return false;
        }

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
        const sheetName = typeof DEFAULT_HISTORICAL_RECORDS_SHEET !== 'undefined'
            ? DEFAULT_HISTORICAL_RECORDS_SHEET
            : "Economic Records";

        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (!sheet) {
            log(`Historical Records sheet "${sheetName}" not found.`, true);
            return false;
        }

        const maxRows = sheet.getMaxRows();
        const maxCols = sheet.getMaxColumns();
        const lastRow = sheet.getLastRow();

        const startRow = typeof HISTORICAL_RECORDS_ROW_START !== 'undefined'
            ? HISTORICAL_RECORDS_ROW_START
            : 11;

        // Guard against out of bounds or empty target ranges
        if (lastRow < startRow) {
            log(`No historical data rows found below row ${startRow} to clear.`, false);
            return true;
        }

        const rowsToClear = (lastRow - startRow) + 1;
        sheet.getRange(startRow, 1, rowsToClear, maxCols).clearContent();

        // Safely prune excess empty structural rows if present
        const startDeleteRow = startRow + 1;
        if (lastRow > startRow && startDeleteRow <= maxRows) {
            const numToDelete = lastRow - startRow;
            const safeToDelete = Math.min(numToDelete, (maxRows - startDeleteRow) + 1);
            if (safeToDelete > 0) {
                sheet.deleteRows(startDeleteRow, safeToDelete);
            }
        }

        return true;
    } catch (error: any) {
        log(`Error in resetHistoricalRecords: ${error.message}`, true);
        return false;
    }
}