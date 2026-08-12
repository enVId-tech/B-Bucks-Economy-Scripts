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
 * Centralized setter with mutex locking to prevent concurrent write collisions.
 * @param cacheKey The key for the cached data to update.
 * @param data The data to cache (will be JSON stringified).
 * @returns {boolean} True if successfully updated, false otherwise.
 */
function setCachedData(cacheKey: string, data: any): boolean {
    try {
        const serializedData = typeof data === 'string' ? data : JSON.stringify(data);
        const lock = LockService.getScriptLock();

        const hasLock = lock.tryLock(WAIT_LOCK_TIME);
        if (!hasLock) {
            log(`Failed to acquire lock for cache update on key ${cacheKey}`, true);
            return false;
        }

        try {
            CacheService.getScriptCache().put(cacheKey, serializedData, SERVER_SIDE_CACHE_AGE);
            PropertiesService.getScriptProperties().setProperty(cacheKey, serializedData);
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
        [SERVICES_CACHED_KEY]: cache.get(SERVICES_CACHED_KEY) || "{}",
        [SETTINGS_CACHED_KEY]: cache.get(SETTINGS_CACHED_KEY) || "{}",
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
    fetchServicesDataCached(JSON.stringify({ forceRefresh: true }));
    fetchSettingsDataCached(JSON.stringify({ forceRefresh: true }));
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

                // Precision Regex: match cell coordinates (e.g. A1, $B$2:$C$10) without modifying function names like SUM or AVERAGE
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

        const lastRow = historySheet.getLastRow();
        const startRow = lastRow < historicalStartRow ? historicalStartRow : lastRow + 1;

        const rowsNeeded = outputRows.length;
        const currentMax = historySheet.getMaxRows();
        if ((startRow + rowsNeeded - 1) > currentMax) {
            historySheet.insertRowsAfter(currentMax, (startRow + rowsNeeded - 1) - currentMax);
        }

        const targetRange = historySheet.getRange(startRow, 1, outputRows.length, outputRows[0].length);
        targetRange.setValues(outputRows);

        if (startRow > historicalStartRow) {
            const templateRange = historySheet.getRange(historicalStartRow, 1, 1, outputRows[0].length);
            templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        }

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