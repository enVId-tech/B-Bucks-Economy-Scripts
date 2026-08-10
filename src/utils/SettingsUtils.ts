/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran. 
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 * 
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 * 
 * This file contains utility functions for various operations in the B-Bucks Economy Scripts project, including functions related to services management and other financial operations.
 */

interface ImportantDates {
    startBankingDate: string;
    endBankingDate: string;
    startInvestingDate: string;
    endInvestingDate: string;
    startQ1Date: string;
    endQ1Date: string;
    startQ2Date: string;
    endQ2Date: string;
    startQ3Date: string;
    endQ3Date: string;
    startQ4Date: string;
    endQ4Date: string;
}

interface StandardPercentages {
    incomeTaxRate: number;
    investmentWithdrawalTaxRate: number;
    weeklyInterestRate: number;
}

interface MandatedPolicies {
    withdrawLockoutTimeframeDays: number;
    allowDebt: boolean;
    enforceServicesLimits: boolean;
}

interface LedgersAndRecords {
    autologTransactions: boolean;
    logBankingDaily: boolean;
    logSheetLastModified: boolean;
    commentExpenditures: boolean;
    autoSortNames: boolean;
}

interface Limits {
    depositMinimum: number;
    depositMaximum: number;
    expendituresMaximum: number;
}

interface AdvancedTechnicalSettings {
    namesCol: number;
    balanceCol: number;
    earningsCol: number;
    netIncomeCol: number;
    expendituresCol: number;
    investmentReturnsCol: number;
    initialDepositCol: number;
    timeDepositCol: number;
    grossInvestmentGainCol: number;
    netInvestmentGainCol: number;
    netPercentageGainCol: number;
    defaultServicesSheet: string;
    servicesRowStart: number;
    defaultSettingsSheet: string;
    settingsRowStart: number;
    defaultTransactionsSheet: string;
    transactionsRowStart: number;
    fillStartingCell: string;
    fillSheetCellName: string;
    defaultHistoricalRecordsSheet: string;
    historicalRecordsRowStart: number;
    metricDataStartCol: number;
    periodColIndex: number;
    headerRowIndex: number;
    sheetNamePattern: string;
    timestampCell: string;
}

interface SettingsData {
    importantDates: ImportantDates;
    standardPercentages: StandardPercentages;
    mandatedPolicies: MandatedPolicies;
    ledgersAndRecords: LedgersAndRecords;
    limits: Limits;
    advancedTechnicalSettings: AdvancedTechnicalSettings;
}

type SettingsSectionKey = keyof SettingsData;
type SettingsPropertyMap = ImportantDates & StandardPercentages & MandatedPolicies & LedgersAndRecords & Limits & AdvancedTechnicalSettings;
type SettingsPropertyKey = keyof SettingsPropertyMap;

/**
 * Fetches settings data with caching. It first checks for cached data to minimize latency, and if not found or if a force refresh is requested, it reads the settings data from the sheet and updates the cache with the new data. This function ensures that the application can quickly access settings data while also providing a mechanism to refresh the data when necessary.
 * @param data A string containing the data for the function, including a forceRefresh flag. Defaults to undefined, meaning it will use cached data if available for faster access.
 * @returns {SettingsData | { error: string }} An object containing the structured settings data or an error message if the sheet is not found or an error occurs. The settings data includes important dates, standard percentages, mandated policies, ledgers and records preferences, and limits, all organized into their respective categories for easy access throughout the application.
 */
function fetchSettingsDataCached(data?: string): SettingsData | { error: string } {
    try {
        if (data && typeof data === 'string') {
            log(`Received data for fetchSettingsDataCached: ${data}`, false);
        } else {
            log("No data received for fetchSettingsDataCached, proceeding with default cache retrieval.", false);
            data = JSON.stringify({ forceRefresh: false });
        }

        const parsedData = data ? JSON.parse(data) : null;
        const forceRefresh = parsedData?.forceRefresh || false;

        const cache = CacheService.getScriptCache();
        const props = PropertiesService.getScriptProperties();

        if (!forceRefresh) {
            const cachedString = getCachedData(SETTINGS_CACHED_KEY);
            if (cachedString && cachedString !== "{}" && cachedString !== "") {
                log(`Cache hit: Settings data loaded from cache. String: ${cachedString}`, false);
                return JSON.parse(cachedString) as SettingsData;
            }

            const savedProperties = props.getProperty(SETTINGS_CACHED_KEY);
            if (savedProperties) {
                log(`Cache hit: Settings data loaded from server properties. String: ${savedProperties}`, false);
                // Repopulate fast RAM cache so the next window open loads even faster
                cache.put(SETTINGS_CACHED_KEY, savedProperties, 21600);
                return JSON.parse(savedProperties);
            }
        }

        log("Cache miss: Re-calculating policy vectors from Settings sheet cells...", false);
        const freshSettings = fetchSettingsData();

        if (freshSettings && !('error' in freshSettings)) {
            setCachedData(SETTINGS_CACHED_KEY, freshSettings);
        }

        return freshSettings;
    } catch (error: any) {
        log(`Error occurred in fetchSettingsDataCached: ${error.message}`, true);
        return { error: `Error occurred in fetchSettingsDataCached: ${error.message}` };
    }
}

/**
 * Fetches settings data from the "Settings" sheet in the active spreadsheet. It reads a predefined range of rows and columns to extract various configuration settings, including important dates, standard percentages, mandated policies, ledgers and records preferences, and limits. The function processes the raw grid data to construct structured objects for each category of settings, handling type conversions for booleans, numbers, and dates as needed. If the sheet is not found or an error occurs during processing, it returns an error message.
 * @returns {SettingsData | { error: string }} An object containing the structured settings data or an error message if the sheet is not found or an error occurs. The settings data includes important dates, standard percentages, mandated policies, ledgers and records preferences, and limits, all organized into their respective categories for easy access throughout the application.
 */
function fetchSettingsData(): SettingsData | { error: string } {
    try {

        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

        const settingsSheet = spreadsheet.getSheetByName(DEFAULT_SETTINGS_SHEET);

        if (!settingsSheet) {
            log("Settings sheet not found.", true);
            return { error: "Settings sheet not found." };
        }

        const rawGrid: any[][] = settingsSheet.getDataRange().getValues();

        const extractSettings = (keyName: string): SettingsData[keyof SettingsData] => {
            const { keyCol, valCol } = SETTINGS_COLUMNS[keyName];
            const resultObject: any = {};
            for (let i = SETTINGS_ROW_START - 1; i < rawGrid.length; i++) {
                // Ensure key exists and isn't whitespace padding
                const key = rawGrid[i][keyCol - 1];
                let val = rawGrid[i][valCol - 1];
                if (key !== undefined && key !== null && String(key).trim() !== "") {
                    const cleanKey = String(key).trim();

                    // Handle native Spreadsheet Boolean Conversions
                    if (val === true || String(val).toUpperCase() === "TRUE") {
                        val = true;
                    } else if (val === false || String(val).toUpperCase() === "FALSE") {
                        val = false;
                    } else if (val instanceof Date) {
                        // Handle Native Timestamp Date strings
                        val = val.toISOString();
                    } else if (cleanKey.toLowerCase().includes("date") && val) {
                        const parsedDate = new Date(val);
                        val = isNaN(parsedDate.getTime()) ? val : parsedDate.toISOString();
                    } else if (typeof val === 'number' && !isNaN(val)) {
                        // Handle absolute numeric handling
                        val = Number(val);
                    } else if (typeof val === 'string' && !isNaN(Number(val))) {
                        // If the value is a string that can be converted to a number, convert it
                        val = Number(val);
                    } else if (val === 'string') {
                        // Handle string values that are explicitly set as 'string'
                        val = String(val);
                    }

                    resultObject[cleanKey] = val;
                }
            }
            return resultObject;
        };

        return {
            importantDates: extractSettings(Object.keys(SETTINGS_COLUMNS)[0]) as ImportantDates,
            standardPercentages: extractSettings(Object.keys(SETTINGS_COLUMNS)[1]) as StandardPercentages,
            mandatedPolicies: extractSettings(Object.keys(SETTINGS_COLUMNS)[2]) as MandatedPolicies,
            ledgersAndRecords: extractSettings(Object.keys(SETTINGS_COLUMNS)[3]) as LedgersAndRecords,
            limits: extractSettings(Object.keys(SETTINGS_COLUMNS)[4]) as Limits,
            advancedTechnicalSettings: extractSettings(Object.keys(SETTINGS_COLUMNS)[5]) as AdvancedTechnicalSettings,
        };
    } catch (err: any) {
        log(`Error in fetchSettingsData: ${err.message}`, true);
        return { error: `Error in fetchSettingsData: ${err.message}` };
    }
}

/**
 * Fetches a specific property from the settings data.
 * @param propertyKey The key of the property to fetch.
 * @param propertySection The section of the settings data to search in.
 * @returns The value of the property or an error object if it's not found.
 */
function fetchProperty(propertyKey: SettingsPropertyKey, propertySection?: keyof SettingsData): string | { error: string } {
    try {
        const settings = fetchSettingsDataCached();

        if ('error' in settings) {
            log(`Failed to fetch settings data: ${settings.error}`, false);
            return { error: `Failed to fetch settings data: ${settings.error}` };
        }

        // If a specific section is provided, check that section first
        if (propertySection) {
            const section = settings[propertySection];
            if (section && typeof section === 'object' && propertyKey in section) {
                return String((section as any)[propertyKey]);
            }
        }

        // Search through each inner configuration section group
        let propertyValue: any = undefined;
        for (const section of Object.values(settings)) {
            if (section && typeof section === 'object' && propertyKey in section) {
                propertyValue = (section as any)[propertyKey];
                break;
            }
        }

        if (propertyValue === undefined) {
            log(`Property key "${propertyKey}" not found in any settings section.`, false);
            return { error: `Property key "${propertyKey}" not found in any settings section.` };
        }
        return String(propertyValue);
    } catch (error: any) {
        log(`Error in fetchProperty for key ${propertyKey}: ${error.message}`, true);
        return { error: `Error in fetchProperty for key ${propertyKey}: ${error.message}` };
    }
}

/**
 * Sets a property in the settings data.
 * @param data The data to set, either as a string (JSON) or an object.
 * @returns { boolean } Returns true if the property was successfully set; otherwise, returns false.
 */
function setSettingsProperty(data: any): boolean {
    const parseResult = typeof data === 'string' ? JSON.parse(data) : data;
    // Key is the column identifer for the setting (importantDates, standardPercentages, mandatedPolicies, ledgersAndRecords, limits, advancedTechnicalSettings)
    // Property value is every setting within that column as an object (e.g. { incomeTaxRate: 0.1, weeklyInterestRate: 0.02 } for standardPercentages)
    const { propertyKey, propertyValue } = parseResult;

    try {
        const settings = fetchSettingsDataCached();
        if ('error' in settings) {
            log(`Failed to fetch settings data: ${settings.error}`, true);
            return false;
        }

        const settingKeys = Object.keys(settings) as (keyof SettingsData)[];
        if (!settingKeys.includes(propertyKey as keyof SettingsData)) {
            log(`Invalid property key "${propertyKey}". Must be one of: ${settingKeys.join(", ")}`, true);
            return false;
        }

        // Update the specific property in the settings data
        settings[propertyKey] = propertyValue;

        // Also update the server properties for persistence across sessions and cache misses
        const props = PropertiesService.getScriptProperties();
        props.setProperty(SETTINGS_CACHED_KEY, JSON.stringify(settings));

        CacheService.getScriptCache().put(SETTINGS_CACHED_KEY, JSON.stringify(settings), 21600);

        // Update the sheet directly to ensure the source of truth is consistent
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const settingsSheet = spreadsheet.getSheetByName(DEFAULT_SETTINGS_SHEET);
        if (!settingsSheet) {
            log(`Settings sheet "${DEFAULT_SETTINGS_SHEET}" not found.`, true);
            return false;
        }

        // Find the row and column for the property key
        const columnInfo = SETTINGS_COLUMNS[propertyKey as keyof typeof SETTINGS_COLUMNS];

        log(`Attempting to update settings for property key "${propertyKey}" with value: ${JSON.stringify(propertyValue)}`, false);
        if (!columnInfo) {
            log(`Column information for property key "${propertyKey}" not found.`, true);
            return false;
        }

        const { keyCol, valCol } = columnInfo;
        const lastRow = settingsSheet.getLastRow();

        log(`Updating settings sheet "${DEFAULT_SETTINGS_SHEET}" for property key "${propertyKey}" in column ${keyCol} with value: ${JSON.stringify(propertyValue)}`, false);

        // Format the values to match the expected object structure in the sheet, which is key-value pairs where the key is the setting name and the value is the setting value. The keys are listed in the keyCol and the values are listed in the valCol, so we need to find the correct row for the propertyKey and update its corresponding value in valCol.
        Object.entries(propertyValue).forEach(([innerKey, innerValue]) => {
            let itemUpdated = false;

            // Look down the column to find the row matching this specific configuration parameter label
            for (let row = SETTINGS_ROW_START; row <= lastRow; row++) {
                const cellValue = settingsSheet.getRange(row, keyCol).getValue().toString().trim();

                if (cellValue === innerKey) {
                    // Normalize boolean expressions to native values so Google Sheets shows clean checkboxes
                    let finalValue = innerValue;
                    if (typeof finalValue === 'string' && finalValue.endsWith('Z') && !isNaN(Date.parse(finalValue))) {
                        // Pass back down as an actual runtime Date so the cell formats natively
                        finalValue = new Date(finalValue);
                    }

                    settingsSheet.getRange(row, valCol).setValue(finalValue);
                    itemUpdated = true;
                    break;
                }
            }

            if (!itemUpdated) {
                log(`Warning: Inner setting key "${innerKey}" was not found in column ${keyCol}.`, true);
            }
        });

        // apply the income tax rate to all users if the standardPercentages property was updated and contains an incomeTaxRate key
        if (propertyKey === "standardPercentages" && "incomeTaxRate" in propertyValue) {
            const applyResult = applyIncomeTaxRateToAll();
            if (!applyResult) {
                log(`Failed to apply income tax rate to all users after updating settings.`, true);
                return false;
            }
        }

        return true;
    } catch (error: any) {
        log(`Error in setSettingsProperty: ${error.message}`, true);
        return false;
    }
}

/**
 * Applies the income tax rate to every person in the sheet as a formula
 */
function applyIncomeTaxRateToAll(): boolean {
    try {
        const incomeTaxRateResult = fetchProperty("incomeTaxRate");
        if (typeof incomeTaxRateResult === 'object' && 'error' in incomeTaxRateResult) {
            log(`Failed to fetch income tax rate: ${incomeTaxRateResult.error}`, true);
            return false;
        }

        const incomeTaxRate = parseFloat(incomeTaxRateResult);

        if (isNaN(incomeTaxRate)) {
            log(`Invalid income tax rate fetched: ${incomeTaxRateResult}`, true);
            return false;
        }

        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const spreadsheetId = spreadsheet.getId();

        const allSheets = spreadsheet.getSheets();
        const periodSheets = allSheets.filter(sheet => sheet.getName().includes('Period'));

        const valueUpdates: GoogleAppsScript.Sheets.Schema.ValueRange[] = [];
        const nativeFallbackUpdates: Array<{ sheet: GoogleAppsScript.Spreadsheet.Sheet; numRows: number }> = [];

        periodSheets.forEach(sheet => {
            const lastRow = sheet.getLastRow();

            if (lastRow < USER_STARTING_ROW) {
                return;
            }

            const numRows = (lastRow - USER_STARTING_ROW) + 1;
            const formulas = Array.from({ length: numRows }, (_, index) => {
                const rowNumber = USER_STARTING_ROW + index;
                return [`=C${rowNumber}*(1-${incomeTaxRate})`];
            });

            nativeFallbackUpdates.push({ sheet, numRows });
            valueUpdates.push({
                range: `${sheet.getName()}!D${USER_STARTING_ROW}:D${lastRow}`,
                values: formulas,
            });
        });

        if (valueUpdates.length === 0) {
            return true;
        }

        if (typeof Sheets !== 'undefined' && Sheets.Spreadsheets && Sheets.Spreadsheets.Values) {
            try {
                Sheets.Spreadsheets.Values.batchUpdate(
                    {
                        valueInputOption: 'USER_ENTERED',
                        data: valueUpdates,
                    },
                    spreadsheetId
                );

                return true;
            } catch (apiError: any) {
                log(`Sheets API batch update failed for applyIncomeTaxRateToAll: ${apiError.message}. Falling back to native writes.`, true);
            }
        }

        nativeFallbackUpdates.forEach(({ sheet, numRows }) => {
            const formulas = Array.from({ length: numRows }, (_, index) => {
                const rowNumber = USER_STARTING_ROW + index;
                return [`=C${rowNumber}*(1-${incomeTaxRate})`];
            });

            sheet.getRange(USER_STARTING_ROW, NET_INCOME_COL, numRows, 1).setFormulas(formulas);
        });

        return true;
    } catch (error: any) {
        log(`Error in applyIncomeTaxRateToAll: ${error.message}`, true);
        return false;
    }
}