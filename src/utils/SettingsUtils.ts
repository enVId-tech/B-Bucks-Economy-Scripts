// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. 
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains utility functions for various operations in the B-Bucks Economy Scripts project, including functions related to services management and other financial operations.

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

interface SettingsData {
    importantDates: ImportantDates;
    standardPercentages: StandardPercentages;
    mandatedPolicies: MandatedPolicies;
    ledgersAndRecords: LedgersAndRecords;
    limits: Limits;
}

// -- DO NOT CHANGE THESE COLUMNS UNLESS YOU KNOW WHAT YOU ARE DOING --
// no but really actually dont pls ty, the sheet is named settings cuz ppl are stupid asf :D
const DEFAULT_SETTINGS_SHEET: string = "Settings";
// change only if u dont want the default option which is like wai bro ts beautiful
const COLUMNS = {
    importantDates: { keyCol: 2, valCol: 3 },
    standardPercentages: { keyCol: 5, valCol: 6 },
    mandatedPolicies: { keyCol: 8, valCol: 9 },
    ledgersAndRecords: { keyCol: 11, valCol: 12 },
    limits: { keyCol: 14, valCol: 15 },
}
const ROW_START = 4;

/**
 * Fetches settings data with caching. It first checks for cached data to minimize latency, and if not found or if a force refresh is requested, it reads the settings data from the sheet and updates the cache with the new data. This function ensures that the application can quickly access settings data while also providing a mechanism to refresh the data when necessary.
 * @param data A string containing the data for the function, including a forceRefresh flag. Defaults to undefined, meaning it will use cached data if available for faster access.
 * @returns {SettingsData | { error: string }} An object containing the structured settings data or an error message if the sheet is not found or an error occurs. The settings data includes important dates, standard percentages, mandated policies, ledgers and records preferences, and limits, all organized into their respective categories for easy access throughout the application.
 */
function fetchSettingsDataCached(data?: string): SettingsData | { error: string } {
    try {
        if (data && typeof data === 'string') {
            Logger.log(`Received data for fetchSettingsDataCached: ${data}`);
        } else {
            Logger.log("No data received for fetchSettingsDataCached, proceeding with default cache retrieval.");
            data = JSON.stringify({ forceRefresh: false });
        }

        const parsedData = data ? JSON.parse(data) : null;
        const forceRefresh = parsedData?.forceRefresh || false;

        const CACHE_KEY = "cachedSettings";
        const cache = CacheService.getScriptCache();
        const props = PropertiesService.getScriptProperties();

        if (!forceRefresh) {
            const cachedString = getCachedData(CACHE_KEY);
            if (cachedString && cachedString !== "{}" && cachedString !== "") {
                // SpreadsheetApp.getUi().alert(`Cache hit: Settings data loaded from cache. String: ${cachedString}`);
                return JSON.parse(cachedString) as SettingsData;
            }

            const savedProperties = props.getProperty(CACHE_KEY);
            if (savedProperties) {
                Logger.log(`Server Cache Hit (Properties) for ${CACHE_KEY}`);
                // Repopulate fast RAM cache so the next window open loads even faster
                cache.put(CACHE_KEY, savedProperties, 21600);
                return JSON.parse(savedProperties);
            }
        }

        console.log("Cache miss: Re-calculating policy vectors from Settings sheet cells...");
        // SpreadsheetApp.getUi().alert("Cache miss: Re-calculating policy vectors from Settings sheet cells...");
        const freshSettings = fetchSettingsData();

        if (freshSettings && !('error' in freshSettings)) {
            setCachedData(CACHE_KEY, freshSettings);
        }

        return freshSettings;
    } catch (error: any) {
        // SpreadsheetApp.getUi().alert(`Error occurred in fetchSettingsDataCached: ${error.message}`);
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
            Logger.log("Settings sheet not found.");
            try { SpreadsheetApp.getUi().alert("Settings sheet not found."); } catch (e) { }
            return { error: "Settings sheet not found." };
        }

        const rawGrid: any[][] = settingsSheet.getDataRange().getValues();

        const extractSettings = (keyName: string): SettingsData[keyof SettingsData] => {
            const { keyCol, valCol } = COLUMNS[keyName];
            const resultObject: any = {};
            for (let i = ROW_START - 1; i < rawGrid.length; i++) {
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
                    }

                    resultObject[cleanKey] = val;
                }
            }
            return resultObject;
        };

        return {
            importantDates: extractSettings("importantDates") as ImportantDates,
            standardPercentages: extractSettings("standardPercentages") as StandardPercentages,
            mandatedPolicies: extractSettings("mandatedPolicies") as MandatedPolicies,
            ledgersAndRecords: extractSettings("ledgersAndRecords") as LedgersAndRecords,
            limits: extractSettings("limits") as Limits,
        };
    } catch (err: any) {
        Logger.log(`Error in fetchSettingsData: ${err.message}`);
        try { SpreadsheetApp.getUi().alert(`Error in fetchSettingsData: ${err.message}`); } catch (e) { }
        return { error: `Error in fetchSettingsData: ${err.message}` };
    }
}

function fetchProperty(propertyKey: string): string | { error: string } {
    try {
        const settings = fetchSettingsDataCached();

        if ('error' in settings) {
            return { error: `Failed to fetch settings data: ${settings.error}` };
        }

        // Access the requested property using the provided key.
        // Each property key is unique and corresponds to a specific setting in the Settings sheet, such as "incomeTaxRate", "startBankingDate", "allowDebt", etc.
        const propertyValue = (settings as any)[propertyKey];

        if (propertyValue === undefined) {
            return { error: `Property key "${propertyKey}" not found in settings data.` };
        }
        return String(propertyValue);
    } catch (error: any) {
        Logger.log(`Error in fetchProperty for key ${propertyKey}: ${error.message}`);
        return { error: `Error in fetchProperty for key ${propertyKey}: ${error.message}` };
    }
}

function setSettingsProperty(data: any): boolean {
    const parseResult = typeof data === 'string' ? JSON.parse(data) : data;
    // Key is the column identifer for the setting (importantDates, standardPercentages, mandatedPolicies, ledgersAndRecords, limits)
    // Property value is every setting within that column as an object (e.g. { incomeTaxRate: 0.1, weeklyInterestRate: 0.02 } for standardPercentages)
    const { propertyKey, propertyValue } = parseResult;

    try {
        const settings = fetchSettingsDataCached();
        if ('error' in settings) {
            Logger.log(`Failed to fetch settings data: ${settings.error}`);
            SpreadsheetApp.getUi().alert(`Failed to fetch settings data: ${settings.error}`);
            return false;
        }

        const settingKeys = Object.keys(settings) as (keyof SettingsData)[];
        if (!settingKeys.includes(propertyKey as keyof SettingsData)) {
            Logger.log(`Invalid property key "${propertyKey}". Must be one of: ${settingKeys.join(", ")}`);
            SpreadsheetApp.getUi().alert(`Invalid property key "${propertyKey}". Must be one of: ${settingKeys.join(", ")}`);
            return false;
        }

        // Update the specific property in the settings data
        settings[propertyKey] = propertyValue;

        // Also update the server properties for persistence across sessions and cache misses
        const props = PropertiesService.getScriptProperties();
        props.setProperty("cachedSettings", JSON.stringify(settings));

        CacheService.getScriptCache().put("cachedSettings", JSON.stringify(settings), 21600);

        // Update the sheet directly to ensure the source of truth is consistent
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const settingsSheet = spreadsheet.getSheetByName(DEFAULT_SETTINGS_SHEET);
        if (!settingsSheet) {
            Logger.log(`Settings sheet "${DEFAULT_SETTINGS_SHEET}" not found.`);
            SpreadsheetApp.getUi().alert(`Settings sheet "${DEFAULT_SETTINGS_SHEET}" not found.`);
            return false;
        }

        // Find the row and column for the property key
        const columnInfo = COLUMNS[propertyKey as keyof typeof COLUMNS];

        // SpreadsheetApp.getUi().alert(`Found column info for property key "${propertyKey}": keyCol=${columnInfo.keyCol}, valCol=${columnInfo.valCol}, columnInfo: ${JSON.stringify(columnInfo)}`);
        if (!columnInfo) {
            Logger.log(`Column information for property key "${propertyKey}" not found.`);
            SpreadsheetApp.getUi().alert(`Column information for property key "${propertyKey}" not found.`);
            return false;
        }

        const { keyCol, valCol } = columnInfo;
        const lastRow = settingsSheet.getLastRow();

        // SpreadsheetApp.getUi().alert(`Attempting to update with ${JSON.stringify({ propertyValue })}}`)

        // Format the values to match the expected object structure in the sheet, which is key-value pairs where the key is the setting name and the value is the setting value. The keys are listed in the keyCol and the values are listed in the valCol, so we need to find the correct row for the propertyKey and update its corresponding value in valCol.
        Object.entries(propertyValue).forEach(([innerKey, innerValue]) => {
            let itemUpdated = false;

            // Look down the column to find the row matching this specific configuration parameter label
            for (let row = ROW_START; row <= lastRow; row++) {
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
                Logger.log(`Warning: Inner setting key "${innerKey}" was not found in column ${keyCol}.`);
                SpreadsheetApp.getUi().alert(`Warning: Inner setting key "${innerKey}" was not found in column ${keyCol}. Please ensure the key exists in the sheet for it to be updated.`);
            }
        });

        return true;
    } catch (error: any) {
        Logger.log(`Error in setSettingsProperty: ${error.message}`);
        SpreadsheetApp.getUi().alert(`Error in setSettingsProperty: ${error.message}`);
        return false;
    }
}