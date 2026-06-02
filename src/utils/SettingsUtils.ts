// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. 
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains utility functions for various operations in the B-Bucks Economy Scripts project, including functions related to services management and other financial operations.

interface ImportantDates {
    startBankingDate: Date;
    endBankingDate: Date;
    startInvestingDate: Date;
    endInvestingDate: Date;
    startQuarterOneDate: Date;
    endQuarterOneDate: Date;
    startQuarterTwoDate: Date;
    endQuarterTwoDate: Date;
    startQuarterThreeDate: Date;
    endQuarterThreeDate: Date;
    startQuarterFourDate: Date;
    endQuarterFourDate: Date;
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

/**
 * Fetches settings data with caching. It first checks for cached data to minimize latency, and if not found or if a force refresh is requested, it reads the settings data from the sheet and updates the cache with the new data. This function ensures that the application can quickly access settings data while also providing a mechanism to refresh the data when necessary.
 * @param forceRefresh A boolean flag indicating whether to bypass the cache and fetch fresh data from the sheet. Defaults to false, meaning it will use cached data if available for faster access. Setting this to true will force the function to read directly from the sheet and update the cache with the new data.
 * @returns {SettingsData | { error: string }} An object containing the structured settings data or an error message if the sheet is not found or an error occurs. The settings data includes important dates, standard percentages, mandated policies, ledgers and records preferences, and limits, all organized into their respective categories for easy access throughout the application.
 */
function fetchSettingsDataCached(forceRefresh: boolean = false): SettingsData | { error: string } {
    try {
        const CACHE_KEY = "cachedSettings";

        if (!forceRefresh) {
            const cachedString = getCachedData(CACHE_KEY);
            if (cachedString && cachedString !== "{}" && cachedString !== "") {
                // SpreadsheetApp.getUi().alert(`Cache hit: Settings data loaded from cache. String: ${cachedString}`);
                return JSON.parse(cachedString) as SettingsData;
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
        const SHEET_NAME = "Settings";

        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

        const settingsSheet = spreadsheet.getSheetByName(SHEET_NAME);

        if (!settingsSheet) {
            Logger.log("Settings sheet not found.");
            try { SpreadsheetApp.getUi().alert("Settings sheet not found."); } catch (e) { }
            return { error: "Settings sheet not found." };
        }

        const rawGrid: any[][] = settingsSheet.getDataRange().getValues();
        /**
         * Safely extracts and maps paired columns from the monolithic raw grid block.
         * Handles native string typecasts for flags, floating numbers, and timestamps.
         * @param keyColIdx Column Index for Keys (e.g., Column A = 0, C = 2, E = 4)
         * @param valColIdx Column Index for Values (e.g., Column B = 1, D = 3, F = 5)
         * @param maxRows The boundary limit of configurations inside this subgroup
         */
        const extractGroup = (keyColIdx: number, valColIdx: number, maxRows: number): any => {
            const resultObject: any = {};
            for (let i = 0; i < maxRows; i++) {
                const row = rawGrid[i];
                if (!row) break;

                const key = row[keyColIdx];
                let val = row[valColIdx];

                // Ensure key exists and isn't whitespace padding
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

        const importantDates = extractGroup(0, 1, 12) as ImportantDates;
        const standardPercentages = extractGroup(2, 3, 5) as StandardPercentages;
        const mandatedPolicies = extractGroup(4, 5, 3) as MandatedPolicies;
        const ledgersAndRecords = extractGroup(6, 7, 5) as LedgersAndRecords;
        const limits = extractGroup(8, 9, 3) as Limits;

        return {
            importantDates,
            standardPercentages,
            mandatedPolicies,
            ledgersAndRecords,
            limits
        };
    } catch (err: any) {
        Logger.log(`Error in fetchSettingsData: ${err.message}`);
        try { SpreadsheetApp.getUi().alert(`Error in fetchSettingsData: ${err.message}`); } catch (e) { }
        return { error: `Error in fetchSettingsData: ${err.message}` };
    }
}