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