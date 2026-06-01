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
            SpreadsheetApp.getUi().alert("Settings sheet not found.");
            return { error: "Settings sheet not found." };
        }

        const bRange: any[][] = settingsSheet.getRange("B2:B13").getValues();
        const dRange: any[][] = settingsSheet.getRange("D2:D4").getValues();
        const fRange: any[][] = settingsSheet.getRange("F2:F4").getValues();
        const hRange: any[][] = settingsSheet.getRange("H2:H6").getValues();
        const jRange: any[][] = settingsSheet.getRange("J2:J4").getValues();

        const mapColumnToObject = (range: any[][]): any => {
            return range.reduce((obj: any, row: any[]) => {
                const key = row[0];
                const value = row[1];
                if (key !== undefined && value !== undefined) {
                    obj[key] = value;
                }
                return obj;
            }, {});
        };

        const importantDates = mapColumnToObject(bRange);
        const standardPercentages = mapColumnToObject(dRange);
        const mandatedPolicies = mapColumnToObject(fRange);
        const ledgersAndRecords = mapColumnToObject(hRange);
        const limits = mapColumnToObject(jRange);

        return {
            importantDates,
            standardPercentages,
            mandatedPolicies,
            ledgersAndRecords,
            limits
        };
    } catch (err: any) {
        Logger.log(`Error in fetchSettingsData: ${err.message}`);
        SpreadsheetApp.getUi().alert(`Error in fetchSettingsData: ${err.message}`);
        return { error: `Error in fetchSettingsData: ${err.message}` };
    }
}