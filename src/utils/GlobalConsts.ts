/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran. 
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 * 
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 * 
 * This file contains global constants used throughout the B-Bucks Economy Scripts project.
 */

// ---- ADVANCED SETTINGS - Don't modify these values unless you know what you're doing (or you're me) -----
const WAIT_LOCK_TIME: number = 5000; // in ms
const SERVER_SIDE_CACHE_AGE: number = 21600; // in seconds

// this is separated and not in an object because separate variables are easier to use since these are used pretty often (and are cleaner)
const SERVICES_CACHED_KEY: string = "cachedServices";
const SETTINGS_CACHED_KEY: string = "cachedSettings";
const TRANSACTIONS_CACHED_KEY: string = "cachedTransactions";
const INVESTMENTS_LEDGER_CACHED_KEY: string = "cachedInvestmentsLedger";

// ----- Dynamic Settings - These values are fetched from the settings sheet and can be modified by the user -----

// --- Main Sheet ---
const USER_STARTING_ROW: number = fetchProperty("userStartingRow", "advancedTechnicalSettings") as number || 7;

// -- columns --
// this isn't in an object btw bc i'm dumb and also bc i didnt think abt it when starting this project
// if you are maintaining this in the future, just put it in a key-value pair object
// also bc im lazy asf so yeah i didnt do it. what are you gonna do abt it
const NAMES_COL: number = fetchProperty("namesCol", "advancedTechnicalSettings") as number || 1;
const BALANCE_COL: number = fetchProperty("balanceCol", "advancedTechnicalSettings") as number || 2;
const EARNINGS_COL: number = fetchProperty("earningsCol", "advancedTechnicalSettings") as number || 3;
const NET_INCOME_COL: number = fetchProperty("netIncomeCol", "advancedTechnicalSettings") as number || 4;
const EXPENDITURES_COL: number = fetchProperty("expendituresCol", "advancedTechnicalSettings") as number || 5;
const INVESTMENT_RETURNS_COL: number = fetchProperty("investmentReturnsCol", "advancedTechnicalSettings") as number || 6;

// investments
const INITIAL_DEPOSIT_COL: number = fetchProperty("initialDepositCol", "advancedTechnicalSettings") as number || 7;
const DATE_DEPOSIT_COL: number = fetchProperty("timeDepositCol", "advancedTechnicalSettings") as number || 8;
const GROSS_INVESTMENT_GAIN_COL: number = fetchProperty("grossInvestmentGainCol", "advancedTechnicalSettings") as number || 9;
const NET_INVESTMENT_GAIN_COL: number = fetchProperty("netInvestmentGainCol", "advancedTechnicalSettings") as number || 10;
const NET_PERCENTAGE_GAIN_COL: number = fetchProperty("netPercentageGainCol", "advancedTechnicalSettings") as number || 11;

// --- Services Sheet ---
const DEFAULT_SERVICES_SHEET: string = fetchProperty("defaultServicesSheet", "advancedTechnicalSettings") as string || "Services";
const SERVICES_ROW_START: number = fetchProperty("servicesRowStart", "advancedTechnicalSettings") as number || 3;

const SERVICES_COLUMNS: { [key: string]: number | number[] } = {
    itemName: 1,
    category: 2,
    // Represents Q1-Q4 pricing columns
    pricing: [3, 4, 5, 6],
    // Represents Q1-Q4 max per person columns
    limit: [7, 8, 9, 10]
}

// --- Settings Sheet ---
// -- DO NOT CHANGE THESE COLUMNS UNLESS YOU KNOW WHAT YOU ARE DOING --
// no but really actually dont pls ty, the sheet is named settings cuz the code likes it and ppl are stupid asf :D
const DEFAULT_SETTINGS_SHEET: string = fetchProperty("defaultSettingsSheet", "advancedTechnicalSettings") as string || "Settings";
const SETTINGS_ROW_START: number = fetchProperty("settingsRowStart", "advancedTechnicalSettings") as number || 4;

// change only if u dont want the default option which is like wai bro ts beautiful
const SETTINGS_COLUMNS: { [key: string]: { keyCol: number; valCol: number } } = {
    importantDates: {
        keyCol: 2,
        valCol: 3
    },
    standardPercentages: {
        keyCol: 5,
        valCol: 6
    },
    mandatedPolicies: {
        keyCol: 8,
        valCol: 9
    },
    ledgersAndRecords: {
        keyCol: 11,
        valCol: 12
    },
    limits: {
        keyCol: 14,
        valCol: 15
    },
    advancedTechnicalSettings: {
        keyCol: 17,
        valCol: 18
    }
}

// --- Transaction Sheet ---
const DEFAULT_TRANSACTIONS_SHEET: string = fetchProperty("defaultTransactionsSheet", "advancedTechnicalSettings") as string || "Transactions";
const TRANSACTIONS_ROW_START: number = fetchProperty("transactionsRowStart", "advancedTechnicalSettings") as number || 3;

// -- Fill Names Utils ---
const STARTING_CELL: string = fetchProperty("fillStartingCell", "advancedTechnicalSettings") as string || "A7";
const SHEET_CELL_NAME: string = fetchProperty("fillSheetCellName", "advancedTechnicalSettings") as string || "A2";

// -- Historical Records Sheet ---
const DEFAULT_HISTORICAL_RECORDS_SHEET: string = fetchProperty("defaultHistoricalRecordsSheet", "advancedTechnicalSettings") as string || "Historical Records";
const HISTORICAL_RECORDS_ROW_START: number = fetchProperty("historicalRecordsRowStart", "advancedTechnicalSettings") as number || 11;
const METRIC_DATA_START_COL: number = fetchProperty("metricDataStartCol", "advancedTechnicalSettings") as number || 3;
const PERIOD_COL_INDEX: number = fetchProperty("periodColIndex", "advancedTechnicalSettings") as number || 2;
const HEADER_ROW_INDEX: number = fetchProperty("headerRowIndex", "advancedTechnicalSettings") as number || 2;
const SHEET_NAME_PATTERN: string = fetchProperty("sheetNamePattern", "advancedTechnicalSettings") as string || "records";

// -- Timestamp Sheets --
const TIMESTAMP_LIST: string[] = [
    "Period 1",
    "Period 2",
    "Period 3",
    "Period 4",
    "Period 5",
    "Period 6",
    "Period 7",
    "Period 8",
    "Period 9",
    "Period 10",
    "Template",
    "Testing"
];

const TIMESTAMP_CELL: string = fetchProperty("timestampCell", "advancedTechnicalSettings") as string || "A5";