// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. 
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains global constants used throughout the B-Bucks Economy Scripts project.

// ADVANCED SETTINGS - Don't modify these values unless you know what you're doing (or you're me)
const WAIT_LOCK_TIME: number = 5000; // in ms
const SERVER_SIDE_CACHE_AGE: number = 21600; // in seconds

// this is separated and not in an object because separate variables are easier to use since these are used pretty often (and are cleaner)
const SERVICES_CACHED_KEY: string = "cachedServices";
const SETTINGS_CACHED_KEY: string = "cachedSettings";
const TRANSACTIONS_CACHED_KEY: string = "cachedTransactions";
const INVESTMENTS_LEDGER_CACHED_KEY: string = "cachedInvestmentsLedger";

// --- Main Sheet ---
const USER_STARTING_ROW: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { userStartingRow?: number } }).advancedTechnicalSettings?.userStartingRow ?? 7;

// -- columns --
// this isn't in an object btw bc i'm dumb and also bc i didnt think abt it when starting this project
// if you are maintaining this in the future, just put it in a key-value pair object
// also bc im lazy asf so yeah i didnt do it. what are you gonna do abt it
const NAMES_COL: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { namesCol?: number } }).advancedTechnicalSettings?.namesCol ?? 1;
const BALANCE_COL: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { balanceCol?: number } }).advancedTechnicalSettings?.balanceCol ?? 2;
const EARNINGS_COL: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { earningsCol?: number } }).advancedTechnicalSettings?.earningsCol ?? 3;
const NET_INCOME_COL: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { netIncomeCol?: number } }).advancedTechnicalSettings?.netIncomeCol ?? 4;
const EXPENDITURES_COL: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { expendituresCol?: number } }).advancedTechnicalSettings?.expendituresCol ?? 5;
const INVESTMENT_RETURNS_COL: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { investmentReturnsCol?: number } }).advancedTechnicalSettings?.investmentReturnsCol ?? 6;

// investments
const INITIAL_DEPOSIT_COL: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { initialDepositCol?: number } }).advancedTechnicalSettings?.initialDepositCol ?? 7;
const DATE_DEPOSIT_COL: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { dateDepositCol?: number } }).advancedTechnicalSettings?.dateDepositCol ?? 8;
const GROSS_INVESTMENT_GAIN_COL: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { grossInvestmentGainCol?: number } }).advancedTechnicalSettings?.grossInvestmentGainCol ?? 9;
const NET_INVESTMENT_GAIN_COL: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { netInvestmentGainCol?: number } }).advancedTechnicalSettings?.netInvestmentGainCol ?? 10;
const NET_PERCENTAGE_GAIN_COL: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { netPercentageGainCol?: number } }).advancedTechnicalSettings?.netPercentageGainCol ?? 11;

// --- Services Sheet ---
const DEFAULT_SERVICES_SHEET: string = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { defaultServicesSheet?: string } }).advancedTechnicalSettings?.defaultServicesSheet ?? "Services";
const SERVICES_ROW_START: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { servicesRowStart?: number } }).advancedTechnicalSettings?.servicesRowStart ?? 3;

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
const DEFAULT_SETTINGS_SHEET: string = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { defaultSettingsSheet?: string } }).advancedTechnicalSettings?.defaultSettingsSheet ?? "Settings";
const SETTINGS_ROW_START: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { settingsRowStart?: number } }).advancedTechnicalSettings?.settingsRowStart ?? 4;

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
const DEFAULT_TRANSACTIONS_SHEET: string = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { defaultTransactionsSheet?: string } }).advancedTechnicalSettings?.defaultTransactionsSheet ?? "Transactions";
const TRANSACTIONS_ROW_START: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { transactionsRowStart?: number } }).advancedTechnicalSettings?.transactionsRowStart ?? 3;

// -- Fill Names Utils ---
const STARTING_CELL: string = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { startingCell?: string } }).advancedTechnicalSettings?.startingCell ?? "A7";
const SHEET_CELL_NAME: string = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { sheetCellName?: string } }).advancedTechnicalSettings?.sheetCellName ?? "A2";

// -- Historical Records Sheet ---
const DEFAULT_HISTORICAL_RECORDS_SHEET: string = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { defaultHistoricalRecordsSheet?: string } }).advancedTechnicalSettings?.defaultHistoricalRecordsSheet ?? "Historical Records";
const HISTORICAL_RECORDS_ROW_START: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { historicalRecordsRowStart?: number } }).advancedTechnicalSettings?.historicalRecordsRowStart ?? 11;
const METRIC_DATA_START_COL: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { metricDataStartCol?: number } }).advancedTechnicalSettings?.metricDataStartCol ?? 3;
const PERIOD_COL_INDEX: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { periodColIndex?: number } }).advancedTechnicalSettings?.periodColIndex ?? 2;
const HEADER_ROW_INDEX: number = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { headerRowIndex?: number } }).advancedTechnicalSettings?.headerRowIndex ?? 2;
const SHEET_NAME_PATTERN: string = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { sheetNamePattern?: string } }).advancedTechnicalSettings?.sheetNamePattern ?? "records";

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

const TIMESTAMP_CELL: string = (fetchSettingsDataCached() as { advancedTechnicalSettings?: { timestampCell?: string } }).advancedTechnicalSettings?.timestampCell ?? "A5";