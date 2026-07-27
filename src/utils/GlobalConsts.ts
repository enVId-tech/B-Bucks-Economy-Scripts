// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. 
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains global constants used throughout the B-Bucks Economy Scripts project.

// --- Main Sheet ---
const USER_STARTING_ROW = 7;

// -- columns --
// this isn't in an object btw bc i'm dumb and also bc i didnt think abt it when starting this project
// if you are maintaining this in the future, just put it in a key-value pair object
// also bc im lazy asf so yeah i didnt do it. what are you gonna do abt it
const BALANCE_COL = 2;
const EARNINGS_COL = 3;
const NET_INCOME_COL = 4;
const EXPENDITURES_COL = 5;
const INVESTMENT_RETURNS_COL = 6;

// investments
const INITIAL_DEPOSIT_COL = 7;
const DATE_DEPOSIT_COL = 8;
const GROSS_INVESTMENT_GAIN_COL = 9;
const NET_INVESTMENT_GAIN_COL = 10;
const NET_PERCENTAGE_GAIN_COL = 11;

// --- Services Sheet ---
const DEFAULT_SERVICES_SHEET = "Services";
const SERVICES_ROW_START = 3;

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
const DEFAULT_SETTINGS_SHEET = "Settings";
const SETTINGS_ROW_START = 4;

// change only if u dont want the default option which is like wai bro ts beautiful
const SETTINGS_COLUMNS = {
    importantDates: { keyCol: 2, valCol: 3 },
    standardPercentages: { keyCol: 5, valCol: 6 },
    mandatedPolicies: { keyCol: 8, valCol: 9 },
    ledgersAndRecords: { keyCol: 11, valCol: 12 },
    limits: { keyCol: 14, valCol: 15 },
}

// --- Transaction Sheet ---
const DEFAULT_TRANSACTIONS_SHEET = "Transactions";
const TRANSACTIONS_ROW_START = 3;

// -- Fill Names Utils ---
const STARTING_CELL = "A7";
const SHEET_CELL_NAME = "A2";

// -- Historical Records Sheet ---
const DEFAULT_HISTORICAL_RECORDS_SHEET = "Economic Records";
const HISTORICAL_RECORDS_ROW_START = 11;


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
]
const TIMESTAMP_CELL = "A5";

// ADVANCED SETTINGS - Don't modify these values unless you know what you're doing (or you're me)
const WAIT_LOCK_TIME = 5000; // in ms
const SERVER_SIDE_CACHE_AGE = 21600; // in seconds

// this is separated and not in an object because separate variables are easier to use since these are used pretty often (and are cleaner)
const SERVICES_CACHED_KEY = "cachedServices";
const SETTINGS_CACHED_KEY = "cachedSettings";
const TRANSACTIONS_CACHED_KEY = "cachedTransactions";
const INVESTMENTS_LEDGER_CACHED_KEY = "cachedInvestmentsLedger";