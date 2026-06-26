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
const SERVICES_SHEET_NAME = "Services";
const SERVICES_ROW_START = 3;

const SERVICES_COLUMNS: { [key: string]: number | number[] } = {
    itemName: 1,
    category: 2,
    // Represents Q1-Q4 pricing columns
    pricing: [3, 4, 5, 6],
    // Represents Q1-Q4 max per person columns
    limit: [7, 8, 9, 10]
}

// ADVANCED SETTINGS - Don't modify these values unless you know what you're doing (or you're me)
const WAIT_LOCK_TIME = 5000; // in ms
const SERVER_SIDE_CACHE_AGE = 21600; // in seconds

const SERVICES_CACHED = "cachedServices";
const SETTINGS_CACHED = "cachedSettings";
const TRANSACTIONS_CACHED = "cachedTransactions";
const INVESTMENTS_LEDGER_CACHED = "cachedInvestmentsLedger";