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

// ---------------------------------------------------------------------------------------------------------
// LOAD-ORDER CONTRACT (read before editing this file)
//
// Apps Script evaluates every file's top-level statements before any entry point runs, in the order the files
// were pushed. Anything executed at the top level of this file therefore runs BEFORE the functions and constants
// of files that load after it. To keep initialisation deterministic this file follows two strict rules:
//
//   TIER 0 (this file): "bootstrap" constants. Pure literals, or expressions built ONLY from constants declared
//   earlier in this same file. They must never call a function, read a sheet, or touch CacheService/PropertiesService.
//   Everything else in the project (including the settings loader itself) depends on these, so they have no
//   dependencies of their own.
//
//   TIER 1 (utils/Config.ts): values that come from the Settings sheet (e.g. userStartingRow). They are resolved
//   lazily through getConfig() the first time they are needed, memoised for the rest of that execution, and fall
//   back to a hard-coded default if the sheet cannot be read. See getUserStartingRow().
//
// The old design evaluated `fetchProperty()` at the top level of this file to populate USER_STARTING_ROW. That call
// needed DEFAULT_SETTINGS_SHEET, SETTINGS_COLUMNS, SETTINGS_ROW_START and the cache keys, which were declared further
// down this very file (still in the temporal dead zone), so settings loading was circular and could fail or return
// an error object. Do not reintroduce a top-level call here.
// ---------------------------------------------------------------------------------------------------------

// ---- ADVANCED SETTINGS - Don't modify these values unless you know what you're doing (or you're me) -----
const WAIT_LOCK_TIME: number = 5000; // in ms
const SERVER_SIDE_CACHE_AGE: number = 21600; // in seconds

// this is separated and not in an object because separate variables are easier to use since these are used pretty often (and are cleaner)
const SETTINGS_CACHED_KEY: string = "cachedSettings";
const SERVICES_CACHED_KEY: string = "cachedServices";
const TRANSACTIONS_CACHED_KEY: string = "cachedTransactions";
const INVESTMENTS_LEDGER_CACHED_KEY: string = "cachedInvestmentsLedger";

// DO NOT MODIFY THESE VALUES, IT WILL BREAK CACHING
// Upper limits based on Apps Script Quotas
const MAX_CACHE_ENTRIES = 50; // Maximum number of distinct key entries tracked
const MAX_PROPERTY_BYTES = 8500; // ~8.5KB safety cap for PropertiesService (Quota is 9KB)
const CACHE_INDEX_KEY = '__GLOBAL_CACHE_INDEX__'; // Key used to maintain insertion order / LRU

// ----- Rolling transaction cache (see utils/TransactionCache.ts) -----
// The transaction log grows without bound, so it is never cached as one payload. Instead the sheet is split into
// fixed-size BLOCKS of rows (anchored at the oldest row, so appending a transaction only ever touches the newest
// block). Only the blocks around the page the admin is looking at are kept, and the window re-centres every time the
// admin switches pages. Tune these three numbers to trade memory for fewer sheet reads.
const TRANSACTION_BLOCK_SIZE: number = 50;               // rows per cache block (~20KB each, far below the 100KB CacheService value cap)
const TRANSACTION_CACHE_MAX_BLOCKS: number = 24;         // hard ceiling on blocks held at once (24 * 50 = 1,200 rows)
const TRANSACTION_CACHE_MAX_BYTES: number = 1200000;     // soft ceiling on the combined serialized size of all cached blocks (~1.2MB)
const TRANSACTION_PREFETCH_BEHIND_PAGES: number = 1;     // pages kept behind the viewed page when the direction is unknown
const TRANSACTION_PREFETCH_AHEAD_PAGES: number = 2;      // pages kept ahead of the viewed page in the direction of travel
const TRANSACTION_MAX_PAGE_SIZE: number = 200;           // largest "Displayed" value the server will honour
const TRANSACTION_COLUMN_COUNT: number = 16;             // A..P on the Transactions sheet
const TRANSACTION_BLOCK_KEY_PREFIX: string = TRANSACTIONS_CACHED_KEY + "_blk_";   // + block number
const TRANSACTION_INDEX_KEY: string = TRANSACTIONS_CACHED_KEY + "_idx";           // block bookkeeping (JSON)
const TRANSACTION_FACETS_KEY: string = TRANSACTIONS_CACHED_KEY + "_facets";       // distinct periods / students / types
const TRANSACTION_STATS_KEY: string = TRANSACTIONS_CACHED_KEY + "_stats";         // incremental totals for the unfiltered log

// ----- Sheet layout DEFAULTS (Tier 0) -----
// Every DEFAULT_* value below is only a fallback. The real value is whatever the Settings sheet says under
// "Advanced Technical Settings", read through the matching accessor in utils/Config.ts (for example getNamesColumn()).
// Never use a DEFAULT_* constant directly outside Config.ts: a test fails if you do.
// Exceptions (BOOTSTRAP): DEFAULT_SETTINGS_SHEET, SETTINGS_ROW_START and SETTINGS_COLUMNS are what the loader uses to
// find the Settings sheet in the first place, so they cannot be overridden by it (Config.ts warns if the sheet disagrees).

// --- Main Sheet ---
// The first student row is NOT a constant any more: it is user-configurable ("userStartingRow" in the Settings sheet).
// Use getUserStartingRow() (utils/Config.ts), which resolves it lazily and always returns a validated integer.
const DEFAULT_USER_STARTING_ROW: number = 7;

// -- columns --
// this isn't in an object btw bc i'm dumb and also bc i didnt think abt it when starting this project
// if you are maintaining this in the future, just put it in a key-value pair object
// also bc im lazy asf so yeah i didnt do it. what are you gonna do abt it
const DEFAULT_NAMES_COL: number = 1; // default for "namesCol" (Settings sheet value wins, read it with getNamesColumn())
const DEFAULT_BALANCE_COL: number = 2; // default for "balanceCol" (Settings sheet value wins, read it with getBalanceColumn())
const DEFAULT_EARNINGS_COL: number = 3; // default for "earningsCol" (Settings sheet value wins, read it with getEarningsColumn())
const DEFAULT_NET_INCOME_COL: number = 4; // default for "netIncomeCol" (Settings sheet value wins, read it with getNetIncomeColumn())
const DEFAULT_EXPENDITURES_COL: number = 5; // default for "expendituresCol" (Settings sheet value wins, read it with getExpendituresColumn())
const DEFAULT_INVESTMENT_RETURNS_COL: number = 6; // default for "investmentReturnsCol" (Settings sheet value wins, read it with getInvestmentReturnsColumn())

// investments
const DEFAULT_INITIAL_DEPOSIT_COL: number = 7; // default for "initialDepositCol" (Settings sheet value wins, read it with getInitialDepositColumn())
const DEFAULT_DATE_DEPOSIT_COL: number = 8; // default for "timeDepositCol" (Settings sheet value wins, read it with getDepositDateColumn())
const DEFAULT_GROSS_INVESTMENT_GAIN_COL: number = 9; // default for "grossInvestmentGainCol" (Settings sheet value wins, read it with getGrossInvestmentGainColumn())
const DEFAULT_NET_INVESTMENT_GAIN_COL: number = 10; // default for "netInvestmentGainCol" (Settings sheet value wins, read it with getNetInvestmentGainColumn())
const DEFAULT_NET_PERCENTAGE_GAIN_COL: number = 11; // default for "netPercentageGainCol" (Settings sheet value wins, read it with getNetPercentageGainColumn())

// --- Services Sheet ---
const DEFAULT_SERVICES_SHEET: string = "Main Services"; // default for "defaultServicesSheet" (Settings sheet value wins, read it with getServicesSheetName())
const DEFAULT_SERVICES_ROW_START: number = 3; // default for "servicesRowStart" (Settings sheet value wins, read it with getServicesRowStart())

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
const DEFAULT_SETTINGS_SHEET: string = "Settings"; // BOOTSTRAP: needed to find the Settings sheet itself, so it can never come from that sheet
const SETTINGS_ROW_START: number = 4; // BOOTSTRAP: needed to read the Settings sheet itself, so it can never come from that sheet

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
const DEFAULT_TRANSACTIONS_SHEET: string = "Transactions"; // default for "defaultTransactionsSheet" (Settings sheet value wins, read it with getTransactionsSheetName())
const DEFAULT_TRANSACTIONS_ROW_START: number = 3; // default for "transactionsRowStart" (Settings sheet value wins, read it with getTransactionsRowStart())

// -- Fill Names Utils ---
const DEFAULT_STARTING_CELL: string = "A7"; // default for "fillStartingCell" (Settings sheet value wins, read it with getFillStartingCell())
const DEFAULT_SHEET_CELL_NAME: string = "A2"; // default for "fillSheetCellName" (Settings sheet value wins, read it with getFillSheetCellName())

// -- Historical Records Sheet ---
const DEFAULT_HISTORICAL_RECORDS_SHEET: string = "Historical Records"; // default for "defaultHistoricalRecordsSheet" (Settings sheet value wins, read it with getHistoricalRecordsSheetName())
const DEFAULT_HISTORICAL_RECORDS_ROW_START: number = 11; // default for "historicalRecordsRowStart" (Settings sheet value wins, read it with getHistoricalRecordsRowStart())
const DEFAULT_METRIC_DATA_START_COL: number = 3; // default for "metricDataStartColumn" (Settings sheet value wins, read it with getMetricDataStartColumn())
const DEFAULT_PERIOD_COL_INDEX: number = 2; // default for "periodColIndex" (Settings sheet value wins, read it with getPeriodColumnIndex())
const DEFAULT_HEADER_ROW_INDEX: number = 2; // default for "headerRowIndex" (Settings sheet value wins, read it with getHeaderRowIndex())
const DEFAULT_SHEET_NAME_PATTERN: string = "records"; // default for "sheetNamePattern" (Settings sheet value wins, read it with getSheetNamePattern())

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

const DEFAULT_TIMESTAMP_CELL: string = "A5"; // default for "timestampCell" (Settings sheet value wins, read it with getTimestampCell())