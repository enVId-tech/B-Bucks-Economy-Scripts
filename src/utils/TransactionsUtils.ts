// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. 
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains utility functions for managing transactions in the B-Bucks Economy Scripts project, including functions for executing balance actions based on user input and fetching transaction data with caching to optimize performance.

// Transaction record interfaces
type TransactionType = "Income" | "Expense" | "Investment" | "Unknown";

interface TransactionRecord {
  id?: number
  individual?: string;
  type?: TransactionType;
  serviceProvided?: string;
  quantity?: number;
  modifiedColumn?: number;
  tenderedMoney?: number;
  initialColumnAmount?: number;
  newColumnAmount?: number;
  initialBalance?: number;
  newBalance?: number;
  timestamp?: string;
}

/**
 * Fetches transactions data with caching. It first checks for cached data to minimize latency, and if not found or if a force refresh is requested, it reads the transactions data from the sheet and updates the cache with the new data. This function ensures that the application can quickly access transactions data while also providing a mechanism to refresh the data when necessary.
 * @param data A string containing the data for the function, including a forceRefresh flag. Defaults to undefined, meaning it will use cached data if available for faster access.
 * @returns {TransactionRecord[] | { error: string }} An array of transaction records or an error message if the sheet is not found or an error occurs.
 */
function fetchTransactionsDataCached(data?: string): TransactionRecord[] | { error: string } {
  try {
    if (data && typeof data === 'string') {
      Logger.log(`Received data for fetchTransactionsDataCached: ${data}`);
    } else {
      Logger.log("No data received for fetchTransactionsDataCached, proceeding with default cache retrieval.");
      data = JSON.stringify({ forceRefresh: false });
    }

    const parsedData = data ? JSON.parse(data) : null;
    const forceRefresh = parsedData?.forceRefresh || false;

    const cache = CacheService.getScriptCache();
    const props = PropertiesService.getScriptProperties();

    if (!forceRefresh) {
      const cachedString = getCachedData(TRANSACTIONS_CACHED_KEY);
      if (cachedString && cachedString !== "{}" && cachedString !== "") {
        // SpreadsheetApp.getUi().alert(`Cache hit: Transactions data loaded from cache. String: ${cachedString}`);
        return JSON.parse(cachedString) as TransactionRecord[];
      }
      const savedProperties = props.getProperty(TRANSACTIONS_CACHED_KEY);
      if (savedProperties) {
        cache.put(TRANSACTIONS_CACHED_KEY, savedProperties, SERVER_SIDE_CACHE_AGE);
        return JSON.parse(savedProperties);
      }
    }

    console.log("Cache miss: Re-extracting transactions from sheet rows...");
    // SpreadsheetApp.getUi().alert("Cache miss: Re-extracting transactions from sheet rows...");
    const freshTransactions = fetchTransactionsData();
    setCachedData(TRANSACTIONS_CACHED_KEY, JSON.stringify(freshTransactions));
    if (!Array.isArray(freshTransactions)) throw new Error("Failed to fetch transactions data from sheet.");
    return freshTransactions;
  } catch (error: any) {
    Logger.log(`Error in fetchTransactionsDataCached: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error in fetchTransactionsDataCached: ${error.message}`);
    return { error: `Error in fetchTransactionsDataCached: ${error.message}` };
  }
}


/**
 * Adds a transaction record to the "Transactions Records" sheet with the provided details, ensuring that all required information is valid and properly formatted.
 * Uses the Google Sheets API for efficient appending of transaction records, with error handling to fall back to the slower method if the API call fails.
 * @param records An array of transaction records to be added, where each record includes the individual's name, transaction type (Income, Expense, or Investment), service description, initial amount, tendered amount, final amount, quantity of services, and timestamp. All fields are required for each record.
 * @returns {boolean} Returns true if the operation was successful, false otherwise.
 */
function addTransactionRecords(records: TransactionRecord[]): boolean {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(DEFAULT_TRANSACTIONS_SHEET);

    // -- Add edge case checking for the sheet and all fields to ensure data integrity when compiled to JavaScript --
    if (!sheet) {
      Logger.log(`Sheet "${DEFAULT_TRANSACTIONS_SHEET}" not found.`);
      SpreadsheetApp.getUi().alert(`Sheet "${DEFAULT_TRANSACTIONS_SHEET}" not found.`);
      return false;
    }

    if (!records || records.length === 0) {
      Logger.log("No records provided to add.");
      return true;
    }

    // Fetch grid boundaries in a single batch to minimize API fetch overhead
    const lastRowWithData = sheet.getLastRow();
    const currentMaxRows = sheet.getMaxRows();
    let biggestId = 0;

    if (lastRowWithData >= (TRANSACTIONS_ROW_START - 1) && lastRowWithData > 0) {
      const rawIdValue = sheet.getRange(lastRowWithData, 1).getValue();
      const parsedId = parseInt(rawIdValue, 10);
      biggestId = isNaN(parsedId) ? 0 : parsedId;
    }

    const insertStartRow = Math.max(lastRowWithData + 1, TRANSACTIONS_ROW_START);
    const rowsNeeded = records.length;

    // Expand the sheet grid at the very last moment if required
    if ((insertStartRow - 1) + rowsNeeded > currentMaxRows) {
      const rowsToAdd = ((insertStartRow - 1) + rowsNeeded) - currentMaxRows;
      sheet.insertRowsAfter(currentMaxRows, rowsToAdd);
    }

    // Prepare for O(1) access with O(N) preprocessing
    const values = new Array(rowsNeeded);

    for (let i = 0; i < rowsNeeded; i++) {
      const record: TransactionRecord = records[i];

      // Native, short-circuiting check. Fast memory lookup.
      if (
        Object.keys(record)
          .some(key => record[key as keyof TransactionRecord] === undefined || record[key as keyof TransactionRecord] === null)
      ) {
        const errMsg = `Validation failed: A required field is missing.`;
        Logger.log(errMsg);
        SpreadsheetApp.getUi().alert(errMsg);
        return false;
      }

      // If valid, map directly to the row matrix array
      values[i] = [
        biggestId + i + 1,
        ...Object.values(record)
      ];
    }

    // Use the Sheets API to efficiently add the record in one request, with error handling to fall back to the slower method if the API call fails
    try {
      if (typeof Sheets !== 'undefined' && Sheets.Spreadsheets && Sheets.Spreadsheets.Values) {
        const spreadsheetId = spreadsheet.getId();

        const resource: GoogleAppsScript.Sheets.Schema.ValueRange = {
          values: values
        };

        // Standardized append structure targeting structural range boundary lookup
        Sheets.Spreadsheets.Values.append(
          resource,
          spreadsheetId,
          `${DEFAULT_TRANSACTIONS_SHEET}!A${insertStartRow}`,
          { valueInputOption: "USER_ENTERED" }
        );

        // Cache invalidation so that transaction records will be refetched
        clearGlobalCache([TRANSACTIONS_CACHED_KEY]);

        return true;
      } else {
        throw new Error("Advanced Sheets API service not enabled in script settings.");
      }
    } catch (apiError: any) {
      Logger.log(`Advanced API pipeline bypassed/failed. Error: ${apiError.message}. Running native fallback setup...`);

      sheet.getRange(
        insertStartRow,
        1,
        values.length,
        values[0].length
      ).setValues(values);

      // Cache invalidation so that transaction records will be refetched
      clearGlobalCache([TRANSACTIONS_CACHED_KEY]);

      return true;
    }
  } catch (error: any) {
    Logger.log(`Error occurred in addTransactionRecords: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error occured in addTransactionRecord: ${error.message}`);
    return false;
  }
}

/**
 * Fetches transaction records from the spreadsheet and returns them as an array of TransactionRecord objects. 
 * If the transactions sheet is not found or an error occurs during the fetch operation, it returns false.
 * @param {string} transactionId - The ID of the transaction to fetch. If not provided, all transactions will be fetched.
 * @returns {TransactionRecord[] | boolean} - An array of TransactionRecord objects if successful, or false if an error occurs or the sheet is not found.
 */
function fetchTransactionsData(): TransactionRecord[] | boolean {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(DEFAULT_TRANSACTIONS_SHEET);
    let sheetData: any | undefined = undefined;

    if (!spreadsheet || !sheet || sheet === undefined || sheet === null) {
      SpreadsheetApp.getUi().alert("Unable to fetch transaction records sheet. Records will not be filled");
      return false;
    }

    try {
      if (typeof Sheets === 'undefined' || !Sheets.Spreadsheets || !Sheets.Spreadsheets.Values) {
        throw new Error("Advanced Sheets API service not enabled in script settings")
      }
      // Use the Sheets API to efficiently fetch all transaction records in one request, with error handling to fall back to the slower method if the API call fails
      sheetData = Sheets.Spreadsheets.Values.get(spreadsheet.getId(), `${DEFAULT_TRANSACTIONS_SHEET}!A${TRANSACTIONS_ROW_START}:L`);

      if (!sheetData.values || sheetData.values.length <= 0) return false;
    } catch (err: any) {
      Logger.log(`Advanced API pipeline bypassed/failed. Error: ${err.message}. Running native fallback setup...`);

      // Fall back to the slower method of fetching all transaction records using the native SpreadsheetApp service
      sheetData = sheet.getRange(TRANSACTIONS_ROW_START, 1, sheet.getLastRow() - TRANSACTIONS_ROW_START + 1, 12).getValues(); // A1:L
    }

    const transactionRecords: TransactionRecord[] = sheetData.values.map((row: any[]) => {
      return {
        id: row[0] as number | 0,
        individual: row[1] as string | "",
        type: row[2] as TransactionType | "Unknown",
        serviceProvided: row[3] as string | "",
        quantity: row[4] as number | 0,
        modifiedColumn: row[5] as number | 0,
        tenderedMoney: row[6] as number | 0,
        initialColumnAmount: row[7] as number | 0,
        newColumnAmount: row[8] as number | 0,
        initialBalance: row[9] as number | 0,
        newBalance: row[10] as number | 0,
        timestamp: row[11] as string | 0
      }
    })

    return transactionRecords;
  } catch (error: any) {
    Logger.log(`An error occurred in fetchTransactionRecords: ${error.message}`);
    SpreadsheetApp.getUi().alert(`An error occurred in fetchTransactionRecords: ${error.message}`);
    return false;
  }
}