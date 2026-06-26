// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. 
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains utility functions for managing transactions in the B-Bucks Economy Scripts project, including functions for executing balance actions based on user input and fetching transaction data with caching to optimize performance.

// Transaction record interface
interface TransactionRecord {
  individual: string;
  type: "Income" | "Expense" | "Investment";
  service: string;
  initialAmount: number;
  tenderedAmount: number;
  tenderedColumn: number;
  quantityOfServices: number;
  timestamp: Date;
}

/**
 * Fetches settings data with caching. It first checks for cached data to minimize latency, and if not found or if a force refresh is requested, it reads the settings data from the sheet and updates the cache with the new data. This function ensures that the application can quickly access settings data while also providing a mechanism to refresh the data when necessary.
 * @param data A string containing the data for the function, including a forceRefresh flag. Defaults to undefined, meaning it will use cached data if available for faster access.
 * @returns {SettingsData | { error: string }} An object containing the structured settings data or an error message if the sheet is not found or an error occurs. The settings data includes important dates, standard percentages, mandated policies, ledgers and records preferences, and limits, all organized into their respective categories for easy access throughout the application.
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

        const CACHE_KEY = "cachedTransactions";
        const cache = CacheService.getScriptCache();
        const props = PropertiesService.getScriptProperties();

        if (!forceRefresh) {
            const cachedString = getCachedData(CACHE_KEY);
            if (cachedString && cachedString !== "{}" && cachedString !== "") {
                // SpreadsheetApp.getUi().alert(`Cache hit: Transactions data loaded from cache. String: ${cachedString}`);
                return JSON.parse(cachedString) as TransactionRecord[];
            }
            const savedProperties = props.getProperty(CACHE_KEY);
            if (savedProperties) {
                cache.put(CACHE_KEY, savedProperties, 21600);
                return JSON.parse(savedProperties);
            }
        }

        console.log("Cache miss: Re-extracting transactions from sheet rows...");
        // SpreadsheetApp.getUi().alert("Cache miss: Re-extracting transactions from sheet rows...");
        const freshTransactions = fetchTransactionsData();
        setCachedData(CACHE_KEY, JSON.stringify(freshTransactions));
        return freshTransactions;
     } catch (error: any) {
        Logger.log(`Error in fetchTransactionsDataCached: ${error.message}`);
        SpreadsheetApp.getUi().alert(`Error in fetchTransactionsDataCached: ${error.message}`);
        return { error: `Error in fetchTransactionsDataCached: ${error.message}` };
    }
}

function fetchTransactionsData(): TransactionRecord[] {
    return [];
}

/**
 * Adds a transaction record to the "Transactions Records" sheet with the provided details, ensuring that all required information is valid and properly formatted.
 * Uses the Google Sheets API for efficient appending of transaction records, with error handling to fall back to the slower method if the API call fails.
 * @param records An array of transaction records to be added, where each record includes the individual's name, transaction type (Income, Expense, or Investment), service description, initial amount, tendered amount, final amount, quantity of services, and timestamp. All fields are required for each record.
 * @returns {boolean} Returns true if the operation was successful, false otherwise.
 */
function addTransactionRecords(records: TransactionRecord[]): boolean {
  try {
    const SHEET_NAME = "Transactions";
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    const ROW_TO_START_FROM = 3;

    // -- Add edge case checking for the sheet and all fields to ensure data integrity when compiled to JavaScript --
    if (!sheet) {
      Logger.log(`Sheet "${SHEET_NAME}" not found.`);
      SpreadsheetApp.getUi().alert(`Sheet "${SHEET_NAME}" not found.`);
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

    if (lastRowWithData >= (ROW_TO_START_FROM - 1) && lastRowWithData > 0) {
      const rawIdValue = sheet.getRange(lastRowWithData, 1).getValue();
      const parsedId = parseInt(rawIdValue, 10);
      biggestId = isNaN(parsedId) ? 0 : parsedId;
    }

    const insertStartRow = Math.max(lastRowWithData + 1, ROW_TO_START_FROM);
    const rowsNeeded = records.length;

    // Expand the sheet grid at the very last moment if required
    if ((insertStartRow - 1) + rowsNeeded > currentMaxRows) {
      const rowsToAdd = ((insertStartRow - 1) + rowsNeeded) - currentMaxRows;
      sheet.insertRowsAfter(currentMaxRows, rowsToAdd);
    }

    // Prepare for O(1) access with O(N) preprocessing
    const values = new Array(rowsNeeded);

    for (let i = 0; i < rowsNeeded; i++) {
      const record = records[i];

      // Native, short-circuiting check. Fast memory lookup.
      if (
        record.individual === undefined || record.individual === null ||
        record.type === undefined || record.type === null ||
        record.service === undefined || record.service === null ||
        record.initialAmount === undefined || record.initialAmount === null ||
        record.tenderedAmount === undefined || record.tenderedAmount === null ||
        record.tenderedColumn === undefined || record.tenderedColumn === null ||
        record.quantityOfServices === undefined || record.quantityOfServices === null ||
        record.timestamp === undefined || record.timestamp === null
      ) {
        const errMsg = `Validation failed: A required field is missing.`;
        Logger.log(errMsg);
        SpreadsheetApp.getUi().alert(errMsg);
        return false;
      }

      // If valid, map directly to the row matrix array
      values[i] = [
        biggestId + i + 1,
        record.individual,
        record.type,
        record.service,
        Number(record.initialAmount.toFixed(2)),
        Number(record.tenderedAmount.toFixed(2)),
        Number(record.tenderedColumn.toFixed(2)),
        Number(record.quantityOfServices.toFixed(2)),
        record.timestamp instanceof Date ? record.timestamp.toISOString() : new Date(record.timestamp).toISOString()
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
          `${SHEET_NAME}!A${insertStartRow}`,
          { valueInputOption: "USER_ENTERED" }
        );

        // Cache invalidation so that transaction records will be refetched
        clearGlobalCache(["transactionRecords"]);

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
      clearGlobalCache(["transactionRecords"]);

      return true;
    }
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(`Error occured in addTransactionRecord: ${error.message}`);
    return false;
  }
}