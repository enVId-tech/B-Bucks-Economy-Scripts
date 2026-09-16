/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran. 
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 * 
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 * 
 * This file contains utility functions for managing transactions in the B-Bucks Economy Scripts project, including functions for executing balance actions based on user input and fetching transaction data with caching to optimize performance.
 */

// Transaction record interfaces
type TransactionType = "Income" | "Expense" | "Investment" | "Unknown";

interface TransactionRecord {
  id?: number
  individual?: string;
  period?: number;
  row?: number;
  type?: TransactionType;
  operation?: "ADD" | "SUBTRACT" | "MULTIPLY" | "DIVIDE";
  serviceProvided?: string;
  unitPrice: number;
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
      log(`Received data for fetchTransactionsDataCached: ${data}`, false);
    } else {
      log("No data received for fetchTransactionsDataCached, proceeding with default cache retrieval.", false);
      data = JSON.stringify({ forceRefresh: false });
    }

    const parsedData = data ? JSON.parse(data) : null;
    const forceRefresh = parsedData?.forceRefresh || false;

    const cache = CacheService.getScriptCache();
    const cacheMetaKey = `${TRANSACTIONS_CACHED_KEY}_meta`;
    const cacheChunkSize = 64 * 1024;

    if (!forceRefresh) {
      const cachedChunkCount = Number(cache.get(cacheMetaKey));
      if (Number.isInteger(cachedChunkCount) && cachedChunkCount > 0) {
        const chunkKeys = Array.from(
          { length: cachedChunkCount },
          (_, index) => `${TRANSACTIONS_CACHED_KEY}_${index}`
        );
        const cachedChunks = cache.getAll(chunkKeys);
        const cachedString = chunkKeys.map(key => cachedChunks[key]).join('');

        if (cachedChunks && cachedString.length > 0 && chunkKeys.every(key => cachedChunks[key])) {
          log(`Cache hit: Loaded ${cachedChunkCount} transaction cache chunks.`, false);
          return cachedString as unknown as TransactionRecord[];
        }
      }
    }

    log(`Cache miss: Re-extracting transactions from sheet rows...`, false);
    const freshTransactions = fetchTransactionsData();
    if (!Array.isArray(freshTransactions)) throw new Error("Failed to fetch transactions data from sheet.");
    const serializedTransactions = JSON.stringify(freshTransactions);
    const previousChunkCount = Number(cache.get(cacheMetaKey));
    const chunkCount = Math.max(1, Math.ceil(serializedTransactions.length / cacheChunkSize));
    const chunkEntries: { [key: string]: string } = {};

    for (let index = 0; index < chunkCount; index++) {
      chunkEntries[`${TRANSACTIONS_CACHED_KEY}_${index}`] = serializedTransactions.slice(
        index * cacheChunkSize,
        (index + 1) * cacheChunkSize
      );
    }

    if (Number.isInteger(previousChunkCount) && previousChunkCount > 0) {
      const obsoleteKeys = Array.from(
        { length: previousChunkCount },
        (_, index) => `${TRANSACTIONS_CACHED_KEY}_${index}`
      ).filter(key => !Object.prototype.hasOwnProperty.call(chunkEntries, key));
      if (obsoleteKeys.length > 0) cache.removeAll(obsoleteKeys);
    }

    cache.putAll(chunkEntries, SERVER_SIDE_CACHE_AGE);
    cache.put(cacheMetaKey, String(chunkCount), SERVER_SIDE_CACHE_AGE);
    return serializedTransactions as unknown as TransactionRecord[];
  } catch (error: any) {
    log(`Error in fetchTransactionsDataCached: ${error.message}`, true);
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
      log(`Sheet "${DEFAULT_TRANSACTIONS_SHEET}" not found.`, true);
      return false;
    }

    if (!records || records.length === 0) {
      log("No transaction records provided to add. Operation aborted.", false);
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
        log(`Validation failed: A required field is missing in record at index ${i}. Record: ${JSON.stringify(record)}. Missing fields: ${Object.keys(record).filter(k => record[k as keyof TransactionRecord] === undefined || record[k as keyof TransactionRecord] === null).join(', ')}`, true);
        return false;
      }

      // If valid, map directly to the row matrix array
      values[i] = [
        biggestId + i + 1,
        record.individual,
        record.period,
        record.type,
        record.serviceProvided,
        record.unitPrice,
        record.quantity,
        record.modifiedColumn,
        record.tenderedMoney,
        record.initialColumnAmount,
        record.newColumnAmount,
        record.initialBalance,
        record.newBalance,
        record.timestamp,
        record.row ?? "",
        record.operation ?? ""
      ];
    }

    // Write to the calculated destination directly. Values.append treats the
    // range as a table boundary and can create extra rows on repeated calls.
    try {
      sheet.getRange(
        insertStartRow,
        1,
        values.length,
        values[0].length
      ).setValues(values);

      clearGlobalCache([TRANSACTIONS_CACHED_KEY]);
      return true;
    } catch (apiError: any) {
      log(`Failed to write transaction records at row ${insertStartRow}: ${apiError.message}`, true);
      return false;
    }
  } catch (error: any) {
    log(`Error occurred in addTransactionRecords: ${error.message}`, true);
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
      log(`Transactions sheet "${DEFAULT_TRANSACTIONS_SHEET}" not found.`, true);
      return false;
    }

    try {
      if (typeof Sheets === 'undefined' || !Sheets.Spreadsheets || !Sheets.Spreadsheets.Values) {
        throw new Error("Advanced Sheets API service not enabled in script settings")
      }
      // Use the Sheets API to efficiently fetch all transaction records in one request, with error handling to fall back to the slower method if the API call fails
      sheetData = Sheets.Spreadsheets.Values.get(spreadsheet.getId(), `${DEFAULT_TRANSACTIONS_SHEET}!A${TRANSACTIONS_ROW_START}:P`);

      if (!sheetData.values || sheetData.values.length <= 0) return [];
    } catch (err: any) {
      log(`Advanced API pipeline bypassed/failed. Error: ${err.message}. Running native fallback setup...`, true);

      // Fall back to the slower method of fetching all transaction records using the native SpreadsheetApp service
      const lastRow = sheet.getLastRow();
      if (lastRow < TRANSACTIONS_ROW_START) return [];
      sheetData = { values: sheet.getRange(TRANSACTIONS_ROW_START, 1, lastRow - TRANSACTIONS_ROW_START + 1, 16).getValues() };
    }

    const transactionRecords: TransactionRecord[] = sheetData.values.map((row: any[]) => {
      return {
        id: row[0] as number | 0,
        individual: row[1] as string | "",
        period: row[2] as number | 0,
        type: row[3] as TransactionType | "Unknown",
        serviceProvided: row[4] as string | "",
        unitPrice: row[5] as number | 0,
        quantity: row[6] as number | 0,
        modifiedColumn: row[7] as number | 0,
        tenderedMoney: row[8] as number | 0,
        initialColumnAmount: row[9] as number | 0,
        newColumnAmount: row[10] as number | 0,
        initialBalance: row[11] as number | 0,
        newBalance: row[12] as number | 0,
        timestamp: row[13] as string | 0,
        row: row[14] ? Number(row[14]) : undefined,
        operation: row[15] as "ADD" | "SUBTRACT" | "MULTIPLY" | "DIVIDE" | undefined
      } as TransactionRecord;
    })

    return transactionRecords;
  } catch (error: any) {
    log(`Error occurred in fetchTransactionRecords: ${error.message}`, true);
    return false;
  }
}

function undoTransactions(data?: string): boolean {
  try {
    if (!data || typeof data !== 'string') {
      log("No data provided for undoTransaction. Operation aborted.", true);
      return false;
    }

    const parsedPayload = JSON.parse(data);
    const transactionIds = Array.isArray(parsedPayload)
      ? parsedPayload
      : parsedPayload?.transactionIds;

    log(`Parsed data for undoTransactions: ${JSON.stringify(parsedPayload)}`, false);

    if (!Array.isArray(transactionIds) || transactionIds.length === 0) {
      log("Parsed data is not a valid array or is empty. Operation aborted.", true);
      return false;
    }

    const normalizedTransactionIds = transactionIds
      .map((id: unknown) => Number(id))
      .filter((id: number) => Number.isFinite(id));

    if (normalizedTransactionIds.length === 0) return false;

    let transactionRecords = fetchTransactionsDataCached() as TransactionRecord[];

    // Convert transactionRecords to an array if it's not already
    if (!Array.isArray(transactionRecords)) {
      log("Fetched transaction records are not in an array format. Attempting convert from string if possible.", false);
      if (typeof transactionRecords === 'string') {
        try {
          const parsedRecords = JSON.parse(transactionRecords);
          if (Array.isArray(parsedRecords)) {
            log("Successfully converted fetched transaction records from string to array.", false);
            transactionRecords = parsedRecords;
          } else {
            log("Fetched transaction records string could not be converted to an array. Operation aborted.", true);
            return false;
          }
        } catch (parseError: any) {
          log(`Error parsing fetched transaction records string: ${parseError.message}. Operation aborted.`, true);
          return false;
        }
      } else {
        log("Fetched transaction records are neither an array nor a string. Operation aborted.", true);
        return false;
      }
    }

    log(`Fetched ${transactionRecords.length} transaction records for undo operation.`, false);
    const recordsToUndo = transactionRecords.filter(record =>
      record.id !== undefined && normalizedTransactionIds.includes(Number(record.id))
    );

    if (recordsToUndo.length === 0) {
      log("No matching transaction records found for the provided IDs. Operation aborted.", true);
      return false
    }

    // Reverse the operations for each record to undo
    for (const record of recordsToUndo) {
      if (!record.individual || !record.type || !record.unitPrice || !record.quantity || !record.modifiedColumn || !record.period) {
        log(`Incomplete transaction record found for ID ${record.id}. Cannot undo.`, true);
        return false;
      }

      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const targetSheet = spreadsheet.getSheets().find(sheet =>
        parseInt(sheet.getName().replace(/\D/g, ""), 10) === Number(record.period)
      );
      if (!targetSheet) {
        log(`No sheet found for period ${record.period} and transaction ID ${record.id}.`, true);
        return false;
      }

      let targetRow = Number(record.row);
      if (!Number.isInteger(targetRow) || targetRow < USER_STARTING_ROW) {
        const lastRow = targetSheet.getLastRow();
        if (lastRow < USER_STARTING_ROW) {
          log(`No student rows found in Period ${record.period} for transaction ID ${record.id}.`, true);
          return false;
        }

        const names = targetSheet
          .getRange(USER_STARTING_ROW, NAMES_COL, lastRow - USER_STARTING_ROW + 1, 1)
          .getValues();
        const studentRowOffset = names.findIndex(row => String(row[0]).trim() === String(record.individual).trim());
        if (studentRowOffset < 0) {
          log(`Student "${record.individual}" not found in Period ${record.period} for transaction ID ${record.id}.`, true);
          return false;
        }
        targetRow = USER_STARTING_ROW + studentRowOffset;
      }

      const rowName = String(targetSheet.getRange(targetRow, NAMES_COL).getValue()).trim();
      if (rowName !== String(record.individual).trim()) {
        log(`Transaction ID ${record.id} row ${targetRow} belongs to "${rowName}", not "${record.individual}".`, true);
        return false;
      }

      const targetRange = targetSheet.getRange(targetRow, Number(record.modifiedColumn), 1, 1);
      const initialCellAmount = Number(record.initialColumnAmount);
      const currentValue = Number(targetRange.getValue());
      if (!Number.isFinite(initialCellAmount)) {
        log(`Transaction ID ${record.id} has no valid initial cell amount.`, true);
        return false;
      }
      if (!Number.isFinite(currentValue)) {
        log(`Target cell ${targetSheet.getName()}!${targetRange.getA1Notation()} is not numeric for transaction ID ${record.id}.`, true);
        return false;
      }

      const originalOperation = record.operation || (
        Number(record.newColumnAmount) >= initialCellAmount ? "ADD" : "SUBTRACT"
      );
      const inverseOperation = {
        ADD: "SUBTRACT",
        SUBTRACT: "ADD",
        MULTIPLY: "DIVIDE",
        DIVIDE: "MULTIPLY"
      }[originalOperation];
      const undoValue = Number(record.unitPrice) * Number(record.quantity);
      if (!Number.isFinite(undoValue) || undoValue === 0) {
        log(`Transaction ID ${record.id} has no valid operation value.`, true);
        return false;
      }

      let restoredValue: number;
      switch (inverseOperation) {
        case "SUBTRACT":
          restoredValue = Number((currentValue - undoValue).toFixed(2));
          break;
        case "ADD":
          restoredValue = Number((currentValue + undoValue).toFixed(2));
          break;
        case "DIVIDE":
          restoredValue = Number((currentValue / undoValue).toFixed(2));
          break;
        case "MULTIPLY":
          restoredValue = Number((currentValue * undoValue).toFixed(2));
          break;
        default:
          log(`Transaction ID ${record.id} has an unsupported operation.`, true);
          return false;
      }

      log(`Undoing transaction ID ${record.id} at ${targetSheet.getName()}!${targetRange.getA1Notation()} with ${inverseOperation} ${undoValue}.`, false);
      targetRange.setValue(restoredValue);
      SpreadsheetApp.flush();

      const updatedBalance = Number(targetSheet.getRange(targetRow, BALANCE_COL).getValue());
      addTransactionRecords([{
        individual: record.individual,
        period: Number(record.period),
        row: targetRow,
        type: inverseOperation === "ADD" || inverseOperation === "MULTIPLY" ? "Income" : "Expense",
        operation: inverseOperation as "ADD" | "SUBTRACT" | "MULTIPLY" | "DIVIDE",
        serviceProvided: `Undo transaction ID ${record.id} - ${record.serviceProvided || "Not specified"}`,
        unitPrice: Number(record.unitPrice),
        quantity: Number(record.quantity),
        modifiedColumn: Number(record.modifiedColumn),
        tenderedMoney: undoValue,
        initialColumnAmount: currentValue,
        newColumnAmount: restoredValue,
        initialBalance: Number.isFinite(updatedBalance) ? updatedBalance : Number(record.initialBalance) || 0,
        newBalance: Number.isFinite(updatedBalance) ? updatedBalance : Number(record.newBalance) || 0,
        timestamp: new Date().toISOString()
      }]);
    }

    clearGlobalCache([TRANSACTIONS_CACHED_KEY]);
    log(`Successfully undone ${recordsToUndo.length} transaction(s).`, true);
    return true;
  } catch (error: any) {
    log(`Error occurred in undoTransaction: ${error.message}`, true);
    return false;
  }
}

function resetAllTransactionRecords(): boolean {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(DEFAULT_TRANSACTIONS_SHEET);

    if (!sheet) {
      log(`Transactions sheet "${DEFAULT_TRANSACTIONS_SHEET}" not found.`, true);
      return false;
    }

    const lastRowWithData = sheet.getLastRow();
    
    if (lastRowWithData >= TRANSACTIONS_ROW_START) {
      sheet.deleteRows(TRANSACTIONS_ROW_START, lastRowWithData - TRANSACTIONS_ROW_START + 1);
    }

    log(`All transaction records have been reset.`, true);
    return true;
  } catch (error: any) {
    log(`Error occurred in resetAllTransactionRecords: ${error.message}`, true);
    return false;
  }
}