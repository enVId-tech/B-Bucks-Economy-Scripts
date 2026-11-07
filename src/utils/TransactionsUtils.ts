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

// NOTE: fetchTransactionsDataCached() / fetchTransactionsData() were removed. They loaded the ENTIRE log into one cached
// payload, which cannot scale. Use the rolling cache in utils/TransactionCache.ts instead:
//   fetchTransactionsPage()   - unfiltered paging with a rolling window
//   queryTransactions()       - searched / filtered paging
//   fetchTransactionFacets()  - distinct values for filter drop-downs
//   fetchTransactionsByIds()  - direct lookup (used by undo)

/**
 * Adds a transaction record to the "Transactions Records" sheet with the provided details, ensuring that all required information is valid and properly formatted.
 * Uses the Google Sheets API for efficient appending of transaction records, with error handling to fall back to the slower method if the API call fails.
 * @param records An array of transaction records to be added, where each record includes the individual's name, transaction type (Income, Expense, or Investment), service description, initial amount, tendered amount, final amount, quantity of services, and timestamp. All fields are required for each record.
 * @returns {boolean} Returns true if the operation was successful, false otherwise.
 */
function addTransactionRecords(records: TransactionRecord[]): boolean {
  try {
    const spreadsheet = getSS();
    const sheet = getSheetNamed(getTransactionsSheetName());

    // -- Add edge case checking for the sheet and all fields to ensure data integrity when compiled to JavaScript --
    if (!sheet) {
      log(`Sheet "${getTransactionsSheetName()}" not found.`, true);
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

    if (lastRowWithData >= (getTransactionsRowStart() - 1) && lastRowWithData > 0) {
      const rawIdValue = sheet.getRange(lastRowWithData, 1).getValue();
      const parsedId = parseInt(rawIdValue, 10);
      biggestId = isNaN(parsedId) ? 0 : parsedId;
    }

    const insertStartRow = Math.max(lastRowWithData + 1, getTransactionsRowStart());
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

      onTransactionsAppended();
      // The expenditure notes are a view of this log, so refresh them from it whenever the log changes.
      syncNotesForRecords(records);
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

    // Look the records up directly by id instead of loading the whole log.
    const recordsToUndo = fetchTransactionsByIds(normalizedTransactionIds);
    log(`Found ${recordsToUndo.length} of ${normalizedTransactionIds.length} requested transaction record(s) for undo.`, false);

    if (recordsToUndo.length === 0) {
      log("No matching transaction records found for the provided IDs. Operation aborted.", true);
      return false
    }

    // A record that was already undone (or an undo that was itself undone) must not be reversed again, otherwise the money
    // moves twice while the log and the notes only count it once. Check everything first so nothing is half-done.
    const alreadyUndone = findRecordsNoLongerInEffect(recordsToUndo.map(record => Number(record.id)));
    if (alreadyUndone.length > 0) {
      log(`Transaction ID(s) ${alreadyUndone.join(", ")} have already been undone. Nothing was changed.`, true);
      return false;
    }

    // Reverse the operations for each record to undo. The expenditure notes are rebuilt once at the end instead of once per record.
    beginNoteSyncBatch();
    try {
    for (const record of recordsToUndo) {
      if (!record.individual || !record.type || !record.unitPrice || !record.quantity || !record.modifiedColumn || !record.period) {
        log(`Incomplete transaction record found for ID ${record.id}. Cannot undo.`, true);
        return false;
      }

      const spreadsheet = getSS();
      const targetSheet = getAllSheets().find(sheet =>
        parseInt(sheet.getName().replace(/\D/g, ""), 10) === Number(record.period)
      );
      if (!targetSheet) {
        log(`No sheet found for period ${record.period} and transaction ID ${record.id}.`, true);
        return false;
      }

      // The stored row is only a hint: rows move when names are sorted, so fall back to finding the student by name.
      let targetRow = Number(record.row);
      const storedRowIsValid = Number.isInteger(targetRow) && targetRow >= getUserStartingRow()
        && String(targetSheet.getRange(targetRow, getNamesColumn()).getValue()).trim() === String(record.individual).trim();
      if (!storedRowIsValid) {
        const lastRow = targetSheet.getLastRow();
        if (lastRow < getUserStartingRow()) {
          log(`No student rows found in Period ${record.period} for transaction ID ${record.id}.`, true);
          return false;
        }

        const names = targetSheet
          .getRange(getUserStartingRow(), getNamesColumn(), lastRow - getUserStartingRow() + 1, 1)
          .getValues();
        const studentRowOffset = names.findIndex(row => String(row[0]).trim() === String(record.individual).trim());
        if (studentRowOffset < 0) {
          log(`Student "${record.individual}" not found in Period ${record.period} for transaction ID ${record.id}.`, true);
          return false;
        }
        targetRow = getUserStartingRow() + studentRowOffset;
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

      const balanceBeforeUndo = Number(targetSheet.getRange(targetRow, getBalanceColumn()).getValue());
      log(`Undoing transaction ID ${record.id} at ${targetSheet.getName()}!${targetRange.getA1Notation()} with ${inverseOperation} ${undoValue}.`, false);
      targetRange.setValue(restoredValue);
      SpreadsheetApp.flush();

      const updatedBalance = Number(targetSheet.getRange(targetRow, getBalanceColumn()).getValue());
      addTransactionRecords([{
        individual: record.individual,
        period: Number(record.period),
        row: targetRow,
        type: record.type === "Investment" ? "Investment" : transactionTypeForChange(inverseOperation as Operation, Number(record.modifiedColumn)),
        operation: inverseOperation as "ADD" | "SUBTRACT" | "MULTIPLY" | "DIVIDE",
        serviceProvided: `Undo transaction ID ${record.id} - ${record.serviceProvided || "Not specified"}`,
        unitPrice: Number(record.unitPrice),
        quantity: Number(record.quantity),
        modifiedColumn: Number(record.modifiedColumn),
        tenderedMoney: undoValue,
        initialColumnAmount: currentValue,
        newColumnAmount: restoredValue,
        initialBalance: Number.isFinite(balanceBeforeUndo) ? balanceBeforeUndo : Number(record.newBalance) || 0,
        newBalance: Number.isFinite(updatedBalance) ? updatedBalance : Number(record.initialBalance) || 0,
        timestamp: new Date().toISOString()
      }]);
    }
    } finally {
      endNoteSyncBatch();
    }

    onTransactionsAppended();
    log(`Successfully undone ${recordsToUndo.length} transaction(s).`, true);
    return true;
  } catch (error: any) {
    log(`Error occurred in undoTransaction: ${error.message}`, true);
    return false;
  }
}

/**
 * Returns the ids (from the list given) that are no longer active: already undone, or an undo that was itself undone.
 * Only the id and description columns of the log are read.
 */
function findRecordsNoLongerInEffect(ids: number[]): number[] {
  const sheet = getTransactionsSheet();
  if (!sheet) return [];
  const total = countTransactionRows(sheet);
  if (total === 0) return [];

  const rows = sheet.getRange(getTransactionsRowStart(), 1, total, 5).getValues();
  const minimalRecords = rows.map(row => ({ id: Number(row[0]), serviceProvided: String(row[4] ?? ""), unitPrice: 0 } as TransactionRecord));
  const activeRecords = resolveActiveRecords(minimalRecords);
  return ids.filter(id => !activeRecords.has(id));
}

function resetAllTransactionRecords(): boolean {
  try {
    const spreadsheet = getSS();
    const sheet = getSheetNamed(getTransactionsSheetName());

    if (!sheet) {
      log(`Transactions sheet "${getTransactionsSheetName()}" not found.`, true);
      return false;
    }

    const lastRowWithData = sheet.getLastRow();
    
    if (lastRowWithData >= getTransactionsRowStart()) {
      sheet.deleteRows(getTransactionsRowStart(), lastRowWithData - getTransactionsRowStart() + 1);
    }

    invalidateTransactionCache();
    log(`All transaction records have been reset.`, true);
    return true;
  } catch (error: any) {
    log(`Error occurred in resetAllTransactionRecords: ${error.message}`, true);
    return false;
  }
}