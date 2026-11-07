/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran. All rights reserved.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 * 
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 * 
 * This file contains utility functions for various operations in the B-Bucks Economy Scripts project, including mathematical operations on cell values, moving values between cells, and adding comments to cells.
 */

// Operation enum 
enum Operation {
  ADD = "ADD",
  SUBTRACT = "SUBTRACT",
  MULTIPLY = "MULTIPLY",
  DIVIDE = "DIVIDE"
}

/**
 * Works out whether a change is income or spending for the STUDENT. Adding to the expenditures column takes money away,
 * so it is spending, and taking from it is a refund, so it is income. Every other column works the other way round.
 */
function transactionTypeForChange(operation: Operation, column: number): TransactionType {
  const isIncreasing = operation === Operation.ADD || operation === Operation.MULTIPLY;
  if (column === getExpendituresColumn()) return isIncreasing ? "Expense" : "Income";
  return isIncreasing ? "Income" : "Expense";
}

/**
 * Executes a balance action based on the provided payload string.
 * @param payloadStr The JSON string containing the operation and amount to apply.
 * @returns The result of the operation or an error message.
 */
function executeBalanceAction(payloadStr: string): string | void {
  try {
    // Check if a string payload was provided
    if (!payloadStr) {
      log("No payload provided for balance action.", true);
      return "No payload provided for balance action.";
    }

    // Parse the clean JSON string into a JSON object for the util function to process
    const payload = JSON.parse(payloadStr);
    const { operation, unitPrice = undefined, quantity = undefined, transactionReason = undefined } = payload;

    if (!operation || !unitPrice || typeof unitPrice !== 'number' || !quantity || typeof quantity !== 'number') {
      log("Invalid payload. Please provide a valid operation and amount.", true);
      return "Invalid payload. Please provide a valid operation and amount.";
    }

    return applyMathToSelection(operation, unitPrice, quantity, true, transactionReason).toString();
  } catch (error: any) {
    log(`Error occurred in executeBalanceAction: ${error.message}`, true);
    return `Error occurred in executeBalanceAction: ${error.message}`;
  }
}

/**
 * Uses an operand and a value to apply to the existing value of all selected cells.
 * @param operation The mathematical operation to apply to the selected cells. Must be one of "ADD", "SUBTRACT", "MULTIPLY", or "DIVIDE".
 * @param unitPrice The unit price to use in the operation. Must be a number.
 * @param quantity The quantity to use in the operation. Must be a number.
 * @param isManualTransaction Whether the transaction is manual or not.
 * @param transactionReason (Optional) The reason for the transaction. Defaults to "Not Specified".
 * @param range (Optional) The range of cells to apply the operation to. If not provided, uses active range list.
 * @param fixedRow (Optional) Target specific row across selected columns.
 * @param fixedCol (Optional) Target specific column across selected rows.
 * @returns {string | boolean} Returns true if the operation was successful, string error otherwise.
 */
function applyMathToSelection(
  operation: Operation | string,
  unitPrice: number,
  quantity: number,
  isManualTransaction: boolean,
  transactionReason?: string,
  range?: GoogleAppsScript.Spreadsheet.Range,
  fixedRow?: number,
  fixedCol?: number
): string | boolean {
  try {
    // must explicitly include undefined check because isManualTransaction is a boolean
    if (
      operation === undefined ||
      unitPrice === undefined ||
      quantity === undefined ||
      isManualTransaction === undefined
    ) {
      log(`You must have an operation, unitPrice, quantity, and isManualTransaction boolean provided. Received - operation: ${operation}, unitPrice: ${unitPrice}, quantity: ${quantity}, isManualTransaction: ${isManualTransaction}`, true);
      return "You must have an operation, unitPrice, quantity, and isManualTransaction boolean provided.";
    }

    if (typeof unitPrice !== 'number' || isNaN(unitPrice)) {
      log("Unit price must be denominated in an integer quantity.", true);
      return "Unit price must be denominated in an integer quantity.";
    }

    if (typeof quantity !== 'number' || isNaN(quantity)) {
      log("Quantity must be denominated in an integer quantity.", true);
      return "Quantity must be denominated in an integer quantity.";
    }

    // Absolutely ENSURE there is no floating point precision issues
    const value = Number((unitPrice * quantity).toFixed(2));

    // Ensure the operation is in an enum format
    let normalMapping: Operation;
    if (typeof operation === 'string') {
      normalMapping = {
        "ADD": Operation.ADD,
        "SUBTRACT": Operation.SUBTRACT,
        "MULTIPLY": Operation.MULTIPLY,
        "DIVIDE": Operation.DIVIDE
      }[operation.toUpperCase()] as Operation;
    } else {
      normalMapping = operation;
    }

    if (!normalMapping || !Object.values(Operation).includes(normalMapping)) {
      log("Invalid operation. Must be one of ADD, SUBTRACT, MULTIPLY, or DIVIDE.", true);
      return "Invalid operation. Must be one of ADD, SUBTRACT, MULTIPLY, or DIVIDE.";
    }

    // Fix: check normalized mapping instead of raw parameter string
    if (normalMapping === Operation.DIVIDE && value === 0) {
      log("You cannot divide by zero.", true);
      return "You cannot divide by zero.";
    }

    // Define the operations in a mapping for cleaner code
    const operations: Record<Operation, (n: number) => number> = {
      [Operation.ADD]: (n) => Number((n + value).toFixed(2)),
      [Operation.SUBTRACT]: (n) => Number((n - value).toFixed(2)),
      [Operation.MULTIPLY]: (n) => Number((n * value).toFixed(2)),
      [Operation.DIVIDE]: (n) => Number((n / value).toFixed(2)),
    };

    const operationFunc = operations[normalMapping];
    const transactionRecords: TransactionRecord[] = [];
    const NOT_ALLOWED_COLS: number[] = [
      getNamesColumn(),
      getBalanceColumn(),
      getNetIncomeColumn(),
      getDepositDateColumn()
    ];

    // Set once the primary path has written to the sheet. After that point a failure must NOT fall back, otherwise the
    // same change would be applied twice.
    let valuesWritten = false;
    // A cell selected twice (for example two ranges on the same row with a fixed column) is only changed once.
    const processedCells = new Set<string>();

    // Primary Execution Method: Batch via Sheets API
    try {
      let rangesToProcess: GoogleAppsScript.Spreadsheet.Range[] = [];
      
      if (range) {
        rangesToProcess = [range];
      } else {
        const activeRangeList = getSS().getActiveRangeList();
        if (!activeRangeList) {
          log("No active range found. Please select cells to apply the operation to.", true);
          return "No active range found. Please select cells to apply the operation to.";
        }
        rangesToProcess = activeRangeList.getRanges();
      }

      const spreadsheetId = getSS().getId();
      const requests: GoogleAppsScript.Sheets.Schema.ValueRange[] = [];
      const affectedRows: number[] = [];

      rangesToProcess.forEach(targetSubRange => {
        const sheet = targetSubRange.getSheet();
        const sheetName = sheet.getName();
        const periodName = parseInt(sheetName.replace(/\D/g, ""), 10);

        // Calculate target dimensions based on selection and overrides
        let startRow = targetSubRange.getRow();
        let numRows = targetSubRange.getNumRows();
        let startCol = targetSubRange.getColumn();
        let numCols = targetSubRange.getNumColumns();

        if (fixedRow !== undefined) {
          startRow = fixedRow;
          numRows = 1;
        }

        if (fixedCol !== undefined) {
          startCol = fixedCol;
          numCols = 1;
        }

        const targetRange = sheet.getRange(startRow, startCol, numRows, numCols);
        const targetA1 = targetRange.getA1Notation();
        const values = targetRange.getValues();

        // Fetch Names and Balances in bulk for the affected row range to prevent API quota exhaustion
        const namesInSheet = sheet.getRange(startRow, getNamesColumn(), numRows, 1).getValues();
        const balancesInSheet = sheet.getRange(startRow, getBalanceColumn(), numRows, 1).getValues();

        log(`Sheet name: ${sheetName}, Range A1: ${targetA1}, Period Name: ${periodName}`, false);

        const updatedValues = values.map((row, rowIndex) => {
          const absoluteRowIndex = startRow + rowIndex;
          const individualName = namesInSheet[rowIndex][0];
          const initialBalance = balancesInSheet[rowIndex][0];

          return row.map((cell, colIndex) => {
            const absoluteColIndex = startCol + colIndex;

            // null tells the Sheets API to leave the cell untouched, so formulas and dates in the selection are never overwritten.
            if (NOT_ALLOWED_COLS.includes(absoluteColIndex) && fixedCol === undefined) {
              log(`Operation not allowed on column ${absoluteColIndex} in range ${sheetName}!${targetA1}. Skipping cell.`, false);
              return null;
            }

            const cellKey = `${sheetName}|${absoluteRowIndex}|${absoluteColIndex}`;
            if (processedCells.has(cellKey)) return null;

            if (typeof cell === 'number' && !isNaN(cell)) {
              if (initialBalance === undefined || initialBalance === null || isNaN(Number(initialBalance))) {
                log(`Balance value is invalid for row ${absoluteRowIndex}. Skipping cell.`, true);
                return null;
              }

              if (!individualName) {
                log(`Individual name is missing for row ${absoluteRowIndex}. Skipping cell.`, true);
                return null;
              }

              processedCells.add(cellKey);
              const result = operationFunc(cell);
              affectedRows.push(absoluteRowIndex);

              transactionRecords.push({
                individual: String(individualName),
                period: periodName || undefined,
                row: absoluteRowIndex,
                type: transactionTypeForChange(normalMapping, absoluteColIndex),
                operation: normalMapping,
                serviceProvided: `${isManualTransaction ? "Manual Balance Adjustment" : ""} ${isManualTransaction && transactionReason ? '-' : ""} ${transactionReason ?? "Not Specified"}`.trim(),
                unitPrice: unitPrice,
                quantity: quantity,
                modifiedColumn: absoluteColIndex,
                tenderedMoney: value,
                initialColumnAmount: Number(cell.toFixed(2)),
                newColumnAmount: result,
                initialBalance: Number(Number(initialBalance).toFixed(2)),
                newBalance: 0, // Calculated post-flush
                timestamp: new Date().toISOString()
              });

              return result;
            }

            log(`Non-numeric value "${cell}" found in range ${sheetName}!${targetA1}. Skipping cell.`, false);
            return null;
          });
        });

        requests.push({
          range: `${sheetName}!${targetA1}`,
          values: updatedValues
        });
      });

      if (requests.length > 0 && typeof Sheets !== 'undefined') {
        Sheets.Spreadsheets.Values.batchUpdate({
          valueInputOption: "USER_ENTERED",
          data: requests
        }, spreadsheetId);
        valuesWritten = true;
      } else {
        log("No valid ranges found or Sheets API not available.", true);
        return false;
      }

      SpreadsheetApp.flush();

      // Re-read each record's balance by its exact sheet period and row.
      if (transactionRecords.length > 0) {
        transactionRecords.forEach(record => {
          const recordSheet = getAllSheets().find(sheet =>
            parseInt(sheet.getName().replace(/\D/g, ""), 10) === Number(record.period)
          );
          const recordRow = Number(record.row);

          if (recordSheet && Number.isInteger(recordRow) && recordRow > 0) {
            const updatedBalance = Number(recordSheet.getRange(recordRow, getBalanceColumn()).getValue());
            if (Number.isFinite(updatedBalance)) {
              record.newBalance = Number(updatedBalance.toFixed(2));
            }
          }
        });

        addTransactionRecords(transactionRecords);
      }

      // Update timestamps for the affected sheet after all operations are complete
      updateTimestampForSheet();

      return true;

    } catch (error: any) {
      if (valuesWritten) {
        log(`The balances were updated, but finishing the transaction failed: ${error.message}. Check the Transactions Records before retrying.`, true);
        return `The balances were updated, but finishing the transaction failed: ${error.message}`;
      }
      log(`Primary Sheets API execution failed, attempting fallback Apps Script path: ${error.message}`, true);

      // Start the fallback from a clean slate.
      transactionRecords.length = 0;
      processedCells.clear();

      // Fallback Method: Standard SpreadsheetApp API
      try {
        let rangesToProcess: GoogleAppsScript.Spreadsheet.Range[] = [];
        if (range) {
          rangesToProcess = [range];
        } else {
          const activeRangeList = getSS().getActiveRangeList();
          if (activeRangeList) {
            rangesToProcess = activeRangeList.getRanges();
          } else {
            return "No active range found. Please select cells to apply the operation to.";
          }
        }

        rangesToProcess.forEach(subRange => {
          const sheet = subRange.getSheet();
          const startRow = fixedRow !== undefined ? fixedRow : subRange.getRow();
          const numRows = fixedRow !== undefined ? 1 : subRange.getNumRows();
          const startCol = fixedCol !== undefined ? fixedCol : subRange.getColumn();
          const numCols = fixedCol !== undefined ? 1 : subRange.getNumColumns();

          const targetRange = sheet.getRange(startRow, startCol, numRows, numCols);
          const values = targetRange.getValues();
          const periodName = parseInt(sheet.getName().replace(/\D/g, ""), 10);

          // Only the cells that actually change are written, so formulas and dates in the selection are left alone.
          values.forEach((row, rowIndex) => {
            const absoluteRowIndex = startRow + rowIndex;

            row.forEach((cell, colIndex) => {
              const absoluteColIndex = startCol + colIndex;

              if (NOT_ALLOWED_COLS.includes(absoluteColIndex) && fixedCol === undefined) {
                return;
              }

              const cellKey = `${sheet.getName()}|${absoluteRowIndex}|${absoluteColIndex}`;
              if (processedCells.has(cellKey)) return;

              if (typeof cell === 'number' && !isNaN(cell)) {
                const balance = sheet.getRange(absoluteRowIndex, getBalanceColumn()).getValue();
                const individualName = sheet.getRange(absoluteRowIndex, getNamesColumn()).getValue();

                if (!individualName || balance === undefined || isNaN(balance)) {
                  return;
                }

                processedCells.add(cellKey);
                const result = operationFunc(cell);

                transactionRecords.push({
                  individual: String(individualName),
                  period: periodName || undefined,
                  row: absoluteRowIndex,
                  type: transactionTypeForChange(normalMapping, absoluteColIndex),
                  operation: normalMapping,
                  serviceProvided: `${isManualTransaction ? "Manual Balance Adjustment - " : ""}${transactionReason ?? "Not Specified"}`.trim(),
                  unitPrice: unitPrice,
                  quantity: quantity,
                  modifiedColumn: absoluteColIndex,
                  tenderedMoney: value,
                  initialColumnAmount: Number(cell.toFixed(2)),
                  newColumnAmount: result,
                  initialBalance: Number(Number(balance).toFixed(2)),
                  newBalance: 0,
                  timestamp: new Date().toISOString()
                });

                sheet.getRange(absoluteRowIndex, absoluteColIndex).setValue(result);
              }
            });
          });
        });

        SpreadsheetApp.flush();

        if (transactionRecords.length > 0) {
          transactionRecords.forEach(record => {
            const recordSheet = getAllSheets().find(sheet =>
              parseInt(sheet.getName().replace(/\D/g, ""), 10) === Number(record.period)
            );
            const recordRow = Number(record.row);

            if (recordSheet && Number.isInteger(recordRow) && recordRow > 0) {
              const updatedBalance = Number(recordSheet.getRange(recordRow, getBalanceColumn()).getValue());
              if (Number.isFinite(updatedBalance)) {
                record.newBalance = Number(updatedBalance.toFixed(2));
              }
            }
          });
          addTransactionRecords(transactionRecords);
        }

        // Update timestamps for the affected sheet after all operations are complete
        updateTimestampForSheet();

        return true;
      } catch (fallbackError: any) {
        log(`Fallback method failed: ${fallbackError.message}`, true);
        return `Error occurred while applying the operation: ${fallbackError.message}`;
      }
    }
  } catch (error: any) {
    log(`Error occurred in applyMathToSelection: ${error.message}`, true);
    return error.message;
  }
}
