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
 * @param commentOnExpenditures Whether to leave expenditure notes on affected rows. Defaults to false.
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
  commentOnExpenditures: boolean = false,
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
      NAMES_COL,
      BALANCE_COL,
      NET_INCOME_COL,
      DATE_DEPOSIT_COL
    ];

    // Primary Execution Method: Batch via Sheets API
    try {
      let rangesToProcess: GoogleAppsScript.Spreadsheet.Range[] = [];
      
      if (range) {
        rangesToProcess = [range];
      } else {
        const activeRangeList = SpreadsheetApp.getActiveSpreadsheet().getActiveRangeList();
        if (!activeRangeList) {
          log("No active range found. Please select cells to apply the operation to.", true);
          return "No active range found. Please select cells to apply the operation to.";
        }
        rangesToProcess = activeRangeList.getRanges();
      }

      const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
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
        const namesInSheet = sheet.getRange(startRow, NAMES_COL, numRows, 1).getValues();
        const balancesInSheet = sheet.getRange(startRow, BALANCE_COL, numRows, 1).getValues();

        log(`Sheet name: ${sheetName}, Range A1: ${targetA1}, Period Name: ${periodName}`, false);

        const updatedValues = values.map((row, rowIndex) => {
          const absoluteRowIndex = startRow + rowIndex;
          const individualName = namesInSheet[rowIndex][0];
          const initialBalance = balancesInSheet[rowIndex][0];

          return row.map((cell, colIndex) => {
            const absoluteColIndex = startCol + colIndex;

            if (NOT_ALLOWED_COLS.includes(absoluteColIndex) && fixedCol === undefined) {
              log(`Operation not allowed on column ${absoluteColIndex} in range ${sheetName}!${targetA1}. Skipping cell.`, true);
              return cell;
            }

            if (typeof cell === 'number' && !isNaN(cell)) {
              if (initialBalance === undefined || initialBalance === null || isNaN(Number(initialBalance))) {
                log(`Balance value is invalid for row ${absoluteRowIndex}. Skipping cell.`, true);
                return cell;
              }

              if (!individualName) {
                log(`Individual name is missing for row ${absoluteRowIndex}. Skipping cell.`, true);
                return cell;
              }

              const result = operationFunc(cell);
              affectedRows.push(absoluteRowIndex);

              transactionRecords.push({
                individual: String(individualName),
                period: periodName || undefined,
                type: normalMapping === Operation.ADD || normalMapping === Operation.MULTIPLY ? "Income" : "Expense",
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

            log(`Non-numeric value "${cell}" found in range ${sheetName}!${targetA1}. Skipping cell.`, true);
            return cell;
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
      } else {
        log("No valid ranges found or Sheets API not available.", true);
        return false;
      }

      SpreadsheetApp.flush();

      // Write expenditure notes if enabled
      if (commentOnExpenditures && affectedRows.length > 0) {
        const uniqueRows = [...new Set(affectedRows)];
        commentExpenditureOnSelection(
          uniqueRows, 
          `$${Number(value.toFixed(2))} - ${new Date().toLocaleDateString("en-US")} ${transactionReason || "Manual Adjustment"}`
        );
      }

      // Re-read updated balances post-flush
      if (transactionRecords.length > 0) {
        rangesToProcess.forEach(targetSubRange => {
          const sheet = targetSubRange.getSheet();
          const startRow = fixedRow !== undefined ? fixedRow : targetSubRange.getRow();
          const numRows = fixedRow !== undefined ? 1 : targetSubRange.getNumRows();
          
          const updatedBalances = sheet.getRange(startRow, BALANCE_COL, numRows, 1).getValues();

          transactionRecords.forEach((record, idx) => {
            if (updatedBalances[idx] && updatedBalances[idx][0] !== undefined) {
              record.newBalance = Number(Number(updatedBalances[idx][0]).toFixed(2));
            }
          });
        });

        addTransactionRecords(transactionRecords);
      }

      // Update timestamps for the affected sheet after all operations are complete
      updateTimestampForSheet();

      return true;

    } catch (error: any) {
      log(`Primary Sheets API execution failed, attempting fallback Apps Script path: ${error.message}`, true);

      // Fallback Method: Standard SpreadsheetApp API
      try {
        let rangesToProcess: GoogleAppsScript.Spreadsheet.Range[] = [];
        if (range) {
          rangesToProcess = [range];
        } else {
          const activeRangeList = SpreadsheetApp.getActiveSpreadsheet().getActiveRangeList();
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

          const updatedValues = values.map((row, rowIndex) => {
            const absoluteRowIndex = startRow + rowIndex;

            return row.map((cell, colIndex) => {
              const absoluteColIndex = startCol + colIndex;

              if (NOT_ALLOWED_COLS.includes(absoluteColIndex) && fixedCol === undefined) {
                return cell;
              }

              if (typeof cell === 'number' && !isNaN(cell)) {
                const balance = sheet.getRange(absoluteRowIndex, BALANCE_COL).getValue();
                const individualName = sheet.getRange(absoluteRowIndex, NAMES_COL).getValue();

                if (!individualName || balance === undefined || isNaN(balance)) {
                  return cell;
                }

                const result = operationFunc(cell);

                transactionRecords.push({
                  individual: String(individualName),
                  period: periodName || undefined,
                  type: normalMapping === Operation.ADD || normalMapping === Operation.MULTIPLY ? "Income" : "Expense",
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

                return result;
              }
              return cell;
            });
          });

          targetRange.setValues(updatedValues);
        });

        SpreadsheetApp.flush();

        if (transactionRecords.length > 0) {
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

/**
 * Adds a note to the specified cell range with the provided text, ensuring that the note is valid,
 * does not exceed character limits, and does not contain line breaks.
 * @param cells The range of cells to which the note will be added.
 * @param comment The text of the note to be added to the cells. Must be a string, cannot exceed 255 characters, and cannot contain line breaks.
 * @returns {boolean} Returns true if the operation was successful, false otherwise.
 */
function commentExpenditureOnSelection(rows: number[], comment: string): boolean {
  try {
    if (!rows || comment === undefined || comment === null) {
      log("You must provide both a cell range and a note string.", false);
      return false;
    }

    if (typeof comment !== 'string') {
      log("Comment must be a string.", false);
      return false;
    }

    if (comment.length > 255) {
      log("Comment cannot exceed 255 characters.", false);
      return false;
    }
    if (comment.includes("\n") || comment.includes("\r")) {
      log("Comment cannot contain line breaks.", false);
      return false;
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    for (const row of rows) {
      const currentNote = sheet.getRange(row, EXPENDITURES_COL).getNote();
      if (currentNote) {
        sheet.getRange(row, EXPENDITURES_COL).setNote(currentNote + "\n" + comment);
      } else {
        sheet.getRange(row, EXPENDITURES_COL).setNote(comment);
      }
    }
    return true;

  } catch (error: any) {
    log(`Error occurred in commentExpenditureOnSelection: ${error.message}`, true);
    return false;
  }
}

function resetAllMoney(): boolean {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow < USER_STARTING_ROW) {
      log("No data rows found to reset.", true);
      return false;
    }

    try {
      // Use the Google Sheets API to batch update the monetary columns to 0 for all rows starting from USER_STARTING_ROW
      const requests: GoogleAppsScript.Sheets.Schema.ValueRange[] = [];

      // Create a request for each row to update the monetary columns to 0
      for (let row = USER_STARTING_ROW; row <= lastRow; row++) {
        requests.push({
          range: `${sheet.getName()}!${row}:${row}`,
          values: [[0, 0, 0, 0]] // Set EARNINGS_COL, EXPENDITURES_COL, INVESTMENT_RETURNS_COL to 0
        });
      }

      // Send a batchUpdate request to the Sheets API to update all rows at once
      if (requests.length > 0 && Sheets) {
        Sheets.Spreadsheets.Values.batchUpdate({
          valueInputOption: "USER_ENTERED",
          data: requests
        }, SpreadsheetApp.getActiveSpreadsheet().getId());
      } else {
        log("No valid rows found to reset.", true);
        return false;
      }
      return true;
    } catch (error: any) {
      log(`Error occurred in resetAllMoney: ${error.message}, resorting to slower method. If this error persists, please contact the developer.`, true);

      // Reset all monetary columns to 0 for all rows starting from USER_STARTING_ROW
      for (let row = USER_STARTING_ROW; row <= lastRow; row++) {
        sheet.getRange(row, EARNINGS_COL).setValue(0);
        sheet.getRange(row, EXPENDITURES_COL).setValue(0);
        sheet.getRange(row, INVESTMENT_RETURNS_COL).setValue(0);
      }

      return true;
    }
  } catch (error: any) {
    log(`Error occurred in resetAllMoney: ${error.message}`, true);
    return false;
  }
}

function resetAccountHolders(): boolean {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow < USER_STARTING_ROW) {
      log("No data rows found to reset.", true);
      return false;
    }

    try {
      // Use the Google Sheets API to batch update the account holder names to empty strings for all rows starting from USER_STARTING_ROW
      const requests: GoogleAppsScript.Sheets.Schema.ValueRange[] = [];

      // Create a request for each row to update the account holder names to empty strings
      for (let row = USER_STARTING_ROW; row <= lastRow; row++) {
        requests.push({
          range: `${sheet.getName()}!${row}:${row}`,
          values: [[""]] // Set account holder names to empty strings
        });
      }

      // Send a batchUpdate request to the Sheets API to update all rows at once
      if (requests.length > 0 && Sheets) {
        Sheets.Spreadsheets.Values.batchUpdate({
          valueInputOption: "USER_ENTERED",
          data: requests
        }, SpreadsheetApp.getActiveSpreadsheet().getId());
      } else {
        log("No valid rows found to reset.", true);
        return false;
      }
      return true;
    } catch (error: any) {
      log(`Error occurred in resetAccountHolders: ${error.message}, resorting to slower method. If this error persists, please contact the developer.`, true);

      // Reset all account holder names to empty strings for all rows starting from USER_STARTING_ROW
      for (let row = USER_STARTING_ROW; row <= lastRow; row++) {
        sheet.getRange(row, 1).setValue(""); // Assuming account holder names are in column 1
      }

      return true;
    }
  } catch (error: any) {
    log(`Error occurred in resetAccountHolders: ${error.message}`, true);
    return false;
  }
}

function resetColumn(data: any): boolean {
  try {
    const { column } = data;

    if (!column) {
      log("No column specified for resetColumn.", true);
      return false;
    }

    const columnMapping: { [key: string]: number } = {
      "EARNINGS": EARNINGS_COL,
      "EXPENDITURES": EXPENDITURES_COL,
      "INVESTMENT_RETURNS": INVESTMENT_RETURNS_COL,
      "INITIAL_DEPOSIT": INITIAL_DEPOSIT_COL
    };

    const targetColumn = columnMapping[column.toUpperCase()];

    if (!targetColumn) {
      log(`Invalid column specified for resetColumn: ${column}. Must be one of EARNINGS, EXPENDITURES, INVESTMENT_RETURNS, or INITIAL_DEPOSIT.`, true);
      return false;
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < USER_STARTING_ROW) {
      log("No data rows found to reset.", true);
      return false;
    }

    try {
      // Use the Google Sheets API to batch update the earnings column to 0 for all rows starting from USER_STARTING_ROW
      const requests: GoogleAppsScript.Sheets.Schema.ValueRange[] = [];

      // Create a request for each row to update the target column to 0
      for (let row = USER_STARTING_ROW; row <= lastRow; row++) {
        const rangeA1 = `${sheet.getName()}!${row}:${row}`;
        const values = Array(sheet.getLastColumn()).fill(0); // Create an array filled with 0s for all columns
        values[targetColumn - 1] = 0; // Set the target column to 0 (adjusting for 0-based index)
        requests.push({
          range: rangeA1,
          values: [values]
        });
      }

      // Send a batchUpdate request to the Sheets API to update all rows at once
      if (requests.length > 0 && Sheets) {
        Sheets.Spreadsheets.Values.batchUpdate({
          valueInputOption: "USER_ENTERED",
          data: requests
        }, SpreadsheetApp.getActiveSpreadsheet().getId());
      } else {
        log("No valid rows found to reset.", true);
        return false;
      }
      return true;
    } catch (error: any) {
      log(`Error occurred in resetColumn: ${error.message}, resorting to slower method. If this error persists, please contact the developer.`, true);

      // Reset all earnings to 0 for all rows starting from USER_STARTING_ROW
      for (let row = USER_STARTING_ROW; row <= lastRow; row++) {
        sheet.getRange(row, targetColumn).setValue(0);
      }

      return true;
    }
  } catch (error: any) {
    log(`Error occurred in resetColumn: ${error.message}`, true);
    return false;
  }
}
