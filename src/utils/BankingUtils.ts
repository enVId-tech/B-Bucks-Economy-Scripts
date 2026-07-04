// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. All rights reserved.
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains utility functions for various operations in the B-Bucks Economy Scripts project, including mathematical operations on cell values, moving values between cells, and adding comments to cells.

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
      Logger.log("No payload provided for balance action.");
      SpreadsheetApp.getUi().alert("No payload provided for balance action.");
      return "No payload provided for balance action.";
    }

    // Parse the clean JSON string into a JSON object for the util function to process
    const payload = JSON.parse(payloadStr);
    const { operation, amount, transactionReason = undefined } = payload;

    if (!operation || !amount || typeof amount !== 'number') {
      Logger.log("Invalid payload. Please provide a valid operation and amount.");
      SpreadsheetApp.getUi().alert(`Invalid payload. Please provide a valid operation and amount. Information received - operation: ${operation}, amount: ${amount}`);
      return "Invalid payload. Please provide a valid operation and amount.";
    }

    return applyMathToSelection(operation, amount, true, transactionReason).toString();
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(`Error occurred in executeBalanceAction: ${error.message}`);
    return `Error occurred in executeBalanceAction: ${error.message}`;
  }
}

/**
 * Uses an operand and a value to apply to the existing value of all selected cells.
 * @param operation The mathematical operation to apply to the selected cells. Must be one of "ADD", "SUBTRACT", "MULTIPLY", or "DIVIDE".
 * @param value The value to use in the mathematical operation on the selected cells. Must be a number.
 * @param isManualTransaction Whether the transaction is manual or not.
 * @param transactionReason (Optional) The reason for the transaction, which can be recorded in the transaction records for auditing purposes. If not provided, it will default to "Not Specified".
 * @param range (Optional) The range of cells to which the operation will be applied. If not provided, the currently active range will be used.
 * @returns {string |boolean} Returns true if the operation was successful, false otherwise.
 */
function applyMathToSelection(operation: Operation | string, value: number, isManualTransaction: boolean, transactionReason?: string, range?: GoogleAppsScript.Spreadsheet.Range, commentOnExpenditures: boolean = false): string | boolean {
  try {
    // must explicitly include undefined check because isManualTransaction is a boolean so it will always evaluate as its boolean value if you check if !isManualTranscation.
    if (
      operation === undefined ||
      value === undefined ||
      isManualTransaction === undefined
    ) {
      Logger.log(`You must have an operation, a value, and a manual transaction flag. Received operation: ${operation}, value: ${value}, isManualTransaction: ${isManualTransaction}`);
      SpreadsheetApp.getUi().alert(`You must have an operation, a value, and a manual transaction flag. Received operation: ${operation}, value: ${value}, isManualTransaction: ${isManualTransaction}`);
      return `You must have an operation, a value, and a manual transaction flag. Received operation: ${operation}, value: ${value}, isManualTransaction: ${isManualTransaction}`;
    }

    if (typeof value !== 'number' || isNaN(value)) {
      Logger.log("Value must be a number.");
      SpreadsheetApp.getUi().alert("Value must be a number.");
      return "Value must be a number.";
    }

    Logger.log(`Applying math operation: ${operation}, value: ${value}`);

    // Absolutely ENSURE there is no floating point precision issues
    // god damn floating point issues kms
    value = Number(value.toFixed(2));

    // Ensure the operation is in an enum format
    // Edge case handling
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

    if (!Object.values(Operation).includes(normalMapping)) {
      Logger.log("Invalid operation. Must be one of ADD, SUBTRACT, MULTIPLY, or DIVIDE.");
      SpreadsheetApp.getUi().alert("Invalid operation. Must be one of ADD, SUBTRACT, MULTIPLY, or DIVIDE.");
      return "Invalid operation. Must be one of ADD, SUBTRACT, MULTIPLY, or DIVIDE.";
    }

    if (operation === Operation.DIVIDE && value === 0) {
      Logger.log("You cannot divide by zero.");
      SpreadsheetApp.getUi().alert("You cannot divide by zero.");
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

    // First try the faster, API efficient method using the Sheets API, with error handling to fall back to the slower method
    try {
      const activeRangeList = SpreadsheetApp.getActiveSpreadsheet().getActiveRangeList();
      if (!activeRangeList) {
        Logger.log("No active range found. Please select cells to apply the operation to.");
        SpreadsheetApp.getUi().alert("No active range found. Please select cells to apply the operation to.");
        return "No active range found. Please select cells to apply the operation to.";
      }

      const ranges = activeRangeList.getRanges();
      const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

      // Use the Sheets API to apply the operation to all cells in the active range list
      const requests: GoogleAppsScript.Sheets.Schema.ValueRange[] = [];

      ranges.forEach(range => {
        // Get the current values of the range
        const rangeA1 = range.getA1Notation();
        const sheetName = range.getSheet().getName();
        const values = range.getValues();

        const startRowIndex = range.getRow();
        const startColIndex = range.getColumn();

        const updatedValues = values.map((row, rowIndex) => {
          const absoluteRowIndex = startRowIndex + rowIndex;

          return row.map(cell => {
            if (typeof cell === 'number' && !isNaN(cell)) {
              const balance = range.getSheet().getRange(absoluteRowIndex, BALANCE_COL).getValue();

              const result = Number(operationFunc(cell).toFixed(2));

              Logger.log(`Applying operation "${normalMapping}" to cell "${cell}" resulted in "${result}".`);

              const individualName = range.getSheet().getRange(absoluteRowIndex, 1).getValue();

              const targetColumnIndex = startColIndex;

              const targetColumnInitValue = range.getSheet().getRange(absoluteRowIndex, targetColumnIndex).getValue();

              transactionRecords.push({
                individual: individualName,
                type: operation === Operation.ADD || operation === Operation.MULTIPLY ? "Income" : "Expense",
                serviceProvided: `${isManualTransaction ? "Manual Balance Adjustment - " : ""}${transactionReason ?? "Not Specified"}`,
                quantity: 1,
                modifiedColumn: targetColumnIndex,
                tenderedMoney: Number(value.toFixed(2)),
                initialColumnAmount: Number(targetColumnInitValue.toFixed(2)),
                newColumnAmount: result,
                initialBalance: Number(balance.toFixed(2)),
                newBalance: 0,
                timestamp: new Date().toISOString()
              });

              return result;
            }
            Logger.log(`Non-numeric value "${cell}" found in range ${sheetName}!${rangeA1}. Skipping this cell.`);
            return cell; // Return the original value if it's not a number
          });
        });

        requests.push({
          range: `${sheetName}!${rangeA1}`,
          values: updatedValues
        });
      });


      // Send ONE batchUpdate request to the Sheets API to update all ranges at once
      if (requests.length > 0 && Sheets) {
        Sheets.Spreadsheets.Values.batchUpdate({
          valueInputOption: "USER_ENTERED",
          data: requests
        }, spreadsheetId);
      } else {
        SpreadsheetApp.getUi().alert("No valid ranges found to update.");
        return false;
      }

      SpreadsheetApp.flush()

      // Comment on the rows affected by the operation with the reason and amount, with error handling to ensure it doesn't interfere with the main operation if it fails
      if (commentOnExpenditures) {
        commentExpenditureOnSelection(ranges.flatMap(range => {
          const startRow = range.getRow();
          const numRows = range.getNumRows();
          return Array.from({ length: numRows }, (_, i) => startRow + i);
        }), `$${Number(value.toFixed(2))} - ${new Date().toLocaleDateString("en-US")} ${transactionReason || "Manual Adjustment"}`);
      }

      // Get the new balance for the selected range after performing the operation
      const newBalances = ranges.map(range => {
        const getRow = range.getRow();

        SpreadsheetApp.getUi().alert("Found row " + getRow + "  and column " + BALANCE_COL + " to get the new balance.");

        return Number(range.getSheet().getRange(getRow, BALANCE_COL).getValue().toFixed(2));
      })

      // Change the transaction balance property to the new calculated balance
      transactionRecords.forEach((record: TransactionRecord, index: number) => {
        record.newBalance = Number(newBalances[index].toFixed(2));
      });


      // After successfully applying the operations, add the transaction records
      if (transactionRecords.length > 0) {
        addTransactionRecords(transactionRecords);
      }

      return true;
    } catch (error: any) {
      // Try using the slower, API heavy method if the faster Sheets API doesn't work, with error handling for both methods
      SpreadsheetApp.getUi().alert(`Error occurred while applying the operation, trying a slower method. If this error persists, please contact the developer. Error details: ${error.message}`);

      try {
        let rangesToProcess: GoogleAppsScript.Spreadsheet.Range[] = [];

        if (range) {
          rangesToProcess.push(range);
        } else {
          const activeRangeList = SpreadsheetApp.getActiveSpreadsheet().getActiveRangeList();
          if (activeRangeList) {
            rangesToProcess = activeRangeList.getRanges();
          } else {
            Logger.log("No active range found. Please select cells to apply the operation to.");
            SpreadsheetApp.getUi().alert("No active range found. Please select cells to apply the operation to.");
            return "No active range found. Please select cells to apply the operation to.";
          }
        }

        // Apply the operation to each cell in each range, with error handling for non-numeric cells
        rangesToProcess.forEach(subRange => {
          const values = subRange.getValues();
          const targetSheet = subRange.getSheet();
          const startRowIndex = subRange.getRow();
          const startColIndex = subRange.getColumn();

          const updatedValues = values.map((row, rowIndex) => {
            const absoluteRowIndex = startRowIndex + rowIndex;

            return row.map(cell => {
              if (typeof cell === 'number' && !isNaN(cell)) {
                const balance = targetSheet.getRange(absoluteRowIndex, 2).getValue();

                const result = Number(operationFunc(cell).toFixed(2));

                Logger.log(`Applying operation "${normalMapping}" to cell "${cell}" resulted in "${result}".`);

                // Absolutely ENSURE there is no floating point precision issues
                const individualName = targetSheet.getRange(absoluteRowIndex, 1).getValue();
                const targetColumnIndex = startColIndex;
                const targetColumnInitValue = targetSheet.getRange(absoluteRowIndex, targetColumnIndex).getValue();

                // If the operation is applied on column 3
                transactionRecords.push({
                  individual: individualName,
                  type: operation === Operation.ADD || operation === Operation.MULTIPLY ? "Income" : "Expense",
                  serviceProvided: `${isManualTransaction ? "Manual Balance Adjustment - " : ""}${transactionReason ?? "Not Specified"}`,
                  quantity: 1,
                  modifiedColumn: targetColumnIndex,
                  tenderedMoney: Number(value.toFixed(2)),
                  initialColumnAmount: Number(targetColumnInitValue.toFixed(2)),
                  newColumnAmount: result,
                  initialBalance: Number(balance.toFixed(2)),
                  newBalance: 0,
                  timestamp: new Date().toISOString()
                });

                return result;
              }
              Logger.log(`Non-numeric value "${cell}" found. Skipping this cell.`);
              return cell; // Return the original value if it's not a number
            });
          });

          subRange.setValues(updatedValues);

          // Update balance column for each row in the range
          // Get the new balance for the selected range after performing the operation
          const newBalances = rangesToProcess.map(range => {
            const getRow = range.getRow();

            SpreadsheetApp.getUi().alert("Found row " + getRow + "  and column " + BALANCE_COL + " to get the new balance.");

            return Number(range.getSheet().getRange(getRow, BALANCE_COL).getValue().toFixed(2));
          });

          // Update transaction records with new balances
          transactionRecords.forEach((record, index) => {
            record.newBalance = Number(newBalances[index].toFixed(2));
          });

          // After successfully applying the operations, add the transaction records
          if (transactionRecords.length > 0) {
            addTransactionRecords(transactionRecords);
          }
        });
      } catch (error: any) {
        SpreadsheetApp.getUi().alert(`Error occurred while applying the operation, trying a slower method. If this error persists, please contact the developer. Error details: ${error.message}`);
        return `Error occurred while applying the operation, trying a slower method. If this error persists, please contact the developer. Error details: ${error.message}`;
      }
    }
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(`Error occured in applyMathToSelection: ${error.message}`);
    return error.message;
  }
  return true;
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
      Logger.log("You must provide both a cell range and a note string.");
      return false;
    }

    if (typeof comment !== 'string') {
      Logger.log("Comment must be a string.");
      return false;
    }

    if (comment.length > 255) {
      Logger.log("Comment cannot exceed 255 characters.");
      return false;
    }
    if (comment.includes("\n") || comment.includes("\r")) {
      Logger.log("Comment cannot contain line breaks.");
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
    Logger.log(`Error occurred in commentExpenditureOnSelection: ${error.message}`);
    try {
      SpreadsheetApp.getUi().alert(`Error occurred in commentExpenditureOnSelection: ${error.message}`);
    } catch (e) {
      Logger.log(`Additionally, failed to show alert in commentExpenditureOnSelection error handling: ${e instanceof Error ? e.message : String(e)}`);
    }
    return false;
  }
}