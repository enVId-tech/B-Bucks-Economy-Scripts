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
 * Uses an operand and a value to apply to the existing value of all selected cells.
 * @param operation The mathematical operation to apply to the selected cells. Must be one of "ADD", "SUBTRACT", "MULTIPLY", or "DIVIDE".
 * @param value The value to use in the mathematical operation on the selected cells. Must be a number.
 * @param transactionReason (Optional) The reason for the transaction, which can be recorded in the transaction records for auditing purposes. If not provided, it will default to "Not Specified".
 * @param range (Optional) The range of cells to which the operation will be applied. If not provided, the currently active range will be used.
 * @returns {string |boolean} Returns true if the operation was successful, false otherwise.
 */
function applyMathToSelection(operation: Operation | string, value: number, isManualTransaction: boolean, transactionReason?: string, range?: GoogleAppsScript.Spreadsheet.Range): string | boolean {
  try {
    if (!operation || !value || !isManualTransaction) {
      Logger.log(`You must have an operation, a value, and a manual transaction flag. Received operation: ${operation}, value: ${value}, isManualTransaction: ${isManualTransaction}`);
      SpreadsheetApp.getUi().alert(`You must have an operation, a value, and a manual transaction flag. Received operation: ${operation}, value: ${value}, isManualTransaction: ${isManualTransaction}`);
      return `You must have an operation, a value, and a manual transaction flag. Received operation: ${operation}, value: ${value}, isManualTransaction: ${isManualTransaction}`;
    }

    if (typeof value !== 'number' || isNaN(value)) {
      Logger.log("Value must be a number.");
      SpreadsheetApp.getUi().alert("Value must be a number.");
      return "Value must be a number.";
    }

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
      [Operation.ADD]: (n) => n + value,
      [Operation.SUBTRACT]: (n) => n - value,
      [Operation.MULTIPLY]: (n) => n * value,
      [Operation.DIVIDE]: (n) => n / value,
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

        const updatedValues = values.map((row, rowIndex) => {

          const absoluteRowIndex = startRowIndex + rowIndex;

          return row.map(cell => {
            if (typeof cell === 'number' && !isNaN(cell)) {
              const result = operationFunc(cell);

              Logger.log(`Applying operation "${normalMapping}" to cell "${cell}" resulted in "${result}".`);

              const individualName = range.getSheet().getRange(absoluteRowIndex, 1).getValue();

              if (result && typeof result === 'number' && !isNaN(result)) {
                transactionRecords.push({
                  individual: individualName,
                  type: operation === Operation.ADD || operation === Operation.MULTIPLY ? "Income" : "Expense", // TODO: Add support for investments and other transaction types in the future
                  service: `${isManualTransaction ? "Manual Balance Adjustment - " : ""}${transactionReason ?? "Not Specified"}`,
                  initialAmount: cell,
                  tenderedAmount: value,
                  finalAmount: result,
                  quantityOfServices: 1,
                  timestamp: new Date()
                });
              } else {
                Logger.log(`Result of operation "${normalMapping}" on cell "${cell}" is not a valid number. Result: "${result}". Skipping transaction record.`);
              }

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
      }

      // Comment on the rows affected by the operation with the reason and amount, with error handling to ensure it doesn't interfere with the main operation if it fails
      commentExpenditureOnSelection(ranges.flatMap(range => {
        const startRow = range.getRow();
        const numRows = range.getNumRows();
        return Array.from({ length: numRows }, (_, i) => startRow + i);
      }), `$${value} - ${new Date().toLocaleDateString("en-US")} ${transactionReason || "Manual Adjustment"}`);

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

          const updatedValues = values.map((row, rowIndex) => {
            const absoluteRowIndex = startRowIndex + rowIndex;

            return row.map(cell => {
              if (typeof cell === 'number' && !isNaN(cell)) {
                const result = operationFunc(cell);

                Logger.log(`Applying operation "${normalMapping}" to cell "${cell}" resulted in "${result}".`);

                if (result && typeof result === 'number' && !isNaN(result)) {
                  const individualName = targetSheet.getRange(absoluteRowIndex, 1).getValue();

                  transactionRecords.push({
                    individual: individualName,
                    type: operation === Operation.ADD || operation === Operation.MULTIPLY ? "Income" : "Expense", // TODO: Add support for investments and other transaction types in the future
                    service: `${isManualTransaction ? "Manual Balance Adjustment - " : ""}${transactionReason ?? "Not Specified"}`,
                    initialAmount: cell,
                    tenderedAmount: value,
                    finalAmount: result,
                    quantityOfServices: 1,
                    timestamp: new Date()
                  });
                } else {
                  Logger.log(`Result of operation "${normalMapping}" on cell "${cell}" is not a valid number. Result: "${result}". Skipping transaction record.`);
                }

                return result;
              }
              Logger.log(`Non-numeric value "${cell}" found. Skipping this cell.`);
              return cell; // Return the original value if it's not a number
            });
          });

          subRange.setValues(updatedValues);

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
 * Moves a specified amount from one cell to another, ensuring that the source cell has enough value and that both cells contain numbers.
 * @param amount The amount to move from the source cell to the final cell.
 * @param finalCells The range of cells to which the amount will be added.
 * @param initialCells (Optional) The range of cells from which the amount will be subtracted. If not provided, the currently active range will be used.
 * @returns {boolean} Returns true if the operation was successful, false otherwise.
 */
function moveToSelection(amount: number, finalCells: GoogleAppsScript.Spreadsheet.Range, initialCells?: GoogleAppsScript.Spreadsheet.Range): boolean {
  try {
    if (!amount || !finalCells) throw new Error("You must have an amount and final cells.");

    let sheet = SpreadsheetApp.getActiveSpreadsheet();

    let sourceCells = initialCells || sheet.getActiveRange();

    if (!sourceCells) {
      Logger.log("No active range found. Please select a cell to move from.");
      return false
    }

    let sourceValues = sourceCells.getValues();

    if (typeof sourceValues !== 'number' || isNaN(sourceValues)) {
      Logger.log("The source cell must contain a number.");
      return false;
    }
    if (sourceValues < amount) {
      Logger.log("The source cell does not have enough value to move.");
      return false;
    }

    let finalValues = finalCells.getValues();

    if (typeof finalValues !== 'number' || isNaN(finalValues)) {
      Logger.log("The final cell must contain a number.");
      return false;
    }

    sourceCells.setValue(sourceValues - amount);
    finalCells.setValue(finalValues + amount);
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(`Error occured in moveToSelection: ${error.message}`);
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
    const EXPENDITURE_COL = 5;
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
      const currentNote = sheet.getRange(row, EXPENDITURE_COL).getNote();
      if (currentNote) {
        sheet.getRange(row, EXPENDITURE_COL).setNote(currentNote + "\n" + comment);
      } else {
        sheet.getRange(row, EXPENDITURE_COL).setNote(comment);
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