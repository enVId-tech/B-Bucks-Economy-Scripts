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

// Transaction record interface
interface TransactionRecord {
  individual: string;
  type: "Income" | "Expense" | "Investment";
  service: string;
  initialAmount: number;
  tenderedAmount: number;
  finalAmount: number;
  quantityOfServices: number;
  timestamp: Date;
}

/**
 * Uses an operand and a value to apply to the existing value of all selected cells.
 * @param operation The mathematical operation to apply to the selected cells. Must be one of "ADD", "SUBTRACT", "MULTIPLY", or "DIVIDE".
 * @param value The value to use in the mathematical operation on the selected cells. Must be a number.
 * @param transactionReason (Optional) The reason for the transaction, which can be recorded in the transaction records for auditing purposes. If not provided, it will default to "Not Specified".
 * @param range (Optional) The range of cells to which the operation will be applied. If not provided, the currently active range will be used.
 * @returns {string |boolean} Returns true if the operation was successful, false otherwise.
 */
function applyMathToSelection(operation: Operation | string, value: number, isManualTransaction: boolean, transactionReason: string = "Not Specified", range?: GoogleAppsScript.Spreadsheet.Range): string | boolean {
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

      const operations: Record<Operation, (n: number) => number> = {
        [Operation.ADD]: (n) => n + value,
        [Operation.SUBTRACT]: (n) => n - value,
        [Operation.MULTIPLY]: (n) => n * value,
        [Operation.DIVIDE]: (n) => n / value,
      };

      const operationFunc = operations[normalMapping];

      ranges.forEach(range => {
        // Get the current values of the range
        const rangeA1 = range.getA1Notation();
        const sheetName = range.getSheet().getName();
        const values = range.getValues();
        const individualName = range.getSheet().getRange(range.getRow(), 1).getValue();

        const updatedValues = values.map(row => row.map(cell => {
          if (typeof cell === 'number' && !isNaN(cell)) {
            const result = operationFunc(cell);
            Logger.log(`Applying operation "${normalMapping}" to cell "${cell}" resulted in "${result}".`);
            if (result && typeof result === 'number' && !isNaN(result)) {
              addTransactionRecord({
                individual: individualName,
                type: operation === Operation.ADD || operation === Operation.MULTIPLY ? "Income" : "Expense", // TODO: Add support for investments and other transaction types in the future
                service: `${isManualTransaction ? "Manual Balance Adjustment - " : ""}${transactionReason}`,
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
        }));

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

        // Define the operations in a more structured map instead of looping
        // Maps the operation enum to a function that takes a number and returns the result of applying the operation with the provided value
        const operations: Record<Operation, (n: number) => number> = {
          [Operation.ADD]: (n) => n + value,
          [Operation.SUBTRACT]: (n) => n - value,
          [Operation.MULTIPLY]: (n) => n * value,
          [Operation.DIVIDE]: (n) => n / value,
        };

        // Use a const function to apply the operation to the cell
        const operationFunc = operations[normalMapping];

        // Apply the operation to each cell in each range, with error handling for non-numeric cells
        rangesToProcess.forEach(subRange => {
          const values = subRange.getValues();

          const updatedValues = values.map(row => row.map(cell => {
            if (typeof cell === 'number' && !isNaN(cell)) {
              const result = operationFunc(cell);
              Logger.log(`Applying operation "${normalMapping}" to cell "${cell}" resulted in "${result}".`);

              if (result && typeof result === 'number' && !isNaN(result)) {
                const individualName = subRange.getSheet().getRange(subRange.getRow(), 1).getValue();
                addTransactionRecord({
                  individual: individualName,
                  type: operation === Operation.ADD || operation === Operation.MULTIPLY ? "Income" : "Expense", // TODO: Add support for investments and other transaction types in the future
                  service: `${isManualTransaction ? "Manual Balance Adjustment - " : ""}${transactionReason}`,
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
          }));

          subRange.setValues(updatedValues);
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
 * Adds a comment to the selected cells with the provided comment text, ensuring that the comment is valid and does not exceed character limits or contain line breaks.
 * @param cells The range of cells to which the comment will be added.
 * @param comment The text of the comment to be added to the cells. Must be a string, cannot exceed 255 characters, and cannot contain line breaks.
 * @returns {boolean} Returns true if the operation was successful, false otherwise.
 */
function commentOnSelection(cells: GoogleAppsScript.Spreadsheet.Range, comment: string): boolean {
  try {
    if (!cells || !comment) {
      Logger.log("You must have cells and a comment.");
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
    if (typeof comment !== 'string') {
      Logger.log("Comment must be a string.");
      return false;
    }

    cells.getValue().setComment(comment);
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(`Error occured in commentOnSelection: ${error.message}`);
  }
  return true;
}

/**
 * Adds a transaction record to the "Transactions Records" sheet with the provided details, ensuring that all required information is valid and properly formatted.
 */
function addTransactionRecord({individual, type, service, initialAmount, tenderedAmount, finalAmount, quantityOfServices, timestamp}: TransactionRecord): boolean | string {
  try {
    const SHEET_NAME = "Transactions"

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    
    // -- Add edge case checking for the sheet and all fields to ensure data integrity when compiled to JavaScript --
    if (!sheet) {
      Logger.log(`Sheet "${SHEET_NAME}" not found. Please ensure the sheet exists.`);
      SpreadsheetApp.getUi().alert(`Sheet "${SHEET_NAME}" not found. Please ensure the sheet exists.`);
      return `Sheet "${SHEET_NAME}" not found. Please ensure the sheet exists.`;
    }

    if (!individual || !type || !service || initialAmount === undefined || tenderedAmount === undefined || finalAmount === undefined || quantityOfServices === undefined || !timestamp) {
      Logger.log("All fields are required and must be valid.");
      SpreadsheetApp.getUi().alert("All fields are required and must be valid.");
      return "All fields are required and must be valid.";
    }

    if (typeof initialAmount !== 'number' || isNaN(initialAmount) || typeof tenderedAmount !== 'number' || isNaN(tenderedAmount) || typeof finalAmount !== 'number' || isNaN(finalAmount) || typeof quantityOfServices !== 'number' || isNaN(quantityOfServices)) {
      Logger.log("All amount fields must be valid numbers.");
      SpreadsheetApp.getUi().alert("All amount fields must be valid numbers.");
      return "All amount fields must be valid numbers.";
    }

    if (typeof timestamp === 'string') {
      timestamp = new Date(timestamp);
    }

    if (!(timestamp instanceof Date) || isNaN(timestamp.getTime())) {
      Logger.log("Timestamp must be a valid date.");
      SpreadsheetApp.getUi().alert("Timestamp must be a valid date.");
      return "Timestamp must be a valid date.";
    }
    
    if (type !== "Income" && type !== "Expense" && type !== "Investment") {
      Logger.log("Type must be one of 'Income', 'Expense', or 'Investment'.");
      SpreadsheetApp.getUi().alert("Type must be one of 'Income', 'Expense', or 'Investment'.");
      return "Type must be one of 'Income', 'Expense', or 'Investment'.";
    }

    if (typeof individual !== 'string' || typeof service !== 'string') {
      Logger.log("Individual and service must be strings.");
      SpreadsheetApp.getUi().alert("Individual and service must be strings.");
      return "Individual and service must be strings.";
    }

    if (individual.length === 0 || service.length === 0) {
      Logger.log("Individual and service cannot be empty.");
      SpreadsheetApp.getUi().alert("Individual and service cannot be empty.");
      return "Individual and service cannot be empty.";
    }

    const prevTransactionId = sheet.getLastRow() > 1 ? sheet.getRange(sheet.getLastRow(), 1).getValue() : 0;
    const newTransactionId = prevTransactionId + 1;

    sheet.appendRow([
      newTransactionId,
      individual,
      type,
      service,
      initialAmount,
      tenderedAmount,
      finalAmount,
      quantityOfServices,
      timestamp
    ]);

    return true;
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(`Error occured in addTransactionRecord: ${error.message}`);
    return false;
  }
}