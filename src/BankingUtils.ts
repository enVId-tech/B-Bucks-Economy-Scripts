// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. All rights reserved.
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains utility functions for various operations in the B-Bucks Economy Scripts project, including mathematical operations on cell values, moving values between cells, and adding comments to cells.

enum Operation {
  ADD = "ADD",
  SUBTRACT = "SUBTRACT",
  MULTIPLY = "MULTIPLY",
  DIVIDE = "DIVIDE"
}
/**
 * Uses an operand and a value to apply to the existing value of all selected cells.
 * @param operation The mathematical operation to apply to the selected cells. Must be one of "ADD", "SUBTRACT", "MULTIPLY", or "DIVIDE".
 * @param value The value to use in the mathematical operation. Must be a number.
 * @param range (Optional) The range of cells to which the operation will be applied. If not provided, the currently active range will be used.
 * @returns {boolean} Returns true if the operation was successful, false otherwise.
 */
function applyMathToSelection(operation: Operation, value: number, range?: GoogleAppsScript.Spreadsheet.Range): boolean {
  try {
    if (!operation || !value) {
      Logger.log("You must have an operation and a value.");
      return false;
    }
    if (typeof value !== 'number' || isNaN(value)) {
      Logger.log("Value must be a number.");
      return false;
    }
    if (!Object.values(Operation).includes(operation)) {
      Logger.log("Invalid operation. Must be one of ADD, SUBTRACT, MULTIPLY, or DIVIDE.");
      return false;
    }

    let sheet = SpreadsheetApp.getActiveSpreadsheet();

    let activeRange = range || sheet.getActiveRange();

    if (!activeRange) throw new Error("You must select a cell.");

    let values = activeRange.getValues();

    if (operation === Operation.DIVIDE && value === 0) {
      Logger.log("You cannot divide by zero.");
      return false;
    }

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        let cellValue = values[r][c];
        if (typeof cellValue === 'number' && !isNaN(cellValue)) {
          switch (operation) {
            case Operation.ADD:
              values[r][c] = cellValue + value;
              break;
            case Operation.SUBTRACT:
              values[r][c] = cellValue - value;
              break;
            case Operation.MULTIPLY:
              values[r][c] = cellValue * value;
              break;
            case Operation.DIVIDE:
              values[r][c] = cellValue / value;
              break;
          }
        }
      }
    }

    activeRange.setValues(values);
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(error.message);
    return false;
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
    SpreadsheetApp.getUi().alert(error.message);
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
    SpreadsheetApp.getUi().alert(error.message);
  }
  return true;
}