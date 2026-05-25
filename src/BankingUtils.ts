// Coded by Erick Tran for Mr. Banderas, 2026
// This file contains utility functions for banking operations in the B-Bucks Economy Scripts project.
// Script licensed under the MIT License.
// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts

/**
 * Uses an operand and a value to apply to the existing value of all selected cells.
 */
function applyMathToSelection(operation: string, value: number, range?: GoogleAppsScript.Spreadsheet.Range): void {
  try {
    if (!operation || !value) throw new Error("You must have an operation and a value");

    let sheet = SpreadsheetApp.getActiveSpreadsheet();

    let activeRange = range || sheet.getActiveRange();

    if (!activeRange) throw new Error("You must select a cell.");

    let values = activeRange.getValues();

    if (operation === "DIVIDE" && value === 0) throw new Error("You cannot divide by zero.");

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        let cellValue = values[r][c];
        if (typeof cellValue === 'number' && !isNaN(cellValue)) {
          values[r][c] = operation === "ADD" ? (cellValue + value) : operation === "SUBTRACT" ? (cellValue - value) : operation === "MULTIPLY" ? (cellValue * value) : operation === "DIVIDE" ? (cellValue / value) : cellValue;
        }
      }
    }

    activeRange.setValues(values);
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(error.message);
  }
}

/**
 * Moves a specified amount from one cell to another, ensuring that the source cell has enough value and that both cells contain numbers.
 * @param amount The amount to move from the source cell to the final cell.
 * @param finalCells The range of cells to which the amount will be added.
 * @param initialCells (Optional) The range of cells from which the amount will be subtracted. If not provided, the currently active range will be used.
 * @returns void
 */
function moveToSelection(amount: number, finalCells: GoogleAppsScript.Spreadsheet.Range, initialCells?: GoogleAppsScript.Spreadsheet.Range): void {
  try {
    if (!amount || !finalCells) throw new Error("You must have an amount and final cells.");

    let sheet = SpreadsheetApp.getActiveSpreadsheet();

    let sourceCells = initialCells || sheet.getActiveRange();

    if (!sourceCells) throw new Error("You must select a cell.");

    let sourceValues = sourceCells.getValues();

    if (typeof sourceValues !== 'number' || isNaN(sourceValues)) throw new Error("The source cell must contain a number.");
    if (sourceValues < amount) throw new Error("The source cell does not have enough value to move.");

    let finalValues = finalCells.getValues();

    if (typeof finalValues !== 'number' || isNaN(finalValues)) throw new Error("The final cell must contain a number.");

    sourceCells.setValue(sourceValues - amount);
    finalCells.setValue(finalValues + amount);
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(error.message);
  }
}

function commentOnSelection(cells: GoogleAppsScript.Spreadsheet.Range, comment: string): void {
  try {
    if (!cells || !comment) throw new Error("You must have cells and a comment.");

    if (comment.length > 255) throw new Error("Comment cannot exceed 255 characters.");
    if (comment.includes("\n") || comment.includes("\r")) throw new Error("Comment cannot contain line breaks.");
    if (typeof comment !== 'string') throw new Error("Comment must be a string.");

    cells.getValue().setComment(comment);
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(error.message);
  }
}