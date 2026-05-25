// --- Coded by Erick Tran for Mr. Banderas, 2026 --- enVId Tech

/**
 * Uses an operand and a value to apply to the existing value of all selected cells.
 */
function applyMathToSelection(operation: string, value: number): void {
  if (!operation || !value) throw new Error("You must have an operation and a value");

  let sheet = SpreadsheetApp.getActiveSpreadsheet();

  let activeRange = sheet.getActiveRange();

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
}

function moveToCell(amount: number, finalCell: GoogleAppsScript.Spreadsheet.Range, initialCell: GoogleAppsScript.Spreadsheet.Range) {

}

function commentOnCell() {

}