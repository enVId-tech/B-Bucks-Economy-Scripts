// --- Coded by Erick Tran for Mr. Banderas, 2026 --- enVId Tech

/**
 * Uses an operand and a value to apply to the existing value of all selected cells.
 */
function applyMathToSelection(operation, value) {
  if (!operation || !value) throw new Error("You must have an operation and a value");

  let sheet = SpreadsheetApp.getActiveSpreadsheet();

  let activeRange = sheet.getActiveRange();

  let values = activeRange.getValues();

  if (!values) throw new Error("You must select a cell.");

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

function moveToCell(amount, initialCell, finalCell) {
  
}

function commentOnCell() {

}