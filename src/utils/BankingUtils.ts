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
 * @param transactionReason (Optional) The reason for the transaction, which can be recorded in the transaction records for auditing purposes. If not provided, it will default to "Not Specified".
 * @param range (Optional) The range of cells to which the operation will be applied. If not provided, the currently active range will be used.
 * @returns {string |boolean} Returns true if the operation was successful, false otherwise.
 */
function applyMathToSelection(operation: Operation | string, unitPrice: number, quantity: number, isManualTransaction: boolean, transactionReason?: string, range?: GoogleAppsScript.Spreadsheet.Range, commentOnExpenditures: boolean = false): string | boolean {
  try {
    // must explicitly include undefined check because isManualTransaction is a boolean so it will always evaluate as its boolean value if you check if !isManualTranscation.
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
    // god damn floating point issues kms
    const value = Number((unitPrice * quantity).toFixed(2));

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
      log("Invalid operation. Must be one of ADD, SUBTRACT, MULTIPLY, or DIVIDE.", true);
      return "Invalid operation. Must be one of ADD, SUBTRACT, MULTIPLY, or DIVIDE.";
    }

    if (operation === Operation.DIVIDE && value === 0) {
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

    // First try the faster, API efficient method using the Sheets API, with error handling to fall back to the slower method
    try {
      const activeRangeList = SpreadsheetApp.getActiveSpreadsheet().getActiveRangeList();
      if (!activeRangeList) {
        log("No active range found. Please select cells to apply the operation to.", true);
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

              if (balance === undefined || balance === null || isNaN(balance)) {
                log(`Balance value is invalid for row ${absoluteRowIndex}. Skipping this cell.`, true);
                return cell; // Return the original value if balance is invalid
              }

              const result = operationFunc(cell);

              const individualName = range.getSheet().getRange(absoluteRowIndex, 1).getValue();

              if (!individualName) {
                log(`Individual name is missing for row ${absoluteRowIndex}. Skipping this cell.`, true);
                return cell; // Return the original value if individual name is missing
              }

              const targetColumnIndex = startColIndex;

              const targetColumnInitValue = range.getSheet().getRange(absoluteRowIndex, targetColumnIndex).getValue();

              if (targetColumnInitValue === undefined || targetColumnInitValue === null || isNaN(targetColumnInitValue)) {
                log(`Target column initial value is invalid for row ${absoluteRowIndex}. Skipping this cell.`, true);
                return cell; // Return the original value if target column initial value is invalid
              }

              const rawTenderedMoney = unitPrice * quantity;
              const tenderedMoney = Number(rawTenderedMoney.toFixed(2));

              if (!tenderedMoney || isNaN(tenderedMoney)) {
                log(`Tendered money calculation is invalid for row ${absoluteRowIndex}. Skipping this cell.`, true);
                return cell; // Return the original value if tendered money is invalid
              }

              transactionRecords.push({
                individual: individualName,
                type: operation === Operation.ADD || operation === Operation.MULTIPLY ? "Income" : "Expense",
                serviceProvided: `${isManualTransaction ? "Manual Balance Adjustment" : ""} ${transactionReason ? '-' : ""} ${transactionReason ?? "Not Specified"}`,
                unitPrice: unitPrice,
                quantity: quantity,
                modifiedColumn: targetColumnIndex,
                tenderedMoney: Number(tenderedMoney.toFixed(2)),
                initialColumnAmount: Number(targetColumnInitValue.toFixed(2)),
                newColumnAmount: result,
                initialBalance: Number(balance.toFixed(2)),
                newBalance: 0,
                timestamp: new Date().toISOString()
              });

              return result;
            }

            log(`Non-numeric value "${cell}" found in range ${sheetName}!${rangeA1}. Skipping this cell.`, true);
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
        log("No valid ranges found to update.", true);
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

      const newBalancesRecords: number[] = [];

      ranges.forEach(range => {
        const values = range.getValues();

        const startRowIndex = range.getRow();

        values.map((row, rowIndex) => {
          const absoluteRowIndex = startRowIndex + rowIndex;

          return row.map(cell => {
            if (typeof cell === 'number' && !isNaN(cell)) {
              const balance = range.getSheet().getRange(absoluteRowIndex, BALANCE_COL).getValue();

              if (balance === undefined || balance === null || isNaN(balance)) {
                log(`Balance value is invalid for row ${absoluteRowIndex}. Skipping this cell.`, true);
                return cell; // Return the original value if balance is invalid
              }
              newBalancesRecords.push(Number(balance.toFixed(2)));
              return cell;
            }
          });
        });
      });

      log(`New balances calculated: ${newBalancesRecords.join(", ")}`, false);

      // Change the transaction balance property to the new calculated balance
      transactionRecords.forEach((record: TransactionRecord, index: number) => {
        record.newBalance = Number(newBalancesRecords[index].toFixed(2));
      });

      // After successfully applying the operations, add the transaction records
      if (transactionRecords.length > 0) {
        addTransactionRecords(transactionRecords);
      }

      return true;
    } catch (error: any) {
      // Try using the slower, API heavy method if the faster Sheets API doesn't work, with error handling for both methods
      log(`Error occurred while applying the operation, trying a slower method. If this error persists, please contact the developer. Error details: ${error.message}`, true);

      try {
        let rangesToProcess: GoogleAppsScript.Spreadsheet.Range[] = [];

        if (range) {
          rangesToProcess.push(range);
        } else {
          const activeRangeList = SpreadsheetApp.getActiveSpreadsheet().getActiveRangeList();
          if (activeRangeList) {
            rangesToProcess = activeRangeList.getRanges();
          } else {
            log("No active range found. Please select cells to apply the operation to.", true);
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

                const result = operationFunc(cell);

                log(`Applying operation "${normalMapping}" to cell "${cell}" resulted in "${result}".`, false);

                // Absolutely ENSURE there is no floating point precision issues
                const individualName = targetSheet.getRange(absoluteRowIndex, 1).getValue();

                if (!individualName) {
                  log(`Individual name is missing for row ${absoluteRowIndex}. Skipping this cell.`, true);
                  return cell; // Return the original value if individual name is missing
                }

                const targetColumnIndex = startColIndex;

                if (targetColumnIndex === undefined || targetColumnIndex === null || isNaN(targetColumnIndex)) {
                  log(`Target column index is invalid for row ${absoluteRowIndex}. Skipping this cell.`, true);
                  return cell; // Return the original value if target column index is invalid
                }

                const targetColumnInitValue = targetSheet.getRange(absoluteRowIndex, targetColumnIndex).getValue();

                if (targetColumnInitValue === undefined || targetColumnInitValue === null || isNaN(targetColumnInitValue)) {
                  log(`Target column initial value is invalid for row ${absoluteRowIndex}. Skipping this cell.`, true);
                  return cell; // Return the original value if target column initial value is invalid
                }

                const rawTenderedMoney = unitPrice * quantity;
                const tenderedMoney = Number(rawTenderedMoney.toFixed(2));

                if (!tenderedMoney || isNaN(tenderedMoney)) {
                  log(`Tendered money calculation is invalid for row ${absoluteRowIndex}. Skipping this cell.`, true);
                  return cell; // Return the original value if tendered money is invalid
                }

                // If the operation is applied on column 3
                transactionRecords.push({
                  individual: individualName,
                  type: operation === Operation.ADD || operation === Operation.MULTIPLY ? "Income" : "Expense",
                  serviceProvided: `${isManualTransaction ? "Manual Balance Adjustment - " : ""}${transactionReason ?? "Not Specified"}`,
                  unitPrice: unitPrice,
                  quantity: quantity,
                  modifiedColumn: targetColumnIndex,
                  tenderedMoney: tenderedMoney,
                  initialColumnAmount: Number(targetColumnInitValue.toFixed(2)),
                  newColumnAmount: result,
                  initialBalance: Number(balance.toFixed(2)),
                  newBalance: 0,
                  timestamp: new Date().toISOString()
                });

                return result;
              }
              log(`Non-numeric value "${cell}" found in range ${subRange.getSheet().getName()}!${subRange.getA1Notation()}. Skipping this cell.`, true);
              return cell; // Return the original value if it's not a number
            });
          });

          subRange.setValues(updatedValues);

          // Update balance column for each row in the range
          // Get the new balance for the selected range after performing the operation

          // Update transaction records with new balances
          const newBalancesRecords: number[] = [];

          rangesToProcess.forEach(range => {
            const values = range.getValues();

            const startRowIndex = range.getRow();

            values.map((row, rowIndex) => {
              const absoluteRowIndex = startRowIndex + rowIndex;

              return row.map(cell => {
                if (typeof cell === 'number' && !isNaN(cell)) {
                  const balance = range.getSheet().getRange(absoluteRowIndex, BALANCE_COL).getValue();

                  if (balance === undefined || balance === null || isNaN(balance)) {
                    log(`Balance value is invalid for row ${absoluteRowIndex}. Skipping this cell.`, true);
                    return cell; // Return the original value if balance is invalid
                  }
                  newBalancesRecords.push(Number(balance.toFixed(2)));
                  return cell;
                }
              });
            });
          });
          
          // After successfully applying the operations, add the transaction records
          if (transactionRecords.length > 0) {
            addTransactionRecords(transactionRecords);
          }
        });
      } catch (error: any) {
        log(`Error occurred while applying the operation, trying a slower method. If this error persists, please contact the developer. Error details: ${error.message}`, true);
        return `Error occurred while applying the operation, trying a slower method. If this error persists, please contact the developer. Error details: ${error.message}`;
      }
    }
  } catch (error: any) {
    log(`Error occurred in applyMathToSelection: ${error.message}`, true);
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