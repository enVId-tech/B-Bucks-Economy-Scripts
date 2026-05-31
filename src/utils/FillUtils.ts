// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. 
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains utility functions for various operations in the B-Bucks Economy Scripts project.

interface CSVData {
    identifier: number;
    individuals: string[];
}

// Sheet is "Reports -> Gb -> Rb Rstr"

/**
 * Fills the sheets with data parsed from the CSV string. Each sheet is named after the identifier and contains the associated individuals starting from a specified cell. If a sheet for an identifier doesn't exist, it creates one by copying a template sheet.
 * @param data  The raw CSV data as a string, where each row represents an identifier and its associated individuals. The identifier can be separated from the individuals by either a hyphen or a comma.
 * @returns {boolean} Returns true if the operation was successful, false otherwise.
 */
function fillSheetsWithData(data: string): boolean {
    try {
        const entry: CSVData = JSON.parse(data);
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const spreadsheetId = spreadsheet.getId();
        let sheets = spreadsheet.getSheets();

        const cellToStartFrom = "A7"; // Starting cell for filling data
        const cellToSetName = "A2"; // Cell to set the identifier name

        let targetSheet = sheets.find(s => s.getName() === "Period " + entry.identifier);

        if (!targetSheet) {
            const templateSheet = sheets.find(s => s.getName() === "Template");
            if (templateSheet) {
                targetSheet = templateSheet.copyTo(spreadsheet);
                targetSheet.setName("Period " + entry.identifier);
                sheets = spreadsheet.getSheets();
            } else {
                Logger.log("Template sheet not found. Skipping period: " + entry.identifier);
                return false;
            }
        }

        // Attempt to add with the Google Sheets API first, only if it fails run the fallback
        try {
            const startRow = targetSheet.getRange(cellToStartFrom).getRow();
            const lengthParsed = entry.individuals.length;
            if (lengthParsed === 0) {
                Logger.log("Length parsed is 0 for identifier: " + entry.identifier + ". Clearing existing data and skipping filling.");
                return true;
            }

            const lastRowWithData = targetSheet.getLastRow();
            const currentMaxRows = targetSheet.getMaxRows();
            const maxColumns = targetSheet.getMaxColumns();
            const targetSheetId = targetSheet.getSheetId();

            const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
            const valueUpdates: GoogleAppsScript.Sheets.Schema.ValueRange[] = [];

            // Clear existing content from row (startRow) down.
            if (lastRowWithData >= startRow) {
                requests.push({
                    updateCells: {
                        range: {
                            sheetId: targetSheetId,
                            startRowIndex: startRow,
                            endRowIndex: lastRowWithData,
                            startColumnIndex: 0,
                            endColumnIndex: maxColumns
                        },
                        fields: "userEnteredValue"
                    }
                });
            }

            valueUpdates.push({
                range: `${targetSheet.getName()}!${cellToSetName}`,
                values: [[`Period ${entry.identifier}`]]
            });

            // The total number of rows needed is the starting row plus the length of the individuals list minus one (since the starting row is inclusive)
            const maxRowsNeeded = startRow + lengthParsed - 1;

            if (maxRowsNeeded > currentMaxRows) {
                requests.push({
                    insertRange: {
                        range: {
                            sheetId: targetSheetId,
                            startRowIndex: startRow,
                            endRowIndex: maxRowsNeeded,
                            startColumnIndex: 0,
                            endColumnIndex: maxColumns
                        },
                        shiftDimension: "ROWS",
                    }
                });
            } else if (currentMaxRows > maxRowsNeeded) {
                requests.push({
                    deleteRange: {
                        range: {
                            sheetId: targetSheetId,
                            startRowIndex: maxRowsNeeded,
                            endRowIndex: currentMaxRows,
                            startColumnIndex: 0,
                            endColumnIndex: maxColumns
                        }
                    }
                });
            }

            if (lengthParsed > 1) {
                requests.push({
                    copyPaste: {
                        source: {
                            sheetId: targetSheetId,
                            startRowIndex: startRow - 1,
                            endRowIndex: startRow,
                            startColumnIndex: 0,
                            endColumnIndex: maxColumns
                        },
                        destination: {
                            sheetId: targetSheetId,
                            startRowIndex: startRow,
                            endRowIndex: startRow + lengthParsed - 1,
                            startColumnIndex: 0,
                            endColumnIndex: maxColumns
                        },
                        pasteType: "PASTE_NORMAL",
                        pasteOrientation: "NORMAL"
                    }
                });
            }

            const outputValues = entry.individuals.map(name => [name]);
            valueUpdates.push({
                range: `'Period ${entry.identifier}'!A${startRow}:A${maxRowsNeeded}`,
                values: outputValues
            });

            if (requests.length > 0 && Sheets && Sheets.Spreadsheets) {
                Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
            }

            // Send matrix values configurations
            if (valueUpdates.length > 0 && Sheets && Sheets.Spreadsheets && Sheets.Spreadsheets.Values) {
                Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: 'USER_ENTERED', data: valueUpdates }, spreadsheetId);
            }

            return true;
        } catch (err: any) {
            const startRow = targetSheet.getRange(cellToStartFrom).getRow();
            let lengthParsed = entry.individuals.length;

            targetSheet.getRange(cellToSetName).setValue("Period " + entry.identifier);

            if (lengthParsed === 0) {
                Logger.log("Length parsed is 0 for identifier: " + entry.identifier + ". Clearing existing data and skipping filling.");
                return true;
            }

            // Clear existing content from row A7 down.
            const lastRowWithData = targetSheet.getLastRow();
            if (lastRowWithData >= startRow) {
                targetSheet.getRange(startRow, 1, lastRowWithData - startRow + 1, 1).clearContent();
            }

            const currentMaxRows = targetSheet.getMaxRows();
            if (startRow + lengthParsed > currentMaxRows) {
                const rowsToAdd = (startRow + lengthParsed) - currentMaxRows;
                targetSheet.insertRowsAfter(startRow, rowsToAdd);
            }

            // Set header name safely
            if (lengthParsed > 1) {
                const sourceRange = targetSheet.getRange(startRow, 1, 1, targetSheet.getMaxColumns());
                const destinationRange = targetSheet.getRange(startRow + 1, 1, lengthParsed - 1, targetSheet.getMaxColumns());
                sourceRange.copyTo(destinationRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
            }

            // Fill data cleanly in a single network call
            const outputValues = entry.individuals.map(name => [name]);
            targetSheet.getRange(startRow, 1, lengthParsed, 1).setValues(outputValues);

            // Dynamically removing excess rows after filling data to keep sheets tidy
            const updatedMaxRows = targetSheet.getMaxRows();
            const lastActiveRow = targetSheet.getLastRow();

            if (updatedMaxRows > lastActiveRow) {
                const rowsToDelete = updatedMaxRows - lastActiveRow;
                targetSheet.deleteRows(lastActiveRow + 1, rowsToDelete);
            }
        }
        return true;
    } catch (err: any) {
        Logger.log(`An error occurred in function fillSheetsWithData: ${err as string}`);
        console.error(`An error occurred in function fillSheetsWithData: ${err as string}`);
        return false;
    }
}