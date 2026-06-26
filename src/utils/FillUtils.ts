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

        let targetSheet = sheets.find(s => s.getName() === "Period " + entry.identifier);

        let isFreshlyCreated = false;
        let forceUnideFreshSheet = false;

        // If the sheet already exists, delete it first to force a clean recreation from scratch
        if (targetSheet) {
            spreadsheet.deleteSheet(targetSheet);
            targetSheet = undefined;
            sheets = spreadsheet.getSheets();
        }

        if (!targetSheet) {
            const templateSheet = sheets.find(s => s.getName() === "Template");
            if (templateSheet) {
                targetSheet = templateSheet.copyTo(spreadsheet);
                targetSheet.setName("Period " + entry.identifier);
                sheets = spreadsheet.getSheets();
                isFreshlyCreated = true;
                forceUnideFreshSheet = true;
            } else {
                Logger.log("Template sheet not found. Skipping period: " + entry.identifier);
                return false;
            }
        }

        // Attempt to add with the Google Sheets API first, only if it fails run the fallback
        try {
            const startRow = targetSheet.getRange(STARTING_CELL).getRow(); // Row 7
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

            // Clear existing data rows starting from row 7 down to prevent old data or formatting ghost artifacts
            if (!isFreshlyCreated && lastRowWithData >= startRow) {
                targetSheet.getRange(startRow, 1, (lastRowWithData - startRow) + 1, maxColumns).clearContent();
            }

            if (forceUnideFreshSheet) {
                requests.push({
                    updateSheetProperties: {
                        properties: {
                            sheetId: targetSheetId,
                            hidden: false
                        },
                        fields: "hidden"
                    }
                });
            }

            valueUpdates.push({
                range: `${targetSheet.getName()}!${SHEET_CELL_NAME}`,
                values: [[`Period ${entry.identifier}`]]
            });

            if (lengthParsed === 0) {
                if (Sheets && Sheets.Spreadsheets && Sheets.Spreadsheets.Values) {
                    Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
                    Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: 'RAW', data: valueUpdates }, spreadsheetId);
                }
                return true;
            }

            // Ensure the sheet has enough rows structurally allocated before applying ranges
            const maxRowsNeeded = startRow + lengthParsed - 1;
            if (maxRowsNeeded > currentMaxRows) {
                targetSheet.insertRowsAfter(currentMaxRows, maxRowsNeeded - currentMaxRows);
            }

            // Native formatting and formula replication pipeline (Bypasses Advanced REST structural issues)
            if (lengthParsed > 1) {
                const sourceTemplateRange = targetSheet.getRange(startRow, 1, 1, maxColumns); // Row 7 master template
                const destinationTargetRange = targetSheet.getRange(startRow + 1, 1, lengthParsed - 1, maxColumns); // Target Rows 8 and below
                
                // Copy formulas and styling layouts securely over the newly assigned rows block
                // Copies formatting, formulas, AND static cell values from Row 7
                sourceTemplateRange.copyTo(destinationTargetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
            }

            // Set the raw names into Column A
            const outputValues = entry.individuals.map(name => [name]);
            targetSheet.getRange(startRow, 1, lengthParsed, 1).setValues(outputValues);

            // Execute metadata updates via the advanced sheet endpoints
            if (requests.length > 0 && Sheets && Sheets.Spreadsheets) {
                Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
            }

            if (valueUpdates.length > 0 && Sheets && Sheets.Spreadsheets && Sheets.Spreadsheets.Values) {
                Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: 'USER_ENTERED', data: valueUpdates }, spreadsheetId);
            }

            // Dynamically removing excess rows after filling data to keep sheets tidy
            SpreadsheetApp.flush();
            const updatedMaxRows = targetSheet.getMaxRows();
            const lastActiveRow = targetSheet.getLastRow();
            if (updatedMaxRows > lastActiveRow && lastActiveRow >= startRow) {
                targetSheet.deleteRows(lastActiveRow + 1, updatedMaxRows - lastActiveRow);
            }

            // Forces an explicit structural recalculation chain on the workbook to re-render formulas and currency formatting
            SpreadsheetApp.flush();
            return true;
        } catch (err: any) {
            Logger.log(`Error occurred in fillSheetWithEntry: ${err.message}`);
            const startRow = targetSheet.getRange(STARTING_CELL).getRow(); // Row 7
            let lengthParsed = entry.individuals.length;

            targetSheet.getRange(SHEET_CELL_NAME).setValue("Period " + entry.identifier);

            targetSheet.showSheet();

            if (lengthParsed === 0) {
                Logger.log("Length parsed is 0 for identifier: " + entry.identifier + ". Clearing existing data and skipping filling.");
                return true;
            }

            // Clear existing content from row A7 down.
            const lastRowWithData = targetSheet.getLastRow();
            if (lastRowWithData >= startRow) {
                targetSheet.getRange(startRow, 1, lastRowWithData - startRow + 1, targetSheet.getMaxColumns()).clearContent();
            }

            const currentMaxRows = targetSheet.getMaxRows();
            if (startRow + lengthParsed > currentMaxRows) {
                const rowsToAdd = (startRow + lengthParsed) - currentMaxRows;
                targetSheet.insertRowsAfter(currentMaxRows, rowsToAdd);
            }

            // Set header name safely
            if (lengthParsed > 1) {
                const sourceRange = targetSheet.getRange(startRow, 1, 1, targetSheet.getMaxColumns()); // Row 7
                const destinationRange = targetSheet.getRange(startRow + 1, 1, lengthParsed - 1, targetSheet.getMaxColumns()); // Row 8 down
                // Use PASTE_NORMAL to accurately clone static text blocks alongside validations
                sourceRange.copyTo(destinationRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
            }

            // Fill data cleanly in a single network call
            const outputValues = entry.individuals.map(name => [name]);
            targetSheet.getRange(startRow, 1, lengthParsed, 1).setValues(outputValues);

            // Dynamically removing excess rows after filling data to keep sheets tidy
            SpreadsheetApp.flush();
            const updatedMaxRows = targetSheet.getMaxRows();
            const lastActiveRow = targetSheet.getLastRow();

            if (updatedMaxRows > lastActiveRow && lastActiveRow >= startRow) {
                const rowsToDelete = updatedMaxRows - lastActiveRow;
                targetSheet.deleteRows(lastActiveRow + 1, rowsToDelete);
            }
            
            SpreadsheetApp.flush();
        }
        return true;
    } catch (err: any) {
        Logger.log(`An error occurred in function fillSheetsWithData: ${err as string}`);
        console.error(`An error occurred in function fillSheetsWithData: ${err as string}`);
        return false;
    }
}