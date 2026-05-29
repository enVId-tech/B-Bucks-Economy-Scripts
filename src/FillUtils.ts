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
 * Parses CSV data into an array of objects.
 * @param csvData The raw CSV data as a string, where each row represents an identifier and its associated individuals. The identifier can be separated from the individuals by either a hyphen or a comma.
 * @returns {CSVData[]} An array of objects, each containing an identifier and an array of associated individuals.
 */
function parseCsvData(csvData: string): CSVData[] {
    try {
        if (!csvData) return [];

        // Split once, but don't mutate the array elements in place
        const rows = csvData.split('\n');
        const data: CSVData[] = [];
        // Initialize the first block container
        let currentSheet: CSVData | null = null;

        const headerStrContains = " - "

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i].trim();
            if (!row) continue;

            // Check for Identifier Row (e.g., "1 - Header Group")
            if (row.includes(headerStrContains)) {
                let cleanRow = row;
                if (row.includes(']')) {
                    cleanRow = row.split(']')[1].trim();
                }

                const identifierPart = row.split(headerStrContains)[0].trim();
                const identifier = parseInt(identifierPart, 10);
                if (!isNaN(identifier)) {
                    if (currentSheet) {
                        data.push(currentSheet);
                    }

                    currentSheet = {
                        identifier: identifier,
                        individuals: []
                    };
                    continue;
                }
            }

            if (!row.includes('"')) {
                continue;
            }

            if (currentSheet) {
                const partsByQuotes = row.split('"');
                if (partsByQuotes.length >= 3) {
                    const name = partsByQuotes[1].trim();
                    currentSheet.individuals.push(name);
                }
            }
        }

        // Push the final sheet if it exists
        if (currentSheet) {
            data.push(currentSheet);
        }

        // Sort the individuals inside each group
        // Time complexity here depends on the number of individuals per group: O(K log K)
        for (const entry of data) {
            entry.individuals.sort((a, b) => a.localeCompare(b));
        }

        return data;
    } catch (err: any) {
        Logger.log(`An error occured in parseCsvData: ${err as string}`);
        console.error(`An error occurred in function fillSheetsWithData: ${err as string}`);
        return [];
    }
}

/**
 * Fills the sheets with data parsed from the CSV string. Each sheet is named after the identifier and contains the associated individuals starting from a specified cell. If a sheet for an identifier doesn't exist, it creates one by copying a template sheet.
 * @param data  The raw CSV data as a string, where each row represents an identifier and its associated individuals. The identifier can be separated from the individuals by either a hyphen or a comma.
 * @returns {string}
 */
function fillSheetsWithData(data: string): string {
    try {
        const parsedData: CSVData[] = parseCsvData(data);
        const spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        let sheets: GoogleAppsScript.Spreadsheet.Sheet[] = spreadsheet.getSheets();

        const cellToStartFrom = "A7"; // Starting cell for filling data
        const cellToSetName = "A2"; // Cell to set the identifier name

        for (const entry of parsedData) {
            let targetSheet = sheets.find(s => s.getName() === "Period " + entry.identifier);

            if (!targetSheet) {
                const templateSheet = sheets.find(s => s.getName() === "Template");
                if (templateSheet) {
                    targetSheet = templateSheet.copyTo(spreadsheet);
                    targetSheet.setName("Period " + entry.identifier);
                    sheets = spreadsheet.getSheets();
                } else {
                    Logger.log("Template sheet not found. Skipping period: " + entry.identifier);
                    continue;
                }
            }

            const startRow = targetSheet.getRange(cellToStartFrom).getRow();
            let lengthParsed = entry.individuals.length;

            targetSheet.getRange(cellToSetName).setValue("Period " + entry.identifier);

            if (lengthParsed === 0) continue;

            // Clear existing content from row A7 down.
            const lastRowWithData = targetSheet.getLastRow();
            if (lastRowWithData >= startRow) {
                targetSheet.getRange(startRow, 1, lastRowWithData - startRow + 1, 1).clearContent();
            }

            const maxRows = targetSheet.getMaxRows();
            if (startRow + lengthParsed > maxRows) {
                const rowsToAdd = (startRow + lengthParsed) - maxRows;
                targetSheet.insertRowsAfter(maxRows, rowsToAdd);
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

            const lastRow = targetSheet.getLastRow();

            if (maxRows - lastRow != 0) targetSheet.deleteRows(lastRow + 1, maxRows - lastRow);
        }

        return JSON.stringify(parsedData, null, 2);
    } catch (err: any) {
        Logger.log(`An error occurred in function fillSheetsWithData: ${err as string}`);
        console.error(`An error occurred in function fillSheetsWithData: ${err as string}`);
        return "";
    }
}