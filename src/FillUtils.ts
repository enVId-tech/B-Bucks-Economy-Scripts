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
    // Split once, but don't mutate the array elements in place
    const rows = csvData.split('\n');
    const data: CSVData[] = [];

    // Initialize the first block container
    let currentSheet: CSVData | null = null;

    // Compiled regexes outside or pre-compiled per row execution
    // Grouping patterns allows us to check AND extract data in one single operation
    const identifierRegex = /^(\d+)\s*-\s*/;
    const individualRegex = /^\d+\s*,\s*([^,]+)/; // Captures the name after the comma

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i].trim();
        if (!row) continue;

        // Check for Identifier Row (e.g., "1 - Header Group")
        const idMatch = row.match(identifierRegex);
        if (idMatch) {
            if (currentSheet) {
                data.push(currentSheet);
            }

            currentSheet = {
                identifier: parseInt(idMatch[1], 10),
                individuals: []
            };
            continue;
        }

        // Check for Individual Row (e.g., "1, John Doe")
        const individualMatch = row.match(individualRegex);
        if (individualMatch && currentSheet) {
            currentSheet.individuals.push(individualMatch[1].trim());
            continue;
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
}

function fillSheetsWithData(data: string): void {
    let parsedData: CSVData[] = parseCsvData(data);
    let sheet: GoogleAppsScript.Spreadsheet.Sheet[] = SpreadsheetApp.getActiveSpreadsheet().getSheets();

    const cellToStartFrom = "A7"; // Starting cell for filling data
    const cellToSetName = "A2"; // Cell to set the identifier name

    for (const entry of parsedData) {
        let targetSheet = sheet.find(s => s.getName().includes(entry.identifier.toString()));
        
        if (!targetSheet) {
            const findSheet = sheet.find(s => s.getName() === "Template");
            if (findSheet) {
                targetSheet = findSheet.copyTo(SpreadsheetApp.getActiveSpreadsheet());
                targetSheet.setName("Period " + entry.identifier);
                targetSheet.getRange(cellToSetName).setValue("Period " + entry.identifier);
            } else {
                Logger.log("Template sheet not found. Cannot create new sheet for identifier: " + entry.identifier);
                continue; // Skip to the next entry if template is missing
            }
        }

        // Clear existing content before filling new data
        targetSheet.clearContents();
        let lengthParsed = entry.individuals.length;
        // Make sure the maximum number of rows available in the sheet isn't exceeded, but extended, starting from the starting cell
        let rowCount = targetSheet.getMaxRows() - targetSheet.getRange(cellToStartFrom).getRow() + 1;
        if (lengthParsed < rowCount) {
            // Copy the starting cell row to the end of the data range to ensure formatting is preserved
            targetSheet.getRange(cellToStartFrom + ":" + cellToStartFrom.replace(/\d+/, (match) => (parseInt(match, 10) + lengthParsed - 1).toString())).copyTo(targetSheet.getRange(cellToStartFrom + ":" + cellToStartFrom.replace(/\d+/, (match) => (parseInt(match, 10) + rowCount - 1).toString())), { formatOnly: true });
        }

        // Fill the sheet with the individuals' names starting from the specified cell
        for (let i = 0; i < lengthParsed && i < rowCount; i++) {
            targetSheet.getRange(cellToStartFrom).offset(i, 0).setValue(entry.individuals[i]);
        }
    }
}