/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran. All rights reserved.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 *
 * Per-execution handles for the spreadsheet and its sheets.
 *
 * Every Apps Script execution starts with a fresh global scope, so remembering the Spreadsheet object and the sheet list
 * for the rest of THIS execution is safe, and it removes dozens of repeated getActiveSpreadsheet() / getSheets() /
 * getSheetByName() round trips from a single menu action.
 *
 * Rule: any code that adds, deletes, renames or copies a sheet must call resetSheetMemo() afterwards.
 */

/** Only touched at call time (never while files load), so file order does not matter. */
const sheetMemo: {
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null;
    sheets: GoogleAppsScript.Spreadsheet.Sheet[] | null;
    byName: { [name: string]: GoogleAppsScript.Spreadsheet.Sheet } | null;
} = { spreadsheet: null, sheets: null, byName: null };

/** The active spreadsheet, fetched once per execution. */
function getSS(): GoogleAppsScript.Spreadsheet.Spreadsheet {
    if (!sheetMemo.spreadsheet) sheetMemo.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    return sheetMemo.spreadsheet;
}

/** Every sheet in the spreadsheet, fetched once per execution (until resetSheetMemo). */
function getAllSheets(): GoogleAppsScript.Spreadsheet.Sheet[] {
    if (!sheetMemo.sheets) {
        sheetMemo.sheets = getSS().getSheets();
        sheetMemo.byName = null;
    }
    return sheetMemo.sheets;
}

/** Same as getSheetByName but answered from the memoized list. Returns null when the sheet does not exist. */
function getSheetNamed(name: string): GoogleAppsScript.Spreadsheet.Sheet | null {
    if (!sheetMemo.byName) {
        const sheetsByName: { [name: string]: GoogleAppsScript.Spreadsheet.Sheet } = {};
        // First sheet with a given name wins, matching getSheetByName
        getAllSheets().forEach(sheet => { const sheetName = sheet.getName(); if (!(sheetName in sheetsByName)) sheetsByName[sheetName] = sheet; });
        sheetMemo.byName = sheetsByName;
    }
    return Object.prototype.hasOwnProperty.call(sheetMemo.byName, name) ? sheetMemo.byName[name] : null;
}

/** Forget the remembered sheets. Call after inserting, deleting, renaming or copying any sheet. */
function resetSheetMemo(): void {
    sheetMemo.sheets = null;
    sheetMemo.byName = null;
}
