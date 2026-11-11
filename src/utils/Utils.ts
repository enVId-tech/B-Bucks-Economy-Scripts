/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran. 
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 * 
 * IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code.
 * 
 * This file contains utility functions for various operations in the B-Bucks Economy Scripts project.
 */

interface CSVData {
    identifier: number;
    /** Student names, sorted. */
    individuals: string[];
    /** Student IDs in the same order as `individuals` ("" when the roster has none). Older imports do not send this. */
    studentIds?: string[];
}

/** What importing one period did, shown to the teacher when the import finishes. */
interface FillSheetResult {
    ok: boolean;
    period: number;
    /** created: new tab. updated: kept the tab and added what was missing. rebuilt: the old tab barely matched, so it was replaced. skipped: nothing to import. */
    mode: "created" | "updated" | "rebuilt" | "skipped";
    added: number;
    kept: number;
    idsFilled: number;
    /** Names that were not added because a student with that name but a different ID already exists on the tab. */
    skippedNames: string[];
    message: string;
}

/** An existing student on a period tab. */
interface RosterRow {
    row: number;
    name: string;
    id: string;
}

// A period tab that matches at least this share of the import (or of its own students) is updated in place instead of rebuilt.
const ROSTER_MERGE_MATCH_SHARE: number = 0.5;

// Sheet is "Reports -> Gb -> Rb Rstr"

/** "Doe,  Jane" and "doe, jane" are the same name. */
function normalizeRosterName_(name: string): string {
    return String(name ?? "").toLowerCase().replace(/\s+/g, " ").trim();
}

/** Student IDs are kept as text so leading zeros survive. */
function normalizeStudentId_(rawValue: any): string {
    return String(rawValue ?? "").trim();
}

/** The IDs of an import, one per name ("" where the roster had none). */
function alignedStudentIds_(entry: CSVData): string[] {
    return entry.individuals.map((_, index) => normalizeStudentId_(entry.studentIds?.[index]));
}

/** Students currently on a period tab, with the ID stored in the hidden ID column. */
function readExistingRoster_(sheet: GoogleAppsScript.Spreadsheet.Sheet, startRow: number): RosterRow[] {
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) return [];

    const rowCount = lastRow - startRow + 1;
    const names = sheet.getRange(startRow, getNamesColumn(), rowCount, 1).getValues();
    const idColumn = getStudentIdColumn();
    const ids = sheet.getMaxColumns() >= idColumn ? sheet.getRange(startRow, idColumn, rowCount, 1).getValues() : [];

    const roster: RosterRow[] = [];
    names.forEach((cells, index) => {
        const name = String(cells[0] ?? "").trim();
        if (name) roster.push({ row: startRow + index, name, id: normalizeStudentId_(ids[index]?.[0]) });
    });
    return roster;
}

/** Something other than a student ID already sits in the ID column of these rows. Returns it, or null when the column is free. */
function findStudentIdColumnConflict_(sheet: GoogleAppsScript.Spreadsheet.Sheet, startRow: number): string | null {
    const idColumn = getStudentIdColumn();
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow || sheet.getMaxColumns() < idColumn) return null;

    const cells = sheet.getRange(startRow, idColumn, lastRow - startRow + 1, 1).getValues();
    for (const cell of cells) {
        const value = normalizeStudentId_(cell[0]);
        if (value && !/^\d{4,}$/.test(value)) return value;
    }
    return null;
}

/** Makes sure the ID column exists, writes the IDs next to the names and hides the column so only editors can see it. */
function writeStudentIds_(sheet: GoogleAppsScript.Spreadsheet.Sheet, firstRow: number, ids: string[]): void {
    if (ids.length === 0 || !ids.some(id => id)) return;
    invalidatePortalRosterCache_(); // the student IDs the portal matches emails against are changing

    const idColumn = getStudentIdColumn();
    if (sheet.getMaxColumns() < idColumn) sheet.insertColumnsAfter(sheet.getMaxColumns(), idColumn - sheet.getMaxColumns());

    const idRange = sheet.getRange(firstRow, idColumn, ids.length, 1);
    idRange.setNumberFormat("@");
    idRange.setValues(ids.map(id => [id]));
    if (!sheet.isColumnHiddenByUser(idColumn)) sheet.hideColumns(idColumn);
}

interface RosterMergePlan {
    shouldMerge: boolean;
    existing: RosterRow[];
    /** Existing rows that an imported student matched, with the ID the import knows for them. */
    matched: { rosterRow: RosterRow; importedId: string }[];
    /** Positions in the import (names/IDs arrays) that have no match and must be added. */
    toAdd: number[];
    skippedNames: string[];
}

/**
 * Decides whether an existing period tab is "the same class" as the import and which students are missing.
 * A student matches by ID when both sides have one, otherwise by name. Two different IDs with the same name are two people:
 * the newcomer is not added (the code keys everything by name) and is reported instead.
 */
function planRosterMerge_(existing: RosterRow[], names: string[], ids: string[]): RosterMergePlan {
    const rowsById = new Map<string, RosterRow>();
    const rowsByName = new Map<string, RosterRow>();
    existing.forEach(rosterRow => {
        if (rosterRow.id && !rowsById.has(rosterRow.id)) rowsById.set(rosterRow.id, rosterRow);
        const nameKey = normalizeRosterName_(rosterRow.name);
        if (!rowsByName.has(nameKey)) rowsByName.set(nameKey, rosterRow);
    });

    const usedRows = new Set<number>();
    const plan: RosterMergePlan = { shouldMerge: false, existing, matched: [], toAdd: [], skippedNames: [] };

    names.forEach((name, index) => {
        const importedId = ids[index];
        let hit = importedId ? rowsById.get(importedId) : undefined;

        if (!hit) {
            const sameName = rowsByName.get(normalizeRosterName_(name));
            if (sameName && sameName.id && importedId && sameName.id !== importedId) {
                plan.skippedNames.push(name);
                return;
            }
            hit = sameName;
        }

        if (!hit) {
            plan.toAdd.push(index);
        } else if (!usedRows.has(hit.row)) {
            usedRows.add(hit.row);
            plan.matched.push({ rosterRow: hit, importedId });
        }
    });

    const matchedCount = plan.matched.length;
    plan.shouldMerge = matchedCount > 0 && (matchedCount / names.length >= ROSTER_MERGE_MATCH_SHARE || matchedCount / existing.length >= ROSTER_MERGE_MATCH_SHARE);
    return plan;
}

/**
 * Updates a period tab in place: students who are already there keep their row and all their data (only a missing ID is filled in),
 * and students who are in the import but not on the tab are added below the last student. Nobody is removed or renamed.
 */
function mergeStudentsIntoSheet_(sheet: GoogleAppsScript.Spreadsheet.Sheet, templateSheet: GoogleAppsScript.Spreadsheet.Sheet | undefined, entry: CSVData, ids: string[], plan: RosterMergePlan, startRow: number): FillSheetResult {
    const result: FillSheetResult = { ok: true, period: entry.identifier, mode: "updated", added: 0, kept: plan.matched.length, idsFilled: 0, skippedNames: plan.skippedNames, message: "" };

    if (plan.toAdd.length > 0 && !templateSheet) {
        log("Template sheet not found. Period " + entry.identifier + " was left unchanged because new students need the Template row.", true);
        return { ...result, ok: false, mode: "skipped", kept: 0, message: `Period ${entry.identifier}: the Template tab is missing, nothing was changed.` };
    }

    // 1. Fill in IDs for students who were imported before IDs existed. Only blanks are written.
    const missingIds = plan.matched.filter(match => match.importedId && !match.rosterRow.id);
    if (missingIds.length > 0) {
        const idColumn = getStudentIdColumn();
        if (sheet.getMaxColumns() < idColumn) sheet.insertColumnsAfter(sheet.getMaxColumns(), idColumn - sheet.getMaxColumns());
        const lastExistingRow = plan.existing[plan.existing.length - 1].row;
        const span = sheet.getRange(startRow, idColumn, lastExistingRow - startRow + 1, 1);
        const currentValues = span.getValues();
        missingIds.forEach(match => { currentValues[match.rosterRow.row - startRow][0] = match.importedId; });
        span.setNumberFormat("@");
        span.setValues(currentValues.map(row => [row[0] === "" ? "" : normalizeStudentId_(row[0])]));
        result.idsFilled = missingIds.length;
    }

    // 2. Add the students who are missing, using the Template row so formulas, formats and starting values match everyone else.
    if (plan.toAdd.length > 0 && templateSheet) {
        const firstNewRow = (plan.existing.length > 0 ? plan.existing[plan.existing.length - 1].row : startRow - 1) + 1;
        const addCount = plan.toAdd.length;
        const lastNeededRow = firstNewRow + addCount - 1;
        if (lastNeededRow > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), lastNeededRow - sheet.getMaxRows());

        const columnCount = Math.min(templateSheet.getMaxColumns(), sheet.getMaxColumns());
        templateSheet.getRange(startRow, 1, 1, columnCount).copyTo(sheet.getRange(firstNewRow, 1, addCount, columnCount), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

        sheet.getRange(firstNewRow, getNamesColumn(), addCount, 1).setValues(plan.toAdd.map(index => [entry.individuals[index]]));
        writeStudentIds_(sheet, firstNewRow, plan.toAdd.map(index => ids[index]));
        result.added = addCount;
    }

    // 3. The column has to stay hidden even when nothing new was written to it.
    const idColumn = getStudentIdColumn();
    if (sheet.getMaxColumns() >= idColumn && ids.some(id => id) && !sheet.isColumnHiddenByUser(idColumn)) sheet.hideColumns(idColumn);

    SpreadsheetApp.flush();
    result.message = `Period ${entry.identifier}: kept ${result.kept}, added ${result.added}` + (result.idsFilled > 0 ? `, filled ${result.idsFilled} ID(s)` : "") + (result.skippedNames.length > 0 ? `, skipped ${result.skippedNames.length} (same name, different ID)` : "") + ".";
    return result;
}

/**
 * Imports one period. If the period's tab already exists and mostly matches the import (same students, by ID or name), the tab is
 * updated in place so nobody loses data. If it exists with students but does not match, it is left untouched and reported.
 * Only a missing (or empty) tab is built fresh from the Template tab.
 * @param data JSON string of CSVData: { identifier, individuals, studentIds? }.
 * @returns {FillSheetResult} What happened to the period.
 */
function fillSheetsWithData(data: string): FillSheetResult {
    const failedResult = (period: number, message: string): FillSheetResult => ({ ok: false, period, mode: "skipped", added: 0, kept: 0, idsFilled: 0, skippedNames: [], message });
    let period = NaN;
    try {
        const entry: CSVData = JSON.parse(data);
        period = entry.identifier;
        const ids = alignedStudentIds_(entry);
        const sheetName = "Period " + entry.identifier;
        const existingSheet = getSheetNamed(sheetName);
        const templateSheet = getSheetNamed("Template") ?? undefined;
        const startRow = (existingSheet ?? templateSheet)?.getRange(getFillStartingCell()).getRow() ?? getUserStartingRow();

        if (existingSheet) {
            if (entry.individuals.length === 0) {
                return { ...failedResult(period, `Period ${period}: the file listed no students, the tab was left alone.`), ok: true };
            }

            const conflict = findStudentIdColumnConflict_(existingSheet, startRow);
            const existingRoster = readExistingRoster_(existingSheet, startRow);
            const plan = planRosterMerge_(existingRoster, entry.individuals, ids);

            if (plan.shouldMerge) {
                if (conflict) {
                    log(`Column ${getStudentIdColumn()} on ${sheetName} already holds "${conflict}", so student IDs can't go there. Change "studentIdCol" in the Advanced Technical Settings to a free column.`, true);
                    return failedResult(period, `Period ${period}: column ${getStudentIdColumn()} is already in use, nothing was changed.`);
                }
                return mergeStudentsIntoSheet_(existingSheet, templateSheet, entry, ids, plan, startRow);
            }

            // A tab that already has students is never deleted and rebuilt here: that would wipe balances, expenditures and CoDs.
            if (existingRoster.length > 0) {
                log(`Period ${period}: the file matches only ${plan.matched.length} of the ${existingRoster.length} student(s) on the tab, so the tab was left untouched.`, false);
                return failedResult(period, `Period ${period}: the file does not match this tab (${plan.matched.length} of ${existingRoster.length} students found), so NOTHING was changed. If it really is a different class, rename or delete the old tab yourself and import again.`);
            }
        } else if (templateSheet && ids.some(id => id)) {
            const templateConflict = findStudentIdColumnConflict_(templateSheet, startRow);
            if (templateConflict) {
                log(`Column ${getStudentIdColumn()} on the Template tab already holds "${templateConflict}", so student IDs can't go there. Change "studentIdCol" in the Advanced Technical Settings to a free column.`, true);
                return failedResult(period, `Period ${period}: column ${getStudentIdColumn()} is already in use on the Template tab, nothing was changed.`);
            }
        }

        const replacedExisting = existingSheet !== null;
        const rebuilt = rebuildPeriodSheet_(entry, ids);
        return {
            ok: rebuilt, period, mode: rebuilt ? (replacedExisting ? "rebuilt" : "created") : "skipped",
            added: rebuilt ? entry.individuals.length : 0, kept: 0, idsFilled: 0, skippedNames: [],
            message: rebuilt ? `Period ${period}: ${replacedExisting ? "rebuilt (it did not match the file)" : "created"} with ${entry.individuals.length} student(s).` : `Period ${period}: could not be imported.`
        };
    } catch (err: any) {
        log(`An error occurred in function fillSheetsWithData: ${err.message ?? err}`, false);
        return failedResult(period, `Period ${period}: ${err.message ?? err}`);
    }
}

/**
 * Builds a period tab from scratch from the Template tab (deleting a tab with the same name first).
 * Only used for new periods and for tabs that do not match the import at all.
 * @returns {boolean} Returns true if the operation was successful, false otherwise.
 */
function rebuildPeriodSheet_(entry: CSVData, studentIds: string[]): boolean {
    try {
        const spreadsheet = getSS();
        const spreadsheetId = spreadsheet.getId();
        let sheets = spreadsheet.getSheets();

        let targetSheet = sheets.find(s => s.getName() === "Period " + entry.identifier);

        let isFreshlyCreated = false;
        let forceUnideFreshSheet = false;

        // If the sheet already exists, delete it first to force a clean recreation from scratch
        if (targetSheet) {
            spreadsheet.deleteSheet(targetSheet);
            resetSheetMemo();
            targetSheet = undefined;
            sheets = spreadsheet.getSheets();
        }

        if (!targetSheet) {
            const templateSheet = sheets.find(s => s.getName() === "Template");
            if (templateSheet) {
                targetSheet = templateSheet.copyTo(spreadsheet);
                targetSheet.setName("Period " + entry.identifier);
                resetSheetMemo();
                sheets = spreadsheet.getSheets();
                isFreshlyCreated = true;
                forceUnideFreshSheet = true;
            } else {
                log("Template sheet not found. Skipping period: " + entry.identifier, true);
                return false;
            }
        }

        // Attempt to add with the Google Sheets API first, only if it fails run the fallback
        try {
            const startRow = targetSheet.getRange(getFillStartingCell()).getRow(); // Row 7
            const lengthParsed = entry.individuals.length;
            if (lengthParsed === 0) {
                log("Length parsed is 0 for identifier: " + entry.identifier + ". Clearing existing data and skipping filling.", false);
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
                range: `${targetSheet.getName()}!${getFillSheetCellName()}`,
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
            writeStudentIds_(targetSheet, startRow, studentIds);

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
            log(`Error occurred in fillSheetWithEntry: ${err.message}`, true);
            const startRow = targetSheet.getRange(getFillStartingCell()).getRow(); // Row 7
            let lengthParsed = entry.individuals.length;

            targetSheet.getRange(getFillSheetCellName()).setValue("Period " + entry.identifier);

            targetSheet.showSheet();

            if (lengthParsed === 0) {
                log("Length parsed is 0 for identifier: " + entry.identifier + ". Clearing existing data and skipping filling.", false);
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
            writeStudentIds_(targetSheet, startRow, studentIds);

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

        log(`An error occurred in function fillSheetsWithData: ${err as string}`, false);
        return false;
    }
}

/**
 * Caller information utility to retrieve the function name and file name of the caller in the stack trace.
 * @returns {object} An object containing the function name and file name of the caller.
 * If the caller information cannot be determined, it returns 'unknown' for both function name and file name.
 */
function getCallerInfo() {
    const stack = new Error().stack?.split('\n') ?? [];
    const callerLine = stack[2] ?? '';

    const match =
        callerLine.match(/at\s+(.*?)\s+\((.*?):(\d+):(\d+)\)/) ??
        callerLine.match(/at\s+(.*?):(\d+):(\d+)/);

    if (!match) return { functionName: 'unknown', fileName: 'unknown' };

    if (match.length === 5) {
        return {
            functionName: match[1] || 'unknown',
            fileName: match[2].split('/').pop() ?? match[2],
        };
    }

    return {
        functionName: 'anonymous',
        fileName: match[1].split('/').pop() ?? match[1],
    };
}

/**
 * Logger for the entire B-Bucks Economy project
 * @param {any} message The message to log, can be of type string, number, boolean, or an array of any type.
 * @param {boolean} isVisibleToClient Whether to show the log message to the client via a toast notification. Default is true.
 * @returns {boolean} Returns true if the logging operation was successful, false otherwise.
 */
function log(message: any[] | string | number | boolean, isVisibleToClient: boolean = true): boolean {
    try {
        const timestamp = new Date().toISOString();
        const formattedMessage = `[${timestamp}] ${message}`;
        const callerInfo = getCallerInfo();

        // Will always log to console for developer purposes
        // Will also always log to Apps Script Logger for debugging
        if (!getSS()) {
            console.error("No active spreadsheet found. Logging aborted (likely an error unless this is intended).");
            Logger.log("No active spreadsheet found. Logging aborted (likely an error unless this is intended).");
            return false;
        }

        console.error(`${formattedMessage}`);
        Logger.log(`${formattedMessage}`);

        if (isVisibleToClient) {
            if (portalCapturedMessages) {
                // A portal request has no spreadsheet window to alert, so its messages go back to the student instead
                portalCapturedMessages.push(String(message));
            } else {
                SpreadsheetApp.getUi().alert(`${formattedMessage}`);
            }
        }

        return true;
    } catch (err: any) {
        Logger.log(`Error formatting log message: ${err.message}`);
        getSS().toast(`Error formatting log message: ${err.message}`, "Logging Error", 5);
        return false;
    }
}

/**
 * Simple trigger that executes automatically when a user alters a cell in the spreadsheet.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The edit event object.
 */
function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
    // Guard against running manually from the Apps Script IDE without an event object
    if (!e || !e.range) return;

    // Pass the edited sheet directly to your stamper function
    const editedSheet = e.range.getSheet();
    stampLastModified(editedSheet);
}

/**
 * Stamps the last modified timestamp and user who made the modification on the specified sheet.
 * @param targetSheet The Google Sheets sheet object where the timestamp should be stamped.
 */
function stampLastModified(targetSheet?: GoogleAppsScript.Spreadsheet.Sheet): void {
    try {
        const sheet = targetSheet || getSS().getActiveSheet();

        if (!sheet) {
            log("No active sheet found to stamp last modified.", true);
            return;
        }

        if (TIMESTAMP_LIST.includes(sheet.getName())) {
            // Avoid infinite trigger loops: don't re-stamp if the edit happened on the timestamp cell itself
            if (targetSheet) {
                const activeCell = targetSheet.getActiveCell();
                if (activeCell && activeCell.getA1Notation() === getTimestampCell()) {
                    return;
                }
            }

            const email = Session.getActiveUser().getEmail();
            const userIdentifier = email ? email : "Authorized User";
            
            sheet.getRange(getTimestampCell()).setValue(`Last Modified by: ${userIdentifier} on ${new Date().toLocaleString()}`);
        }
    } catch (err: any) {
        log(`Error in stampLastModified: ${err.message}`, true);
    }
}