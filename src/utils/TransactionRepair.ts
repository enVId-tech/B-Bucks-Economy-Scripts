/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 *
 * TRANSACTION LOG REPAIR (Banking Settings -> Maintenance -> Transaction Log Repair)
 * ==================================================================================
 * Every row of the Transactions sheet stores the cell that was changed (columns J and K: before and after) AND the student's balance
 * before and after (columns L and M). The student portal and Transaction Records show a row's change as M - L. If an older version of
 * the script (or a manual edit) wrote wrong balances, a row can show the wrong amount, or even the wrong sign, for example an income
 * of $1 x 2 that reads "-$28.70".
 *
 * This file finds those rows and rebuilds them, in three steps that the settings page exposes as three buttons:
 *   1. scanTransactionLog()    READ ONLY. Lists the rows whose reported balance change disagrees with what the changed cell says.
 *   2. repairTransactionLog()  Rewrites ONLY column M ("balance after") of the rows the admin picked, to column L plus the expected
 *                              change. Every old value is first copied to the "Transaction Repair Backup" tab.
 *   3. undoTransactionRepair() Puts the values of the LAST repair back from that backup tab.
 * Nothing here ever touches a student's Period sheet, so no real balance changes: only the log's report of it is corrected.
 *
 * HOW A ROW IS JUDGED (judgeTransactionRow_)
 *   expected balance change = +(K - J) when the changed column is Earnings or Balance, and -(K - J) when it is Expenditures.
 *   Rows that changed any other column (investments, CoD) are skipped: their cash movement is not a single cell.
 *   reported change         = M - L.
 *   They must agree within one cent. If they are exactly opposite the problem is "wrongSign", otherwise "wrongChange". Both are
 *   repairable because the expected value comes from the changed cell, which is the more trustworthy number.
 *   The changed cell is only trusted when |K - J| equals price x quantity, or that minus the income tax. If it does not, the row is
 *   listed as "review" and NEVER rewritten: then the changed cell itself looks wrong, and only a human can say what really happened.
 */

/** Name of the tab that keeps the old values of every repaired row, so a repair can be undone. */
const TRANSACTION_REPAIR_BACKUP_SHEET: string = "Transaction Repair Backup";
/** Header of the backup tab. "Batch" groups the rows of one repair run (its timestamp). */
const TRANSACTION_REPAIR_BACKUP_HEADERS: string[] = ["Batch", "Backed Up At", "Sheet Row", "Transaction ID", "Old Balance After", "New Balance After", "Reason", "Restored"];
/** Script Property bumped by every repair or restore. The portal includes it in its history cache key, so repaired rows show at once. */
const TRANSACTION_LOG_EPOCH_PROPERTY: string = "transactionLogEpoch";
/** Rows read per sheet request while scanning (keeps each read small on a very long log). */
const TRANSACTION_REPAIR_CHUNK_ROWS: number = 2000;
/** How many problem rows the scan sends back to the page to preview (the counts always cover every row). */
const TRANSACTION_REPAIR_PREVIEW_LIMIT: number = 100;
/** Two money values closer than this (one cent, with room for float error) count as equal. */
const TRANSACTION_REPAIR_TOLERANCE: number = 0.011;

type TransactionProblemKind = "wrongSign" | "wrongChange" | "review";

/** One row of the log that did not pass judgeTransactionRow_. */
interface TransactionProblem {
    /** The row on the Transactions sheet (1-based). */
    sheetRow: number;
    id: number;
    student: string;
    period: string;
    type: string;
    description: string;
    /** M - L: what the log currently says the balance changed by. */
    reportedChange: number;
    /** What the changed cell (K - J) says it should be, or null for "review" rows. */
    expectedChange: number | null;
    kind: TransactionProblemKind;
    /** Why the row was flagged, in plain words. */
    reason: string;
}

/** The current epoch (a counter in Script Properties). Part of the portal's history cache key. */
function getTransactionLogEpoch_(): string {
    try { return String(PropertiesService.getScriptProperties().getProperty(TRANSACTION_LOG_EPOCH_PROPERTY) || "0"); } catch (error) { return "0"; }
}

/** Bumps the epoch and clears the transaction caches, so the portal and Transaction Records reload the repaired rows. */
function bumpTransactionLogEpoch_(): void {
    try {
        const next = String(Number(getTransactionLogEpoch_()) + 1);
        PropertiesService.getScriptProperties().setProperty(TRANSACTION_LOG_EPOCH_PROPERTY, next);
    } catch (error: any) {
        log(`Could not bump the transaction log epoch: ${error.message}`, false);
    }
    invalidateTransactionCache();
}

/**
 * Judges one log row (the 16 values of columns A..P). Returns the problem, or null when the row is fine or not checkable.
 * @param row the values of one row of the Transactions sheet, A..P
 * @param sheetRow the row number on the sheet (for the report)
 */
function judgeTransactionRow_(row: any[], sheetRow: number): TransactionProblem | null {
    const id = Number(row[0]);
    if (!Number.isFinite(id) || id <= 0) return null; // not a transaction row (blank or a stray note)

    const modifiedColumn = Number(row[7]);
    const unitPrice = Number(row[5]);
    const quantity = Number(row[6]);
    const columnBefore = Number(row[9]);
    const columnAfter = Number(row[10]);
    const balanceBefore = Number(row[11]);
    const balanceAfter = Number(row[12]);
    const operation = String(row[15] || "").toUpperCase();

    const base = { sheetRow, id, student: String(row[1] ?? ""), period: String(row[2] ?? ""), type: String(row[3] ?? ""), description: String(row[4] ?? "") };

    if (![columnBefore, columnAfter, balanceBefore, balanceAfter].every(Number.isFinite)) {
        return { ...base, reportedChange: 0, expectedChange: null, kind: "review", reason: "One of the amounts (columns J to M) is missing or not a number." };
    }

    const reportedChange = Number((balanceAfter - balanceBefore).toFixed(2));
    const columnStep = Number((columnAfter - columnBefore).toFixed(2));

    // Only rows that changed a single money column with + or - can be rebuilt from the changed cell.
    if (operation === "MULTIPLY" || operation === "DIVIDE") return null;
    let expectedChange: number | null = null;
    if (modifiedColumn === getExpendituresColumn()) expectedChange = -columnStep;
    else if (modifiedColumn === getEarningsColumn() || modifiedColumn === getBalanceColumn()) expectedChange = columnStep;
    if (expectedChange === null) return null;
    expectedChange = Number(expectedChange.toFixed(2));

    // Is the changed cell itself believable? Its step must be price x quantity, or that minus the income tax. Only then is it a safe
    // source of truth to rebuild the balance report from; otherwise the row is listed for review and never rewritten.
    const step = Math.abs(columnStep);
    const full = Math.abs(unitPrice * quantity);
    const incomeTax = Number(fetchProperty("incomeTaxRate", "standardPercentages"));
    const taxed = Number.isFinite(incomeTax) ? full * (1 - incomeTax) : full;
    const cellIsBelievable = Number.isFinite(full) && full > 0 && step > 0 && (Math.abs(step - full) <= TRANSACTION_REPAIR_TOLERANCE || Math.abs(step - taxed) <= TRANSACTION_REPAIR_TOLERANCE);
    if (!cellIsBelievable) {
        return { ...base, reportedChange, expectedChange: null, kind: "review", reason: Number.isFinite(full) && full > 0 ? `The changed cell moved by ${step.toFixed(2)}, which is neither price x quantity (${full.toFixed(2)}) nor that after tax.` : "The price or quantity is missing, so the changed cell cannot be checked." };
    }

    if (Math.abs(reportedChange - expectedChange) <= TRANSACTION_REPAIR_TOLERANCE) return null;
    const opposite = Math.abs(reportedChange + expectedChange) <= TRANSACTION_REPAIR_TOLERANCE;
    return {
        ...base, reportedChange, expectedChange, kind: opposite ? "wrongSign" : "wrongChange",
        reason: opposite ? "The log shows the opposite sign to the cell that changed." : `The log shows ${reportedChange.toFixed(2)} but the cell that changed moved by ${expectedChange.toFixed(2)}.`
    };
}

/** Reads the whole log in chunks and returns every problem row plus how many rows were checked. */
function scanTransactionLogInside_(): { problems: TransactionProblem[]; rowsChecked: number } {
    const sheet = getTransactionsSheet();
    if (!sheet) throw new Error(`The "${getTransactionsSheetName()}" tab was not found.`);
    const total = countTransactionRows(sheet);
    const problems: TransactionProblem[] = [];
    for (let offset = 0; offset < total; offset += TRANSACTION_REPAIR_CHUNK_ROWS) {
        const count = Math.min(TRANSACTION_REPAIR_CHUNK_ROWS, total - offset);
        const firstRow = getTransactionsRowStart() + offset;
        sheet.getRange(firstRow, 1, count, TRANSACTION_COLUMN_COUNT).getValues().forEach((row, index) => {
            const problem = judgeTransactionRow_(row, firstRow + index);
            if (problem) problems.push(problem);
        });
    }
    return { problems, rowsChecked: total };
}

/** Counts problems by kind. */
function countTransactionProblems_(problems: TransactionProblem[]): { wrongSign: number; wrongChange: number; review: number } {
    const counts = { wrongSign: 0, wrongChange: 0, review: 0 };
    problems.forEach(problem => { counts[problem.kind]++; });
    return counts;
}

/**
 * STEP 1 (read only). Scans the Transactions sheet and reports the rows that need repair.
 * Returns JSON: { rowsChecked, counts: { wrongSign, wrongChange, review }, repairable, preview: TransactionProblem[] (first 100) }.
 */
function scanTransactionLog(): string {
    requireEditor_();
    try {
        const { problems, rowsChecked } = scanTransactionLogInside_();
        const counts = countTransactionProblems_(problems);
        return JSON.stringify({ ok: true, rowsChecked, counts, repairable: counts.wrongSign + counts.wrongChange, preview: problems.slice(0, TRANSACTION_REPAIR_PREVIEW_LIMIT), previewLimit: TRANSACTION_REPAIR_PREVIEW_LIMIT });
    } catch (error: any) {
        log(`Error in scanTransactionLog: ${error.message}`, false);
        return JSON.stringify({ ok: false, message: error.message });
    }
}

/** The backup tab, created with its header row when it does not exist yet. */
function ensureTransactionRepairBackupSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
    let sheet = getSheetNamed(TRANSACTION_REPAIR_BACKUP_SHEET);
    if (!sheet) {
        sheet = getSS().insertSheet(TRANSACTION_REPAIR_BACKUP_SHEET);
        sheet.getRange(1, 1, 1, TRANSACTION_REPAIR_BACKUP_HEADERS.length).setValues([TRANSACTION_REPAIR_BACKUP_HEADERS]).setFontWeight("bold");
        sheet.setFrozenRows(1);
    }
    return sheet;
}

/**
 * STEP 2. Rebuilds the "balance after" (column M) of the problem rows of the kinds asked for, after backing up their old values.
 * The scan is repeated inside the transaction lock so the rows repaired are exactly the ones that are wrong right now.
 * @param data JSON string: { kinds?: ("wrongSign" | "wrongChange")[] } (default both). "review" rows are never changed.
 * Returns JSON: { ok, message, repaired }.
 */
function repairTransactionLog(data?: string): string {
    requireEditor_();
    return runWithTransactionLock(() => {
        try {
            let kinds: string[] = ["wrongSign", "wrongChange"];
            try { const parsed = data ? JSON.parse(data) : null; if (parsed && Array.isArray(parsed.kinds)) kinds = parsed.kinds.filter((kind: string) => kind === "wrongSign" || kind === "wrongChange"); } catch (e) { /* default */ }
            if (kinds.length === 0) return JSON.stringify({ ok: false, message: "Choose at least one kind of problem to repair.", repaired: 0 });

            const { problems } = scanTransactionLogInside_();
            const toFix = problems.filter(problem => kinds.indexOf(problem.kind) >= 0 && problem.expectedChange !== null);
            if (toFix.length === 0) return JSON.stringify({ ok: true, message: "Nothing to repair: every checked row is consistent.", repaired: 0 });

            const sheet = getTransactionsSheet()!;
            const firstRow = getTransactionsRowStart();
            const lastRow = sheet.getLastRow();
            const balanceBeforeColumn = sheet.getRange(firstRow, 12, lastRow - firstRow + 1, 1).getValues();
            const balanceAfterRange = sheet.getRange(firstRow, 13, lastRow - firstRow + 1, 1);
            const balanceAfterColumn = balanceAfterRange.getValues();

            // 1) Back up the old values first. If this fails nothing has been changed.
            const batch = new Date().toISOString();
            const backupRows: any[][] = [];
            toFix.forEach(problem => {
                const index = problem.sheetRow - firstRow;
                const rebuilt = Number((Number(balanceBeforeColumn[index][0]) + (problem.expectedChange as number)).toFixed(2));
                backupRows.push([batch, batch, problem.sheetRow, problem.id, balanceAfterColumn[index][0], rebuilt, problem.reason, ""]);
                balanceAfterColumn[index][0] = rebuilt;
            });
            const backup = ensureTransactionRepairBackupSheet_();
            backup.getRange(backup.getLastRow() + 1, 1, backupRows.length, TRANSACTION_REPAIR_BACKUP_HEADERS.length).setValues(backupRows);
            SpreadsheetApp.flush();

            // 2) One write for the whole column (the other rows are written back exactly as read).
            balanceAfterRange.setValues(balanceAfterColumn);
            SpreadsheetApp.flush();
            bumpTransactionLogEpoch_();
            log(`Transaction log repair: rebuilt ${toFix.length} row(s). Old values are in "${TRANSACTION_REPAIR_BACKUP_SHEET}".`, false);
            return JSON.stringify({ ok: true, message: `Repaired ${toFix.length} row(s). The old values are saved on the "${TRANSACTION_REPAIR_BACKUP_SHEET}" tab and "Undo last repair" puts them back.`, repaired: toFix.length });
        } catch (error: any) {
            log(`Error in repairTransactionLog: ${error.message}`, false);
            return JSON.stringify({ ok: false, message: `The repair stopped: ${error.message}`, repaired: 0 });
        }
    }, JSON.stringify({ ok: false, message: "Another transaction is still running. Nothing was changed, please try again.", repaired: 0 }));
}

/**
 * STEP 3. Puts back the old "balance after" values of the most recent repair (every backup row of the newest batch not restored yet).
 * A row is only restored when its transaction id still matches, so rows that moved or were deleted since are skipped, not guessed.
 * Returns JSON: { ok, message, restored, skipped }.
 */
function undoTransactionRepair(): string {
    requireEditor_();
    return runWithTransactionLock(() => {
        try {
            const backup = getSheetNamed(TRANSACTION_REPAIR_BACKUP_SHEET);
            if (!backup || backup.getLastRow() < 2) return JSON.stringify({ ok: false, message: "There is no repair to undo.", restored: 0, skipped: 0 });

            const backupValues = backup.getRange(2, 1, backup.getLastRow() - 1, TRANSACTION_REPAIR_BACKUP_HEADERS.length).getValues();
            let newestBatch = "";
            backupValues.forEach(row => { if (String(row[7]) === "" && String(row[0]) > newestBatch) newestBatch = String(row[0]); });
            if (!newestBatch) return JSON.stringify({ ok: false, message: "Every repair on the backup tab has already been undone.", restored: 0, skipped: 0 });

            const sheet = getTransactionsSheet();
            if (!sheet) return JSON.stringify({ ok: false, message: `The "${getTransactionsSheetName()}" tab was not found.`, restored: 0, skipped: 0 });
            const firstRow = getTransactionsRowStart();
            const lastRow = sheet.getLastRow();
            const idColumn = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, 1).getValues();
            const balanceAfterRange = sheet.getRange(firstRow, 13, lastRow - firstRow + 1, 1);
            const balanceAfterColumn = balanceAfterRange.getValues();

            let restored = 0, skipped = 0;
            const restoredAt = new Date().toISOString();
            backupValues.forEach((row, index) => {
                if (String(row[0]) !== newestBatch || String(row[7]) !== "") return;
                const position = Number(row[2]) - firstRow;
                if (position < 0 || position >= idColumn.length || Number(idColumn[position][0]) !== Number(row[3])) { skipped++; return; }
                balanceAfterColumn[position][0] = row[4];
                backup.getRange(index + 2, 8).setValue(`Restored ${restoredAt}`);
                restored++;
            });
            if (restored > 0) { balanceAfterRange.setValues(balanceAfterColumn); SpreadsheetApp.flush(); bumpTransactionLogEpoch_(); }
            log(`Transaction log repair undone: ${restored} row(s) restored, ${skipped} skipped.`, false);
            return JSON.stringify({ ok: true, message: `Restored ${restored} row(s) from the last repair${skipped ? `; ${skipped} skipped because their transaction id no longer matches the row` : ""}.`, restored, skipped });
        } catch (error: any) {
            log(`Error in undoTransactionRepair: ${error.message}`, false);
            return JSON.stringify({ ok: false, message: `The undo stopped: ${error.message}`, restored: 0, skipped: 0 });
        }
    }, JSON.stringify({ ok: false, message: "Another transaction is still running. Nothing was changed, please try again.", restored: 0, skipped: 0 }));
}
