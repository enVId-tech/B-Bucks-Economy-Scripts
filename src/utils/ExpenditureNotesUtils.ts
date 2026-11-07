/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 *
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 *
 * The note on a student's expenditure cell is a VIEW of the transaction log, never a separate record.
 * Every time the log changes for a student (a purchase, an undo, a CoD deposit) the whole note is rebuilt from the log, so the
 * note always matches the log and an undone transaction disappears from the note.
 */

interface NoteTarget {
    individual: string;
    period?: number;
}

// Cell notes can hold far more than this, but a note this long is unreadable anyway. The newest entries are always kept.
const MAX_EXPENDITURE_NOTE_LINES: number = 60;
// How many log rows are read per step while collecting a student's history.
const NOTE_SCAN_CHUNK_SIZE: number = 2000;

// While a batch is open (for example during a multi-transaction undo) the targets are collected and synced once at the end.
let pendingNoteTargets: Map<string, NoteTarget> | null = null;

function noteTargetKey(target: NoteTarget): string {
    return `${String(target.individual).trim()}|${target.period ?? ""}`;
}

/** Starts collecting note updates instead of applying them one by one. Always pair with endNoteSyncBatch() in a finally block. */
function beginNoteSyncBatch(): void {
    pendingNoteTargets = new Map<string, NoteTarget>();
}

/** Applies every note update collected since beginNoteSyncBatch() in one pass over the log. */
function endNoteSyncBatch(): void {
    const targets = pendingNoteTargets ? Array.from(pendingNoteTargets.values()) : [];
    pendingNoteTargets = null;
    if (targets.length > 0) syncExpenditureNotes(targets);
}

/** True when a log record changes (or reverses a change to) the expenditure total shown on the student's row. */
function recordAffectsExpenditureNote(record: TransactionRecord): boolean {
    const modifiedColumn = Number(record.modifiedColumn);
    return modifiedColumn === getExpendituresColumn() || modifiedColumn === getInitialDepositColumn();
}

/**
 * Called after records were appended to the log. Rebuilds the notes of every student whose expenditures were touched.
 * Inside a batch the work is postponed until endNoteSyncBatch().
 */
function syncNotesForRecords(records: TransactionRecord[]): void {
    try {
        const targets = records
            .filter(recordAffectsExpenditureNote)
            .filter(record => record.individual)
            .map(record => ({ individual: String(record.individual), period: record.period }));
        if (targets.length === 0) return;

        if (pendingNoteTargets) {
            targets.forEach(target => pendingNoteTargets!.set(noteTargetKey(target), target));
            return;
        }
        syncExpenditureNotes(targets);
    } catch (error: any) {
        // A failed note must never undo or hide a transaction that already succeeded.
        log(`Could not update expenditure notes: ${error.message}`, true);
    }
}

/** Formats a log timestamp down to the second in the spreadsheet's own time zone, e.g. "11/05/2026 14:03:27". */
function formatNoteTimestamp(isoTimestamp: any): string {
    const parsedMs = Date.parse(String(isoTimestamp));
    if (!Number.isFinite(parsedMs)) return "unknown time";
    return Utilities.formatDate(new Date(parsedMs), getSS().getSpreadsheetTimeZone(), "MM/dd/yyyy HH:mm:ss");
}

/** One note line for one log record, or null when the record is not an expenditure entry. */
function formatExpenditureNoteLine(record: TransactionRecord): string | null {
    const amount = Number(Number(record.tenderedMoney).toFixed(2));
    if (!Number.isFinite(amount)) return null;
    const timestamp = formatNoteTimestamp(record.timestamp);
    const reason = String(record.serviceProvided ?? "").trim();

    if (Number(record.modifiedColumn) === getInitialDepositColumn()) {
        if (!/^Deposit/i.test(reason)) return null;
        const overrideSuffix = /\(Override\)/i.test(reason) ? " (Override)" : "";
        return `$${amount} - ${timestamp} CoDs/Investments${overrideSuffix}`;
    }

    if (Number(record.modifiedColumn) !== getExpendituresColumn()) return null;
    if (record.operation === "ADD") return `$${amount} - ${timestamp} ${reason || "Manual Adjustment"}`;
    if (record.operation === "SUBTRACT") return `-$${amount} - ${timestamp} ${reason || "Manual Adjustment"}`;
    return null;
}

/**
 * Reads the log once and returns, per student, the records that are still in effect (undone ones removed) oldest first.
 * @param wantedKeys keys from noteTargetKey(). Pass null to collect every student.
 */
function collectExpenditureRecords(wantedKeys: Set<string> | null): Map<string, TransactionRecord[]> {
    const recordsByStudent = new Map<string, TransactionRecord[]>();
    const sheet = getTransactionsSheet();
    if (!sheet) return recordsByStudent;

    const total = countTransactionRows(sheet);
    for (let chunkStart = 0; chunkStart < total; chunkStart += NOTE_SCAN_CHUNK_SIZE) {
        const chunk = readTransactionRows(sheet, chunkStart, Math.min(NOTE_SCAN_CHUNK_SIZE, total - chunkStart));
        chunk.forEach(record => {
            if (!record.individual || !recordAffectsExpenditureNote(record)) return;
            const key = noteTargetKey({ individual: record.individual, period: record.period });
            if (wantedKeys && !wantedKeys.has(key)) return;
            const studentRecords = recordsByStudent.get(key) ?? [];
            studentRecords.push(record);
            recordsByStudent.set(key, studentRecords);
        });
    }

    recordsByStudent.forEach((studentRecords, key) => {
        const inEffect = resolveRecordsInEffect(studentRecords);
        recordsByStudent.set(key, studentRecords.filter(record => inEffect.has(Number(record.id))));
    });
    return recordsByStudent;
}

/** Finds the row of a student on a period sheet by NAME (rows move when names are sorted, so a stored row number can be stale). */
function findStudentRow(sheet: GoogleAppsScript.Spreadsheet.Sheet, individual: string): number | null {
    const lastRow = sheet.getLastRow();
    if (lastRow < getUserStartingRow()) return null;
    const names = sheet.getRange(getUserStartingRow(), getNamesColumn(), lastRow - getUserStartingRow() + 1, 1).getValues();
    const offset = names.findIndex(row => String(row[0]).trim() === individual.trim());
    return offset < 0 ? null : getUserStartingRow() + offset;
}

/** Turns a student's records into the final note text (newest entries win when the note would get too long). */
function buildExpenditureNoteText(studentRecords: TransactionRecord[]): string {
    const lines = studentRecords
        .map(formatExpenditureNoteLine)
        .filter((line): line is string => line !== null);

    if (lines.length <= MAX_EXPENDITURE_NOTE_LINES) return lines.join("\n");
    const hiddenCount = lines.length - MAX_EXPENDITURE_NOTE_LINES;
    return [`(${hiddenCount} older entries are in the transaction log)`, ...lines.slice(hiddenCount)].join("\n");
}

/**
 * Rebuilds the expenditure note of each listed student from the transaction log and writes it on the right period sheet.
 * A student whose every expenditure was undone gets an empty note.
 * @returns true when every note was written.
 */
function syncExpenditureNotes(targets: NoteTarget[]): boolean {
    try {
        const uniqueTargets = new Map<string, NoteTarget>();
        targets.forEach(target => uniqueTargets.set(noteTargetKey(target), target));

        const recordsByStudent = collectExpenditureRecords(new Set(uniqueTargets.keys()));
        let allWritten = true;

        uniqueTargets.forEach((target, key) => {
            const sheet = getAllSheets().find(candidate =>
                parseInt(candidate.getName().replace(/\D/g, ""), 10) === Number(target.period)
            );
            const row = sheet ? findStudentRow(sheet, target.individual) : null;
            if (!sheet || row === null) {
                log(`Could not place the expenditure note for "${target.individual}" (period ${target.period}).`, true);
                allWritten = false;
                return;
            }

            const noteText = buildExpenditureNoteText(recordsByStudent.get(key) ?? []);
            sheet.getRange(row, getExpendituresColumn()).setNote(noteText);
            log(`Expenditure note rebuilt from the log for ${target.individual} (period ${target.period}): ${noteText.split("\n").filter(Boolean).length} entries.`, false);
        });
        return allWritten;
    } catch (error: any) {
        log(`Error in syncExpenditureNotes: ${error.message}`, true);
        return false;
    }
}

/**
 * Maintenance: rebuilds the expenditure notes of every student that has expenditure entries in the log.
 * Students with no entries in the log keep whatever note they have. Run it from the menu after upgrading, or any time notes look off.
 */
function rebuildAllExpenditureNotes(): string {
    try {
        const recordsByStudent = collectExpenditureRecords(null);
        const targets: NoteTarget[] = [];
        recordsByStudent.forEach((_studentRecords, key) => {
            const separatorIndex = key.lastIndexOf("|");
            const periodText = key.slice(separatorIndex + 1);
            targets.push({ individual: key.slice(0, separatorIndex), period: periodText === "" ? undefined : Number(periodText) });
        });

        const success = syncExpenditureNotes(targets);
        const message = `Rebuilt expenditure notes for ${targets.length} student(s) from the transaction log.${success ? "" : " Some students could not be found, see the log."}`;
        log(message, true);
        return message;
    } catch (error: any) {
        log(`Error in rebuildAllExpenditureNotes: ${error.message}`, true);
        return `Error occurred in rebuildAllExpenditureNotes: ${error.message}`;
    }
}

/** Menu entry: rebuilds every expenditure note from the log and tells the user how it went. */
function rebuildExpenditureNotesFromMenu(): void {
    const message = rebuildAllExpenditureNotes();
    SpreadsheetApp.getActive().toast(message, "Expenditure notes", 8);
}
