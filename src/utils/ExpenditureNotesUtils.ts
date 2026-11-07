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
        log(`Could not update expenditure notes: ${error.message}`, false);
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

// Each student's expenditure history is cached (key below + "<name>|<period>") together with how many log rows it covers, so a
// purchase only has to read the rows added since then instead of the whole log. The log stays the source of truth: an entry is
// only trusted while the log still has at least that many rows and the id on its last covered row still matches.
const EXPENDITURE_NOTE_CACHE_PREFIX: string = "expenditureNoteRecords:";

interface CachedNoteRecords {
    /** How many log rows (from the oldest) this entry has already taken into account. */
    rows: number;
    /** Id of the log row at index rows - 1, used to notice that the log was cleared or rewritten. */
    lastId: number;
    records: TransactionRecord[];
}

/** Keeps only the fields a note needs, so the cached entries stay small. */
function toNoteRecord(record: TransactionRecord): TransactionRecord {
    return {
        id: Number(record.id),
        serviceProvided: String(record.serviceProvided ?? ""),
        unitPrice: 0,
        tenderedMoney: record.tenderedMoney,
        timestamp: record.timestamp,
        modifiedColumn: Number(record.modifiedColumn),
        operation: record.operation,
    };
}

/** Id stored on the log row at a chronological index (0 = oldest). */
function readTransactionIdAt(sheet: GoogleAppsScript.Spreadsheet.Sheet, index: number): number {
    return Number(sheet.getRange(getTransactionsRowStart() + index, 1).getValue());
}

/**
 * Reads log rows [fromIndex, toIndex) and hands every expenditure-related record to `onRecord` with its index.
 * @returns the id of the last row read, or null when nothing was read.
 */
function scanExpenditureLog(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    fromIndex: number,
    toIndex: number,
    onRecord: (key: string, record: TransactionRecord, index: number) => void
): number | null {
    let lastIdRead: number | null = null;
    for (let chunkStart = fromIndex; chunkStart < toIndex; chunkStart += NOTE_SCAN_CHUNK_SIZE) {
        const chunk = readTransactionRows(sheet, chunkStart, Math.min(NOTE_SCAN_CHUNK_SIZE, toIndex - chunkStart));
        chunk.forEach((record, offset) => {
            lastIdRead = Number(record.id);
            if (!record.individual || !recordAffectsExpenditureNote(record)) return;
            onRecord(noteTargetKey({ individual: record.individual, period: record.period }), toNoteRecord(record), chunkStart + offset);
        });
    }
    return lastIdRead;
}

/** Loads the cached histories of the wanted students, or null when any of them is missing or no longer matches the log. */
function loadCachedNoteRecords(
    cache: GoogleAppsScript.Cache.Cache,
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    wantedKeys: Set<string>,
    total: number
): Map<string, CachedNoteRecords> | null {
    const keys = Array.from(wantedKeys);
    const cachedRaw = cache.getAll(keys.map(key => EXPENDITURE_NOTE_CACHE_PREFIX + key));
    const entries = new Map<string, CachedNoteRecords>();

    for (const key of keys) {
        const raw = cachedRaw[EXPENDITURE_NOTE_CACHE_PREFIX + key];
        if (!raw) return null;
        try {
            const entry = JSON.parse(raw) as CachedNoteRecords;
            if (!entry || !Array.isArray(entry.records) || !(entry.rows >= 0 && entry.rows <= total)) return null;
            entries.set(key, entry);
        } catch (parseError) {
            return null;
        }
    }

    // Every distinct anchor is checked once (usually all entries share one).
    const idsAtAnchor = new Map<number, number>();
    for (const entry of Array.from(entries.values())) {
        if (entry.rows === 0) continue;
        if (!idsAtAnchor.has(entry.rows)) idsAtAnchor.set(entry.rows, readTransactionIdAt(sheet, entry.rows - 1));
        if (idsAtAnchor.get(entry.rows) !== Number(entry.lastId)) return null;
    }
    return entries;
}

/** Saves student histories to the cache. A history too large for one cache value is simply left out (it is rebuilt next time). */
function saveCachedNoteRecords(cache: GoogleAppsScript.Cache.Cache, entries: Map<string, CachedNoteRecords>): void {
    const values: { [key: string]: string } = {};
    entries.forEach((entry, key) => { values[EXPENDITURE_NOTE_CACHE_PREFIX + key] = JSON.stringify(entry); });
    try {
        cache.putAll(values, SERVER_SIDE_CACHE_AGE);
    } catch (bulkError) {
        Object.keys(values).forEach(cacheKey => {
            try { cache.put(cacheKey, values[cacheKey], SERVER_SIDE_CACHE_AGE); } catch (sizeError) { /* too large to cache */ }
        });
    }
}

/**
 * Returns, per student, the expenditure records that are still in effect (undone ones removed), oldest first.
 * Uses the cached histories when they still match the log and only reads the rows appended since; otherwise reads the whole log
 * once and refreshes every student's cached history.
 * @param wantedKeys keys from noteTargetKey(). Pass null to collect every student (always reads the whole log).
 */
function collectExpenditureRecords(wantedKeys: Set<string> | null): Map<string, TransactionRecord[]> {
    const recordsByStudent = new Map<string, TransactionRecord[]>();
    const sheet = getTransactionsSheet();
    if (!sheet) return recordsByStudent;

    const total = countTransactionRows(sheet);
    let cache: GoogleAppsScript.Cache.Cache | null = null;
    try { cache = CacheService.getScriptCache(); } catch (cacheError) { cache = null; }

    let entries: Map<string, CachedNoteRecords> | null = null;
    if (wantedKeys && cache && total > 0) {
        try { entries = loadCachedNoteRecords(cache, sheet, wantedKeys, total); } catch (cacheError) { entries = null; }
    }

    if (entries) {
        // Fast path: only the rows the cached histories have not seen yet.
        const cachedEntries = entries;
        const fromIndex = Math.min(...Array.from(cachedEntries.values()).map(entry => entry.rows));
        const lastIdRead = scanExpenditureLog(sheet, fromIndex, total, (key, record, index) => {
            const entry = cachedEntries.get(key);
            if (entry && index >= entry.rows) entry.records.push(record);
        });
        if (lastIdRead !== null) {
            cachedEntries.forEach(entry => { entry.rows = total; entry.lastId = lastIdRead; });
            if (cache) saveCachedNoteRecords(cache, cachedEntries);
        }
    } else {
        // Slow path: whole log, every student, then refresh the whole cache so the next purchases take the fast path.
        entries = new Map<string, CachedNoteRecords>();
        const allEntries = entries;
        const lastIdRead = scanExpenditureLog(sheet, 0, total, (key, record) => {
            const entry = allEntries.get(key) ?? { rows: 0, lastId: 0, records: [] };
            entry.records.push(record);
            allEntries.set(key, entry);
        });
        if (lastIdRead !== null) {
            allEntries.forEach(entry => { entry.rows = total; entry.lastId = lastIdRead; });
            if (cache) saveCachedNoteRecords(cache, allEntries);
        }
    }

    entries.forEach((entry, key) => {
        if (wantedKeys && !wantedKeys.has(key)) return;
        const inEffect = resolveRecordsInEffect(entry.records);
        recordsByStudent.set(key, entry.records.filter(record => inEffect.has(Number(record.id))));
    });
    return recordsByStudent;
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

        // Group by period sheet: one read of the names, one read and one write of the notes per sheet.
        const targetsBySheet = new Map<GoogleAppsScript.Spreadsheet.Sheet, { key: string; target: NoteTarget }[]>();
        uniqueTargets.forEach((target, key) => {
            const sheet = getAllSheets().find(candidate =>
                parseInt(candidate.getName().replace(/\D/g, ""), 10) === Number(target.period)
            );
            if (!sheet) {
                log(`Could not place the expenditure note for "${target.individual}" (period ${target.period}).`, false);
                allWritten = false;
                return;
            }
            const sheetTargets = targetsBySheet.get(sheet) ?? [];
            sheetTargets.push({ key, target });
            targetsBySheet.set(sheet, sheetTargets);
        });

        targetsBySheet.forEach((sheetTargets, sheet) => {
            const lastRow = sheet.getLastRow();
            if (lastRow < getUserStartingRow()) {
                allWritten = false;
                return;
            }
            const names = sheet.getRange(getUserStartingRow(), getNamesColumn(), lastRow - getUserStartingRow() + 1, 1).getValues();
            const rowByName = new Map<string, number>();
            names.forEach((nameRow, offset) => {
                const name = String(nameRow[0]).trim();
                if (name && !rowByName.has(name)) rowByName.set(name, getUserStartingRow() + offset); // first match wins, like before
            });

            const noteByRow = new Map<number, string>();
            sheetTargets.forEach(({ key, target }) => {
                const row = rowByName.get(String(target.individual).trim());
                if (row === undefined) {
                    log(`Could not place the expenditure note for "${target.individual}" (period ${target.period}).`, false);
                    allWritten = false;
                    return;
                }
                noteByRow.set(row, buildExpenditureNoteText(recordsByStudent.get(key) ?? []));
            });
            if (noteByRow.size === 0) return;

            // Notes of rows in between are read and written back unchanged (safe: this runs under the transaction lock).
            const rows = Array.from(noteByRow.keys());
            const firstRow = Math.min(...rows);
            const noteRange = sheet.getRange(firstRow, getExpendituresColumn(), Math.max(...rows) - firstRow + 1, 1);
            const notes = noteRange.getNotes();
            noteByRow.forEach((noteText, row) => { notes[row - firstRow][0] = noteText; });
            noteRange.setNotes(notes);
            log(`Expenditure notes rebuilt from the log for ${noteByRow.size} student(s) on ${sheet.getName()}.`, false);
        });
        return allWritten;
    } catch (error: any) {
        log(`Error in syncExpenditureNotes: ${error.message}`, false);
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
        log(message, false); // the menu entry shows it as a toast
        return message;
    } catch (error: any) {
        log(`Error in rebuildAllExpenditureNotes: ${error.message}`, false);
        return `Error occurred in rebuildAllExpenditureNotes: ${error.message}`;
    }
}

/** Menu entry: rebuilds every expenditure note from the log and tells the user how it went. */
function rebuildExpenditureNotesFromMenu(): void {
    const message = runWithTransactionLock(() => rebuildAllExpenditureNotes(), "Another transaction is still running, try again in a moment.");
    SpreadsheetApp.getActive().toast(message, "Expenditure notes", 8);
}
