/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 *
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 *
 * Enforces the "Max/Person" columns of the services price sheet.
 * The transaction log is the only source of truth: how many times a person already bought an item in a quarter is counted from
 * the log, and any transaction that was undone is not counted.
 */

interface ServiceLimitStatus {
    individual: string;
    period?: number;
    /** How many units the person already bought this quarter (undone transactions excluded). */
    used: number;
    /** The cap from the services sheet. */
    limit: number;
    /** limit - used, never below zero. */
    remaining: number;
}

interface ClassLimitStatus {
    /** Period number taken from the sheet name. */
    period: number;
    /** Units of the item the whole class already bought this quarter (undone transactions excluded). */
    used: number;
    /** The Max/Class cap from the services sheet. */
    limit: number;
    remaining: number;
    /** How many students of this class are currently selected. */
    selectedStudents: number;
}

interface ServiceLimitReport {
    itemName: string;
    quarter: QuarterId;
    /** Max/Person from the services sheet, or null when this item has none for this quarter. */
    limit: number | null;
    /** Max/Class from the services sheet, or null when this item has none for this quarter. */
    classLimit: number | null;
    classes: ClassLimitStatus[];
    /** False when the "enforceServicesLimits" setting is turned off, the counts are then informational only. */
    enforced: boolean;
    people: ServiceLimitStatus[];
}

// A purchase is logged as "$<price> - <item name> (x<quantity>)", see addBalance() in ServicesManager.html.
const SERVICE_PURCHASE_LOG_PATTERN: RegExp = /^\$-?[\d.,]+ - (.+) \(x\d+\)$/;
// Undoing writes "Undo transaction ID <id> - <original text>".
const UNDO_LOG_PATTERN: RegExp = /^Undo transaction ID (\d+) - /;
// How many log rows are read per step while walking backwards through the log.
const LIMIT_SCAN_CHUNK_SIZE: number = 500;

/** True unless the "enforceServicesLimits" setting is explicitly switched off. */
function isServicesLimitEnforced(): boolean {
    const setting = fetchProperty("enforceServicesLimits", "mandatedPolicies");
    if (setting === false || String(setting).toLowerCase() === "false") return false;
    return true;
}

/** Extracts the item name from a purchase log text, or null when the text is not a purchase. */
function parsePurchaseItemName(serviceProvided: any): string | null {
    const match = SERVICE_PURCHASE_LOG_PATTERN.exec(String(serviceProvided ?? "").trim());
    return match ? match[1].trim() : null;
}

/** Id of the transaction an undo record reverses, or null when the record is not an undo. */
function parseUndoTargetId(serviceProvided: any): number | null {
    const match = UNDO_LOG_PATTERN.exec(String(serviceProvided ?? "").trim());
    return match ? Number(match[1]) : null;
}

/** Start of a quarter in epoch milliseconds, or 0 when the quarter has no configured start date. */
function getQuarterStartMs(quarter: QuarterId): number {
    const settings = fetchSettingsDataCached();
    if ('error' in settings) return 0;
    const dates = settings.importantDates as any;
    return readQuarterBoundary(dates, QUARTER_IDS.indexOf(quarter), "start") ?? 0;
}

/**
 * Works out which log records are still "in effect". An undo record switches its target off, and undoing an undo switches the
 * original back on, so this walks the records oldest-first.
 * @returns the set of record ids that are still in effect (undo records themselves are never in effect).
 */
function resolveRecordsInEffect(records: TransactionRecord[]): Set<number> {
    const activeRecords = resolveActiveRecords(records);
    records.forEach(record => {
        if (parseUndoTargetId(record.serviceProvided) !== null) activeRecords.delete(Number(record.id));
    });
    return activeRecords;
}

/**
 * Same walk as resolveRecordsInEffect(), but undo records that are still active (not undone themselves) are kept in the set.
 * Undo uses this to refuse reversing something twice.
 */
function resolveActiveRecords(records: TransactionRecord[]): Set<number> {
    const inEffect = new Set<number>();
    const undoTargets = new Map<number, number>();

    records.forEach(record => {
        const id = Number(record.id);
        inEffect.add(id);
        const targetId = parseUndoTargetId(record.serviceProvided);
        if (targetId !== null) undoTargets.set(id, targetId);
    });

    // oldest first, so a later undo overrides an earlier state
    const switchOff = (id: number): void => {
        if (!inEffect.has(id)) return;
        inEffect.delete(id);
        const reversedTarget = undoTargets.get(id);
        if (reversedTarget !== undefined) inEffect.add(reversedTarget); // an undone undo brings its target back
    };
    records.forEach(record => {
        const targetId = undoTargets.get(Number(record.id));
        if (targetId !== undefined && inEffect.has(Number(record.id))) switchOff(targetId);
    });

    return inEffect;
}

/**
 * Counts how many units of one item each person bought during a quarter, straight from the transaction log.
 * Walks the log from the newest row backwards and stops as soon as a chunk is entirely older than the quarter, so a long log stays cheap.
 * @returns a map "<individual>|<period>" -> units bought.
 */
function countPurchasesFromLog(itemName: string, quarter: QuarterId): Map<string, number> {
    const counts = new Map<string, number>();
    const sheet = getTransactionsSheet();
    if (!sheet) return counts;

    const total = countTransactionRows(sheet);
    const quarterStartMs = getQuarterStartMs(quarter);
    const settings = fetchSettingsDataCached();
    const importantDates = ('error' in settings ? {} : settings.importantDates) as any;

    const recordsNewestChunkFirst: TransactionRecord[][] = [];
    let chunkEnd = total;
    while (chunkEnd > 0) {
        const chunkStart = Math.max(0, chunkEnd - LIMIT_SCAN_CHUNK_SIZE);
        const chunk = readTransactionRows(sheet, chunkStart, chunkEnd - chunkStart);
        recordsNewestChunkFirst.push(chunk);
        const oldestTimestampMs = Date.parse(String(chunk[0].timestamp));
        if (quarterStartMs > 0 && Number.isFinite(oldestTimestampMs) && oldestTimestampMs < quarterStartMs) break;
        chunkEnd = chunkStart;
    }

    const records = recordsNewestChunkFirst.reverse().reduce((all, chunk) => all.concat(chunk), [] as TransactionRecord[]);
    const inEffect = resolveRecordsInEffect(records);

    // There are only 8 quarter boundaries, so their "is it midnight" answers are remembered instead of asking the
    // spreadsheet's time zone again for every record.
    const midnightAnswers = new Map<number, boolean>();
    const rememberedMidnightCheck = (timestampMs: number): boolean => {
        if (!midnightAnswers.has(timestampMs)) midnightAnswers.set(timestampMs, isMidnight(timestampMs));
        return midnightAnswers.get(timestampMs) as boolean;
    };

    records.forEach(record => {
        if (!inEffect.has(Number(record.id))) return;
        if (parsePurchaseItemName(record.serviceProvided) !== itemName) return;

        const timestampMs = Date.parse(String(record.timestamp));
        if (!Number.isFinite(timestampMs)) return;
        const recordQuarter = resolveCurrentQuarter(importantDates, timestampMs, rememberedMidnightCheck).quarter;
        if (recordQuarter !== quarter) return;

        const key = `${String(record.individual).trim()}|${record.period ?? ""}`;
        counts.set(key, (counts.get(key) ?? 0) + (Number(record.quantity) || 0));
    });

    return counts;
}

/** The people on the currently selected rows of the active sheet (name + period), without duplicates. */
function getSelectedIndividuals(): { individual: string; period?: number }[] {
    const people: { individual: string; period?: number }[] = [];
    const seen = new Set<string>();

    const activeRangeList = getSS().getActiveRangeList();
    if (!activeRangeList) return people;

    activeRangeList.getRanges().forEach(range => {
        const sheet = range.getSheet();
        const periodNumber = parseInt(sheet.getName().replace(/\D/g, ""), 10);
        for (let row = range.getRow(); row < range.getRow() + range.getNumRows(); row++) {
            if (row < getUserStartingRow()) continue;
            const individual = String(sheet.getRange(row, getNamesColumn()).getValue()).trim();
            if (!individual) continue;
            const key = `${individual}|${periodNumber}`;
            if (seen.has(key)) continue;
            seen.add(key);
            people.push({ individual, period: Number.isFinite(periodNumber) ? periodNumber : undefined });
        }
    });
    return people;
}

/** Looks up the Max/Person and Max/Class caps of an item for one quarter from the services sheet (null = no cap). */
function lookupItemLimits(itemName: string, quarter: QuarterId): { personLimit: number | null; classLimit: number | null } {
    const asCap = (value: any): number | null => typeof value === "number" && Number.isFinite(value) && value > 0 ? value : null;
    const services = fetchServicesDataCached();
    if (!Array.isArray(services)) return { personLimit: null, classLimit: null };
    const item = services.find(candidate => candidate.itemName === itemName);
    return { personLimit: asCap(item?.limit?.[quarter]), classLimit: asCap(item?.classLimit?.[quarter]) };
}

/**
 * Server endpoint for the Services dialog: how many purchases of an item are left for each selected person?
 * @param data JSON string: { itemName: string, quarter: "Q1".."Q4" }
 */
function fetchServiceLimitStatus(data?: string): ServiceLimitReport | { error: string } {
    try {
        const parsed = data ? JSON.parse(data) : {};
        const itemName = String(parsed.itemName ?? "").trim();
        const quarter = parsed.quarter as QuarterId;
        if (!itemName || !QUARTER_IDS.includes(quarter)) return { error: "An item name and a quarter are required." };

        const { personLimit: limit, classLimit } = lookupItemLimits(itemName, quarter);
        const report: ServiceLimitReport = { itemName, quarter, limit, classLimit, classes: [], enforced: isServicesLimitEnforced(), people: [] };
        if (limit === null && classLimit === null) return report;

        const counts = countPurchasesFromLog(itemName, quarter);
        const selectedPeople = getSelectedIndividuals();

        if (limit !== null) {
            report.people = selectedPeople.map(person => {
                const used = counts.get(`${person.individual}|${person.period ?? ""}`) ?? 0;
                return { individual: person.individual, period: person.period, used, limit, remaining: Math.max(0, limit - used) };
            });
        }

        if (classLimit !== null) {
            // A class is a period sheet, so its total is every student's purchases of this item on that period.
            const classTotals = new Map<number, number>();
            counts.forEach((units, key) => {
                const period = Number(key.slice(key.lastIndexOf("|") + 1));
                classTotals.set(period, (classTotals.get(period) ?? 0) + units);
            });

            const selectedPerPeriod = new Map<number, number>();
            selectedPeople.forEach(person => {
                if (person.period !== undefined) selectedPerPeriod.set(person.period, (selectedPerPeriod.get(person.period) ?? 0) + 1);
            });

            selectedPerPeriod.forEach((selectedStudents, period) => {
                const used = classTotals.get(period) ?? 0;
                report.classes.push({ period, used, limit: classLimit, remaining: Math.max(0, classLimit - used), selectedStudents });
            });
        }
        return report;
    } catch (error: any) {
        // Not shown as a dialog: the Services page polls this every few seconds.
        log(`Error in fetchServiceLimitStatus: ${error.message}`, false);
        return { error: `Error in fetchServiceLimitStatus: ${error.message}` };
    }
}

/**
 * Checks a purchase against the Max/Person cap for every selected person.
 * @returns an error message when somebody would go over their cap, otherwise null.
 */
function checkServiceLimit(itemName: string, quarter: QuarterId, quantity: number): string | null {
    if (!isServicesLimitEnforced()) return null;

    const report = fetchServiceLimitStatus(JSON.stringify({ itemName, quarter }));
    if ('error' in report) return null;

    const messages: string[] = [];

    const overPersonLimit = report.people.filter(person => quantity > person.remaining);
    if (overPersonLimit.length > 0) {
        const summary = overPersonLimit.map(person => `${person.individual} (${person.remaining} of ${person.limit} left)`).join(" - ");
        messages.push(`Max per person reached for ${itemName} in ${quarter}: ${summary}.`);
    }

    // Every selected student of a class buys the item, so the class uses up quantity x selected students.
    const overClassLimit = report.classes.filter(classStatus => quantity * classStatus.selectedStudents > classStatus.remaining);
    if (overClassLimit.length > 0) {
        const summary = overClassLimit.map(classStatus => `Period ${classStatus.period} (${classStatus.remaining} of ${classStatus.limit} left)`).join(" - ");
        messages.push(`Max per class reached for ${itemName} in ${quarter}: ${summary}.`);
    }

    if (messages.length === 0) return null;
    return `${messages.join(" ")} Use "Override Restrictions" to allow it anyway.`;
}
