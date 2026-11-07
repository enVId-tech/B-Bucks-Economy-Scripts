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

interface ServiceLimitReport {
    itemName: string;
    quarter: QuarterId;
    /** Cap from the services sheet, or null when this item has no cap for this quarter. */
    limit: number | null;
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

    undoTargets.forEach((_target, undoId) => inEffect.delete(undoId));
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

    records.forEach(record => {
        if (!inEffect.has(Number(record.id))) return;
        if (parsePurchaseItemName(record.serviceProvided) !== itemName) return;

        const timestampMs = Date.parse(String(record.timestamp));
        const recordQuarter = resolveCurrentQuarter(importantDates, timestampMs).quarter;
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

/** Looks up the Max/Person cap of an item for one quarter from the services sheet, or null when it has none. */
function lookupItemLimit(itemName: string, quarter: QuarterId): number | null {
    const services = fetchServicesDataCached();
    if (!Array.isArray(services)) return null;
    const item = services.find(candidate => candidate.itemName === itemName);
    const limit = item?.limit?.[quarter];
    return typeof limit === "number" && Number.isFinite(limit) && limit > 0 ? limit : null;
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

        const limit = lookupItemLimit(itemName, quarter);
        const report: ServiceLimitReport = { itemName, quarter, limit, enforced: isServicesLimitEnforced(), people: [] };
        if (limit === null) return report;

        const counts = countPurchasesFromLog(itemName, quarter);
        report.people = getSelectedIndividuals().map(person => {
            const used = counts.get(`${person.individual}|${person.period ?? ""}`) ?? 0;
            return { individual: person.individual, period: person.period, used, limit, remaining: Math.max(0, limit - used) };
        });
        return report;
    } catch (error: any) {
        log(`Error in fetchServiceLimitStatus: ${error.message}`, true);
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
    if ('error' in report || report.limit === null) return null;

    const overLimit = report.people.filter(person => quantity > person.remaining);
    if (overLimit.length === 0) return null;

    const summary = overLimit.map(person => `${person.individual} (${person.remaining} of ${person.limit} left)`).join(" - ");
    return `Max per person reached for ${itemName} in ${quarter}: ${summary}. Use "Override Restrictions" to allow it anyway.`;
}
