/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 *
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 *
 * Works out which grading quarter (Q1-Q4) is active for a given moment, based on the start/end dates in the Settings sheet
 * ("Important Dates": startQ1Date, endQ1Date ... startQ4Date, endQ4Date). This is the single source of truth: the dialogs
 * call getCurrentQuarterInfo() instead of re-implementing the date maths on the client.
 */

type QuarterId = "Q1" | "Q2" | "Q3" | "Q4";

/**
 * - active:       today falls inside a quarter's start/end dates.
 * - between:      a quarter has ended and the next one has not started yet (e.g. a break). The most recently STARTED quarter stays selected.
 * - before:       no quarter has started yet. Defaults to the earliest quarter.
 * - after:        the last configured quarter has ended. It stays selected.
 * - unconfigured: no usable start dates were found. Defaults to Q1.
 */
type QuarterStatus = "active" | "between" | "before" | "after" | "unconfigured";

interface QuarterResolution {
    quarter: QuarterId;
    status: QuarterStatus;
    /** ISO timestamp the answer is valid for. */
    asOf: string;
    /** ISO start of the resolved quarter, when known. */
    startsOn?: string;
    /** ISO end (inclusive) of the resolved quarter, when known. */
    endsOn?: string;
}

const QUARTER_IDS: QuarterId[] = ["Q1", "Q2", "Q3", "Q4"];
// Older builds of the client looked for these names, so they are still accepted as aliases.
const QUARTER_WORDS: string[] = ["One", "Two", "Three", "Four"];

/** Parses a settings value (ISO string, Date, epoch ms or locale string) into epoch ms. Returns null when unusable. */
function parseQuarterBoundary(value: any): number | null {
    if (value === undefined || value === null || value === "") return null;
    const ms = value instanceof Date ? value.getTime() : typeof value === 'number' ? value : Date.parse(String(value));
    return Number.isFinite(ms) ? ms : null;
}

/** Reads one boundary from the important-dates object, trying the canonical key first and then the legacy alias. */
function readQuarterBoundary(dates: { [key: string]: any }, index: number, edge: "start" | "end"): number | null {
    const canonical = `${edge}Q${index + 1}Date`;
    const legacy = `${edge}Quarter${QUARTER_WORDS[index]}Date`;
    return parseQuarterBoundary(dates[canonical]) ?? parseQuarterBoundary(dates[legacy]);
}

/** True when a timestamp sits exactly on local midnight, i.e. the sheet cell held a date with no time-of-day. */
function isMidnight(ms: number): boolean {
    try {
        if (typeof Utilities !== 'undefined' && typeof SpreadsheetApp !== 'undefined') {
            const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
            return Utilities.formatDate(new Date(ms), tz, "HH:mm:ss") === "00:00:00";
        }
    } catch (e) { /* fall through to the runtime time zone */ }
    const d = new Date(ms);
    return d.getHours() === 0 && d.getMinutes() === 0 && d.getSeconds() === 0;
}

/**
 * Pure quarter resolver (no Apps Script calls apart from the optional time-zone lookup in isMidnight).
 * An end date typed as a plain date means "through the end of that day", so a midnight end boundary is extended by 24h.
 * Where windows overlap, the quarter that started most recently wins.
 * @param dates The importantDates settings object.
 * @param nowMs The moment to evaluate. Defaults to now.
 * @param midnightCheck Injectable for tests.
 */
function resolveCurrentQuarter(dates: { [key: string]: any } | undefined | null, nowMs: number = Date.now(), midnightCheck: (ms: number) => boolean = isMidnight): QuarterResolution {
    const asOf = new Date(nowMs).toISOString();
    const source = dates || {};

    const quarters = QUARTER_IDS.map((id, index) => {
        const start = readQuarterBoundary(source, index, "start");
        const rawEnd = readQuarterBoundary(source, index, "end");
        const end = rawEnd === null ? null : (midnightCheck(rawEnd) ? rawEnd + 24 * 60 * 60 * 1000 - 1 : rawEnd);
        return { id, start, end };
    }).filter(q => q.start !== null) as { id: QuarterId; start: number; end: number | null }[];

    if (quarters.length === 0) return { quarter: "Q1", status: "unconfigured", asOf };

    const describe = (q: { start: number; end: number | null }) => ({
        startsOn: new Date(q.start).toISOString(),
        ...(q.end !== null ? { endsOn: new Date(q.end).toISOString() } : {}),
    });

    const started = quarters.filter(q => q.start <= nowMs).sort((a, b) => a.start - b.start);

    if (started.length === 0) {
        const earliest = quarters.slice().sort((a, b) => a.start - b.start)[0];
        return { quarter: earliest.id, status: "before", asOf, ...describe(earliest) };
    }

    const current = started[started.length - 1];
    if (current.end === null || nowMs <= current.end) {
        return { quarter: current.id, status: "active", asOf, ...describe(current) };
    }

    const hasLater = quarters.some(q => q.start > nowMs);
    return { quarter: current.id, status: hasLater ? "between" : "after", asOf, ...describe(current) };
}

/**
 * Server endpoint for the dialogs: which quarter is it right now?
 * @param data Optional JSON string: { nowMs?: number } (nowMs is for testing/debugging only).
 */
function getCurrentQuarterInfo(data?: string): QuarterResolution | { error: string } {
    try {
        let nowMs = Date.now();
        if (data && typeof data === 'string') {
            const parsed = JSON.parse(data);
            if (parsed && Number.isFinite(parsed.nowMs)) nowMs = Number(parsed.nowMs);
        }

        const settings = fetchSettingsDataCached();
        if ('error' in settings) return { error: settings.error };

        return resolveCurrentQuarter(settings.importantDates as any, nowMs);
    } catch (error: any) {
        log(`Error in getCurrentQuarterInfo: ${error.message}`, false);
        return { error: `Error in getCurrentQuarterInfo: ${error.message}` };
    }
}
