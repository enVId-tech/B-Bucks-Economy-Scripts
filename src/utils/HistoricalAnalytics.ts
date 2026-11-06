/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran. 
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 * 
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 * 
 * This file builds the data behind the Historical Analytics chart: every snapshot stored on the live Historical Records sheet and on any archived
 * "Records_*" sheets, PLUS a live data point computed from the current state of the period sheets. It provides a structured JSON payload for client-side chart parsing and analysis.
 */

interface MetricHeader {
  key: string;
  label: string;
  colIndex: number;
}

interface HistoricalRecord {
  /** yyyy-MM-dd in the spreadsheet's time zone. Records are grouped into chart points by this. */
  day: string;
  /** Epoch ms of the snapshot. Used for ordering and to pick the newest snapshot of a day. */
  ts: number;
  /** ISO calendar date kept for backwards compatibility. */
  date: string;
  /** Human-readable date for the chart axis. */
  dateStr: string;
  sheet: string;
  /** True for the point computed right now from the period sheets (it is not stored anywhere). */
  live: boolean;
  /** Which sheet it came from: "live" (Historical Records), "archive" (Records_*) or "current" (computed now). */
  source: "live" | "archive" | "current";
  metrics: Record<string, number | null>;
}

interface HistoricalRecordPayload {
  headers: MetricHeader[];
  records: HistoricalRecord[];
  meta: {
    generatedAt: string;
    sheetsRead: string[];
    duplicatesRemoved: number;
    liveIncluded: boolean;
    liveError?: string;
  };
}

/** Turns header text into a stable JSON key. */
function sanitizeHistoricalKey(text: string, colIdx: number): string {
  const cleaned = text.replace(/[^a-zA-Z0-9]/g, '');
  return cleaned ? cleaned.charAt(0).toLowerCase() + cleaned.slice(1) : `col_${colIdx}`;
}

/** Reads one cell value as a timestamp (epoch ms). Accepts Date objects and date strings. */
function toTimestampMs(raw: any): number | null {
  if (raw instanceof Date) return isNaN(raw.getTime()) ? null : raw.getTime();
  if (raw === '' || raw === null || raw === undefined) return null;
  const ms = Date.parse(String(raw));
  return isNaN(ms) ? null : ms;
}

/** Formats an epoch-ms timestamp in the spreadsheet's time zone (falls back to UTC if Apps Script services are unavailable). */
function formatHistoricalDay(ms: number, tz: string): { day: string; dateStr: string } {
  const d = new Date(ms);
  try {
    return { day: Utilities.formatDate(d, tz, 'yyyy-MM-dd'), dateStr: Utilities.formatDate(d, tz, 'MMM d, yyyy') };
  } catch (e) {
    return { day: d.toISOString().split('T')[0], dateStr: d.toDateString() };
  }
}

/**
 * Returns the data for the Historical Analytics chart as a JSON string.
 *
 * Sources, in order:
 *  1. the live "Historical Records" sheet (always included by name, regardless of SHEET_NAME_PATTERN),
 *  2. every archived sheet whose name matches SHEET_NAME_PATTERN (e.g. "Records_2026-09-01T..."),
 *  3. a "current" snapshot computed right now from the period sheets using the same formulas the daily recorder uses, so the chart
 *     reaches today even if today's snapshot hasn't been recorded yet.
 *
 * Snapshots that appear on more than one sheet (an archive is a copy of the live sheet) are de-duplicated by timestamp + period.
 * @param data Optional JSON string: { includeLive?: boolean } (default true).
 */
function fetchHistoricalRecordbookData(data?: string): string {
  let includeLive = true;
  try {
    if (data && typeof data === 'string') {
      const parsed = JSON.parse(data);
      if (parsed && parsed.includeLive === false) includeLive = false;
    }
  } catch (e) { /* malformed payload: use defaults */ }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const liveSheetName = typeof DEFAULT_HISTORICAL_RECORDS_SHEET !== 'undefined' ? DEFAULT_HISTORICAL_RECORDS_SHEET : "Historical Records";
  const patternRegex = new RegExp(SHEET_NAME_PATTERN, 'i');

  // Live sheet first, then archives (matched by pattern) in tab order.
  const recordSheets = ss.getSheets().filter(sheet => sheet.getName() === liveSheetName || patternRegex.test(sheet.getName()));
  recordSheets.sort((a, b) => (a.getName() === liveSheetName ? -1 : 0) - (b.getName() === liveSheetName ? -1 : 0));

  const globalHeadersMap = new Map<string, MetricHeader>();
  const collected: HistoricalRecord[] = [];
  const seen = new Set<string>();
  let duplicatesRemoved = 0;
  let liveHeaderKeys: string[] | null = null;

  recordSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const isLiveSheet = sheetName === liveSheetName;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // Headers are read even when there are no data rows yet, so the live sheet still defines the metric columns for the current point.
    if (lastCol < METRIC_DATA_START_COL) return;

    const numHeaderCols = lastCol - (METRIC_DATA_START_COL - 1);
    const headerValues = sheet.getRange(HEADER_ROW_INDEX, METRIC_DATA_START_COL, 1, numHeaderCols).getValues()[0];
    const sheetHeaderKeys: string[] = [];

    headerValues.forEach((headerVal, idx) => {
      const colIndex = idx + METRIC_DATA_START_COL;
      const rawHeaderStr = String(headerVal || '').trim();

      if (rawHeaderStr) {
        const key = sanitizeHistoricalKey(rawHeaderStr, colIndex);
        sheetHeaderKeys[idx] = key;
        if (!globalHeadersMap.has(key)) {
          globalHeadersMap.set(key, { key, label: rawHeaderStr, colIndex });
        }
      } else {
        sheetHeaderKeys[idx] = `col_${colIndex}`;
      }
    });

    if (isLiveSheet) liveHeaderKeys = sheetHeaderKeys;
    if (lastRow < HISTORICAL_RECORDS_ROW_START) return;

    const numDataRows = lastRow - (HISTORICAL_RECORDS_ROW_START - 1);
    const dataMatrix = sheet.getRange(HISTORICAL_RECORDS_ROW_START, 1, numDataRows, lastCol).getValues();

    for (let i = 0; i < dataMatrix.length; i++) {
      const row = dataMatrix[i];
      const ts = toTimestampMs(row[0]);                       // Col A (Date)
      const sheetPeriod = String(row[PERIOD_COL_INDEX - 1] || '').trim(); // Col B (Sheet/Period)

      if (ts === null || !sheetPeriod) continue;

      // An archive is a copy of the live sheet, so the same snapshot can exist twice. Keep the first one seen (live wins).
      const dedupeKey = `${ts}|${sheetPeriod}`;
      if (seen.has(dedupeKey)) { duplicatesRemoved++; continue; }
      seen.add(dedupeKey);

      const metricsMap: Record<string, number | null> = {};
      for (let colIdx = METRIC_DATA_START_COL - 1; colIdx < row.length; colIdx++) {
        const metricKey = sheetHeaderKeys[colIdx - (METRIC_DATA_START_COL - 1)];
        if (metricKey) {
          const rawCellVal = row[colIdx];
          const numVal = Number(rawCellVal);
          // Assign null if cell is empty or non-numeric (handling missing/cutoff data)
          metricsMap[metricKey] = (rawCellVal !== '' && rawCellVal !== null && !isNaN(numVal)) ? numVal : null;
        }
      }

      const { day, dateStr } = formatHistoricalDay(ts, tz);
      collected.push({
        day, ts, date: day, dateStr, sheet: sheetPeriod, live: false,
        source: isLiveSheet ? "live" : "archive",
        metrics: metricsMap
      });
    }
  });

  // --- The current, not-yet-recorded data point ---
  let liveIncluded = false;
  let liveError: string | undefined;

  if (includeLive) {
    const liveSheet = ss.getSheetByName(liveSheetName);
    if (!liveSheet) {
      liveError = `Sheet "${liveSheetName}" not found, so the current values can't be computed.`;
    } else {
      const lock = LockService.getScriptLock();
      const gotLock = lock.tryLock(5000); // the evaluation briefly creates a scratch sheet, so don't overlap with the daily recorder
      try {
        if (!gotLock) throw new Error("Another process is updating records right now.");
        const evaluated = evaluatePeriodFormulas(ss, liveSheet);
        const keys: string[] = liveHeaderKeys || [];
        const nowMs = Date.now();
        const { day, dateStr } = formatHistoricalDay(nowMs, tz);

        evaluated.forEach(row => {
          const metricsMap: Record<string, number | null> = {};
          row.values.forEach((val, idx) => {
            const key = keys[idx];
            if (!key) return;
            const numVal = Number(val);
            metricsMap[key] = (val !== null && val !== '' && !isNaN(numVal)) ? numVal : null;
          });
          collected.push({ day, ts: nowMs, date: day, dateStr, sheet: row.name, live: true, source: "current", metrics: metricsMap });
        });
        liveIncluded = evaluated.length > 0;
      } catch (error: any) {
        liveError = `Current values could not be computed: ${error.message}`;
        Logger.log(`fetchHistoricalRecordbookData: ${liveError}`);
      } finally {
        if (gotLock) lock.releaseLock();
      }
    }
  }

  // Chronological, oldest first, no matter how many sheets the rows came from.
  collected.sort((a, b) => a.ts - b.ts);

  const payload: HistoricalRecordPayload = {
    headers: Array.from(globalHeadersMap.values()),
    records: collected,
    meta: {
      generatedAt: new Date().toISOString(),
      sheetsRead: recordSheets.map(s => s.getName()),
      duplicatesRemoved,
      liveIncluded,
      ...(liveError ? { liveError } : {})
    }
  };

  return JSON.stringify(payload);
}
