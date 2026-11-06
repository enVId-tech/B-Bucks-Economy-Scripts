/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 *
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 *
 * ROLLING TRANSACTION CACHE
 *
 * The transaction log only ever grows, so caching "all transactions" as a single payload eventually blows through the
 * CacheService limits no matter how it is chunked. This module caches only a moving WINDOW of the log instead:
 *
 *   sheet rows ──► fixed-size BLOCKS (TX_BLOCK_SIZE rows, anchored at the OLDEST row) ──► CacheService
 *
 *  - Anchoring at the oldest row means appending a transaction only changes the newest (partial) block; every full block
 *    is immutable and stays valid forever.
 *  - Pages are shown newest-first. When the admin switches pages, planTxWindow() decides which blocks the page and its
 *    neighbours need (biased in the direction they are travelling), the missing ones are read from the sheet in as few
 *    range reads as possible, and the cache is then trimmed back to TX_CACHE_MAX_BLOCKS / TX_CACHE_MAX_BYTES by evicting the
 *    blocks FARTHEST from the current window first. Memory use is therefore bounded no matter how large the log grows.
 *  - The whole window is returned to the client, which keeps it in RAM (utils/TransactionWindow.html) so that the next
 *    page switch usually needs no server call at all. Think of it as RecyclerView-style lazy loading for the ledger.
 *
 * Searching/filtering can't use blocks (matches are scattered), so queryTransactions() scans the sheet in slices and returns
 * the same window shape. Totals for the stat tiles are maintained incrementally (only newly appended rows are ever scanned).
 */

// ----------------------------------------------------------------------------------------------------------------------
// Types
// ----------------------------------------------------------------------------------------------------------------------

type TxDirection = "forward" | "backward" | "jump";

interface TxFilters {
    /** Free-text search, matched against the fields selected by `scope`. */
    search?: string;
    scope?: "all" | "student" | "item" | "type" | "period" | "id" | "column";
    /** Exact student name (case-insensitive). */
    student?: string;
    /** Student-name substring (case-insensitive). */
    studentContains?: string;
    /** Period number as a string, e.g. "3". */
    period?: string;
    /** UI type label: "income" | "spending" | "deposit". */
    type?: string;
    /** Inclusive lower bound / inclusive upper bound as epoch ms (computed on the client in the user's time zone). */
    fromMs?: number | null;
    toMs?: number | null;
}

interface TxWindowPlan {
    total: number;
    totalPages: number;
    page: number;
    pageSize: number;
    firstPage: number;
    lastPage: number;
    /** Newest-first index of the first record in the window. */
    windowStart: number;
    /** Number of records in the window. */
    windowCount: number;
    requiredBlocks: number[];
    windowBlocks: number[];
}

interface TxStats {
    total: number;
    income: number;
    spending: number;
    /** Net balance change, in dollars. */
    net: number;
}

interface TxCacheIndexEntry { rows: number; bytes: number; used: number; }
interface TxCacheIndex { v: number; blockSize: number; clock: number; blocks: { [block: string]: TxCacheIndexEntry }; }

interface TxPageResponse {
    ok: true;
    total: number;
    totalPages: number;
    page: number;
    pageSize: number;
    /** The cached window, newest first. The client slices the requested page out of it. */
    window: { startIndex: number; records: TransactionRecord[] };
    stats?: TxStats;
    cache?: { blocksCached: number; bytes: number; maxBlocks: number; maxBytes: number; hits: number; misses: number; evicted: number; writable: boolean };
    filtered: boolean;
    serverMs: number;
}

// ----------------------------------------------------------------------------------------------------------------------
// Pure helpers (no Apps Script services)
// ----------------------------------------------------------------------------------------------------------------------

function normalizeTxText(value: any): string {
    return String(value === undefined || value === null ? '' : value).toLowerCase().trim();
}

/** Maps a transaction `type` to the label the UI filters use. */
function txTypeLabel(type: any): string {
    const mapping: { [key: string]: string } = { Income: 'income', Expense: 'spending', Investment: 'deposit' };
    return mapping[String(type)] || 'unknown';
}

/** Number of rows block `block` holds when the log has `total` rows. */
function txRowsInBlock(block: number, total: number, blockSize: number): number {
    return Math.max(0, Math.min(blockSize, total - block * blockSize));
}

/**
 * Works out which page, which neighbouring pages and which blocks a request needs.
 * @param total Total number of transaction rows.
 * @param requestedPage 1-based page number (newest transactions are page 1).
 * @param pageSize Records per page.
 * @param direction Which way the admin is paging. "forward" = towards older pages, so more pages are kept ahead (higher numbers)
 *   than behind; "backward" mirrors that; "jump" keeps a symmetric window.
 */
function planTxWindow(
    total: number,
    requestedPage: number,
    pageSize: number,
    direction: TxDirection = "jump",
    blockSize: number = TX_BLOCK_SIZE,
    maxBlocks: number = TX_CACHE_MAX_BLOCKS,
    behindPages: number = TX_PREFETCH_BEHIND_PAGES,
    aheadPages: number = TX_PREFETCH_AHEAD_PAGES
): TxWindowPlan {
    const size = Math.max(1, Math.min(TX_MAX_PAGE_SIZE, Math.floor(pageSize) || 25));
    const totalPages = Math.max(1, Math.ceil(total / size));
    const page = Math.min(totalPages, Math.max(1, Math.floor(requestedPage) || 1));

    let lower = direction === "backward" ? aheadPages : behindPages;  // pages with smaller numbers (newer records)
    let higher = direction === "forward" ? aheadPages : behindPages;  // pages with larger numbers (older records)

    // Near either end of the log one side has nothing to prefetch (e.g. nothing is newer than page 1). Give that side's
    // allowance to the other side instead of wasting it, so the window stays the same size.
    const availableLower = page - 1;
    const availableHigher = totalPages - page;
    if (lower > availableLower) { higher += lower - availableLower; lower = availableLower; }
    if (higher > availableHigher) { lower = Math.min(availableLower, lower + (higher - availableHigher)); higher = availableHigher; }

    // newest-first index range -> chronological row range -> block range
    const blocksFor = (firstPage: number, lastPage: number): number[] => {
        const startNewest = (firstPage - 1) * size;
        const endNewest = Math.min(total, lastPage * size) - 1;
        if (total === 0 || endNewest < startNewest) return [];
        const hiChron = total - 1 - startNewest;
        const loChron = total - 1 - endNewest;
        const out: number[] = [];
        for (let b = Math.floor(loChron / blockSize); b <= Math.floor(hiChron / blockSize); b++) out.push(b);
        return out;
    };

    const requiredBlocks = blocksFor(page, page);
    let firstPage = Math.max(1, page - lower);
    let lastPage = Math.min(totalPages, page + higher);
    let windowBlocks = blocksFor(firstPage, lastPage);

    // Respect the hard block ceiling: shrink the side that is farther from the page first (a page itself is never dropped).
    while (windowBlocks.length > maxBlocks && (firstPage < page || lastPage > page)) {
        if (page - firstPage >= lastPage - page && firstPage < page) firstPage++; else if (lastPage > page) lastPage--; else firstPage++;
        windowBlocks = blocksFor(firstPage, lastPage);
    }

    const windowStart = (firstPage - 1) * size;
    const windowCount = Math.max(0, Math.min(total, lastPage * size) - windowStart);

    return { total, totalPages, page, pageSize: size, firstPage, lastPage, windowStart, windowCount, requiredBlocks, windowBlocks };
}

/**
 * Chooses which cached blocks to evict so the cache fits its limits.
 * Blocks outside the window go first, farthest from `centerBlock` first (least recently used breaks ties). Window blocks are never evicted.
 * Blocks that no longer exist (log was shortened) are always evicted.
 */
function chooseTxEvictions(index: TxCacheIndex, windowBlocks: number[], centerBlock: number, totalBlocks: number, maxBlocks: number, maxBytes: number): number[] {
    const inWindow = new Set(windowBlocks);
    const evict: number[] = [];
    const entries = Object.keys(index.blocks).map(k => ({ block: Number(k), entry: index.blocks[k] }));

    const doomed = new Set<number>();
    entries.forEach(e => { if (e.block >= totalBlocks || e.block < 0) doomed.add(e.block); });
    let remaining = entries.filter(e => !doomed.has(e.block));
    let bytes = remaining.reduce((sum, e) => sum + e.entry.bytes, 0);
    let count = remaining.length;

    const candidates = remaining
        .filter(e => !inWindow.has(e.block))
        .sort((a, b) => (Math.abs(b.block - centerBlock) - Math.abs(a.block - centerBlock)) || (a.entry.used - b.entry.used));

    for (const c of candidates) {
        if (count <= maxBlocks && bytes <= maxBytes) break;
        evict.push(c.block);
        count--;
        bytes -= c.entry.bytes;
    }
    return Array.from(doomed).concat(evict);
}

/** The same matching rules the dialogs used to apply on the client, now applied on the server. */
function matchesTxFilters(tx: TransactionRecord, f: TxFilters | null | undefined): boolean {
    if (!f) return true;

    if (f.period && String(tx.period) !== String(f.period)) return false;
    if (f.student && normalizeTxText(tx.individual) !== normalizeTxText(f.student)) return false;
    if (f.studentContains && !normalizeTxText(tx.individual).includes(normalizeTxText(f.studentContains))) return false;
    if (f.type && txTypeLabel(tx.type) !== f.type) return false;

    if (f.fromMs !== undefined && f.fromMs !== null || f.toMs !== undefined && f.toMs !== null) {
        const ms = tx.timestamp ? Date.parse(String(tx.timestamp)) : NaN;
        if (!isNaN(ms)) { // rows without a readable timestamp are never excluded by a date filter (matches the old behaviour)
            if (f.fromMs !== undefined && f.fromMs !== null && ms < f.fromMs) return false;
            if (f.toMs !== undefined && f.toMs !== null && ms > f.toMs) return false;
        }
    }

    const query = normalizeTxText(f.search);
    if (query) {
        const fields: { [scope: string]: any[] } = {
            all: [tx.id, tx.individual, tx.period, tx.type, tx.serviceProvided, tx.unitPrice, tx.quantity, tx.modifiedColumn, tx.tenderedMoney, tx.timestamp],
            student: [tx.individual],
            item: [tx.serviceProvided],
            type: [tx.type],
            period: [tx.period],
            id: [tx.id],
            column: [tx.modifiedColumn],
        };
        const haystack = (fields[f.scope || 'all'] || fields.all).map(normalizeTxText).join(' ');
        if (!haystack.includes(query)) return false;
    }
    return true;
}

/** Dollar change a record made to the student's balance, rounded to cents (same maths as the stat tiles always used). */
function txBalanceDelta(tx: { initialBalance?: any; newBalance?: any }): number {
    const d = Number((Number(tx.newBalance) - Number(tx.initialBalance)).toFixed(2));
    return isNaN(d) ? 0 : d;
}

/** Turns one sheet row (A..P) into a TransactionRecord. Dates are normalised to ISO strings so they survive serialisation. */
function mapTransactionRow(row: any[]): TransactionRecord {
    const ts = row[13];
    return {
        id: row[0] as number,
        individual: row[1] as string,
        period: row[2] as number,
        type: (row[3] || "Unknown") as TransactionType,
        serviceProvided: row[4] as string,
        unitPrice: row[5] as number,
        quantity: row[6] as number,
        modifiedColumn: row[7] as number,
        tenderedMoney: row[8] as number,
        initialColumnAmount: row[9] as number,
        newColumnAmount: row[10] as number,
        initialBalance: row[11] as number,
        newBalance: row[12] as number,
        timestamp: ts instanceof Date ? ts.toISOString() : (ts as string),
        row: row[14] ? Number(row[14]) : undefined,
        operation: (row[15] || undefined) as TransactionRecord["operation"],
    } as TransactionRecord;
}

// ----------------------------------------------------------------------------------------------------------------------
// Sheet access
// ----------------------------------------------------------------------------------------------------------------------

function getTransactionsSheet(): GoogleAppsScript.Spreadsheet.Sheet | null {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DEFAULT_TRANSACTIONS_SHEET);
}

/** Number of transaction rows currently on the sheet (the header rows above TRANSACTIONS_ROW_START are excluded). */
function countTransactionRows(sheet: GoogleAppsScript.Spreadsheet.Sheet): number {
    return Math.max(0, sheet.getLastRow() - TRANSACTIONS_ROW_START + 1);
}

/** Reads `count` rows starting at chronological index `firstIndex` (0 = oldest) in a single range read. */
function readTransactionRows(sheet: GoogleAppsScript.Spreadsheet.Sheet, firstIndex: number, count: number): TransactionRecord[] {
    if (count <= 0) return [];
    return sheet.getRange(TRANSACTIONS_ROW_START + firstIndex, 1, count, TX_COLUMN_COUNT).getValues().map(mapTransactionRow);
}

// ----------------------------------------------------------------------------------------------------------------------
// Cache index
// ----------------------------------------------------------------------------------------------------------------------

function emptyTxIndex(): TxCacheIndex {
    return { v: 1, blockSize: TX_BLOCK_SIZE, clock: 0, blocks: {} };
}

function readTxIndex(cache: GoogleAppsScript.Cache.Cache): TxCacheIndex {
    try {
        const raw = cache.get(TX_INDEX_KEY);
        if (!raw) return emptyTxIndex();
        const parsed = JSON.parse(raw) as TxCacheIndex;
        // If TX_BLOCK_SIZE was changed, every cached block has the wrong shape: start over.
        if (!parsed || parsed.v !== 1 || parsed.blockSize !== TX_BLOCK_SIZE || typeof parsed.blocks !== 'object') {
            if (parsed && parsed.blocks) cache.removeAll(Object.keys(parsed.blocks).map(b => TX_BLOCK_KEY_PREFIX + b));
            return emptyTxIndex();
        }
        return parsed;
    } catch (e) {
        return emptyTxIndex();
    }
}

function writeTxIndex(cache: GoogleAppsScript.Cache.Cache, index: TxCacheIndex): void {
    cache.put(TX_INDEX_KEY, JSON.stringify(index), SERVER_SIDE_CACHE_AGE);
}

/** Removes every cached transaction block, the bookkeeping, and the derived facets/stats. Call after bulk edits or a reset. */
function invalidateTransactionCache(): boolean {
    try {
        const cache = CacheService.getScriptCache();
        const index = readTxIndex(cache);
        const keys = Object.keys(index.blocks).map(b => TX_BLOCK_KEY_PREFIX + b);
        cache.removeAll(keys.concat([TX_INDEX_KEY, TX_FACETS_KEY, TX_STATS_KEY]));
        return true;
    } catch (error: any) {
        log(`Failed to invalidate the transaction cache: ${error.message}`, false);
        return false;
    }
}

/** Called after rows are appended. Blocks self-validate by row count, so only the derived facets need dropping. */
function onTransactionsAppended(): void {
    try { CacheService.getScriptCache().remove(TX_FACETS_KEY); } catch (e) { /* non-critical */ }
}

// ----------------------------------------------------------------------------------------------------------------------
// Public endpoints
// ----------------------------------------------------------------------------------------------------------------------

function parseTxPayload(data?: string): any {
    if (!data || typeof data !== 'string') return {};
    try { return JSON.parse(data) || {}; } catch (e) { return {}; }
}

/**
 * Returns the requested page of the (unfiltered) transaction log together with a rolling window of neighbouring pages, and
 * re-centres the server cache on that window.
 * @param data JSON: { page?: number, pageSize?: number, direction?: "forward"|"backward"|"jump", forceRefresh?: boolean, includeStats?: boolean }
 */
function fetchTransactionsPage(data?: string): TxPageResponse | { error: string } {
    const startedAt = Date.now();
    try {
        const payload = parseTxPayload(data);
        const sheet = getTransactionsSheet();
        if (!sheet) return { error: `Transactions sheet "${DEFAULT_TRANSACTIONS_SHEET}" not found.` };

        const cache = CacheService.getScriptCache();
        const total = countTransactionRows(sheet);
        const plan = planTxWindow(total, Number(payload.page) || 1, Number(payload.pageSize) || 25, (payload.direction as TxDirection) || "jump");
        const forceRefresh = payload.forceRefresh === true;

        // Mutating the shared cache index is serialised with a lock. If we can't get it we still answer from the sheet,
        // we just don't touch the cache (never block the admin on housekeeping).
        const lock = LockService.getScriptLock();
        const writable = lock.tryLock(2000);
        let hits = 0, misses = 0, evicted = 0;

        const blockData: { [block: number]: TransactionRecord[] } = {};

        try {
            let index = readTxIndex(cache);
            if (forceRefresh && writable) {
                cache.removeAll(Object.keys(index.blocks).map(b => TX_BLOCK_KEY_PREFIX + b));
                index = emptyTxIndex();
            }

            // 1. Satisfy what we can from the cache.
            const cachedRaw = forceRefresh ? {} : cache.getAll(plan.windowBlocks.map(b => TX_BLOCK_KEY_PREFIX + b));
            const missing: number[] = [];
            plan.windowBlocks.forEach(b => {
                const entry = index.blocks[String(b)];
                const raw = (cachedRaw as { [k: string]: string })[TX_BLOCK_KEY_PREFIX + b];
                if (entry && raw && entry.rows === txRowsInBlock(b, total, TX_BLOCK_SIZE)) {
                    try { blockData[b] = JSON.parse(raw); hits++; return; } catch (e) { /* corrupt entry: refetch */ }
                }
                missing.push(b);
            });

            // 2. Read missing blocks from the sheet, merging adjacent blocks into one range read.
            let runStart = 0;
            while (runStart < missing.length) {
                let runEnd = runStart;
                while (runEnd + 1 < missing.length && missing[runEnd + 1] === missing[runEnd] + 1) runEnd++;
                const firstBlock = missing[runStart], lastBlock = missing[runEnd];
                const firstIndex = firstBlock * TX_BLOCK_SIZE;
                const lastIndexExclusive = Math.min(total, (lastBlock + 1) * TX_BLOCK_SIZE);
                const rows = readTransactionRows(sheet, firstIndex, lastIndexExclusive - firstIndex);
                for (let b = firstBlock; b <= lastBlock; b++) {
                    blockData[b] = rows.slice((b - firstBlock) * TX_BLOCK_SIZE, (b - firstBlock + 1) * TX_BLOCK_SIZE);
                    misses++;
                }
                runStart = runEnd + 1;
            }

            // 3. Store the new blocks, then roll the cache: evict whatever is farthest from the window until it fits.
            if (writable) {
                const entries: { [key: string]: string } = {};
                missing.forEach(b => {
                    const serialized = JSON.stringify(blockData[b]);
                    entries[TX_BLOCK_KEY_PREFIX + b] = serialized;
                    index.blocks[String(b)] = { rows: blockData[b].length, bytes: serialized.length, used: 0 };
                });
                if (Object.keys(entries).length > 0) {
                    try { cache.putAll(entries, SERVER_SIDE_CACHE_AGE); }
                    catch (e: any) { // e.g. a block exceeded the per-value limit: serve the data, just don't cache it
                        missing.forEach(b => { delete index.blocks[String(b)]; });
                        log(`Transaction cache write skipped: ${e.message}`, false);
                    }
                }

                index.clock++;
                plan.windowBlocks.forEach(b => { if (index.blocks[String(b)]) index.blocks[String(b)].used = index.clock; });

                const totalBlocks = Math.ceil(total / TX_BLOCK_SIZE);
                const centerBlock = plan.requiredBlocks.length ? plan.requiredBlocks[Math.floor(plan.requiredBlocks.length / 2)] : 0;
                const toEvict = chooseTxEvictions(index, plan.windowBlocks, centerBlock, totalBlocks, TX_CACHE_MAX_BLOCKS, TX_CACHE_MAX_BYTES);
                if (toEvict.length > 0) {
                    cache.removeAll(toEvict.map(b => TX_BLOCK_KEY_PREFIX + b));
                    toEvict.forEach(b => { delete index.blocks[String(b)]; });
                    evicted = toEvict.length;
                }
                writeTxIndex(cache, index);
            }

            // 4. Assemble the window, newest first.
            const windowRecords: TransactionRecord[] = [];
            for (let newestIdx = plan.windowStart; newestIdx < plan.windowStart + plan.windowCount; newestIdx++) {
                const chron = total - 1 - newestIdx;
                const block = Math.floor(chron / TX_BLOCK_SIZE);
                const record = blockData[block] ? blockData[block][chron - block * TX_BLOCK_SIZE] : undefined;
                if (record) windowRecords.push(record);
            }

            const response: TxPageResponse = {
                ok: true,
                total,
                totalPages: plan.totalPages,
                page: plan.page,
                pageSize: plan.pageSize,
                window: { startIndex: plan.windowStart, records: windowRecords },
                filtered: false,
                serverMs: 0,
            };
            if (payload.includeStats) response.stats = getUnfilteredTxStats(sheet, total, cache, writable);

            const finalIndex = writable ? index : readTxIndex(cache);
            const keys = Object.keys(finalIndex.blocks);
            response.cache = {
                blocksCached: keys.length,
                bytes: keys.reduce((sum, k) => sum + finalIndex.blocks[k].bytes, 0),
                maxBlocks: TX_CACHE_MAX_BLOCKS,
                maxBytes: TX_CACHE_MAX_BYTES,
                hits, misses, evicted, writable,
            };
            response.serverMs = Date.now() - startedAt;
            return response;
        } finally {
            if (writable) lock.releaseLock();
        }
    } catch (error: any) {
        log(`Error in fetchTransactionsPage: ${error.message}`, false);
        return { error: `Error in fetchTransactionsPage: ${error.message}` };
    }
}

/**
 * Incrementally maintained totals for the whole log. Only rows appended since the last call are scanned.
 * Net is accumulated in whole cents to avoid floating point drift.
 */
function getUnfilteredTxStats(sheet: GoogleAppsScript.Spreadsheet.Sheet, total: number, cache: GoogleAppsScript.Cache.Cache, writable: boolean): TxStats {
    let state = { rows: 0, income: 0, spending: 0, netCents: 0 };
    try {
        const raw = cache.get(TX_STATS_KEY);
        if (raw) {
            const parsed = JSON.parse(raw);
            if (parsed && parsed.rows <= total) state = parsed;
        }
    } catch (e) { /* recompute from scratch */ }

    if (state.rows < total) {
        // Columns D..M: type (D) and the balance columns (L, M)
        const values = sheet.getRange(TRANSACTIONS_ROW_START + state.rows, 4, total - state.rows, 10).getValues();
        values.forEach(r => {
            if (r[0] === 'Income') state.income++;
            if (r[0] === 'Expense') state.spending++;
            state.netCents += Math.round(txBalanceDelta({ initialBalance: r[8], newBalance: r[9] }) * 100);
        });
        state.rows = total;
        if (writable) {
            try { cache.put(TX_STATS_KEY, JSON.stringify(state), SERVER_SIDE_CACHE_AGE); } catch (e) { /* non-critical */ }
        }
    }
    return { total, income: state.income, spending: state.spending, net: state.netCents / 100 };
}

/**
 * Filtered / searched view of the log. Matches are scattered through the sheet so this scans it newest-to-oldest in slices
 * (never holding the whole sheet in memory at once), counts every match, and returns the requested page plus its neighbours.
 * @param data JSON: { filters?: TxFilters, page?: number, pageSize?: number, direction?: TxDirection }
 */
function queryTransactions(data?: string): TxPageResponse | { error: string } {
    const startedAt = Date.now();
    try {
        const payload = parseTxPayload(data);
        const filters: TxFilters = payload.filters || {};
        const sheet = getTransactionsSheet();
        if (!sheet) return { error: `Transactions sheet "${DEFAULT_TRANSACTIONS_SHEET}" not found.` };

        const total = countTransactionRows(sheet);
        const SLICE = 1000;
        const size = Math.max(1, Math.min(TX_MAX_PAGE_SIZE, Math.floor(Number(payload.pageSize)) || 25));
        const direction = (payload.direction as TxDirection) || "jump";

        let page = Math.max(1, Math.floor(Number(payload.page)) || 1);
        let attempt = 0;

        while (true) {
            attempt++;
            const lower = direction === "backward" ? TX_PREFETCH_AHEAD_PAGES : TX_PREFETCH_BEHIND_PAGES;
            const higher = direction === "forward" ? TX_PREFETCH_AHEAD_PAGES : TX_PREFETCH_BEHIND_PAGES;
            const wantStart = Math.max(0, (page - 1 - lower) * size);
            const wantEnd = (page + higher) * size; // exclusive

            const windowRecords: TransactionRecord[] = [];
            const stats = { total: 0, income: 0, spending: 0, netCents: 0 };

            for (let hi = total; hi > 0; hi -= SLICE) {
                const lo = Math.max(0, hi - SLICE);
                const rows = readTransactionRows(sheet, lo, hi - lo);
                for (let i = rows.length - 1; i >= 0; i--) { // newest first
                    const tx = rows[i];
                    if (!matchesTxFilters(tx, filters)) continue;
                    if (stats.total >= wantStart && stats.total < wantEnd) windowRecords.push(tx);
                    stats.total++;
                    if (tx.type === 'Income') stats.income++;
                    if (tx.type === 'Expense') stats.spending++;
                    stats.netCents += Math.round(txBalanceDelta(tx) * 100);
                }
            }

            const totalPages = Math.max(1, Math.ceil(stats.total / size));
            if (page > totalPages && attempt < 2) { page = totalPages; continue; } // requested page is past the end: re-collect for the last page

            return {
                ok: true,
                total: stats.total,
                totalPages,
                page: Math.min(page, totalPages),
                pageSize: size,
                window: { startIndex: wantStart, records: windowRecords },
                stats: { total: stats.total, income: stats.income, spending: stats.spending, net: stats.netCents / 100 },
                filtered: true,
                serverMs: Date.now() - startedAt,
            };
        }
    } catch (error: any) {
        log(`Error in queryTransactions: ${error.message}`, false);
        return { error: `Error in queryTransactions: ${error.message}` };
    }
}

/**
 * Distinct values for the filter drop-downs in the Transaction Records dialog, without shipping the whole log to the browser.
 * @param data JSON: { fromMs?: number, toMs?: number } restricts the facets to a date range.
 * @returns { periods: number[], types: string[], studentsByPeriod: { [period: string]: string[] }, allStudents: string[] }
 */
function fetchTransactionFacets(data?: string): any {
    try {
        const payload = parseTxPayload(data);
        const hasRange = (payload.fromMs !== undefined && payload.fromMs !== null) || (payload.toMs !== undefined && payload.toMs !== null);
        const sheet = getTransactionsSheet();
        if (!sheet) return { error: `Transactions sheet "${DEFAULT_TRANSACTIONS_SHEET}" not found.` };

        const cache = CacheService.getScriptCache();
        const total = countTransactionRows(sheet);

        if (!hasRange) {
            try {
                const raw = cache.get(TX_FACETS_KEY);
                if (raw) {
                    const cached = JSON.parse(raw);
                    if (cached && cached.rows === total) return cached.facets;
                }
            } catch (e) { /* recompute */ }
        }

        const periods = new Set<number>();
        const types = new Set<string>();
        const students = new Set<string>();
        const byPeriod: { [period: string]: Set<string> } = {};

        if (total > 0) {
            // Columns B..D (student, period, type) and N (timestamp) are all that is needed.
            const core = sheet.getRange(TRANSACTIONS_ROW_START, 2, total, 3).getValues();
            const stamps = hasRange ? sheet.getRange(TRANSACTIONS_ROW_START, 14, total, 1).getValues() : null;

            for (let i = 0; i < core.length; i++) {
                if (stamps) {
                    const raw = stamps[i][0];
                    const ms = raw instanceof Date ? raw.getTime() : (raw ? Date.parse(String(raw)) : NaN);
                    if (!isNaN(ms)) {
                        if (payload.fromMs !== undefined && payload.fromMs !== null && ms < payload.fromMs) continue;
                        if (payload.toMs !== undefined && payload.toMs !== null && ms > payload.toMs) continue;
                    }
                }
                const student = String(core[i][0] || '').trim();
                const period = core[i][1];
                const type = core[i][2];
                if (period !== '' && period !== null && period !== undefined) periods.add(Number(period));
                if (type) types.add(String(type));
                if (student) {
                    students.add(student);
                    const key = String(period);
                    (byPeriod[key] = byPeriod[key] || new Set<string>()).add(student);
                }
            }
        }

        const sortedStudents = (set: Set<string>) => Array.from(set).sort((a, b) => a.localeCompare(b));
        const facets = {
            periods: Array.from(periods).sort((a, b) => a - b),
            types: Array.from(types),
            allStudents: sortedStudents(students),
            studentsByPeriod: Object.keys(byPeriod).reduce((acc: { [k: string]: string[] }, k) => { acc[k] = sortedStudents(byPeriod[k]); return acc; }, {}),
        };

        if (!hasRange) {
            try { cache.put(TX_FACETS_KEY, JSON.stringify({ rows: total, facets }), SERVER_SIDE_CACHE_AGE); } catch (e) { /* too large to cache: fine */ }
        }
        return facets;
    } catch (error: any) {
        log(`Error in fetchTransactionFacets: ${error.message}`, false);
        return { error: `Error in fetchTransactionFacets: ${error.message}` };
    }
}

/**
 * Looks up specific transactions by id WITHOUT loading the log (used by undo). Only column A is scanned to find the rows.
 */
function fetchTransactionsByIds(ids: number[]): TransactionRecord[] {
    const sheet = getTransactionsSheet();
    if (!sheet || !ids || ids.length === 0) return [];

    const total = countTransactionRows(sheet);
    if (total === 0) return [];

    const wanted = new Set(ids.map(Number));
    const idColumn = sheet.getRange(TRANSACTIONS_ROW_START, 1, total, 1).getValues();
    const found: TransactionRecord[] = [];

    for (let i = 0; i < idColumn.length; i++) {
        if (wanted.has(Number(idColumn[i][0]))) {
            found.push(readTransactionRows(sheet, i, 1)[0]);
        }
    }
    return found;
}

/** Warms the cache with the newest page so the first dialog open is instant. */
function preloadTransactionCache(): boolean {
    const result = fetchTransactionsPage(JSON.stringify({ page: 1, pageSize: 25, forceRefresh: true, includeStats: true }));
    return !('error' in result);
}

/** Diagnostics for developers: what the rolling cache currently holds. Run from the Apps Script editor. */
function getTransactionCacheStatus(): any {
    const cache = CacheService.getScriptCache();
    const index = readTxIndex(cache);
    const keys = Object.keys(index.blocks).map(Number).sort((a, b) => a - b);
    return {
        blockSize: index.blockSize,
        blocksCached: keys.length,
        blocks: keys,
        bytes: keys.reduce((sum, k) => sum + index.blocks[String(k)].bytes, 0),
        limits: { maxBlocks: TX_CACHE_MAX_BLOCKS, maxBytes: TX_CACHE_MAX_BYTES },
    };
}
