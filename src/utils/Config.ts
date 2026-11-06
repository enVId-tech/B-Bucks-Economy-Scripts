/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 *
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 *
 * TIER 1 CONFIGURATION: values that live in the Settings sheet ("Advanced Technical Settings") and are resolved lazily.
 *
 * Why this file exists: these values used to be computed at the top level of GlobalConsts.ts by calling
 * fetchProperty(). That ran during script load, before the constants the settings loader itself needs were
 * initialised (a circular dependency), and it returned strings/error objects instead of numbers.
 *
 * Rules:
 *  - Nothing in this file executes at load time except trivial literal declarations.
 *  - getConfig() is the ONLY way to read a sheet-driven value. It validates, coerces, memoises (per execution) and
 *    falls back to a default, so callers can never receive a string, undefined or an error object.
 *  - A re-entrancy guard makes accidental recursion (getConfig -> settings loader -> getConfig) return the default
 *    instead of looping forever.
 */

/**
 * Hard-coded fallbacks used when the Settings sheet is unreachable or a value is missing/invalid.
 * This is a FUNCTION on purpose: it reads Tier 0 constants from GlobalConsts.ts, and a top-level object literal here would
 * evaluate them while the files are still loading, in whatever order they happen to be pushed (Config.ts sorts before
 * GlobalConsts.ts). Reading them at call time is always safe.
 */
function getConfigDefault(key: string): number | string | undefined {
    switch (key) {
        case 'userStartingRow': return DEFAULT_USER_STARTING_ROW;
        default: return undefined;
    }
}

/** Validators: a resolved value is only accepted when its validator returns true. */
const ADVANCED_SETTING_VALIDATORS: { [key: string]: (value: any) => boolean } = {
    userStartingRow: (value: any) => Number.isInteger(value) && value >= 1,
};

// Per-execution memo. Apps Script re-evaluates the project for every server call, so this never goes stale across calls.
let configMemo: { [key: string]: any } = {};
let configResolving: boolean = false;

/**
 * Resolves a value from the "advancedTechnicalSettings" section of the Settings sheet.
 * @param key The setting name, e.g. "userStartingRow".
 * @returns The validated value, or the default when the sheet value is missing or invalid.
 */
function getConfig(key: string): any {
    if (Object.prototype.hasOwnProperty.call(configMemo, key)) return configMemo[key];

    const fallback = getConfigDefault(key);

    // Re-entrancy guard: if the settings loader (indirectly) asks for config while we are resolving, answer with the default
    // and do NOT memoise it, so the real value is still resolved on the next top-level call.
    if (configResolving) return fallback;

    configResolving = true;
    let resolved: any = fallback;
    try {
        const settings = fetchSettingsDataCached();
        if (settings && !('error' in settings)) {
            const raw = (settings.advancedTechnicalSettings as any)?.[key];
            const candidate = typeof fallback === 'number' ? coerceToNumber(raw) : (raw === undefined || raw === null || raw === '' ? undefined : String(raw));
            const validator = ADVANCED_SETTING_VALIDATORS[key];
            if (candidate !== undefined && (!validator || validator(candidate))) {
                resolved = candidate;
            }
        }
    } catch (error: any) {
        Logger.log(`getConfig("${key}") fell back to its default: ${error && error.message ? error.message : error}`);
    } finally {
        configResolving = false;
    }

    configMemo[key] = resolved;
    return resolved;
}

/** Converts sheet values ("7", 7, 7.0) to a number, returning undefined for blanks and non-numeric text. */
function coerceToNumber(raw: any): number | undefined {
    if (raw === undefined || raw === null) return undefined;
    if (typeof raw === 'string' && raw.trim() === '') return undefined;
    const parsed = Number(raw);
    return Number.isFinite(parsed) ? parsed : undefined;
}

/** Clears the per-execution memo. Call after writing to the Settings sheet so later reads in the same run see the new value. */
function resetConfigMemo(): void {
    configMemo = {};
}

/**
 * The first row that contains a student on every Period sheet. Always a positive integer (never a string).
 * Replaces the old USER_STARTING_ROW constant.
 */
function getUserStartingRow(): number {
    return getConfig('userStartingRow');
}
