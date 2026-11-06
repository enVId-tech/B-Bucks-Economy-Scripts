/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran.
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 *
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 *
 * TIER 1 CONFIGURATION: every value in the "Advanced Technical Settings" block of the Settings sheet (column layout, row starts,
 * sheet names, cell addresses...). They are resolved lazily and read through the accessor functions at the bottom of this file
 * (getNamesColumn(), getServicesSheetName(), ...). Never read the DEFAULT_* constants directly: they are only the fallbacks.
 *
 * Why this file exists: these values used to be computed at the top level of GlobalConsts.ts by calling fetchProperty().
 * That ran during script load, before the constants the settings loader itself needs were initialised (a circular dependency),
 * and it returned strings/error objects instead of numbers.
 *
 * Rules:
 *  - Nothing in this file executes at load time except trivial literal declarations.
 *  - getConfig() is the ONLY way to read a sheet-driven value. It validates, coerces, memoises (per execution) and falls back to a
 *    default, so callers can never receive a string where a number is expected, undefined or an error object.
 *  - A re-entrancy guard makes accidental recursion (getConfig -> settings loader -> getConfig) return the default instead of looping.
 *  - defaultSettingsSheet and settingsRowStart are BOOTSTRAP values (the loader needs them to find the Settings sheet), so the code
 *    always uses its built-in values for those two. verifyConstantsAreReadable() reports when the sheet disagrees.
 */

type AdvancedSettingKind = 'positiveInteger' | 'cellAddress' | 'sheetName' | 'text';

interface AdvancedSettingDefinition {
    /** The key as written in the Settings sheet (middle column of "Advanced Technical Settings"). */
    settingKey: string;
    defaultValue: number | string;
    kind: AdvancedSettingKind;
}

/** How a setting was resolved; shown by verifyConstantsAreReadable(). */
interface ResolvedAdvancedSetting {
    settingKey: string;
    value: number | string;
    source: 'settings sheet' | 'default';
    /** Why the default was used, empty when the sheet value was accepted. */
    reason: string;
}

let advancedSettingDefinitionsMemo: { [settingKey: string]: AdvancedSettingDefinition } | null = null;

/**
 * Every Advanced Technical Setting with its fallback and the kind of value it must hold.
 * This is a FUNCTION on purpose: it reads Tier 0 constants from GlobalConsts.ts, and a top-level object literal here would evaluate
 * them while the files are still loading, in whatever order they happen to be pushed (Config.ts sorts before GlobalConsts.ts).
 * Reading them at call time is always safe.
 */
function getAdvancedSettingDefinitions(): { [settingKey: string]: AdvancedSettingDefinition } {
    if (advancedSettingDefinitionsMemo) return advancedSettingDefinitionsMemo;

    advancedSettingDefinitionsMemo = {
        userStartingRow: { settingKey: 'userStartingRow', defaultValue: DEFAULT_USER_STARTING_ROW, kind: 'positiveInteger' },
        namesCol: { settingKey: 'namesCol', defaultValue: DEFAULT_NAMES_COL, kind: 'positiveInteger' },
        balanceCol: { settingKey: 'balanceCol', defaultValue: DEFAULT_BALANCE_COL, kind: 'positiveInteger' },
        earningsCol: { settingKey: 'earningsCol', defaultValue: DEFAULT_EARNINGS_COL, kind: 'positiveInteger' },
        netIncomeCol: { settingKey: 'netIncomeCol', defaultValue: DEFAULT_NET_INCOME_COL, kind: 'positiveInteger' },
        expendituresCol: { settingKey: 'expendituresCol', defaultValue: DEFAULT_EXPENDITURES_COL, kind: 'positiveInteger' },
        investmentReturnsCol: { settingKey: 'investmentReturnsCol', defaultValue: DEFAULT_INVESTMENT_RETURNS_COL, kind: 'positiveInteger' },
        initialDepositCol: { settingKey: 'initialDepositCol', defaultValue: DEFAULT_INITIAL_DEPOSIT_COL, kind: 'positiveInteger' },
        timeDepositCol: { settingKey: 'timeDepositCol', defaultValue: DEFAULT_DATE_DEPOSIT_COL, kind: 'positiveInteger' },
        grossInvestmentGainCol: { settingKey: 'grossInvestmentGainCol', defaultValue: DEFAULT_GROSS_INVESTMENT_GAIN_COL, kind: 'positiveInteger' },
        netInvestmentGainCol: { settingKey: 'netInvestmentGainCol', defaultValue: DEFAULT_NET_INVESTMENT_GAIN_COL, kind: 'positiveInteger' },
        netPercentageGainCol: { settingKey: 'netPercentageGainCol', defaultValue: DEFAULT_NET_PERCENTAGE_GAIN_COL, kind: 'positiveInteger' },
        defaultServicesSheet: { settingKey: 'defaultServicesSheet', defaultValue: DEFAULT_SERVICES_SHEET, kind: 'sheetName' },
        servicesRowStart: { settingKey: 'servicesRowStart', defaultValue: DEFAULT_SERVICES_ROW_START, kind: 'positiveInteger' },
        defaultTransactionsSheet: { settingKey: 'defaultTransactionsSheet', defaultValue: DEFAULT_TRANSACTIONS_SHEET, kind: 'sheetName' },
        transactionsRowStart: { settingKey: 'transactionsRowStart', defaultValue: DEFAULT_TRANSACTIONS_ROW_START, kind: 'positiveInteger' },
        fillStartingCell: { settingKey: 'fillStartingCell', defaultValue: DEFAULT_STARTING_CELL, kind: 'cellAddress' },
        fillSheetCellName: { settingKey: 'fillSheetCellName', defaultValue: DEFAULT_SHEET_CELL_NAME, kind: 'cellAddress' },
        defaultHistoricalRecordsSheet: { settingKey: 'defaultHistoricalRecordsSheet', defaultValue: DEFAULT_HISTORICAL_RECORDS_SHEET, kind: 'sheetName' },
        historicalRecordsRowStart: { settingKey: 'historicalRecordsRowStart', defaultValue: DEFAULT_HISTORICAL_RECORDS_ROW_START, kind: 'positiveInteger' },
        metricDataStartColumn: { settingKey: 'metricDataStartColumn', defaultValue: DEFAULT_METRIC_DATA_START_COL, kind: 'positiveInteger' },
        periodColIndex: { settingKey: 'periodColIndex', defaultValue: DEFAULT_PERIOD_COL_INDEX, kind: 'positiveInteger' },
        headerRowIndex: { settingKey: 'headerRowIndex', defaultValue: DEFAULT_HEADER_ROW_INDEX, kind: 'positiveInteger' },
        sheetNamePattern: { settingKey: 'sheetNamePattern', defaultValue: DEFAULT_SHEET_NAME_PATTERN, kind: 'text' },
        timestampCell: { settingKey: 'timestampCell', defaultValue: DEFAULT_TIMESTAMP_CELL, kind: 'cellAddress' },
    };
    return advancedSettingDefinitionsMemo;
}

// Per-execution memo. Apps Script re-evaluates the project for every server call, so this never goes stale across calls.
let configMemo: { [settingKey: string]: ResolvedAdvancedSetting } = {};
let configResolving: boolean = false;

/** Converts sheet values ("7", 7, 7.0) to a number, returning undefined for blanks and non-numeric text. */
function coerceToNumber(rawValue: any): number | undefined {
    if (rawValue === undefined || rawValue === null) return undefined;
    if (typeof rawValue === 'string' && rawValue.trim() === '') return undefined;
    const parsedNumber = Number(rawValue);
    return Number.isFinite(parsedNumber) ? parsedNumber : undefined;
}

/**
 * Checks one raw sheet value against the kind of value the setting must hold.
 * @returns The cleaned value, or an object explaining why it was rejected.
 */
function validateAdvancedSetting(definition: AdvancedSettingDefinition, rawValue: any): { value: number | string } | { rejectedBecause: string } {
    const isBlank = rawValue === undefined || rawValue === null || (typeof rawValue === 'string' && rawValue.trim() === '');
    if (isBlank) return { rejectedBecause: 'the setting is blank or missing in the Settings sheet' };

    // The settings loader turns an empty cell into 0, so for text-like settings a bare 0 means "left blank".
    if (definition.kind !== 'positiveInteger' && rawValue === 0) {
        return { rejectedBecause: 'the setting is blank in the Settings sheet' };
    }

    switch (definition.kind) {
        case 'positiveInteger': {
            const numericValue = coerceToNumber(rawValue);
            if (numericValue === undefined || !Number.isInteger(numericValue) || numericValue < 1) {
                return { rejectedBecause: `"${rawValue}" is not a whole number of 1 or more` };
            }
            return { value: numericValue };
        }
        case 'cellAddress': {
            const cellAddress = String(rawValue).trim().toUpperCase();
            if (!/^[A-Z]{1,3}[1-9][0-9]*$/.test(cellAddress)) {
                return { rejectedBecause: `"${rawValue}" is not a cell address like A7` };
            }
            return { value: cellAddress };
        }
        case 'sheetName': {
            const sheetName = String(rawValue).trim();
            if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)) {
                return { rejectedBecause: `there is no sheet tab named "${sheetName}"` };
            }
            return { value: sheetName };
        }
        default:
            return { value: String(rawValue).trim() };
    }
}

/**
 * Works out the value of one Advanced Technical Setting and explains where it came from.
 * Never throws and never returns undefined.
 */
function resolveAdvancedSetting(settingKey: string): ResolvedAdvancedSetting {
    const definition = getAdvancedSettingDefinitions()[settingKey];
    if (!definition) {
        throw new Error(`Unknown advanced setting "${settingKey}". Add it to getAdvancedSettingDefinitions() in Config.ts.`);
    }

    const useDefault = (reason: string): ResolvedAdvancedSetting =>
        ({ settingKey, value: definition.defaultValue, source: 'default', reason });

    try {
        const settings = fetchSettingsDataCached();
        if (!settings || 'error' in settings) {
            return useDefault('the Settings sheet could not be read');
        }

        const rawValue = (settings.advancedTechnicalSettings as any)?.[settingKey];
        const validation = validateAdvancedSetting(definition, rawValue);
        if ('rejectedBecause' in validation) {
            return useDefault(validation.rejectedBecause);
        }
        return { settingKey, value: validation.value, source: 'settings sheet', reason: '' };
    } catch (error: any) {
        return useDefault(`reading it failed: ${error && error.message ? error.message : error}`);
    }
}

/**
 * Resolves a value from the "advancedTechnicalSettings" section of the Settings sheet.
 * Prefer the named accessors below (getNamesColumn(), ...) over calling this directly.
 * @param settingKey The setting name as written in the sheet, e.g. "namesCol".
 * @returns The validated value, or the default when the sheet value is missing or invalid.
 */
function getConfig(settingKey: string): any {
    const memoizedSetting = configMemo[settingKey];
    if (memoizedSetting) return memoizedSetting.value;

    // Re-entrancy guard: if the settings loader (indirectly) asks for config while we are resolving, answer with the default
    // and do NOT memoise it, so the real value is still resolved on the next top-level call.
    if (configResolving) return getAdvancedSettingDefinitions()[settingKey].defaultValue;

    configResolving = true;
    let resolvedSetting: ResolvedAdvancedSetting;
    try {
        resolvedSetting = resolveAdvancedSetting(settingKey);
    } finally {
        configResolving = false;
    }

    if (resolvedSetting.source === 'default' && resolvedSetting.reason) {
        log(`Setting "${settingKey}" is using its default (${resolvedSetting.value}) because ${resolvedSetting.reason}.`, false);
    }

    configMemo[settingKey] = resolvedSetting;
    return resolvedSetting.value;
}

/** Clears the per-execution memo. Call after writing to the Settings sheet so later reads in the same run see the new value. */
function resetConfigMemo(): void {
    configMemo = {};
}

/**
 * Compares the two bootstrap values the Settings sheet also lists with the values the code really uses.
 * @returns One sentence per disagreement (empty when they match).
 */
function findBootstrapSettingMismatches(): string[] {
    const mismatches: string[] = [];
    const settings = fetchSettingsDataCached();
    if (!settings || 'error' in settings) return mismatches;

    const sheetValues = settings.advancedTechnicalSettings as any;
    if (sheetValues?.defaultSettingsSheet !== undefined && String(sheetValues.defaultSettingsSheet).trim() !== DEFAULT_SETTINGS_SHEET) {
        mismatches.push(`defaultSettingsSheet says "${sheetValues.defaultSettingsSheet}" but the code must use "${DEFAULT_SETTINGS_SHEET}" (it is needed to find the Settings sheet).`);
    }
    if (sheetValues?.settingsRowStart !== undefined && Number(sheetValues.settingsRowStart) !== SETTINGS_ROW_START) {
        mismatches.push(`settingsRowStart says ${sheetValues.settingsRowStart} but the code must use ${SETTINGS_ROW_START} (it is needed to read the Settings sheet).`);
    }
    return mismatches;
}

/**
 * Diagnostic: reads EVERY advanced setting and every bootstrap constant, and tells the admin which ones come from the Settings
 * sheet, which fell back to a default and why. Run it from the script editor, or call it from a menu item, when "a constant can't be read".
 * @returns The same report that is logged and shown.
 */
function verifyConstantsAreReadable(): string {
    resetConfigMemo();
    const reportLines: string[] = [];
    let problemCount = 0;

    const settings = fetchSettingsDataCached(JSON.stringify({ forceRefresh: true }));
    if (!settings || 'error' in settings) {
        const reason = settings && 'error' in settings ? settings.error : 'no data returned';
        const failureReport = `The Settings sheet could not be read: ${reason}. Every setting is using its built-in default.`;
        log(failureReport, true);
        return failureReport;
    }

    Object.keys(getAdvancedSettingDefinitions()).forEach(settingKey => {
        const resolvedSetting = resolveAdvancedSetting(settingKey);
        if (resolvedSetting.source === 'default') {
            problemCount++;
            reportLines.push(`DEFAULT  ${settingKey} = ${resolvedSetting.value}  (${resolvedSetting.reason})`);
        } else {
            reportLines.push(`SHEET    ${settingKey} = ${resolvedSetting.value}`);
        }
    });

    findBootstrapSettingMismatches().forEach(mismatch => {
        problemCount++;
        reportLines.push(`BOOTSTRAP ${mismatch}`);
    });

    const summary = problemCount === 0
        ? 'All settings were read from the Settings sheet.'
        : `${problemCount} setting(s) are not coming from the Settings sheet (see lines marked DEFAULT / BOOTSTRAP).`;
    const fullReport = [summary, ...reportLines].join('\n');
    log(fullReport, true);
    return fullReport;
}

// ----- Accessors: use these everywhere instead of a constant -----

/** Setting "userStartingRow" (positiveInteger). */
function getUserStartingRow(): number {
    return getConfig('userStartingRow');
}

/** Setting "namesCol" (positiveInteger). */
function getNamesColumn(): number {
    return getConfig('namesCol');
}

/** Setting "balanceCol" (positiveInteger). */
function getBalanceColumn(): number {
    return getConfig('balanceCol');
}

/** Setting "earningsCol" (positiveInteger). */
function getEarningsColumn(): number {
    return getConfig('earningsCol');
}

/** Setting "netIncomeCol" (positiveInteger). */
function getNetIncomeColumn(): number {
    return getConfig('netIncomeCol');
}

/** Setting "expendituresCol" (positiveInteger). */
function getExpendituresColumn(): number {
    return getConfig('expendituresCol');
}

/** Setting "investmentReturnsCol" (positiveInteger). */
function getInvestmentReturnsColumn(): number {
    return getConfig('investmentReturnsCol');
}

/** Setting "initialDepositCol" (positiveInteger). */
function getInitialDepositColumn(): number {
    return getConfig('initialDepositCol');
}

/** Setting "timeDepositCol" (positiveInteger). */
function getDepositDateColumn(): number {
    return getConfig('timeDepositCol');
}

/** Setting "grossInvestmentGainCol" (positiveInteger). */
function getGrossInvestmentGainColumn(): number {
    return getConfig('grossInvestmentGainCol');
}

/** Setting "netInvestmentGainCol" (positiveInteger). */
function getNetInvestmentGainColumn(): number {
    return getConfig('netInvestmentGainCol');
}

/** Setting "netPercentageGainCol" (positiveInteger). */
function getNetPercentageGainColumn(): number {
    return getConfig('netPercentageGainCol');
}

/** Setting "defaultServicesSheet" (sheetName). */
function getServicesSheetName(): string {
    return getConfig('defaultServicesSheet');
}

/** Setting "servicesRowStart" (positiveInteger). */
function getServicesRowStart(): number {
    return getConfig('servicesRowStart');
}

/** Setting "defaultTransactionsSheet" (sheetName). */
function getTransactionsSheetName(): string {
    return getConfig('defaultTransactionsSheet');
}

/** Setting "transactionsRowStart" (positiveInteger). */
function getTransactionsRowStart(): number {
    return getConfig('transactionsRowStart');
}

/** Setting "fillStartingCell" (cellAddress). */
function getFillStartingCell(): string {
    return getConfig('fillStartingCell');
}

/** Setting "fillSheetCellName" (cellAddress). */
function getFillSheetCellName(): string {
    return getConfig('fillSheetCellName');
}

/** Setting "defaultHistoricalRecordsSheet" (sheetName). */
function getHistoricalRecordsSheetName(): string {
    return getConfig('defaultHistoricalRecordsSheet');
}

/** Setting "historicalRecordsRowStart" (positiveInteger). */
function getHistoricalRecordsRowStart(): number {
    return getConfig('historicalRecordsRowStart');
}

/** Setting "metricDataStartColumn" (positiveInteger). */
function getMetricDataStartColumn(): number {
    return getConfig('metricDataStartColumn');
}

/** Setting "periodColIndex" (positiveInteger). */
function getPeriodColumnIndex(): number {
    return getConfig('periodColIndex');
}

/** Setting "headerRowIndex" (positiveInteger). */
function getHeaderRowIndex(): number {
    return getConfig('headerRowIndex');
}

/** Setting "sheetNamePattern" (text). */
function getSheetNamePattern(): string {
    return getConfig('sheetNamePattern');
}

/** Setting "timestampCell" (cellAddress). */
function getTimestampCell(): string {
    return getConfig('timestampCell');
}
