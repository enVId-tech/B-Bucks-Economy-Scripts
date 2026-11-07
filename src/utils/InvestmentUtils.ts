/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran. 
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 * 
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 *
 * This file contains utility functions for various operations in the B-Bucks Economy Scripts project, including functions related to investments management and other financial operations.
 */

interface PeriodData {
    periodName: string;
    investments: InvestmentRecord[];
}

interface InvestmentRecord {
    individualName: string;
    balance: number;
    earnings?: number;
    netEarnings?: number;
    expenditure?: number;
    returnsAmount?: number;
    initAmount?: number;
    date?: string;
    grossCurrentAmount?: number;
    netCurrentAmount?: number;
    netPercentageGain?: number;
}

/**
 *  Fetches investments ledger data with caching. It first checks for cached data to minimize latency, and if not found or if a force refresh is requested, it reads the investments ledger data from the sheet and updates the cache with the new data. This function ensures that the application can quickly access investments ledger data while also providing a mechanism to refresh the data when necessary.
 * @param data A string containing the data for the function, including a forceRefresh flag.
 * @returns {PeriodData[] | { error: string }} An array of period data objects containing the period name and an array of investment records for that period, or an error message if the sheet is not found or an error occurs. Each period data object is structured to allow easy access to its details throughout the application, facilitating operations such as investments management and performance tracking. The investment records include details such as the date of the investment, the individual's name, the initial amount invested, the returns amount, and optionally the current amount for ongoing investments.
 */
function fetchInvestmentsDataCached(data?: string): PeriodData[] | { error: string } {
    try {
        if (data && typeof data === 'string') {
            log(`Received data for fetchInvestmentsDataCached: ${data}`, false);
        } else {
            log("No data received for fetchInvestmentsDataCached, proceeding with default cache retrieval.", false);
            data = JSON.stringify({ forceRefresh: false });
        }


        const parsedData = data ? JSON.parse(data) : null;
        const forceRefresh = parsedData?.forceRefresh || false;
        // Optional: refresh only these periods and keep the cached copy of every other one
        const onlyPeriods: string[] | null = Array.isArray(parsedData?.periods) && parsedData.periods.length > 0
            ? parsedData.periods.map((p: any) => String(p))
            : null;

        const cache = CacheService.getScriptCache();
        const indexKey = "cachedInvestmentPeriodIndex";

        const storedIndex = PropertiesService.getScriptProperties().getProperty(indexKey);
        const periodKeys: Array<{ periodName: string; cacheKey: string }> | null = storedIndex ? JSON.parse(storedIndex) : null;
        // One batched cache read for all periods (previously one read per period)
        const cachedByKey: { [key: string]: string } = periodKeys ? cache.getAll(periodKeys.map(periodKey => periodKey.cacheKey)) : {};
        const allPeriodsCached = !!periodKeys && periodKeys.length > 0 && periodKeys.every(periodKey => cachedByKey[periodKey.cacheKey]);

        if (!forceRefresh && periodKeys && allPeriodsCached) {
            log(`Cache hit: Loaded ${periodKeys.length} investment periods individually.`, false);
            return periodKeys.map(periodKey => JSON.parse(cachedByKey[periodKey.cacheKey])) as PeriodData[];
        }

        // Targeted refresh: every period is cached and every requested period is a known one, so re-read just those sheets
        if (forceRefresh && onlyPeriods && periodKeys && allPeriodsCached
            && onlyPeriods.every(name => periodKeys.some(periodKey => periodKey.periodName === name))) {
            const fresh = fetchInvestmentsLedgerData(onlyPeriods);
            if (!('error' in fresh) && fresh.length === new Set(onlyPeriods).size) {
                const freshByName: { [name: string]: PeriodData } = {};
                fresh.forEach(freshPeriod => { freshByName[freshPeriod.periodName] = freshPeriod; });

                const toStore: { [key: string]: string } = {};
                const merged = periodKeys.map(periodKey => {
                    const replacement = freshByName[periodKey.periodName];
                    if (!replacement) return JSON.parse(cachedByKey[periodKey.cacheKey]) as PeriodData;
                    toStore[periodKey.cacheKey] = JSON.stringify(replacement);
                    return replacement;
                });
                cache.putAll(toStore, SERVER_SIDE_CACHE_AGE);
                log(`Targeted refresh of ${fresh.length} of ${periodKeys.length} investment periods.`, false);
                return merged;
            }
            // Anything unexpected falls through to the full refresh below
        }

        log("Cache miss: Re-extracting investment periods from sheets...", false);
        const periodDataArray: PeriodData[] | { error: string } = fetchInvestmentsLedgerData();

        if ('error' in periodDataArray) {
            log(`Error fetching investments ledger data: ${periodDataArray.error}`, true);
            return { error: periodDataArray.error };
        }

        log(`Fetched fresh investments ledger data from sheet. Data: ${JSON.stringify(periodDataArray)}`, false);

        if (Array.isArray(periodDataArray)) {
            const newKeys = periodDataArray.map(period => ({
                periodName: period.periodName,
                cacheKey: `investmentPeriod_${Utilities.base64EncodeWebSafe(period.periodName)}`
            }));
            const toStore: { [key: string]: string } = {};
            periodDataArray.forEach((period, i) => { toStore[newKeys[i].cacheKey] = JSON.stringify(period); });
            cache.putAll(toStore, SERVER_SIDE_CACHE_AGE);
            PropertiesService.getScriptProperties().setProperty(indexKey, JSON.stringify(newKeys));
        }

        return periodDataArray as PeriodData[];
    } catch (err) {
        log(`Error in fetchInvestmentsDataCached: ${err instanceof Error ? err.message : String(err)}`, true);
        return { error: "Failed to fetch cached investments data" };
    }
}

/**
 * Reads every student's investment columns from one period sheet with a single range call.
 * Shared by the full ledger fetch and the single-student refresh so both read exactly the same way.
 */
function readInvestmentsFromSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): InvestmentRecord[] {
    const values = sheet.getDataRange().getValues();
    const investments: InvestmentRecord[] = [];

    for (let row = getUserStartingRow() - 1; row < values.length; row++) {
        const individualName = values[row][0];
        const balance = Number(values[row][1].toFixed(2));
        const returnsAmount = Number(values[row][5].toFixed(2));
        const initAmount = Number(values[row][6].toFixed(2));
        const dateValue = values[row][7];
        const grossCurrentAmount = Number(values[row][8].toFixed(2));
        const netCurrentAmount = Number(values[row][9].toFixed(2));
        // the percentage with 2 decimal places is represented as a number
        // with 4 decimal places since the lowest you can represent is 0.01% = 0.0001
        const netPercentageGain = Number(values[row][10].toFixed(4));

        if (individualName) {
            investments.push({
                date: dateValue instanceof Date ? dateValue.toISOString() : undefined,
                individualName: individualName,
                balance: balance && typeof balance === 'number' ? balance : 0,
                initAmount: initAmount && typeof initAmount === 'number' ? initAmount : undefined,
                returnsAmount: returnsAmount && typeof returnsAmount === 'number' ? returnsAmount : undefined,
                grossCurrentAmount: grossCurrentAmount && typeof grossCurrentAmount === 'number' ? grossCurrentAmount : undefined,
                netCurrentAmount: netCurrentAmount && typeof netCurrentAmount === 'number' ? netCurrentAmount : undefined,
                netPercentageGain: netPercentageGain && typeof netPercentageGain === 'number' ? netPercentageGain : undefined
            });
        }
    }
    return investments;
}

/**
 * Fetches investments ledger data from the sheet. It reads all sheets in the active spreadsheet, filters for those that include "Period" in their name, and extracts investment records from a predefined range of rows and columns. The function constructs structured objects for each period, containing the period name and an array of investment records with details such as date, individual name, initial amount, returns amount, and optionally current amount. If the sheet is not found or an error occurs during processing, it returns an empty array.
 * @returns {PeriodData[] | { error: string }} An array of period data objects containing the period name and an array of investment records for that period, or an error object if an error occurs. Each period data object is structured to allow easy access to its details throughout the application, facilitating operations such as investments management and performance tracking. The investment records include details such as the date of the investment, the individual's name, the initial amount invested, the returns amount, and optionally the current amount for ongoing investments.
 */
function fetchInvestmentsLedgerData(onlyPeriods?: string[] | null): PeriodData[] | { error: string } {
    try {
        const allSheets = getAllSheets();
        const periodSheets = allSheets.filter(sheet => sheet.getName().includes("Period") && (!onlyPeriods || onlyPeriods.includes(sheet.getName())));
        let periodDataArray: PeriodData[] = [];

        periodSheets.forEach(sheet => {
            periodDataArray.push({
                periodName: sheet.getName(),
                investments: readInvestmentsFromSheet(sheet)
            });
        });

        // Cache the transactions data for use in the Investments Manager dialog
        return periodDataArray;
    } catch (err) {
        log(`Error in fetchInvestmentsLedgerData: ${err instanceof Error ? err.message : String(err)}`, true);
        return { error: "Failed to fetch investments ledger data from sheet" };
    }
}

/**
 * Refreshes the data for all investments across all periods.
 * @returns A boolean indicating whether the operation was successful.
 */
function refreshAllInvestments(): boolean {
    try {
        const periodDataArray = fetchInvestmentsLedgerData();
        if ('error' in periodDataArray) return false;

        const interestRateProp = fetchProperty("weeklyInterestRate");
        const taxRateProp = fetchProperty("investmentWithdrawalTaxRate");

        if ((typeof interestRateProp === 'object' && 'error' in interestRateProp) ||
            (typeof taxRateProp === 'object' && 'error' in taxRateProp)) {
            return false;
        }

        const interestRate = parseFloat(interestRateProp);
        const withdrawalTaxRate = parseFloat(taxRateProp);

        if (isNaN(interestRate) || isNaN(withdrawalTaxRate)) return false;

        const nowTime = new Date().getTime();
        const allSheets = getAllSheets();

        periodDataArray.forEach(periodData => {
            const sheet = allSheets.find(s => s.getName() === periodData.periodName);
            if (!sheet) return;

            // read from columns 9 (I), 10 (J), and 11 (K)
            const lastRow = sheet.getLastRow();
            if (lastRow < getUserStartingRow()) return;

            const totalRowsToProcess = (lastRow - getUserStartingRow()) + 1;
            // Grab current values I, J and K
            const range = sheet.getRange(getUserStartingRow(), getGrossInvestmentGainColumn(), totalRowsToProcess, 3);
            const values = range.getValues();

            // Names come from one narrow column read (previously the entire sheet was read a second time)
            const namesColumn = sheet.getRange(getUserStartingRow(), getNamesColumn(), totalRowsToProcess, 1).getValues();

            // First record per name wins, exactly like the previous Array.find
            const investmentByName = new Map<string, InvestmentRecord>();
            periodData.investments.forEach(investmentRecord => {
                if (!investmentByName.has(investmentRecord.individualName)) investmentByName.set(investmentRecord.individualName, investmentRecord);
            });

            for (let i = 0; i < totalRowsToProcess; i++) {
                const studentName = namesColumn[i][0];
                if (!studentName) continue;

                const investment = investmentByName.get(studentName);
                if (investment && investment.initAmount !== undefined) {
                    const depositDate = investment.date ? new Date(investment.date) : null;
                    const weeksSinceDeposit = depositDate && !isNaN(depositDate.getTime())
                        ? Math.floor((nowTime - depositDate.getTime()) / (1000 * 60 * 60 * 24 * 7))
                        : 0;
                    const effectiveInterestRate = interestRate * weeksSinceDeposit;

                    const grossCurrentAmount = investment.initAmount * (1 + effectiveInterestRate);
                    const netValue = grossCurrentAmount * (1 - withdrawalTaxRate);
                    const percentageGain = investment.initAmount > 0 ? ((netValue - investment.initAmount) / investment.initAmount) : 0;

                    // Update the values in the range
                    values[i][0] = grossCurrentAmount;
                    values[i][1] = netValue;
                    values[i][2] = percentageGain;
                }
            }

            // Write back columns I, J, and K for all rows simultaneously in one call
            range.setValues(values);
        });

        SpreadsheetApp.flush();
        return true;
    } catch (err) {
        log(`Error in refreshAllInvestments: ${err instanceof Error ? err.message : String(err)}`, true);
        return false;
    }
}

/**
 * Refreshes the data for a single investment within a specific period.
 * @param individualName The name of the individual whose investment needs to be refreshed.
 * @param periodName The name of the period in which the investment is located.
 * @returns A boolean indicating whether the operation was successful.
 */
function refreshSingleInvestment(individualName: string, periodName: string): boolean {
    try {
        const targetSheet = getSheetNamed(periodName);
        if (!targetSheet || !targetSheet.getName().includes("Period")) return false;

        const interestRateProp = fetchProperty("weeklyInterestRate", "standardPercentages");
        const taxRateProp = fetchProperty("investmentWithdrawalTaxRate", "standardPercentages");

        if ((typeof interestRateProp === 'object' && 'error' in interestRateProp) ||
            (typeof taxRateProp === 'object' && 'error' in taxRateProp)) {
            return false;
        }

        const interestRate = parseFloat(interestRateProp);
        const withdrawalTaxRate = parseFloat(taxRateProp);
        if (isNaN(interestRate) || isNaN(withdrawalTaxRate)) return false;

        const investment = readInvestmentsFromSheet(targetSheet).find(inv => inv.individualName === individualName);
        if (!investment) return false;

        if (investment.initAmount !== undefined) {
            const depositDate = investment.date ? new Date(investment.date) : null;
            const weeksSinceDeposit = depositDate && !isNaN(depositDate.getTime())
                ? Math.floor((new Date().getTime() - depositDate.getTime()) / (1000 * 60 * 60 * 24 * 7))
                : 0;
            const effectiveInterestRate = interestRate * weeksSinceDeposit;
            const grossCurrentAmount = investment.initAmount * (1 + effectiveInterestRate);
            const netValue = grossCurrentAmount * (1 - withdrawalTaxRate);
            const percentageGain = investment.initAmount > 0 ? ((netValue - investment.initAmount) / investment.initAmount) : 0;

            // Locate the student with a single narrow column read
            const sheet = targetSheet;
            const firstRow = getUserStartingRow();
            const rowCount = sheet.getLastRow() - firstRow + 1;
            let targetRow = -1;

            if (rowCount > 0) {
                const names = sheet.getRange(firstRow, getNamesColumn(), rowCount, 1).getValues();
                for (let i = 0; i < names.length; i++) {
                    if (names[i][0] === individualName) {
                        targetRow = firstRow + i;
                        break;
                    }
                }
            }

            if (targetRow !== -1) {
                // Write the calculated values to columns 9, 10, and 11 for the target row
                sheet.getRange(targetRow, getGrossInvestmentGainColumn(), 1, 3).setValues([[grossCurrentAmount, netValue, percentageGain]]);
                SpreadsheetApp.flush();
                return true;
            }
        }
        return true;
    } catch (err) {
        log(`Error in refreshSingleInvestment: ${err instanceof Error ? err.message : String(err)}`, true);
        return false;
    }
}

/**
 * Writes the updated investment data back to the sheet.
 * @param periodData The data for the period.
 * @param investment The updated investment record.
 * @returns A boolean indicating whether the operation was successful.
 */
function writeSingleInvestmentToSheet(periodData: PeriodData, investment: InvestmentRecord): boolean {
    try {
        const sheet = getSheetNamed(periodData.periodName);
        if (!sheet) return false;

        const firstRow = getUserStartingRow();
        const rowCount = sheet.getLastRow() - firstRow + 1;
        let investmentRow = -1;

        // Locate the individual's row with a single narrow column read
        if (rowCount > 0) {
            const names = sheet.getRange(firstRow, getNamesColumn(), rowCount, 1).getValues();
            for (let i = 0; i < names.length; i++) {
                if (names[i][0] === investment.individualName) {
                    investmentRow = firstRow + i;
                    break;
                }
            }
        }

        if (investmentRow === -1) return false;

        // Pre-parse the date for investment
        // This prevents reparsing like in the old code
        let dateObj: any = '';
        if (investment.date) {
            const parsedDate = new Date(investment.date);
            dateObj = !isNaN(parsedDate.getTime()) ? parsedDate : investment.date;
        }

        const inputsBatchRow = [[
            Number((investment.expenditure !== undefined ? investment.expenditure : 0.00).toFixed(2)),  // Col 5
            Number((investment.returnsAmount !== undefined ? investment.returnsAmount : 0.00).toFixed(2)),  // Col 6
            Number((investment.initAmount !== undefined ? investment.initAmount : 0.00).toFixed(2)),        // Col 7
            dateObj                                                                    // Col 8
        ]];

        sheet.getRange(investmentRow, getExpendituresColumn(), 1, 4).setValues(inputsBatchRow);

        SpreadsheetApp.flush();
        return true;
    } catch (err) {
        log(`Error in writeSingleInvestmentToSheet: ${err instanceof Error ? err.message : String(err)}`, true);
        return false;
    }
}

/**
 * Writes the updated investment data for all periods back to the respective sheets.
 * @param periodDataArray An array of period data containing the updated investment records.
 * @returns A boolean indicating whether the operation was successful.
 */
function writeInvestmentsDataToSheet(periodDataArray: PeriodData[]): boolean {
    try {
        const allSheets = getAllSheets();

        periodDataArray.forEach(periodData => {
            const sheet = allSheets.find(s => s.getName() === periodData.periodName);
            if (!sheet) return;

            const firstRow = getUserStartingRow();
            const rowCount = sheet.getLastRow() - firstRow + 1;
            if (rowCount <= 0) return;

            // One read of the names and of the four input columns, one write back (previously 2 calls per student).
            const names = sheet.getRange(firstRow, getNamesColumn(), rowCount, 1).getValues();
            const block = sheet.getRange(firstRow, getExpendituresColumn(), rowCount, 4);
            const blockValues = block.getValues();
            const blockFormulas = block.getFormulas();

            // First row per name wins (same as the old per-student search from the top)
            const rowByName = new Map<any, number>();
            names.forEach((nameRow, rowOffset) => { if (!rowByName.has(nameRow[0])) rowByName.set(nameRow[0], rowOffset); });

            let touched = false;
            periodData.investments.forEach(investment => {
                const rowOffset = rowByName.get(investment.individualName);
                if (rowOffset === undefined) return;

                let dateObj: any = '';
                if (investment.date) {
                    const parsedDate = new Date(investment.date);
                    dateObj = !isNaN(parsedDate.getTime()) ? parsedDate : investment.date;
                }

                blockValues[rowOffset] = [
                    Number((investment.expenditure !== undefined ? investment.expenditure : 0.00).toFixed(2)),
                    Number((investment.returnsAmount !== undefined ? investment.returnsAmount : 0.00).toFixed(2)),
                    Number((investment.initAmount !== undefined ? investment.initAmount : 0.00).toFixed(2)),
                    dateObj
                ];
                blockFormulas[rowOffset] = ['', '', '', ''];
                touched = true;
            });

            if (!touched) return;

            // Rows we did not change keep any formula they had
            const mergedValues = blockValues.map((valueRow, rowOffset) => valueRow.map((cellValue, columnOffset) => (blockFormulas[rowOffset][columnOffset] !== '' ? blockFormulas[rowOffset][columnOffset] : cellValue)));
            block.setValues(mergedValues);
        });

        SpreadsheetApp.flush();
        return true;
    } catch (err) {
        log(`Error in writeInvestmentsDataToSheet: ${err instanceof Error ? err.message : String(err)}`, true);
        return false;
    }
}

/**
 * Handles a deposit transaction for a student in a specific period.
 * @param data The JSON string containing the deposit information.
 * @returns A boolean indicating whether the operation was successful.
 */
function handleDeposit(data: string): boolean {
    // Money moves only while holding the transaction lock, see LockUtils.ts
    return runWithTransactionLock(() => handleDepositInsideLock_(data), false);
}

/** Body of handleDeposit(). Only call it through handleDeposit(), which holds the transaction lock. */
function handleDepositInsideLock_(data: string): boolean {
    try {
        const parsedData = JSON.parse(data);
        const { amount, student, period, override } = parsedData;

        if (typeof amount !== 'number' || typeof student !== 'string' || typeof period !== 'string') {
            log(`Invalid data format for deposit. Received data: ${data}. Please check the input`, true);
            return false;
        }

        if (amount <= 0) {
            log(`Deposit amount must be greater than zero. Received amount: ${amount}.`, true);
            return false;
        }

        const sheet = getSheetNamed(period);
        if (!sheet) {
            log(`Sheet for period "${period}" not found. Please check the period name.`, true);
            return false;
        }

        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        let studentRow = -1;

        for (let i = getUserStartingRow() - 1; i < values.length; i++) {
            if (values[i][getNamesColumn() - 1] === student) {
                studentRow = i + 1;
                break;
            }
        }

        if (studentRow === -1) {
            log(`Student "${student}" not found in period "${period}". Please check the student name and period.`, true);
            return false;
        }

        const previousBalanceCell = sheet.getRange(studentRow, getBalanceColumn());
        let previousBalance = parseFloat(previousBalanceCell.getValue()) || 0;

        // The existing CoD (net investment gain) is rolled into the student's cash as part of every deposit, so it
        // counts toward the funds available. Reading the balance column alone under-reports what the student has.
        const netCurrentAmountCell = sheet.getRange(studentRow, getNetInvestmentGainColumn());
        const currentNetAmount = parseFloat(netCurrentAmountCell.getValue());
        const hasRolloverCod = !isNaN(currentNetAmount) && currentNetAmount !== 0;

        // Rolling the CoD over is only allowed once the withdrawal window is open, or when the admin overrides it.
        if (hasRolloverCod && !override) {
            const lockoutSetting = Number(fetchProperty("withdrawLockoutTimeframeDays", "mandatedPolicies"));
            const lockoutDays = Number.isFinite(lockoutSetting) && lockoutSetting >= 0 ? lockoutSetting : 14;
            const depositDateValue = sheet.getRange(studentRow, getDepositDateColumn()).getValue();
            const depositDate = depositDateValue ? new Date(depositDateValue) : null;
            const elapsedDays = depositDate && !isNaN(depositDate.getTime())
                ? (Date.now() - depositDate.getTime()) / (1000 * 60 * 60 * 24)
                : Infinity;

            if (elapsedDays < lockoutDays) {
                log(`Deposit blocked for "${student}": their existing CoD ($${currentNetAmount.toFixed(2)}) is still inside the ${lockoutDays}-day withdrawal window (${Math.floor(elapsedDays)} day(s) elapsed). Wait for the window to close or use Admin Override.`, true);
                return false;
            }
        }

        // Funds available = cash in the balance column + the CoD that is rolled over, unless the admin overrides.
        const availableFunds = Number((previousBalance + (hasRolloverCod ? currentNetAmount : 0)).toFixed(2));
        if (!override && (availableFunds - amount) < -0.005) {
            log(`Depositing $${amount.toFixed(2)} isn't possible for "${student}". Cash: $${previousBalance.toFixed(2)}, CoD: $${(hasRolloverCod ? currentNetAmount : 0).toFixed(2)}, total available: $${availableFunds.toFixed(2)}. Lower the amount or use Admin Override.`, true);
            return false;
        }

        // fetch the net amount first, then add it to the returns if it exists
        if (!isNaN(currentNetAmount) && currentNetAmount !== 0) {
            const currentReturnsCell = sheet.getRange(studentRow, getInvestmentReturnsColumn());
            const currentReturns = parseFloat(currentReturnsCell.getValue()) || 0;
            currentReturnsCell.setValue(currentReturns + currentNetAmount);

            const withdrawalTransactionRecord: TransactionRecord[] = [{
                individual: student,
                type: "Investment",
                period: parseInt(period.replace(/\D/g, '')) || undefined,
                row: studentRow,
                operation: "ADD",
                serviceProvided: `Withdrawal ${override ? '(Override)' : ''}`,
                unitPrice: Number(currentNetAmount.toFixed(2)),
                quantity: 1,
                modifiedColumn: getInvestmentReturnsColumn(),
                tenderedMoney: Number(currentNetAmount.toFixed(2)),
                initialColumnAmount: Number(currentReturns.toFixed(2)),
                newColumnAmount: Number((currentReturns + currentNetAmount).toFixed(2)),
                initialBalance: Number(previousBalance.toFixed(2)),
                newBalance: Number((previousBalance + currentNetAmount).toFixed(2)),
                timestamp: new Date().toISOString()
            }];

            addTransactionRecords(withdrawalTransactionRecord);

            previousBalance += currentNetAmount;
            previousBalance = Number(previousBalance.toFixed(2));
        }

        // Update the expenditure for the student
        const expenditureCell = sheet.getRange(studentRow, getExpendituresColumn());
        if (isNaN(expenditureCell.getValue())) {
            expenditureCell.setValue(Number(amount.toFixed(2)));
        } else {
            const currentExpenditure = parseFloat(expenditureCell.getValue()) || 0;
            expenditureCell.setValue(Number((currentExpenditure + amount).toFixed(2)));
        }

        const initDepositCell = sheet.getRange(studentRow, getInitialDepositColumn());
        initDepositCell.setValue(Number(amount.toFixed(2)));

        // Update the date for the student to the current date (which is also the transaction date)
        const today = new Date();
        sheet.getRange(studentRow, getDepositDateColumn()).setValue(today);

        // The expenditure note is rebuilt from the transaction log when the deposit record below is added.

        refreshSingleInvestment(student, period);

        const transactionRecord: TransactionRecord[] = [{
            individual: student,
            period: parseInt(period.replace(/\D/g, '')) || undefined,
            row: studentRow,
            operation: "ADD",
            type: "Investment",
            serviceProvided: `Deposit ${override ? '(Override)' : ''}`,
            unitPrice: Number(amount.toFixed(2)),
            quantity: 1,
            modifiedColumn: getInitialDepositColumn(),
            tenderedMoney: Number(amount.toFixed(2)),
            initialColumnAmount: 0,
            newColumnAmount: Number(amount.toFixed(2)),
            initialBalance: Number(previousBalance.toFixed(2)),
            newBalance: Number((previousBalance - amount).toFixed(2)),
            timestamp: new Date().toISOString()
        }];

        addTransactionRecords(transactionRecord);

        return true;
    } catch (err) {
        log(`Error in handleDeposit: ${err instanceof Error ? err.message : String(err)}`, true);
        return false;
    }
}

/**
 * Handles the withdrawal of funds from a student's investment account.
 * @param data The withdrawal data in JSON format.
 * @returns A boolean indicating whether the withdrawal was successful.
 */
function handleWithdraw(data: string): boolean {
    // Money moves only while holding the transaction lock, see LockUtils.ts
    return runWithTransactionLock(() => handleWithdrawInsideLock_(data), false);
}

/** Body of handleWithdraw(). Only call it through handleWithdraw(), which holds the transaction lock. */
function handleWithdrawInsideLock_(data: string): boolean {
    try {
        const withdrawData = JSON.parse(data);
        const { amount, student, period, override } = withdrawData;;

        if (typeof amount !== 'number' || typeof student !== 'string' || typeof period !== 'string') {
            log(`Invalid data format for withdrawal. Received data: ${data}. Please check the input`, true);
            return false;
        }

        if (amount <= 0) {
            log(`Withdrawal amount must be greater than zero. Received amount: ${amount}.`, true);
            return false;
        }

        const sheet = getSheetNamed(period);
        if (!sheet) {
            log(`Sheet for period "${period}" not found. Please check the period name.`, true);
            return false;
        }

        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        let studentRow = -1;

        for (let i = getUserStartingRow() - 1; i < values.length; i++) {
            if (values[i][getNamesColumn() - 1] === student) {
                studentRow = i + 1;
                break;
            }
        }

        if (studentRow === -1) {
            log(`Student "${student}" not found in period "${period}". Please check the student name and period.`, true);
            return false;
        }

        const previousBalanceCell = sheet.getRange(studentRow, getBalanceColumn());
        let previousBalance = parseFloat(previousBalanceCell.getValue());
        if (isNaN(previousBalance)) {
            log(`Invalid balance for student "${student}". Please check the balance value in the sheet.`, true);
            return false;
        }
        previousBalance = Number(previousBalance.toFixed(2));

        // fetch the net value of the student's account
        const netCurrentAmountCell = sheet.getRange(studentRow, getNetInvestmentGainColumn());
        const currentNetAmount = parseFloat(netCurrentAmountCell.getValue());
        if (isNaN(currentNetAmount)) {
            log(`Invalid net amount for student "${student}". Please check the net investment gain value in the sheet.`, true);
            return false;
        }
        if (currentNetAmount === 0) {
            log(`"${student}" has no active CoD to withdraw.`, true);
            return false;
        }

        // add net amount to returns
        const returnsAmountCell = sheet.getRange(studentRow, getInvestmentReturnsColumn());
        const currentReturnsAmount = parseFloat(returnsAmountCell.getValue());
        if (isNaN(currentReturnsAmount)) {
            log(`Invalid returns amount for student "${student}". Please check the investment returns value in the sheet.`, true);
            return false;
        }

        const newReturnsAmount = currentReturnsAmount + currentNetAmount;
        returnsAmountCell.setValue(Number(newReturnsAmount.toFixed(2)));

        // clear the initial dep, date, gross value, net value and percentage gain
        sheet.getRange(studentRow, getInitialDepositColumn(), 1, 5).setValues([[0, '', 0, 0, 0]]);

        refreshSingleInvestment(student, period);

        const transactionRecord: TransactionRecord[] = [{
            individual: student,
            period: parseInt(period.replace(/\D/g, '')) || undefined,
            row: studentRow,
            operation: "SUBTRACT",
            type: "Investment",
            serviceProvided: `Withdrawal ${override ? '(Override)' : ''}`,
            unitPrice: Number(currentNetAmount.toFixed(2)),
            quantity: 1,
            modifiedColumn: getNetInvestmentGainColumn(),
            tenderedMoney: Number(currentNetAmount.toFixed(2)),
            initialColumnAmount: Number(currentNetAmount.toFixed(2)),
            newColumnAmount: 0,
            initialBalance: Number(previousBalance.toFixed(2)),
            // The returns column feeds the balance, so the balance goes up by exactly the net CoD value.
            newBalance: Number((previousBalance + currentNetAmount).toFixed(2)),
            timestamp: new Date().toISOString()
        }];

        addTransactionRecords(transactionRecord);
        return true;
    } catch (err) {
        log(`Error in handleWithdraw: ${err instanceof Error ? err.message : String(err)}`, true);
        return false;
    }
}