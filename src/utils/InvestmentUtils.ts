/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran. 
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 * 
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 *
 *  This file contains utility functions for various operations in the B-Bucks Economy Scripts project, including functions related to investments management and other financial operations.
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

        const CACHE_KEY = "cachedInvestmentsLedger";
        const cache = CacheService.getScriptCache();
        const props = PropertiesService.getScriptProperties();

        if (!forceRefresh) {
            const cachedData = getCachedData(CACHE_KEY);
            if (cachedData && cachedData !== "{}" && cachedData !== "") {
                log(`Cache hit: Investments ledger data loaded from cache. String: ${cachedData}`, false);
                return JSON.parse(cachedData) as PeriodData[];
            }
            
            const savedProperties = props.getProperty(CACHE_KEY);
            if (savedProperties) {
                log(`Cache hit: Investments ledger data loaded from script properties for cache key ${CACHE_KEY}. String: ${savedProperties}`, false);
                // Repopulate fast RAM cache so the next window open loads even faster
                cache.put(CACHE_KEY, savedProperties, SERVER_SIDE_CACHE_AGE);
                return JSON.parse(savedProperties);
            }
        }

        log("Cache miss: Re-extracting items from Investments sheet rows...", false);
        const periodDataArray: PeriodData[] | { error: string } = fetchInvestmentsLedgerData();

        if ('error' in periodDataArray) {
            log(`Error fetching investments ledger data: ${periodDataArray.error}`, true);
            return { error: periodDataArray.error };
        }

        log(`Fetched fresh investments ledger data from sheet. Data: ${JSON.stringify(periodDataArray)}`, false);

        if (Array.isArray(periodDataArray)) {
            try {
                setCachedData(CACHE_KEY, periodDataArray);
            } catch (cacheErr) {
                log(`Warning: Failed to set cache payload (likely size limit), continuing return: ${cacheErr}`, true);
            }
        }

        log(`Fetched fresh investments ledger data from sheet. Data: ${JSON.stringify(periodDataArray)}`, false);

        return periodDataArray as PeriodData[];
    } catch (err) {
        log(`Error in fetchInvestmentsDataCached: ${err instanceof Error ? err.message : String(err)}`, true);
        return { error: "Failed to fetch cached investments data" };
    }
}

/**
 * Fetches investments ledger data from the sheet. It reads all sheets in the active spreadsheet, filters for those that include "Period" in their name, and extracts investment records from a predefined range of rows and columns. The function constructs structured objects for each period, containing the period name and an array of investment records with details such as date, individual name, initial amount, returns amount, and optionally current amount. If the sheet is not found or an error occurs during processing, it returns an empty array.
 * @returns {PeriodData[] | { error: string }} An array of period data objects containing the period name and an array of investment records for that period, or an error object if an error occurs. Each period data object is structured to allow easy access to its details throughout the application, facilitating operations such as investments management and performance tracking. The investment records include details such as the date of the investment, the individual's name, the initial amount invested, the returns amount, and optionally the current amount for ongoing investments.
 */
function fetchInvestmentsLedgerData(): PeriodData[] | { error: string } {
    try {
        const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
        const periodSheets = allSheets.filter(sheet => sheet.getName().includes("Period"));
        let periodDataArray: PeriodData[] = [];

        periodSheets.forEach(sheet => {
            const dataRange = sheet.getDataRange();
            const values = dataRange.getValues();

            let investments: InvestmentRecord[] = [];

            for (let row = USER_STARTING_ROW - 1; row < values.length; row++) {
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

            periodDataArray.push({
                periodName: sheet.getName(),
                investments: investments
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
        const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

        periodDataArray.forEach(periodData => {
            const sheet = allSheets.find(s => s.getName() === periodData.periodName);
            if (!sheet) return;

            // read from columns 9 (I), 10 (J), and 11 (K)
            const lastRow = sheet.getLastRow();
            if (lastRow < USER_STARTING_ROW) return;

            const totalRowsToProcess = (lastRow - USER_STARTING_ROW) + 1;
            // Grab current values I, J and K
            const range = sheet.getRange(USER_STARTING_ROW, GROSS_INVESTMENT_GAIN_COL || 9, totalRowsToProcess, 3);
            const values = range.getValues();

            // match values and only modify the necessary columns
            const currentSheetValues = sheet.getDataRange().getValues();

            for (let i = 0; i < totalRowsToProcess; i++) {
                const sheetRowIndex = USER_STARTING_ROW - 1 + i;
                const studentName = currentSheetValues[sheetRowIndex][0];
                if (!studentName) continue;

                const investment = periodData.investments.find(inv => inv.individualName === studentName);
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
        const periodDataArray = fetchInvestmentsLedgerData();
        if ('error' in periodDataArray) return false;

        const interestRateProp = fetchProperty("weeklyInterestRate", "standardPercentages");
        const taxRateProp = fetchProperty("investmentWithdrawalTaxRate", "standardPercentages");

        if ((typeof interestRateProp === 'object' && 'error' in interestRateProp) ||
            (typeof taxRateProp === 'object' && 'error' in taxRateProp)) {
            return false;
        }

        const interestRate = parseFloat(interestRateProp);
        const withdrawalTaxRate = parseFloat(taxRateProp);
        if (isNaN(interestRate) || isNaN(withdrawalTaxRate)) return false;

        const periodData = periodDataArray.find(p => p.periodName === periodName);
        if (!periodData) return false;

        const investment = periodData.investments.find(inv => inv.individualName === individualName);
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

            // Write the calculated values to columns 9, 10, and 11 for the target row
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(periodName);
            if (!sheet) return false;

            const values = sheet.getDataRange().getValues();
            let targetRow = -1;
            const totalRows = values.length;

            for (let i = USER_STARTING_ROW - 1; i < totalRows; i++) {
                if (values[i][0] === individualName) {
                    targetRow = i + 1;
                    break;
                }
            }

            if (targetRow !== -1) {
                // Write the calculated values to columns 9, 10, and 11 for the target row
                sheet.getRange(targetRow, 9, 1, 3).setValues([[grossCurrentAmount, netValue, percentageGain]]);
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
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(periodData.periodName);
        if (!sheet) return false;

        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        let investmentRow = -1;

        // Locate individual row
        const totalRows = values.length;
        for (let i = USER_STARTING_ROW - 1; i < totalRows; i++) {
            if (values[i][0] === investment.individualName) {
                investmentRow = i + 1;
                break;
            }
        }

        if (investmentRow === -1) return false;

        sheet.getRange(investmentRow, 1).setValue(investment.individualName);

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

        sheet.getRange(investmentRow, EXPENDITURES_COL, 1, 4).setValues(inputsBatchRow);

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
        const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

        periodDataArray.forEach(periodData => {
            const sheet = allSheets.find(s => s.getName() === periodData.periodName);
            if (!sheet) return;

            const investments = periodData.investments;
            const values = sheet.getDataRange().getValues();
            const totalRows = values.length;

            investments.forEach(investment => {
                let targetRow = -1;
                for (let i = USER_STARTING_ROW - 1; i < totalRows; i++) {
                    if (values[i][0] === investment.individualName) {
                        targetRow = i + 1;
                        break;
                    }
                }

                if (targetRow !== -1) {
                    sheet.getRange(targetRow, 1).setValue(investment.individualName);

                    let dateObj: any = '';
                    if (investment.date) {
                        const parsedDate = new Date(investment.date);
                        dateObj = !isNaN(parsedDate.getTime()) ? parsedDate : investment.date;
                    }

                    const batchData = [[
                        Number((investment.expenditure !== undefined ? investment.expenditure : 0.00).toFixed(2)),
                        Number((investment.returnsAmount !== undefined ? investment.returnsAmount : 0.00).toFixed(2)),
                        Number((investment.initAmount !== undefined ? investment.initAmount : 0.00).toFixed(2)),
                        dateObj
                    ]];

                    sheet.getRange(targetRow, EXPENDITURES_COL, 1, 4).setValues(batchData);
                }
            });
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

        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(period);
        if (!sheet) {
            log(`Sheet for period "${period}" not found. Please check the period name.`, true);
            return false;
        }

        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        let studentRow = -1;

        for (let i = USER_STARTING_ROW - 1; i < values.length; i++) {
            if (values[i][0] === student) {
                studentRow = i + 1;
                break;
            }
        }

        if (studentRow === -1) {
            log(`Student "${student}" not found in period "${period}". Please check the student name and period.`, true);
            return false;
        }

        const previousBalanceCell = sheet.getRange(studentRow, BALANCE_COL);
        let previousBalance = parseFloat(previousBalanceCell.getValue()) || 0;

        // Add a check to ensure that the deposit does not cause the balance to go negative unless override is true
        if (!override && (previousBalance - amount) < 0) {
            log(`Deposit of ${amount} would result in a negative balance for student "${student}". Current balance: ${previousBalance}. Override not enabled.`, true);
            return false;
        }

        // fetch the net amount first, then add it to the returns if it exists
        const netCurrentAmountCell = sheet.getRange(studentRow, NET_INVESTMENT_GAIN_COL);
        const currentNetAmount = parseFloat(netCurrentAmountCell.getValue());

        if (!isNaN(currentNetAmount) && currentNetAmount !== 0) {
            const currentReturnsCell = sheet.getRange(studentRow, INVESTMENT_RETURNS_COL);
            const currentReturns = parseFloat(currentReturnsCell.getValue()) || 0;
            currentReturnsCell.setValue(currentReturns + currentNetAmount);

            const withdrawalTransactionRecord: TransactionRecord[] = [{
                individual: student,
                type: "Investment",
                period: parseInt(period.replace(/\D/g, '')) || undefined,
                serviceProvided: `Withdrawal ${override ? '(Override)' : ''}`,
                unitPrice: Number(currentNetAmount.toFixed(2)),
                quantity: 1,
                modifiedColumn: INVESTMENT_RETURNS_COL,
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
        const expenditureCell = sheet.getRange(studentRow, EXPENDITURES_COL);
        if (isNaN(expenditureCell.getValue())) {
            expenditureCell.setValue(Number(amount.toFixed(2)));
        } else {
            const currentExpenditure = parseFloat(expenditureCell.getValue()) || 0;
            expenditureCell.setValue(Number((currentExpenditure + amount).toFixed(2)));
        }

        const initDepositCell = sheet.getRange(studentRow, INITIAL_DEPOSIT_COL);
        initDepositCell.setValue(Number(amount.toFixed(2)));

        // Update the date for the student to the current date (which is also the transaction date)
        const today = new Date();
        sheet.getRange(studentRow, DATE_DEPOSIT_COL).setValue(today);

        // Add a comment to the expenditure cell with the transaction details
        const dateString = today.toLocaleDateString("en-US");
        commentExpenditureOnSelection([studentRow], `$${amount} - ${dateString} CoDs/Investments ${override ? '(Override)' : ''}`);

        refreshSingleInvestment(student, period);

        const transactionRecord: TransactionRecord[] = [{
            individual: student,
            period: parseInt(period.replace(/\D/g, '')) || undefined,
            type: "Investment",
            serviceProvided: `Deposit ${override ? '(Override)' : ''}`,
            unitPrice: Number(amount.toFixed(2)),
            quantity: 1,
            modifiedColumn: INITIAL_DEPOSIT_COL,
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

        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(period);
        if (!sheet) {
            log(`Sheet for period "${period}" not found. Please check the period name.`, true);
            return false;
        }

        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        let studentRow = -1;

        for (let i = USER_STARTING_ROW - 1; i < values.length; i++) {
            if (values[i][0] === student) {
                studentRow = i + 1;
                break;
            }
        }

        if (studentRow === -1) {
            log(`Student "${student}" not found in period "${period}". Please check the student name and period.`, true);
            return false;
        }

        const previousBalanceCell = sheet.getRange(studentRow, BALANCE_COL);
        let previousBalance = parseFloat(previousBalanceCell.getValue());
        if (isNaN(previousBalance)) {
            log(`Invalid balance for student "${student}". Please check the balance value in the sheet.`, true);
            return false;
        }
        previousBalance = Number(previousBalance.toFixed(2));

        // fetch the net value of the student's account
        const netCurrentAmountCell = sheet.getRange(studentRow, NET_INVESTMENT_GAIN_COL);
        const currentNetAmount = parseFloat(netCurrentAmountCell.getValue());
        if (isNaN(currentNetAmount)) {
            log(`Invalid net amount for student "${student}". Please check the net investment gain value in the sheet.`, true);
            return false;
        }

        // add net amount to returns
        const returnsAmountCell = sheet.getRange(studentRow, INVESTMENT_RETURNS_COL);
        const currentReturnsAmount = parseFloat(returnsAmountCell.getValue());
        if (isNaN(currentReturnsAmount)) {
            log(`Invalid returns amount for student "${student}". Please check the investment returns value in the sheet.`, true);
            return false;
        }

        const newReturnsAmount = currentReturnsAmount + currentNetAmount;
        returnsAmountCell.setValue(Number(newReturnsAmount.toFixed(2)));

        // clear the initial dep, date, gross value, net value and percentage gain
        sheet.getRange(studentRow, INITIAL_DEPOSIT_COL, 1, 5).setValues([[0, '', 0, 0, 0]]);

        refreshSingleInvestment(student, period);

        const transactionRecord: TransactionRecord[] = [{
            individual: student,
            period: parseInt(period.replace(/\D/g, '')) || undefined,
            type: "Investment",
            serviceProvided: `Withdrawal ${override ? '(Override)' : ''}`,
            unitPrice: Number(currentNetAmount.toFixed(2)),
            quantity: 1,
            modifiedColumn: NET_INVESTMENT_GAIN_COL,
            tenderedMoney: Number(currentNetAmount.toFixed(2)),
            initialColumnAmount: Number(currentNetAmount.toFixed(2)),
            newColumnAmount: 0,
            initialBalance: Number(previousBalance.toFixed(2)),
            newBalance: Number(newReturnsAmount.toFixed(2)),
            timestamp: new Date().toISOString()
        }];

        addTransactionRecords(transactionRecord);
        return true;
    } catch (err) {
        log(`Error in handleWithdraw: ${err instanceof Error ? err.message : String(err)}`, true);
        return false;
    }
}