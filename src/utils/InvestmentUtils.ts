// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. 
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains utility functions for various operations in the B-Bucks Economy Scripts project, including functions related to investments management and other financial operations.

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

const STARTING_ROW = 7;

const INITIAL_AMOUNT_COL = 7;
const EXPENDITURE_COL = 5;
const DATE_COL = 8;
const RETURNS_COL = 6;
const NET_CURRENT_AMOUNT_COL = 10;
const GROSS_CURRENT_AMOUNT_COL = 9;
const NET_PERCENTAGE_GAIN_COL = 11;

/**
 *  Fetches investments ledger data with caching. It first checks for cached data to minimize latency, and if not found or if a force refresh is requested, it reads the investments ledger data from the sheet and updates the cache with the new data. This function ensures that the application can quickly access investments ledger data while also providing a mechanism to refresh the data when necessary.
 * @param data A string containing the data for the function, including a forceRefresh flag.
 * @returns {PeriodData[] | { error: string }} An array of period data objects containing the period name and an array of investment records for that period, or an error message if the sheet is not found or an error occurs. Each period data object is structured to allow easy access to its details throughout the application, facilitating operations such as investments management and performance tracking. The investment records include details such as the date of the investment, the individual's name, the initial amount invested, the returns amount, and optionally the current amount for ongoing investments.
 */
function fetchInvestmentsDataCached(data?: string): PeriodData[] | { error: string } {
    try {
        if (data && typeof data === 'string') {
            Logger.log(`Received data for fetchInvestmentsDataCached: ${data}`);
        } else {
            Logger.log("No data received for fetchInvestmentsDataCached, proceeding with default cache retrieval.");
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
                // SpreadsheetApp.getUi().alert(`Cache hit: Investments ledger data loaded from cache. String: ${cachedData}`);
                return JSON.parse(cachedData) as PeriodData[];
            }
        }

        if (!forceRefresh) {
            const savedProperties = props.getProperty(CACHE_KEY);
            if (savedProperties) {
                Logger.log(`Server Cache Hit (Properties) for ${CACHE_KEY}`);
                // Repopulate fast RAM cache so the next window open loads even faster
                cache.put(CACHE_KEY, savedProperties, 21600);
                return JSON.parse(savedProperties);
            }
        }

        console.log("Cache miss: Re-extracting items from Investments sheet rows...");
        // SpreadsheetApp.getUi().alert("Cache miss: Re-extracting items from Investments sheet rows...");
        const periodDataArray: PeriodData[] | { error: string } = fetchInvestmentsLedgerData();

        if ('error' in periodDataArray) {
            Logger.log(`Error fetching investments ledger data: ${periodDataArray.error}`);
            SpreadsheetApp.getUi().alert(`Error fetching investments ledger data: ${periodDataArray.error}`);
            return { error: periodDataArray.error };
        }

        // SpreadsheetApp.getUi().alert(`Fetched fresh investments ledger data from sheet. Data: ${JSON.stringify(periodDataArray)}`);

        if (Array.isArray(periodDataArray)) {
            try {
                setCachedData(CACHE_KEY, periodDataArray);
            } catch (cacheErr) {
                Logger.log(`Warning: Failed to set cache payload (likely size limit), continuing return: ${cacheErr}`);
                SpreadsheetApp.getUi().alert(`Warning: Failed to set cache payload (likely size limit), continuing return: ${cacheErr}`);
            }
        }

        // SpreadsheetApp.getUi().alert(`Fetched fresh investments ledger data from sheet. Data: ${JSON.stringify(periodDataArray)}`);

        return periodDataArray as PeriodData[];
    } catch (err) {
        Logger.log(`Error in fetchInvestmentsDataCached: ${err instanceof Error ? err.message : String(err)}`);
        SpreadsheetApp.getUi().alert(`Error in fetchInvestmentsDataCached: ${err instanceof Error ? err.message : String(err)}`);
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

        const startingRow = 7;

        periodSheets.forEach(sheet => {
            const dataRange = sheet.getDataRange();
            const values = dataRange.getValues();

            let investments: InvestmentRecord[] = [];

            for (let row = startingRow - 1; row < values.length; row++) {
                const individualName = values[row][0];
                const balance = values[row][1];
                const returnsAmount = values[row][5];
                const initAmount = values[row][6];
                const dateValue = values[row][7];
                const grossCurrentAmount = values[row][8];
                const netCurrentAmount = values[row][9];
                const netPercentageGain = values[row][10];

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
        Logger.log(`Error in fetchInvestmentsLedgerData: ${err instanceof Error ? err.message : String(err)}`);
        SpreadsheetApp.getUi().alert(`Error in fetchInvestmentsLedgerData: ${err instanceof Error ? err.message : String(err)}`);
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
        if ('error' in periodDataArray) {
            Logger.log(`Error fetching investments ledger data for refresh: ${periodDataArray.error}`);
            SpreadsheetApp.getUi().alert(`Error fetching investments ledger data for refresh: ${periodDataArray.error}`);
            return false;
        }

        const interestRateProp = fetchProperty("weeklyInterestRate");
        const taxRateProp = fetchProperty("investmentWithdrawalTaxRate");

        if ((typeof interestRateProp === 'object' && 'error' in interestRateProp) ||
            (typeof taxRateProp === 'object' && 'error' in taxRateProp)) {
            SpreadsheetApp.getUi().alert("Error fetching economic configurations for calculation updates.");
            return false;
        }

        const interestRate = parseFloat(interestRateProp);
        const withdrawalTaxRate = parseFloat(taxRateProp);

        if (isNaN(interestRate) || isNaN(withdrawalTaxRate)) {
            SpreadsheetApp.getUi().alert("Failed to parse banking rates. Check your Settings worksheet formulas.");
            return false;
        }

        // Process loops ultra-fast entirely in local memory variables
        periodDataArray.forEach(periodData => {
            periodData.investments.forEach(investment => {
                if (investment.initAmount !== undefined) {
                    // calculate interest based on weeks since deposit
                    const depositDate = investment.date ? new Date(investment.date) : null;
                    const now = new Date();
                    const weeksSinceDeposit = depositDate ? Math.floor((now.getTime() - depositDate.getTime()) / (1000 * 60 * 60 * 24 * 7)) : 0;
                    const effectiveInterestRate = interestRate * weeksSinceDeposit;

                    const grossCurrentAmount = investment.initAmount * (1 + effectiveInterestRate);
                    const netValue = grossCurrentAmount * (1 - withdrawalTaxRate);
                    const percentageGain = investment.initAmount > 0 ? ((netValue - investment.initAmount) / investment.initAmount) : 0;
                    investment.grossCurrentAmount = grossCurrentAmount;
                    investment.netCurrentAmount = netValue;
                    investment.netPercentageGain = percentageGain;
                }
            });
        });

        const writeSuccess = writeInvestmentsDataToSheet(periodDataArray);
        if (!writeSuccess) {
            SpreadsheetApp.getUi().alert("Failed to write updated investments data back to sheets.");
            return false;
        }

        return true;
    } catch (err) {
        Logger.log(`Error in refreshAllInvestments: ${err instanceof Error ? err.message : String(err)}`);
        SpreadsheetApp.getUi().alert(`Error in refreshAllInvestments: ${err instanceof Error ? err.message : String(err)}`);
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
        if ('error' in periodDataArray) {
            Logger.log(`Error fetching investments ledger data for single refresh: ${periodDataArray.error}`);
            SpreadsheetApp.getUi().alert(`Error fetching investments ledger data for single refresh: ${periodDataArray.error}`);
            return false;
        }

        // Fetch the necessary economic properties for the calculation.
        const interestRateProp = fetchProperty("weeklyInterestRate");
        const taxRateProp = fetchProperty("investmentWithdrawalTaxRate");
        if ((typeof interestRateProp === 'object' && 'error' in interestRateProp) ||
            (typeof taxRateProp === 'object' && 'error' in taxRateProp)) {
            SpreadsheetApp.getUi().alert("Error fetching economic configurations for calculation updates.");
            return false;
        }

        // Parse the economic properties into numbers.
        const interestRate = parseFloat(interestRateProp);
        const withdrawalTaxRate = parseFloat(taxRateProp);
        if (isNaN(interestRate) || isNaN(withdrawalTaxRate)) {
            SpreadsheetApp.getUi().alert("Failed to parse banking rates. Check your Settings worksheet formulas.");
            return false;
        }

        const periodData = periodDataArray.find(p => p.periodName === periodName);
        if (!periodData) {
            SpreadsheetApp.getUi().alert(`Period "${periodName}" not found.`);
            return false;
        }
        const investment = periodData.investments.find(inv => inv.individualName === individualName);
        if (!investment) {
            SpreadsheetApp.getUi().alert(`Investment for "${individualName}" not found in period "${periodName}".`);
            return false;
        }

        // Calculate the updated investment values based on the current date and the economic properties.
        if (investment.initAmount !== undefined) {
            const depositDate = investment.date ? new Date(investment.date) : null;
            const now = new Date();
            const weeksSinceDeposit = depositDate ? Math.floor((now.getTime() - depositDate.getTime()) / (1000 * 60 * 60 * 24 * 7)) : 0;
            const effectiveInterestRate = interestRate * weeksSinceDeposit;
            const grossCurrentAmount = investment.initAmount * (1 + effectiveInterestRate);
            const netValue = grossCurrentAmount * (1 - withdrawalTaxRate);
            const percentageGain = investment.initAmount > 0 ? ((netValue - investment.initAmount) / investment.initAmount) : 0;
            investment.grossCurrentAmount = grossCurrentAmount;
            investment.netCurrentAmount = netValue;
            investment.netPercentageGain = percentageGain;

            // Write the updated investment data back to the sheet.
            const writeSuccess = writeSingleInvestmentToSheet(periodData, investment);
            if (!writeSuccess) {
                SpreadsheetApp.getUi().alert("Failed to write updated investment data back to sheet.");
                return false;
            }
            return true;
        } else {
            SpreadsheetApp.getUi().alert(`Initial amount for "${individualName}" in period "${periodName}" is undefined, cannot calculate returns.`);
            return false;
        }
    } catch (err) {
        Logger.log(`Error in refreshSingleInvestment: ${err instanceof Error ? err.message : String(err)}`);
        SpreadsheetApp.getUi().alert(`Error in refreshSingleInvestment: ${err instanceof Error ? err.message : String(err)}`);
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

        // Find the row of the investment in the sheet.
        for (let i = STARTING_ROW - 1; i < values.length; i++) {
            if (values[i][0] === investment.individualName) {
                investmentRow = i + 1;
                break;
            }
        }

        if (investmentRow === -1) return false;

        // Update the investment data in the sheet.
        sheet.getRange(investmentRow, 1).setValue(investment.individualName);
        sheet.getRange(investmentRow, RETURNS_COL).setValue(investment.returnsAmount !== undefined ? investment.returnsAmount : 0.00);
        sheet.getRange(investmentRow, INITIAL_AMOUNT_COL).setValue(investment.initAmount !== undefined ? investment.initAmount : 0.00);
        sheet.getRange(investmentRow, GROSS_CURRENT_AMOUNT_COL).setValue(investment.grossCurrentAmount !== undefined ? investment.grossCurrentAmount : 0.00);
        sheet.getRange(investmentRow, NET_CURRENT_AMOUNT_COL).setValue(investment.netCurrentAmount !== undefined ? investment.netCurrentAmount : 0.00);
        sheet.getRange(investmentRow, NET_PERCENTAGE_GAIN_COL).setValue(investment.netPercentageGain !== undefined ? investment.netPercentageGain : 0.00);

        if (investment.date) {
            const parsedDate = new Date(investment.date);
            sheet.getRange(investmentRow, DATE_COL).setValue(!isNaN(parsedDate.getTime()) ? parsedDate : investment.date);
        } else {
            sheet.getRange(investmentRow, DATE_COL).setValue('');
        }

        SpreadsheetApp.flush();
        return true;
    } catch (err) {
        Logger.log(`Error in writeSingleInvestmentToSheet: ${err instanceof Error ? err.message : String(err)}`);
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
            if (sheet) {
                const investments = periodData.investments;
                const values = sheet.getDataRange().getValues();

                investments.forEach(inv => {
                    let targetRow = -1;
                    for (let i = STARTING_ROW - 1; i < values.length; i++) {
                        if (values[i][0] === inv.individualName) {
                            targetRow = i + 1;
                            break;
                        }
                    }

                    // Update the investment data in the sheet.
                    if (targetRow !== -1) {
                        sheet.getRange(targetRow, 1).setValue(inv.individualName);
                        sheet.getRange(targetRow, INITIAL_AMOUNT_COL).setValue(inv.initAmount !== undefined ? inv.initAmount : 0.00);
                        sheet.getRange(targetRow, GROSS_CURRENT_AMOUNT_COL).setValue(inv.grossCurrentAmount !== undefined ? inv.grossCurrentAmount : 0.00);
                        sheet.getRange(targetRow, NET_CURRENT_AMOUNT_COL).setValue(inv.netCurrentAmount !== undefined ? inv.netCurrentAmount : 0.00);
                        sheet.getRange(targetRow, NET_PERCENTAGE_GAIN_COL).setValue(inv.netPercentageGain !== undefined ? inv.netPercentageGain : 0.00);

                        if (inv.date) {
                            const parsedDate = new Date(inv.date);
                            sheet.getRange(targetRow, DATE_COL).setValue(!isNaN(parsedDate.getTime()) ? parsedDate : inv.date);
                        } else {
                            sheet.getRange(targetRow, DATE_COL).setValue('');
                        }
                    }
                });
            }
        });

        SpreadsheetApp.flush();
        return true;
    } catch (err) {
        Logger.log(`Error in writeInvestmentsDataToSheet: ${err instanceof Error ? err.message : String(err)}`);
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
        const { amount, student, period } = parsedData;

        if (typeof amount !== 'number' || typeof student !== 'string' || typeof period !== 'string') {
            SpreadsheetApp.getUi().alert("Invalid data format for deposit. Please check the input.");
            return false;
        }

        if (amount <= 0) {
            SpreadsheetApp.getUi().alert("Deposit amount must be greater than zero.");
            return false;
        }

        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(period);
        if (!sheet) {
            SpreadsheetApp.getUi().alert(`Sheet for period "${period}" not found.`);
            return false;
        }

        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        let studentRow = -1;

        for (let i = STARTING_ROW - 1; i < values.length; i++) {
            if (values[i][0] === student) {
                studentRow = i + 1;
                break;
            }
        }

        if (studentRow === -1) {
            SpreadsheetApp.getUi().alert(`Student "${student}" not found in period "${period}".`);
            return false;
        }

        // fetch the net amount first, then add it to the returns if it exists
        const netCurrentAmountCell = sheet.getRange(studentRow, NET_CURRENT_AMOUNT_COL);
        const currentNetAmount = parseFloat(netCurrentAmountCell.getValue());

        if (!isNaN(currentNetAmount) && currentNetAmount !== 0) {
            const currentReturnsCell = sheet.getRange(studentRow, RETURNS_COL);
            const currentReturns = parseFloat(currentReturnsCell.getValue()) || 0;
            currentReturnsCell.setValue(currentReturns + currentNetAmount);
        }

        // Update the expenditure for the student
        const expenditureCell = sheet.getRange(studentRow, EXPENDITURE_COL);
        if (isNaN(expenditureCell.getValue())) {
            expenditureCell.setValue(amount);
        }

        const currentExpenditure = parseFloat(expenditureCell.getValue()) || 0;
        expenditureCell.setValue(currentExpenditure + amount);

        const initDepositCell = sheet.getRange(studentRow, INITIAL_AMOUNT_COL);
        initDepositCell.setValue(amount);

        // Update the date for the student to the current date (which is also the transaction date)
        const today = new Date();
        sheet.getRange(studentRow, DATE_COL).setValue(today);

        // Add a comment to the expenditure cell with the transaction details
        const dateString = today.toLocaleDateString("en-US");
        commentExpenditureOnSelection([studentRow], `$${amount} - ${dateString} CoDs/Investments`);

        SpreadsheetApp.flush();

        refreshSingleInvestment(student, period);

        return true;
    } catch (err) {
        Logger.log(`Error in handleDeposit: ${err instanceof Error ? err.message : String(err)}`);
        try {
            SpreadsheetApp.getUi().alert(`Error in handleDeposit: ${err instanceof Error ? err.message : String(err)}`);
        } catch (e) {
            Logger.log(`Additionally, failed to show alert for error in handleDeposit: ${e instanceof Error ? e.message : String(e)}`);
        }
        return false;
    }
}