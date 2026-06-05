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
    date?: string;
    initAmount?: number;
    returnsAmount?: number;
    currentAmount?: number;
}

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
                const currentAmount = values[row][9];

                if (individualName) {
                    investments.push({
                        date: dateValue instanceof Date ? dateValue.toISOString() : undefined,
                        individualName: individualName,
                        balance: balance && typeof balance === 'number' ? balance : 0,
                        initAmount: initAmount && typeof initAmount === 'number' ? initAmount : undefined,
                        returnsAmount: returnsAmount && typeof returnsAmount === 'number' ? returnsAmount : undefined,
                        currentAmount: currentAmount && typeof currentAmount === 'number' ? currentAmount : undefined
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

// const selectedItems = document.querySelectorAll('[data-selectable="true"].selected');