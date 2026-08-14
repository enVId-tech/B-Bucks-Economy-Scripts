/* Coded by Erick Tran for Mr. Banderas, 2026
 * Copyright (c) 2026 Erick Tran. 
 * This file is licensed under the MIT License, check the LICENSE file for details.
 *
 * GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
 * 
 * !!!--- IMPORTANT NOTE: This file was originally written in TypeScript, go to the GitHub to see the original non-compiled code. ---!!!
 * 
 * This file contains utility functions for various operations in the B-Bucks Economy Scripts project, including functions related to services management and other financial operations.
 */

interface QuarterlyData {
    Q1?: number;
    Q2?: number;
    Q3?: number;
    Q4?: number;
}

interface ItemData {
    itemName: string;
    category: "Income" | "Expense";
    pricing: QuarterlyData;
    limit: QuarterlyData;
}

/**
 * Fetches services data from the "Services" sheet in the active spreadsheet. It first checks for cached data to minimize latency, and if not found or if a force refresh is requested, it reads a predefined range of rows and columns to extract various details about services, including item names, categories, pricing for each quarter, and limits for each quarter. The function processes the raw grid data to construct structured objects for each service item, handling type conversions for numbers as needed. If the sheet is not found or an error occurs during processing, it returns an error message.
 * @param data A string containing the data for the function, including a forceRefresh flag. Defaults to undefined, meaning it will use cached data if available for faster access.
 * @returns {ItemData[] | { error: string }} An array of service item objects containing the item name, category, pricing for each quarter, and limits for each quarter, or an error message if the sheet is not found or an error occurs. Each service item is structured to allow easy access to its details throughout the application, facilitating operations such as pricing management and service categorization.
 */
function fetchServicesDataCached(data?: string): ItemData[] | { error: string } {
    try {
        if (data && typeof data === 'string') {
            log(`Received data for fetchServicesDataCached: ${data}`, false);
        } else {
            log("No data received for fetchServicesDataCached, proceeding with default cache retrieval.", false);
            data = JSON.stringify({ forceRefresh: false, servicesSheet: fetchServicesSheetNames()[0] });
        }

        const parsedData = data ? JSON.parse(data) : null;
        const forceRefresh = parsedData?.forceRefresh || false;

        const cache = CacheService.getScriptCache();
        const props = PropertiesService.getScriptProperties();

        log(`Services Sheet returns: ${JSON.stringify(parsedData)}`, false);

        const servicesSheet = parsedData?.servicesSheet === null || parsedData?.servicesSheet === undefined ? fetchServicesSheetNames()[0] : parsedData?.servicesSheet;

        log(`Using servicesSheet: ${servicesSheet}`, false);

        if (!forceRefresh) {
            const cachedString = getCachedData(SERVICES_CACHED_KEY);
            if (cachedString && cachedString !== "{}" && cachedString !== "") {
                log(`Cache hit: Services data loaded from cache. String: ${cachedString}`, false);
                return JSON.parse(cachedString) as ItemData[];
            }

            const savedProperties = props.getProperty(SERVICES_CACHED_KEY);
            if (savedProperties) {
                cache.put(SERVICES_CACHED_KEY, savedProperties, SERVER_SIDE_CACHE_AGE);
                return JSON.parse(savedProperties);
            }
        }

        log("Cache miss: Re-extracting items from Services sheet rows...", false);
        const freshServices = fetchServicesData(servicesSheet);

        if (Array.isArray(freshServices)) {
            setCachedData(SERVICES_CACHED_KEY, freshServices);
        }

        return freshServices;
    } catch (error: any) {
        log(`Error in fetchServicesDataCached: ${error.message}`, true);
        return { error: `Error occurred in fetchServicesDataCached: ${error.message}` };
    }
}

/**
 * Fetches services data from the "Services" sheet in the active spreadsheet. It reads a predefined range of rows and columns to extract various details about services, including item names, categories, pricing for each quarter, and limits for each quarter. The function processes the raw grid data to construct structured objects for each service item, handling type conversions for numbers as needed. If the sheet is not found or an error occurs during processing, it returns an error message.
 * @param sheetName The name of the sheet to fetch services data from. Defaults to the constant DEFAULT_SERVICES_SHEET, which is "Services". This allows for flexibility in case there are multiple services sheets or if the sheet name changes in the future.
 * @returns {ItemData[] | { error: string }} An array of service item objects containing the item name, category, pricing for each quarter, and limits for each quarter, or an error message if the sheet is not found or an error occurs. Each service item is structured to allow easy access to its details throughout the application, facilitating operations such as pricing management and service categorization.
 */
function fetchServicesData(sheetName: string = fetchServicesSheetNames()[0]): ItemData[] | { error: string } {
    try {
        // Fetch all sheets with the name "Services" to implement multi-service sheet support in the future
        const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

        if (allSheets.length === 0) {
            log("No sheets found in the active spreadsheet.", true);
            return { error: "No sheets found in the active spreadsheet." };
        }

        log("Found sheets: " + allSheets.map(sheet => sheet.getName()).join(", "), true);

        const servicesSheet = allSheets.find(sheet => sheet.getName() === sheetName);
        if (!servicesSheet) {
            log("Services sheet not found.", true);
            return { error: "Services sheet not found." };
        }

        const lastRow = servicesSheet.getLastRow();
        if (lastRow < SERVICES_ROW_START) {
            log("No services data found.", true);
            return { error: "No services data found." };
        }

        const servicesData: ItemData[] = [];

        // Extract data from each row
        for (let row = SERVICES_ROW_START; row <= lastRow; row++) {
            const itemName = servicesSheet.getRange(row, SERVICES_COLUMNS.itemName as number).getValue().toString().trim();
            const category = servicesSheet.getRange(row, SERVICES_COLUMNS.category as number).getValue().toString().trim() as "Income" | "Expense";
            const pricing: QuarterlyData = {};
            const limit: QuarterlyData = {};

            const pricingCols = SERVICES_COLUMNS.pricing as number[];
            const limitCols = SERVICES_COLUMNS.limit as number[];

            // Find boundaries
            const minCol = Math.min(...pricingCols, ...limitCols);
            const maxCol = Math.max(...pricingCols, ...limitCols);
            const colCount = Number((maxCol - minCol + 1).toFixed(2));

            // Get all relevant cells in a single batch to minimize API calls
            const rowValues = servicesSheet.getRange(row, minCol, 1, colCount).getValues()[0];

            // A single helper loop to populate both objects instantly from cache
            for (let i = 0; i < 4; i++) {
                const quarterKey = `Q${i + 1}` as keyof QuarterlyData;

                // Map the absolute column index to the local rowValues array index
                const priceVal = rowValues[pricingCols[i] - minCol];
                const limitVal = rowValues[limitCols[i] - minCol];

                if (priceVal !== "") pricing[quarterKey] = Number(priceVal);
                if (limitVal !== "") limit[quarterKey] = Number(limitVal);
            }

            // Only include items that have a name, category, and at least one pricing
            if (itemName && category && Object.keys(pricing).length > 0) {
                servicesData.push({
                    itemName,
                    category,
                    pricing,
                    limit
                });
            }
        }

        if (servicesData.length === 0) {
            log("No valid services data found.", true);
            return { error: "No valid services data found." };
        }

        return servicesData;
    } catch (error: any) {
        log(`Error in fetchServicesData: ${error.message}`, true);
        return { error: `Error occurred in fetchServicesData: ${error.message}` };
    }
}

/**
 * Executes a balance action based on the provided payload string.
 * @param payloadStr The JSON string containing the operation and amount to apply.
 * @returns The result of the operation or an error message.
 */
function executeServiceAction(payloadStr: string): string | void {
    try {
        // Check if a string payload was provided
        if (!payloadStr) {
            log("No payload provided for balance action.", true);
            return "No payload provided for balance action.";
        }

        // Parse the clean JSON string into a JSON object for the util function to process
        const payload = JSON.parse(payloadStr);
        const { operation, unitPrice, quantity, transactionReason = undefined, isNegativeOverride } = payload;

        if (!operation || !unitPrice || typeof unitPrice !== 'number' || !quantity || typeof quantity !== 'number') {
            log("Invalid payload. Please provide a valid operation and amount. Received - operation: " + operation + ", unitPrice: " + unitPrice + ", quantity: " + quantity, true);
            return "Invalid payload. Please provide a valid operation and amount.";
        }

        if (isNegativeOverride === true || operation === "ADD" || operation === "MULTIPLY") {
            log("Operation is additive, multiplicative, or a negative override is enabled.", false);
        } else {
            log("Operation is subtractive and no negative override is enabled. Checking for minimum balance constraints.", false);
            const personsNegative: string[] = [];

            // Fetch all selected cells in the active sheet, if they have a balance below the minimum required to remove, return a message to the user that they cannot remove more than the minimum balance
            const activeRangeList = SpreadsheetApp.getActiveSpreadsheet().getActiveRangeList();
            if (!activeRangeList) {
                log("No active range found. Please select cells to apply the operation to.", true);
                return "No active range found. Please select cells to apply the operation to.";
            }

            const ranges = activeRangeList.getRanges();

            ranges.forEach(range => {
                const sheet = range.getSheet();
                const sheetName = sheet.getName();
                // Only get numeric characters out of sheetName
                const periodName = parseInt(sheetName.replace(/\D/g, ""), 10);

                // Calculate target dimensions based on selection and overrides
                let startRow = range.getRow();
                let numRows = range.getNumRows();
                let startCol = range.getColumn();
                let numCols = range.getNumColumns();

                // Get the target range on the sheet and its corresponding A1 notation
                const targetRange = sheet.getRange(startRow, startCol, numRows, numCols);
                const targetA1 = targetRange.getA1Notation();
                const values = targetRange.getValues();

                log(`Sheet name: ${sheetName}, Range A1: ${targetA1}, Period Name: ${periodName}`, false);

                values.map((row, rowIndex) => {
                    const absoluteRowIndex = startRow + rowIndex;

                    return row.map((cell) => {
                        if (typeof cell === 'number' && !isNaN(cell)) {
                            const balance = sheet.getRange(absoluteRowIndex, BALANCE_COL).getValue();

                            if (balance === undefined || balance === null || isNaN(balance)) {
                                log(`Balance value is invalid for row ${absoluteRowIndex}. Skipping this cell.`, true);
                                return cell; // Return the original value if balance is invalid
                            }

                            const individualName = sheet.getRange(absoluteRowIndex, NAMES_COL).getValue();

                            // If the operation is subtractive and the balance minus the amount is less than the minimum required, add the individual to the list of persons with negative balances
                            if (operation === "SUBTRACT" && balance - (unitPrice * quantity) < 0) {
                                personsNegative.push(individualName);
                            }
                        }
                    });
                });
            });

            log(`Some selected cells have negative balances. Operation will not proceed unless 'Negative Override' is enabled. 
                Cells for [${personsNegative.join(" - ")}] have negative balances.`, true);
            return `Some selected cells have negative balances. Operation will not proceed unless 'Negative Override' is enabled. Cells for [${personsNegative.join(" - ")}] have negative balances.`;
        }

        // If the operation is additive, always add to EARNINGS_COL
        // If the operation is subtractive, always add to EXPENDITURES_COL
        if (operation === "ADD" || operation === "MULTIPLY") {
            return applyMathToSelection(
                operation,
                Number(unitPrice),
                Number(quantity),
                false,
                transactionReason,
                undefined,
                false,
                undefined,
                EARNINGS_COL
            ).toString();
        } else if (operation === "SUBTRACT" || operation === "DIVIDE") {
            // For subtractive operations, we will add to the EXPENDITURES_COL instead of removing from the EARNINGS_COL, as this is a more accurate representation of the financial action being taken.
            // Divide stays the same since we are still applying the same logic of adding to expenditures, just with a different operation.
            return applyMathToSelection(
                operation == "SUBTRACT" ? "ADD" : "DIVIDE",
                Number(unitPrice),
                Number(quantity),
                false,
                transactionReason,
                undefined,
                true,
                undefined,
                EXPENDITURES_COL
            ).toString();
        }

    } catch (error: any) {
        log(`Error in executeBalanceAction: ${error.message}`, true);
        return `Error occurred in executeBalanceAction: ${error.message}`;
    }
}

/**
 * Fetches the names of all sheets in the active spreadsheet that contain the word "services" in their name. This function is useful for dynamically populating a dropdown or selection list with available services sheets, allowing users to switch between different services data sets easily. If no such sheets are found or an error occurs during the process, it returns an error message.
 * @returns {string[] | { error: string }} An array of sheet names that contain the word "services" in their name, or an error message if no such sheets are found or an error occurs. This allows for flexible management of multiple services sheets within the application, enabling users to select and view different sets of services data as needed.
 */
function fetchServicesSheetNames(): string[] | { error: string } {
    try {
        const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

        if (allSheets.length === 0) {
            log("No sheets found in the active spreadsheet.", true);
            return { error: "No sheets found in the active spreadsheet." };
        }

        let servicesSheetNames = allSheets.filter(sheet => sheet.getName().toLowerCase().includes("services")).map(sheet => sheet.getName());

        if (servicesSheetNames.length === 0) {
            log("No services sheets found in the active spreadsheet.", true);
            return { error: "No services sheets found in the active spreadsheet." };
        }

        // The first option in the dropdown should be the default services sheet, if it exists, followed by any other services sheets found

        if (servicesSheetNames.includes(DEFAULT_SERVICES_SHEET)) {
            servicesSheetNames = [DEFAULT_SERVICES_SHEET, ...servicesSheetNames.filter(name => name !== DEFAULT_SERVICES_SHEET)];
        }

        return servicesSheetNames;
    } catch (error: any) {
        log(`Error in fetchServicesSheetNames: ${error.message}`, true);
        return { error: `Error occurred in fetchServicesSheetNames: ${error.message}` };
    }
}