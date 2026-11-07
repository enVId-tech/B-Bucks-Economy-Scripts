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
    classLimit: QuarterlyData;
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
 * @param sheetName The name of the sheet to fetch services data from. Defaults to the constant getServicesSheetName(), which is "Services". This allows for flexibility in case there are multiple services sheets or if the sheet name changes in the future.
 * @returns {ItemData[] | { error: string }} An array of service item objects containing the item name, category, pricing for each quarter, and limits for each quarter, or an error message if the sheet is not found or an error occurs. Each service item is structured to allow easy access to its details throughout the application, facilitating operations such as pricing management and service categorization.
 */
function fetchServicesData(sheetName: string = fetchServicesSheetNames()[0]): ItemData[] | { error: string } {
    try {
        // Fetch all sheets with the name "Services" to implement multi-service sheet support in the future
        const allSheets = getAllSheets();

        if (allSheets.length === 0) {
            log("No sheets found in the active spreadsheet.", true);
            return { error: "No sheets found in the active spreadsheet." };
        }

        log("Found sheets: " + allSheets.map(sheet => sheet.getName()).join(", "), false);

        const servicesSheet = allSheets.find(sheet => sheet.getName() === sheetName);
        if (!servicesSheet) {
            log("Services sheet not found.", true);
            return { error: "Services sheet not found." };
        }

        const lastRow = servicesSheet.getLastRow();
        if (lastRow < getServicesRowStart()) {
            log("No services data found.", true);
            return { error: "No services data found." };
        }

        const servicesData: ItemData[] = [];

        // Read the whole table with ONE range call instead of three calls per row.
        const pricingCols = SERVICES_COLUMNS.pricing as number[];
        const limitCols = SERVICES_COLUMNS.limit as number[];
        const classLimitCols = SERVICES_COLUMNS.classLimit as number[];
        const nameCol = SERVICES_COLUMNS.itemName as number;
        const categoryCol = SERVICES_COLUMNS.category as number;

        const firstCol = Math.min(nameCol, categoryCol, ...pricingCols, ...limitCols, ...classLimitCols);
        const lastCol = Math.max(nameCol, categoryCol, ...pricingCols, ...limitCols, ...classLimitCols);
        const tableValues = servicesSheet
            .getRange(getServicesRowStart(), firstCol, lastRow - getServicesRowStart() + 1, lastCol - firstCol + 1)
            .getValues();

        for (const rowValues of tableValues) {
            const itemName = rowValues[nameCol - firstCol].toString().trim();
            const category = rowValues[categoryCol - firstCol].toString().trim() as "Income" | "Expense";
            const pricing: QuarterlyData = {};
            const limit: QuarterlyData = {};
            const classLimit: QuarterlyData = {};

            for (let i = 0; i < 4; i++) {
                const quarterKey = `Q${i + 1}` as keyof QuarterlyData;

                const priceVal = rowValues[pricingCols[i] - firstCol];
                const limitVal = rowValues[limitCols[i] - firstCol];
                const classLimitVal = rowValues[classLimitCols[i] - firstCol];

                if (priceVal !== "") pricing[quarterKey] = Number(priceVal);
                if (limitVal !== "") limit[quarterKey] = Number(limitVal);
                if (classLimitVal !== "") classLimit[quarterKey] = Number(classLimitVal);
            }

            // Only include items that have a name, category, and at least one pricing
            if (itemName && category && Object.keys(pricing).length > 0) {
                servicesData.push({
                    itemName,
                    category,
                    pricing,
                    limit,
                    classLimit
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
    // Money moves only while holding the transaction lock, see LockUtils.ts
    return runWithTransactionLock(() => executeServiceActionInsideLock_(payloadStr), "Another transaction is still running. Nothing was changed, please try again.");
}

/** Body of executeServiceAction(). Only call it through executeServiceAction(), which holds the transaction lock. */
function executeServiceActionInsideLock_(payloadStr: string): string | void {
    try {
        // Check if a string payload was provided
        if (!payloadStr) {
            log("No payload provided for balance action.", true);
            return "No payload provided for balance action.";
        }

        // Parse the clean JSON string into a JSON object for the util function to process
        const payload = JSON.parse(payloadStr);
        const { operation, unitPrice, quantity, transactionReason = undefined, isRestrictionOverride, itemName, quarter } = payload;

        if (!operation || !unitPrice || typeof unitPrice !== 'number' || !quantity || typeof quantity !== 'number') {
            log("Invalid payload. Please provide a valid operation and amount. Received - operation: " + operation + ", unitPrice: " + unitPrice + ", quantity: " + quantity, true);
            return "Invalid payload. Please provide a valid operation and amount.";
        }

        // Max/Person cap: counted from the transaction log, skipped when the admin ticked "Override Restrictions".
        if (isRestrictionOverride !== true && itemName && quarter) {
            const limitMessage = checkServiceLimit(String(itemName), quarter as QuarterId, quantity);
            if (limitMessage) {
                log(limitMessage, true); // shown to the admin as a dialog, the Services page does not repeat it
                return limitMessage;
            }
        }

        if (isRestrictionOverride === true || operation === "ADD" || operation === "MULTIPLY") {
            log("Operation is additive, multiplicative, or 'Override Restrictions' is enabled.", false);
        } else {
            log("Operation is subtractive and no restriction override is enabled. Checking for minimum balance constraints.", false);
            const personsNegative = new Set<string>();

            // Fetch all selected cells in the active sheet, if they have a balance below the minimum required to remove, return a message to the user that they cannot remove more than the minimum balance
            const activeRangeList = getSS().getActiveRangeList();
            if (!activeRangeList) {
                log("No active range found. Please select cells to apply the operation to.", true);
                return "No active range found. Please select cells to apply the operation to.";
            }

            const ranges = activeRangeList.getRanges();

            // The charge goes on the expenditures cell of every selected ROW (whatever cells were clicked), so the check is per row too.
            // Selecting a name cell used to skip the check because only numeric cells were looked at.
            const checkedRows = new Set<string>();

            ranges.forEach(range => {
                const sheet = range.getSheet();
                const sheetName = sheet.getName();
                const startRow = range.getRow();
                const numRows = range.getNumRows();

                log(`Checking balances on ${sheetName}!${range.getA1Notation()}`, false);

                // One read per column instead of two reads per cell
                const names = sheet.getRange(startRow, getNamesColumn(), numRows, 1).getValues();
                const balances = sheet.getRange(startRow, getBalanceColumn(), numRows, 1).getValues();
                const expenditures = sheet.getRange(startRow, getExpendituresColumn(), numRows, 1).getValues();

                for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
                    const absoluteRowIndex = startRow + rowOffset;
                    const rowKey = `${sheetName}|${absoluteRowIndex}`;
                    if (checkedRows.has(rowKey)) continue;
                    checkedRows.add(rowKey);

                    // Rows the charge would skip anyway (no student, or no numeric expenditures cell) are not checked.
                    const individualName = String(names[rowOffset][0] ?? "").trim();
                    const expenditureCell = expenditures[rowOffset][0];
                    if (!individualName || typeof expenditureCell !== 'number' || isNaN(expenditureCell)) continue;

                    const balance = balances[rowOffset][0];
                    if (typeof balance !== "number" || !Number.isFinite(balance)) {
                        log(`Balance value is invalid for ${individualName}.`, false);
                        personsNegative.add(`${individualName} (invalid balance)`);
                        continue;
                    }

                    // If the operation is subtractive and the balance minus the amount is less than zero, the student can't afford it
                    if (operation === "SUBTRACT" && balance - (unitPrice * quantity) < 0) {
                        personsNegative.add(individualName);
                    }
                }
            });

            if (personsNegative.size > 0) {
                const negativePeople = Array.from(personsNegative).join(" - ");
                log(`Some selected cells have negative balances or invalid account data. Operation will not proceed unless 'Override Restrictions' is enabled. Cells for [${negativePeople}] cannot be processed.`, true);
                return `Some selected cells have negative balances or invalid account data. Operation will not proceed unless 'Override Restrictions' is enabled. Cells for [${negativePeople}] cannot be processed.`;
            }
        }

        // If the operation is additive, always add to getEarningsColumn()
        // If the operation is subtractive, always add to getExpendituresColumn()
        if (operation === "ADD" || operation === "MULTIPLY") {
            return applyMathToSelection(
                operation,
                Number(unitPrice),
                Number(quantity),
                false,
                transactionReason,
                undefined,
                undefined,
                getEarningsColumn()
            ).toString();
        } else if (operation === "SUBTRACT" || operation === "DIVIDE") {
            // For subtractive operations, we will add to the getExpendituresColumn() instead of removing from the getEarningsColumn(), as this is a more accurate representation of the financial action being taken.
            // Divide stays the same since we are still applying the same logic of adding to expenditures, just with a different operation.
            return applyMathToSelection(
                operation == "SUBTRACT" ? "ADD" : "DIVIDE",
                Number(unitPrice),
                Number(quantity),
                false,
                transactionReason,
                undefined,
                undefined,
                getExpendituresColumn()
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
        const allSheets = getAllSheets();

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

        if (servicesSheetNames.includes(getServicesSheetName())) {
            servicesSheetNames = [getServicesSheetName(), ...servicesSheetNames.filter(name => name !== getServicesSheetName())];
        }

        return servicesSheetNames;
    } catch (error: any) {
        log(`Error in fetchServicesSheetNames: ${error.message}`, true);
        return { error: `Error occurred in fetchServicesSheetNames: ${error.message}` };
    }
}