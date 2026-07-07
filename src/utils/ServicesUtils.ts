// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. 
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains utility functions for various operations in the B-Bucks Economy Scripts project, including functions related to services management and other financial operations.

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
            data = JSON.stringify({ forceRefresh: false });
        }

        const parsedData = data ? JSON.parse(data) : null;
        const forceRefresh = parsedData?.forceRefresh || false;

        const cache = CacheService.getScriptCache();
        const props = PropertiesService.getScriptProperties();

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
        const freshServices = fetchServicesData();

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
 * @returns {ItemData[] | { error: string }} An array of service item objects containing the item name, category, pricing for each quarter, and limits for each quarter, or an error message if the sheet is not found or an error occurs. Each service item is structured to allow easy access to its details throughout the application, facilitating operations such as pricing management and service categorization.
 */
function fetchServicesData(): ItemData[] | { error: string } {
    try {
        const servicesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DEFAULT_SERVICES_SHEET);
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

                if (priceVal !== "") pricing[quarterKey] = Number(priceVal.toFixed(2));
                if (limitVal !== "") limit[quarterKey] = Number(limitVal.toFixed(2));
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
        const { operation, unitPrice, quantity, transactionReason = undefined } = payload;
        
        if (!operation || !unitPrice || typeof unitPrice !== 'number' || !quantity || typeof quantity !== 'number') {
            log("Invalid payload. Please provide a valid operation and amount. Received - operation: " + operation + ", unitPrice: " + unitPrice + ", quantity: " + quantity, true);
            return "Invalid payload. Please provide a valid operation and amount.";
        }

        const commentOnExpenditures = payload.operation === "ADD" || payload.operation === "MULTIPLY" ? true : false;

        return applyMathToSelection(
            operation, 
            unitPrice,
            quantity,
            false,
            transactionReason,
            undefined,  
            commentOnExpenditures).toString();
    } catch (error: any) {
        log(`Error in executeBalanceAction: ${error.message}`, true);
        return `Error occurred in executeBalanceAction: ${error.message}`;
    }
}