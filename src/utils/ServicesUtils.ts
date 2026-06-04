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
 * @param forceRefresh A boolean flag indicating whether to bypass the cache and fetch fresh data from the sheet. Defaults to false, meaning it will use cached data if available for faster access. Setting this to true will force the function to read directly from the sheet and update the cache with the new data.
 * @returns {ItemData[] | { error: string }} An array of service item objects containing the item name, category, pricing for each quarter, and limits for each quarter, or an error message if the sheet is not found or an error occurs. Each service item is structured to allow easy access to its details throughout the application, facilitating operations such as pricing management and service categorization.
 */
function fetchServicesDataCached(forceRefresh: boolean = false): ItemData[] | { error: string } {
    try {
        const CACHE_KEY = "cachedServices";
        const cache = CacheService.getScriptCache();
        const props = PropertiesService.getScriptProperties();

        if (!forceRefresh) {
            const cachedString = getCachedData(CACHE_KEY);
            if (cachedString && cachedString !== "{}" && cachedString !== "") {
                // SpreadsheetApp.getUi().alert(`Cache hit: Services data loaded from cache. String: ${cachedString}`);
                return JSON.parse(cachedString) as ItemData[];
            }

            const savedProperties = props.getProperty(CACHE_KEY);
            if (savedProperties) {
                cache.put(CACHE_KEY, savedProperties, 21600);
                return JSON.parse(savedProperties);
            }
        }

        console.log("Cache miss: Re-extracting items from Services sheet rows...");
        // SpreadsheetApp.getUi().alert("Cache miss: Re-extracting items from Services sheet rows...");
        const freshServices = fetchServicesData();

        if (Array.isArray(freshServices)) {
            setCachedData(CACHE_KEY, freshServices);
        }

        return freshServices;
    } catch (error: any) {
        Logger.log(`Error occurred in fetchServicesDataCached: ${error.message}`);
        SpreadsheetApp.getUi().alert(`Error occurred in fetchServicesDataCached: ${error.message}`);
        return { error: `Error occurred in fetchServicesDataCached: ${error.message}` };
    }
}

/**
 * Fetches services data from the "Services" sheet in the active spreadsheet. It reads a predefined range of rows and columns to extract various details about services, including item names, categories, pricing for each quarter, and limits for each quarter. The function processes the raw grid data to construct structured objects for each service item, handling type conversions for numbers as needed. If the sheet is not found or an error occurs during processing, it returns an error message.
 * @returns {ItemData[] | { error: string }} An array of service item objects containing the item name, category, pricing for each quarter, and limits for each quarter, or an error message if the sheet is not found or an error occurs. Each service item is structured to allow easy access to its details throughout the application, facilitating operations such as pricing management and service categorization.
 */
function fetchServicesData(): ItemData[] | { error: string } {
    try {
        const SHEET_NAME = "Services";
        const ROW_START = 3;

        const columns: { [key: string]: number | number[] } = {
            itemName: 1,
            category: 2,
            // Represents Q1-Q4 pricing columns
            pricing: [3, 4, 5, 6],
            // Represents Q1-Q4 max per person columns
            limit: [7, 8, 9, 10]
        }

        const servicesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
        if (!servicesSheet) {
            Logger.log("Services sheet not found.");
            SpreadsheetApp.getUi().alert("Services sheet not found.");
            return { error: "Services sheet not found." };
        }

        const lastRow = servicesSheet.getLastRow();
        if (lastRow < ROW_START) {
            Logger.log("No services data found.");
            SpreadsheetApp.getUi().alert("No services data found.");
            return { error: "No services data found." };
        }

        const servicesData: ItemData[] = [];

        // Extract data from each row
        for (let row = ROW_START; row <= lastRow; row++) {
            const itemName = servicesSheet.getRange(row, columns.itemName as number).getValue().toString().trim();
            const category = servicesSheet.getRange(row, columns.category as number).getValue().toString().trim() as "Income" | "Expense";
            const pricing: QuarterlyData = {};
            const limit: QuarterlyData = {};

            // Extract pricing for Q1-Q4
            (columns.pricing as number[]).forEach((col, index) => {
                const value = servicesSheet.getRange(row, col).getValue();
                if (value !== "") {
                    pricing[`Q${index + 1}` as keyof QuarterlyData] = Number(value);
                }
            });

            // Extract limits for Q1-Q4
            (columns.limit as number[]).forEach((col, index) => {
                const value = servicesSheet.getRange(row, col).getValue();
                if (value !== "") {
                    limit[`Q${index + 1}` as keyof QuarterlyData] = Number(value);
                }
            });

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
            Logger.log("No valid services data found.");
            SpreadsheetApp.getUi().alert("No valid services data found.");
            return { error: "No valid services data found." };
        }

        return servicesData;
    } catch (error: any) {
        SpreadsheetApp.getUi().alert(`Error occurred in fetchServicesData: ${error.message}`);
        return { error: `Error occurred in fetchServicesData: ${error.message}` };
    }
}