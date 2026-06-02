// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. All rights reserved.
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains utility functions for data caching and dialog management in the B-Bucks Economy Scripts project. It provides functions to retrieve cached data, launch modeless dialogs with initial data payloads, and centralizes the logic for opening various dialogs across the application.

/**
 * Retrieves cached data for a given key. It first checks the in-memory cache, then the script properties for permanent storage. If found in properties, it updates the cache for future quick access. If no data is found, it returns an error message as a JSON string.
 * @param cacheKey The key for the cached data to retrieve. This should correspond to a specific dataset like "cachedIndividuals", "cachedServices", etc.
 * @returns A JSON string containing the cached data or an error message if no data is found. The expected structure of the returned data depends on the cacheKey used, but it generally includes relevant information for the application's operations, such as lists of individuals, services, settings, or transactions.
 */
function getCachedData(cacheKey: string): string {
    try {
        const cache = CacheService.getScriptCache();

        const cached = cache.get(cacheKey);
        if (cached) return cached;

        const permanentData = PropertiesService.getScriptProperties().getProperty(cacheKey);
        if (permanentData) {
            cache.put(cacheKey, permanentData, 1500);
            return permanentData;
        }

        return JSON.stringify({ error: "No cached data found" });
    } catch (error: any) {
        Logger.log(`Error in getCachedData for key ${cacheKey}: ${error.message}`);
        return JSON.stringify({ error: `Error retrieving cached data: ${error.message}` });
    }
}

/**
 * Launches a modeless dialog in the Google Sheets UI using a specified HTML template. The dialog is populated with initial data fetched from the cache, which includes individuals, services, settings, and transactions. The function takes parameters for the template name, dialog title, and dimensions (width and height) to customize the appearance of the dialog. This utility function centralizes the logic for opening various dialogs across the application, ensuring consistency in how data is passed and how dialogs are displayed.
 * @param templateName The name of the HTML template file (without the .html extension) to use for the dialog's content. This should correspond to a file in the project's HTML directory, such as "BalanceManager", "InvestmentsManager", etc.
 * @param title The title to display on the dialog window. This should be descriptive of the dialog's purpose, such as "Bank of Banderas - Manual Balance Manager".
 * @param width The width of the dialog in pixels. This should be set based on the expected content and layout of the dialog to ensure a good user experience without excessive scrolling or wasted space.
 * @param height The height of the dialog in pixels. Similar to width, this should be chosen to accommodate the content of the dialog while maintaining usability and aesthetics.
 */
function launchModelessDialog(templateName: string, title: string, width: number, height: number): void {
    const template = HtmlService.createTemplateFromFile(templateName);

    const globalData = {
        "individuals": JSON.parse(getCachedData("cachedIndividuals")),
        "services": JSON.parse(getCachedData("cachedServices")),
        "settings": JSON.parse(getCachedData("cachedSettings")),
        "transactions": JSON.parse(getCachedData("cachedTransactions"))
    }

    template.initialServerPayload = JSON.stringify(globalData);

    const html = template.evaluate()
        .setTitle(title)
        .setWidth(width)
        .setHeight(height);

    SpreadsheetApp.getUi().showModelessDialog(html, title);
}