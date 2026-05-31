// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. 
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains functions related to the manual balance manager, allowing users to perform operations on selected cells in the Google Sheets UI based on a provided payload that specifies the operation and amount to apply.

function executeBalanceAction(payloadStr: string): string | void {
    try {
        // Check if a string payload was provided
        if (!payloadStr) {
            Logger.log("No payload provided for balance action.");
            SpreadsheetApp.getUi().alert("No payload provided for balance action.");
            return "No payload provided for balance action.";
        }

        // Parse the clean JSON string into a JSON object for the util function to process
        const payload = JSON.parse(payloadStr);
        const operation = payload.operation;
        const value = payload.amount;
        const transactionReason: string | undefined = payload.transactionReason;

        if (!operation || !value || typeof value !== 'number') {
            Logger.log("Invalid payload. Please provide a valid operation and amount.");
            SpreadsheetApp.getUi().alert(`Invalid payload. Please provide a valid operation and amount. Information received - operation: ${operation}, amount: ${value}`);
            return "Invalid payload. Please provide a valid operation and amount.";
        }

        return applyMathToSelection(operation, value, true, transactionReason).toString();
    } catch (error: any) {
        SpreadsheetApp.getUi().alert(`Error occurred in executeBalanceAction: ${error.message}`);
        return `Error occurred in executeBalanceAction: ${error.message}`;
    }
}