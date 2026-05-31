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

        if (!operation || !value || typeof value !== 'number') {
            Logger.log("Invalid payload. Please provide a valid operation and amount.");
            SpreadsheetApp.getUi().alert(`Invalid payload. Please provide a valid operation and amount. Information received - operation: ${operation}, amount: ${value}`);
            return "Invalid payload. Please provide a valid operation and amount.";
        }

        return applyMathToSelection(operation, value).toString();
    } catch (error: any) {
        SpreadsheetApp.getUi().alert(`Error occurred in executeBalanceAction: ${error.message}`);
        return `Error occurred in executeBalanceAction: ${error.message}`;
    }
}