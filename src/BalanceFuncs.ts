function executeBalanceAction(payloadStr: string): void {
    try {
        // Check if a string payload was provided
        if (!payloadStr) {
            Logger.log("No payload provided for balance action.");
            SpreadsheetApp.getUi().alert("No payload provided for balance action.");
            return;
        }

        // Parse the clean JSON string into a JSON object for the util function to process
        const payload = JSON.parse(payloadStr);
        const operation = payload.operation;
        const value = payload.amount;

        if (!operation || value === undefined) {
            Logger.log("Invalid payload. Please provide a valid operation and amount.");
            SpreadsheetApp.getUi().alert("Invalid payload. Please provide a valid operation and amount.");
            return;
        }

        applyMathToSelection(operation, value);
    } catch (error: any) {
        SpreadsheetApp.getUi().alert(error.message);
    }
}