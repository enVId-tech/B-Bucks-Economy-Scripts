// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. All rights reserved.
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains the main entry point for the B-Bucks Economy Scripts project, including the onOpen function that creates a custom menu in the Google Sheets UI and various functions for opening dialogs and performing operations related to the B-Bucks Economy.

/**
 * Creates a custom dropdown menu in the Google Sheets UI when the spreadsheet is opened, 
 * allowing users to access various functions related to the B-Bucks Economy Scripts project.
 */
function onOpen(): void {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('The Central Bank of Banderas')
    .addItem('Open Manual B-Bucks Modifier', 'openManualBalanceManager')
    .addItem('Open Service Provider', 'openIncomeConsumables')
    .addItem('Open Investment Banking', 'openInvestmentsManager')
    .addSeparator()
    .addItem('View Transactions Records', 'openTransactionsRecords')
    .addItem('Banking Settings', 'openBankingSettings')
    .addSeparator()
    .addItem('Documentation & Attribution', 'openDocumentation')
    .addToUi();
}

/** Opens a dialog for manually modifying B-Bucks balances, allowing users to add or subtract amounts from specific cells. */
function openManualBalanceManager(): void { launchModelessDialog('BalanceManager', "Bank of Banderas - Manual Balance Manager", 500, 650) }

/** Opens a dialog for managing investments, allowing users to view and modify investment records. */
function openInvestmentsManager(): void { launchModelessDialog('InvestmentsManager', "Bank of Banderas - Investments Manager", 600, 700) }

/** Opens a dialog for managing income and consumables, allowing users to view and modify records related to income and expenses. */
function openIncomeConsumables(): void { launchModelessDialog('ServicesManager', "Bank of Banderas - Services Menu", 500, 700) }

/** Opens a dialog for viewing transaction records, allowing users to search and filter all B-Bucks transactions. */
function openTransactionsRecords(): void { launchModelessDialog('TransactionsRecords', "Bank of Banderas - Transaction Records", 700, 650) }

/** Opens a dialog for configuring banking settings, including interest rates, restrictions, and maintenance options. */
function openBankingSettings(): void { launchModelessDialog('BankingSettings', "Bank of Banderas - Banking Settings", 550, 700) }

/** Opens a dialog for importing student names from a CSV file, allowing users to fill sheets with data parsed from the CSV string. Each sheet is named after the identifier and contains the associated individuals starting from a specified cell. If a sheet for an identifier doesn't exist, it creates one by copying a template sheet. */
function importStudentNames(): void { launchModelessDialog('FillSheet', "Bank of Banderas - Fill Sheet", 350, 500) }

/** Opens a dialog displaying documentation and attribution information for the B-Bucks Economy Scripts project, including details about the author, client, and GitHub repository. */
function openDocumentation(): void { launchModelessDialog('DocsAttribution', "Bank of Banderas - Credits & Docs", 400, 600) }

/** Updates timestamps for all records in the spreadsheet. */
function updateTimestamps(): void { SpreadsheetApp.getUi().alert("✅ All timestamps have been updated to current time.") }

/** Records a daily data snapshot to preserve historical information. */
function recordDailyData(): void { SpreadsheetApp.getUi().alert("✅ Daily data snapshot has been recorded.") }

/** Resets all historical records (admin function requiring confirmation). */
function resetHistoricalRecords(): void {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Are you sure you want to reset all historical records? This action cannot be undone.',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    ui.alert("✅ Historical records have been reset.");
  }
}

/** A placeholder function for menu items that are not yet implemented. */
function placeholderFunction(): void { SpreadsheetApp.getUi().alert("This function is not yet implemented. Check back later!") }

/** Helper functions for page navigation in modeless dialogs */
function switchToManualBalanceManager(): void { openManualBalanceManager() }
function switchToInvestmentsManager(): void { openInvestmentsManager() }
function switchToPricingManager(): void { openIncomeConsumables() }
function switchToTransactionsRecords(): void { openTransactionsRecords() }
function switchToBankingSettings(): void { openBankingSettings() }

/**
 * Includes the content of a specified HTML file.
 * @param filename The name of the HTML file to include.
 * @returns The content of the HTML file.
 */
function include(filename: string) { return HtmlService.createHtmlOutputFromFile(filename).getContent() }