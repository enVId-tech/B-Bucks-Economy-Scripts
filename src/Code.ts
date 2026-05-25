// Coded by Erick Tran for Mr. Banderas, 2026
// This file contains the main code for the B-Bucks Economy Scripts project, including menu creation and dialog management.
// Script licensed under the MIT License.
// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts

/**
 * Creates a custom dropdown menu in the Google Sheets UI when the spreadsheet is opened, 
 * allowing users to access various functions related to the B-Bucks Economy Scripts project.
 */
function onOpen(): void {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('The Central Bank of Banderas')
    .addItem('Open Manual B-Bucks Modifier', 'openManualModifier')
    .addItem('Open Service Provider', 'placeholderFunction')
    .addItem('Open Investment Banking', 'openInvestmentsManager')
    .addSeparator()
    .addItem('View Transactions Records', 'placeholderFunction')
    .addSeparator()
    .addItem('Manual Timestamp Update', 'placeholderFunction')
    .addItem('Record Daily Data', 'placeholderFunction')
    .addItem('Reset Historical Records', 'placeholderFunction')
    .addSeparator()
    .addItem('Banking Settings', 'placeholderFunction')
    .addToUi();
}

/**
 * Opens a dialog for manually modifying B-Bucks balances, allowing users to add or subtract amounts from specific cells.
 */
function openManualModifier(): void {
  var html = HtmlService.createHtmlOutputFromFile('ManualModifier')
    .setTitle("Bank of Banderas - Manual Balance Manager")
    .setWidth(390)
    .setHeight(550);

  SpreadsheetApp.getUi().showModelessDialog(html, "Bank of Banderas - Manual Balance Manager");
}

/**
 * Opens a dialog for managing investments, allowing users to view and modify investment records.
 */
function openInvestmentsManager(): void {
  var html = HtmlService.createHtmlOutputFromFile('InvestmentsManager')
    .setTitle("Bank of Banderas - Investments Manager")
    .setWidth(600)
    .setHeight(700);

  SpreadsheetApp.getUi().showModelessDialog(html, "Bank of Banderas - Investments Manager");
}

/**
 * A placeholder function for menu items that are not yet implemented.
 */
function placeholderFunction(): void {
  SpreadsheetApp.getUi().alert("This function is not yet implemented. Check back later!");
}