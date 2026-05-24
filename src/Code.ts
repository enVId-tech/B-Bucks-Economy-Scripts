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

function openManualModifier(): void {
  var html = HtmlService.createHtmlOutputFromFile('ManualModifier')
    .setTitle("Bank of Banderas - Manual Balance Manager")
    .setWidth(390)  
    .setHeight(550);

  SpreadsheetApp.getUi().showModelessDialog(html, "Bank of Banderas - Manual Balance Manager");
}

function openInvestmentsManager(): void {
  var html = HtmlService.createHtmlOutputFromFile('InvestmentsManager')
    .setTitle("Bank of Banderas - Investments Manager")
    .setWidth(600)  
    .setHeight(700);

  SpreadsheetApp.getUi().showModelessDialog(html, "Bank of Banderas - Investments Manager");
}

function placeholderFunction(): void {
  SpreadsheetApp.getUi().alert("This function is not yet implemented. Check back later!");
}