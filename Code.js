function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('The Central Bank of Banderas')
      .addItem('Open Manual B-Bucks Modifier', 'openManualModifier')
      .addItem('Open Service Provider', 'mySecondFunction')
      .addItem('Open Investment Banking', 'openInvestmentsManager')
      .addSeparator()
      .addItem('View Transactions Records', 'mySecondFunction')
      .addSeparator()
      .addItem('Manual Timestamp Update', 'mySecondFunction')
      .addItem('Record Daily Data', 'mySecondFunction')
      .addItem('Reset Historical Records', 'mySecondFunction')
      .addSeparator()
      .addItem('Banking Settings', 'mySecondFunction')
      .addToUi();
}

function openManualModifier() {
  var html = HtmlService.createHtmlOutputFromFile('ManualModifier')
    .setTitle("Bank of Banderas - Manual Balance Manager")
    .setWidth(390)  
    .setHeight(550);

  SpreadsheetApp.getUi().showModelessDialog(html, "Bank of Banderas - Manual Balance Manager");
}

function openInvestmentsManager() {
  var html = HtmlService.createHtmlOutputFromFile('InvestmentsManager')
    .setTitle("Bank of Banderas - Investments Manager")
    .setWidth(600)  
    .setHeight(700);

  SpreadsheetApp.getUi().showModelessDialog(html, "Bank of Banderas - Investments Manager");
}



function myFirstFunction() {
  SpreadsheetApp.getUi().alert('You clicked the first item!');
}

function mySecondFunction() {
  SpreadsheetApp.getUi().alert('You clicked the second item!');
}

function mySubFunction() {
  SpreadsheetApp.getUi().alert('You clicked the sub-menu item!');
}