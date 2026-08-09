/**
 * Reads the Historical Economic Recordbook sheet starting at row 11 and maps
 * the structural matrix directly for client-side Chart parsing.
 */
function fetchHistoricalRecordbookData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Economic Records') || ss.getSheets()[0];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 11) {
    return JSON.stringify([]);
  }

  // Fetch block from A11 through Column P
  const rangeValues = sheet.getRange(11, 1, lastRow - 10, 16).getValues();
  const records: Array<{
    date: string;
    dateStr: string;
    sheet: string;
    totalBalances: number;
    totalGross: number;
    totalNet: number;
    totalDebt: number;
    totalReturns: number;
    totalInvested: number;
    m1Mean: number;
    m2Mean: number;
    avgGross: number;
    avgNet: number;
    avgDebt: number;
    avgReturns: number;
    avgInvested: number;
  }> = [];

  for (let i = 0; i < rangeValues.length; i++) {
    const row = rangeValues[i];
    const rawDate = row[0]; // Col A (Date)
    const sheetPeriod = row[1]; // Col B (Sheet)

    if (!rawDate || !sheetPeriod) continue;

    const formattedDate = rawDate instanceof Date ? rawDate.toISOString().split('T')[0] : rawDate;
    const displayDateStr = rawDate instanceof Date ? rawDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) : String(rawDate);

    records.push({
      date: formattedDate,
      dateStr: displayDateStr,
      sheet: String(sheetPeriod).trim(),
      totalBalances: Number(row[2]) || 0,   // Col C
      totalGross: Number(row[3]) || 0,      // Col D
      totalNet: Number(row[4]) || 0,        // Col E
      totalDebt: Number(row[5]) || 0,       // Col F
      totalReturns: Number(row[6]) || 0,    // Col G
      totalInvested: Number(row[7]) || 0,   // Col H
      m1Mean: Number(row[9]) || 0,          // Col J
      m2Mean: Number(row[10]) || 0,         // Col K
      avgGross: Number(row[11]) || 0,       // Col L
      avgNet: Number(row[12]) || 0,         // Col M
      avgDebt: Number(row[13]) || 0,        // Col N
      avgReturns: Number(row[14]) || 0,     // Col O
      avgInvested: Number(row[15]) || 0     // Col P
    });
  }

  return JSON.stringify(records);
}