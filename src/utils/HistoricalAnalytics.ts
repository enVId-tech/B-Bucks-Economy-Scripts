// Coded by Erick Tran for Mr. Banderas, 2026
// Copyright (c) 2026 Erick Tran. 
// This file is licensed under the MIT License, check the LICENSE file for details.

// GitHub Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts
// This file contains utility functions for fetching and processing historical economic data from the "Economic Records" sheet in the B-Bucks Economy Scripts project. It provides a structured JSON payload for client-side chart parsing and analysis.

interface MetricHeader {
  key: string;
  label: string;
  colIndex: number;
}

interface HistoricalRecordPayload {
  headers: MetricHeader[];
  records: Array<{
    date: string;
    dateStr: string;
    sheet: string;
    metrics: Record<string, number | null>;
  }>;
}

function fetchHistoricalRecordbookData(): string {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  // Target all sheets matching the "records" search pattern
  const recordSheets = allSheets.filter(sheet => {
    const sheetNameRegex = new RegExp(SHEET_NAME_PATTERN, 'i');
    return sheetNameRegex.test(sheet.getName());
  });

  if (recordSheets.length === 0) {
    return JSON.stringify({ headers: [], records: [] });
  }

  const globalHeadersMap = new Map<string, MetricHeader>();
  const allRecords: HistoricalRecordPayload['records'] = [];

  // Helper to convert header text into a consistent JSON key
  const sanitizeKey = (text: string, colIdx: number): string => {
    const cleaned = text.replace(/[^a-zA-Z0-9]/g, '');
    return cleaned ? cleaned.charAt(0).toLowerCase() + cleaned.slice(1) : `col_${colIdx}`;
  };

  recordSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // Ensure the sheet has enough rows and columns to meet minimum structural requirements
    if (lastRow < HISTORICAL_RECORDS_ROW_START || lastCol < METRIC_DATA_START_COL) return;

    // Fetch header row starting at HEADER_ROW_INDEX from METRIC_DATA_START_COL to end of sheet
    const numHeaderCols = lastCol - (METRIC_DATA_START_COL - 1);
    const headerValues = sheet.getRange(HEADER_ROW_INDEX, METRIC_DATA_START_COL, 1, numHeaderCols).getValues()[0];
    const sheetHeaderKeys: string[] = [];

    headerValues.forEach((headerVal, idx) => {
      const colIndex = idx + METRIC_DATA_START_COL;
      const rawHeaderStr = String(headerVal || '').trim();
      
      if (rawHeaderStr) {
        const key = sanitizeKey(rawHeaderStr, colIndex);
        sheetHeaderKeys[idx] = key;

        if (!globalHeadersMap.has(key)) {
          globalHeadersMap.set(key, {
            key: key,
            label: rawHeaderStr,
            colIndex: colIndex
          });
        }
      } else {
        sheetHeaderKeys[idx] = `col_${colIndex}`;
      }
    });

    // Fetch all data rows starting at DATA_START_ROW_INDEX
    const numDataRows = lastRow - (HISTORICAL_RECORDS_ROW_START - 1);
    const dataMatrix = sheet.getRange(HISTORICAL_RECORDS_ROW_START, 1, numDataRows, lastCol).getValues();

    for (let i = 0; i < dataMatrix.length; i++) {
      const row = dataMatrix[i];
      const rawDate = row[0]; // Col A (Date)
      const sheetPeriod = row[PERIOD_COL_INDEX - 1]; // Col B (Sheet/Period) using 0-indexed lookup

      if (!rawDate || !sheetPeriod) continue;

      const formattedDate = rawDate instanceof Date 
        ? rawDate.toISOString().split('T')[0] 
        : String(rawDate);
        
      const displayDateStr = rawDate instanceof Date 
        ? rawDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) 
        : String(rawDate);

      const metricsMap: Record<string, number | null> = {};

      // Parse metric data columns starting from METRIC_DATA_START_COL (converted to 0-indexed offset)
      for (let colIdx = METRIC_DATA_START_COL - 1; colIdx < row.length; colIdx++) {
        const metricKey = sheetHeaderKeys[colIdx - (METRIC_DATA_START_COL - 1)];
        if (metricKey) {
          const rawCellVal = row[colIdx];
          const numVal = Number(rawCellVal);
          
          // Assign null if cell is empty or non-numeric (handling missing/cutoff data)
          metricsMap[metricKey] = (rawCellVal !== '' && !isNaN(numVal)) ? numVal : null;
        }
      }

      allRecords.push({
        date: formattedDate,
        dateStr: displayDateStr,
        sheet: String(sheetPeriod).trim(),
        metrics: metricsMap
      });
    }
  });

  const payload: HistoricalRecordPayload = {
    headers: Array.from(globalHeadersMap.values()),
    records: allRecords
  };

  return JSON.stringify(payload);
}