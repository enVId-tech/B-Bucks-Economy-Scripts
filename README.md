# B-Bucks Economy Scripts

B-Bucks Economy Scripts is a TypeScript-based Google Apps Script project for running a classroom economy system directly from Google Sheets. It combines spreadsheet automation, custom dialogs, and transaction utilities into a single workflow for managing student balances, income and spending, investments, and historical financial activity.

- License: MIT
- Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts

---

## Current Project Status

This project is now a functioning Google Sheets add-on experience built around Apps Script dialogs and TypeScript utilities. The current implementation includes:

- A custom spreadsheet menu named “The Central Bank of Banderas”
- Multiple modeless dialog pages for banking tasks
- Server-side utilities for reading and writing spreadsheet data
- Client-side UI flows for balance changes, services, investments, records, analytics, and settings
- Caching and historical-data support for smoother interaction

---

## What It Does

The project is designed for classroom or club economy management. It helps a teacher or administrator:

- Adjust student balances manually
- Record income and expense-style transactions
- Track investment-style deposits and earnings
- Review transaction history and filter by date, period, student, or type
- Visualize historical economic data over time
- Configure banking settings and maintenance functions

---

## Core Features

### Manual Balance Management
- Quick balance adjustment buttons for common amounts
- Math-style operations such as multiply and divide
- Custom amount entry with format support
- Activity log with search and date filtering
- Undo support for selected transactions

### Service Provider / Income-Expense Flow
- Add income or expense-style transactions from a dedicated dialog
- Validate balances and support override flows when needed
- Review recent activity and transaction history from the same interface

### Investment Banking
- Manage investment-style deposits and withdrawals
- Track balances across periods
- Review earnings projections and investment state

### Transaction Records
- View transaction records in a searchable and filterable interface
- Filter by period, student, transaction type, and date range
- Review detailed transaction information and perform batch actions

### Historical Analytics
- Visualize historical economic data through chart-based analytics
- Filter records by metric columns, date range, and sheet/period scope
- Compare multiple metrics over time

### Banking Settings
- Configure interest-related parameters and restrictions
- Set important dates and period boundaries
- Access maintenance actions for data management and historical record handling

---

## Tech Stack

- Google Apps Script (V8 runtime)
- TypeScript
- HTML/CSS/JavaScript for dialog UIs
- Google Sheets API access
- Clasp for deployment

---

## Project Structure

```text
B-Bucks-Economy-Scripts/
├── src/
│   ├── Code.ts                  # Main Apps Script entry point and menu setup
│   ├── BalanceManager.html      # Manual balance manager dialog UI
│   ├── ServicesManager.html     # Income/expense service dialog UI
│   ├── InvestmentsManager.html  # Investment management dialog UI
│   ├── TransactionsRecords.html # Transaction history viewer UI
│   ├── HistoricalAnalytics.html # Historical analytics dialog UI
│   ├── BankingSettings.html     # Banking configuration dialog UI
│   ├── FillSheet.html           # Sheet import helper dialog
│   ├── DocsAttribution.html     # Documentation and attribution dialog
│   ├── utils/                   # Apps Script utility modules
│   └── styles/                  # UI styling for each dialog
├── package.json                 # Build and deployment scripts
├── tsconfig.json                # TypeScript configuration
├── appsscript.json              # Apps Script manifest
└── LICENSE                      # MIT License
```

---

## Development and Deployment

### Prerequisites
- Node.js and npm
- A Google account with access to Google Sheets
- A Google Apps Script project or a connected Google Sheet
- Clasp configured for deployment

### Install dependencies

```bash
npm install
```

### Build the project

```bash
npm run build
```

### Push to Apps Script

```bash
npm run push
```

This uses the configured Apps Script project and pushes the compiled output through Clasp.

---

## How to Use It

1. Open a Google Sheet.
2. Open the Apps Script editor from Extensions → Apps Script.
3. Deploy or push this project to the active spreadsheet project.
4. Refresh the spreadsheet.
5. Use the “The Central Bank of Banderas” menu to open the available tools.

---

## Notes

The implementation is designed around the spreadsheet structure and utility modules in the repository. The exact sheet layout and columns expected by the scripts are defined in the source under the utilities and Apps Script entry points.

---

## License

This project is licensed under the **MIT License** - see [LICENSE](LICENSE) file for details.

---

## Support

For questions, issues, or suggestions, use the GitHub repository issues page for the project.