# B-Bucks Economy Scripts

This repository contains a Google Apps Script project built for a classroom economy system inside Google Sheets. The scripts add a custom spreadsheet menu titled “The Central Bank of Banderas” and provide a collection of dialogs and utilities for managing student balances, income and expenses, investments, historical records, and system settings.

The project is designed around a spreadsheet-based “banking” workflow for teachers and administrators rather than a standalone app. It relies on a specific workbook structure, sheet names, and settings data that are defined in the Apps Script code.

- License: MIT
- Repository: https://github.com/enVId-tech/B-Bucks-Economy-Scripts

---

## What this project does

The app is a spreadsheet-driven classroom economy manager. It lets a teacher:

- adjust student balances manually
- record service, income, and expenditure activity
- manage student investment accounts and returns
- review all transactions in a searchable record viewer
- analyze historical performance across reporting periods
- configure banking rules, dates, and limits from a settings panel

This is not a generic finance dashboard. It is built specifically to support a B-Bucks-style student economy with school-period-based accounting.

---

## Menu-driven workflow

When the spreadsheet opens, the project creates this custom menu:

- Open Manual B-Bucks Modifier
- Open Service Provider
- Open Investment Banking
- View Transactions Records
- View Historical Analytics
- Banking Settings
- Documentation & Attribution

These menu items correspond to the dialogs in the project and are defined in [src/Code.ts](src/Code.ts).

---

## Core features

### Manual balance management

The balance manager allows for direct adjustment of student balances. It is intended for quick corrections, custom value edits, and operational accounting tasks.

### Service and transaction workflow

The service provider dialog supports classroom transactions such as income and expense-style events. These are tied to the spreadsheet-backed services and transaction logic in the utility files.

### Investment banking

Investment-related functions track deposits, returns, gains, and net investment behavior. The project includes logic for investment balances and timing-related accounting.

### Transaction records

The transaction record viewer is a central log for reviewing entries, filtering by date or type, and investigating prior activity across the economy system.

### Historical analytics

Historical analytics reads recorded data and makes it available for visual analysis across defined periods and metrics.

### Banking settings

The settings tools manage the rules and constants used throughout the system, including dates, percentages, policies, limits, and technical defaults. These values are read from a dedicated Settings sheet and cached for runtime use.

---

## Project structure

```text
B-Bucks-Economy-Scripts/
├── src/
│   ├── Code.ts                     # Spreadsheet menu setup and dialog entry points
│   ├── BalanceManager.html         # Manual balance adjustment dialog
│   ├── ServicesManager.html        # Service / income / expenses dialog
│   ├── InvestmentsManager.html     # Investment management dialog
│   ├── TransactionsRecords.html    # Transaction history and filtering UI
│   ├── HistoricalAnalytics.html    # Historical report / chart UI
│   ├── BankingSettings.html        # Banking settings UI
│   ├── FillSheet.html              # Spreadsheet population helper
│   ├── DocsAttribution.html        # Documentation and attribution UI
│   ├── styles/                     # HTML/CSS styling for dialogs
│   └── utils/                      # Apps Script logic for settings, caching, transactions, services, analytics, and helpers
├── appsscript.json                 # Apps Script manifest
├── package.json                    # TypeScript build and deployment scripts
├── tsconfig.json                   # TypeScript configuration
├── LICENSE                         # MIT License
├── README.md                       # Project documentation
└── dist/                           # Generated build output (after npm run build)
```

---

## Technology stack

- Google Apps Script
- TypeScript
- Google Sheets data model and UI
- HTML/CSS/JavaScript for modal and modeless dialogs
- Clasp for Apps Script deployment

---

## Requirements

Before deploying this project, you should have:

- Node.js and npm installed
- a Google account with access to Google Sheets
- a Google Apps Script project linked to the spreadsheet you want to manage
- clasp configured for pushing Apps Script code

---

## Local development

Install dependencies:

```bash
npm install
```

Build the project:

```bash
npm run build
```

Deploy the compiled project to Apps Script:

```bash
npm run push
```

The push script runs the TypeScript build first and then uploads the result to the connected Apps Script project via clasp.

---

## How to use it

1. Open or create the Google Sheet that will serve as the economy workbook.
2. Ensure the workbook contains the sheet structure expected by the scripts.
3. Load this project into Google Apps Script for that spreadsheet.
4. Refresh the spreadsheet and open the Apps Script menu.
5. Use the “The Central Bank of Banderas” menu to work with balances, services, investments, records, analytics, and settings.

> This project depends on workbook-specific sheet names and settings values. The exact assumptions are defined in the source under the utility files and constants configuration.

---

## Notes on the project design

The codebase is organized around a specific classroom economy workflow rather than a generic accounting system. In particular, the app expects standard spreadsheet sections such as:

- a student names / balance area
- a services sheet
- a settings sheet
- transaction logs
- historical records
- investment-related columns and period tracking

Those conventions are encoded in the constants and helper modules in the utilities directory.

---

## License

This project is licensed under the **MIT License**. See [LICENSE](LICENSE) for details.

---

## Support

For bugs, feature requests, or discussion, use the repository issues page on GitHub.