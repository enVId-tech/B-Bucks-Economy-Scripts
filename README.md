# B-Bucks Economy Scripts

A comprehensive Google Sheets Apps Script solution for managing a virtual classroom economy system. Create, manage, and track student transactions, investments, and financial activity with an intuitive interface built directly into Google Sheets.
 
**License:** MIT  
**Repository:** [github.com/enVId-tech/B-Bucks-Economy-Scripts](https://github.com/enVId-tech/B-Bucks-Economy-Scripts)

---

## Features

### **Core Management Tools**

- **Manual B-Bucks Modifier** - Instantly adjust student balances with quick buttons, custom amounts, and full transaction history with undo functionality
- **Income / Consumables Manager** - Add income transactions from earnable items and spending from consumable items with balance validation
- **Investment Banking** - Manage Certificates of Deposit (CoD) with simple interest calculations, earning projections, and deposit/withdrawal tracking
- **Transaction Records** - View, filter, and manage all B-Bucks transactions with advanced search and period-based filtering
- **Banking Settings** - Configure economy parameters including interest rates, restrictions, important dates, and maintenance operations

### **Key Capabilities**

- **Click-to-Select Transactions** - Select multiple transactions by clicking on entries (hover shows selection state)
- **Batch Undo Operations** - Undo multiple selected transactions at once from centralized undo button
- **Transaction Equations** - View complete transaction history as equations showing original → operation → result
- **Multi-Period Support** - Organize transactions and students across 4 quarters with configurable date ranges
- **Activity Logging** - Comprehensive audit trails with timestamps, transaction IDs, and automatic logging
- **Data Management** - Historical record saving, daily snapshots, timestamp updates, and selective column resets
- **Confirmation Modals** - Multi-step verification for destructive operations to prevent accidental data loss

---

## Installation

### Prerequisites
- Google Account with access to Google Sheets
- Permission to create/edit Google Sheets
- Permission to authorize Apps Scripts

### Setup Steps

1. **Clone or Download the Repository**
   ```bash
   git clone https://github.com/enVId-tech/B-Bucks-Economy-Scripts.git
   ```

2. **Create a New Google Sheet**
   - Go to [sheets.google.com](https://sheets.google.com)
   - Create a new spreadsheet
   - Name it (e.g., "B-Bucks Economy")

3. **Open Apps Script Editor**
   - In your Google Sheet, go to **Extensions → Apps Script**
   - This opens the Apps Script editor in a new tab

4. **Add Project Files**
   - Delete the default `Code.gs` file
   - Create new files for each component:
     - `Code.ts` - Main TypeScript file with menu and dialog functions
     - `AccountingManager.html` - Manual balance modifier interface
     - `PricingManager.html` - Income/Consumables manager interface
     - `InvestmentsManager.html` - Investment banking interface
     - `TransactionsRecords.html` - Transaction records viewer
     - `BankingSettings.html` - Settings and configuration panel
   - Copy-paste the contents from the repository files

5. **Configure Google Sheets Structure**
   - Create columns in your spreadsheet for:
     - Student names
     - Account balances
     - Transaction history
     - Earnings
     - Expenditures
     - Other tracking columns as needed

6. **Authorize the Script**
   - Save the project (Ctrl+S / Cmd+S)
   - Refresh your Google Sheet
   - A new menu **"The Central Bank of Banderas"** will appear
   - Click any menu item to trigger authorization
   - Grant the required permissions

---

## Usage Guide

### The Central Bank of Banderas Menu

Access all features from the menu in your Google Sheet:

#### 1. **Open Manual B-Bucks Modifier** (390×550)
**Purpose:** Quickly adjust student balances

- **Quick Adjustments:** Buttons for common amounts (±$1, ±$2, ±$5, ±$10)
- **Math Operations:** Multiply (×2, ×5) or divide (÷2, ÷4) balances
- **Custom Amount:** Enter exact amounts; checkbox toggles auto-submit
- **Format Guide:** Shows syntax for operations
- **Activity Log:** Search by name, filter by date, click to select, undo selected transactions

**Keyboard Syntax (Custom Amount):**
- `+10` = Add $10
- `-5` = Subtract $5
- `*2` = Double the balance
- `/4` = Quarter the balance

#### 2. **Open Service Provider** (500×700)
**Purpose:** Manage income and spending transactions

**Left Panel:**
- **Add Income:** Select quarter, item, quantity, and submit
- **Add Spending:** Select quarter, item, quantity with insufficient funds warning and override option

**Right Panel:**
- **Activity Log:** View recent income/spending/override entries with selected transaction undo option

#### 3. **Open Investment Banking** (600×700)
**Purpose:** Manage Certificate of Deposit investments

**Left Panel:**
- **Select Student:** Choose period and student
- **Account Overview:** Display cash, CoD balance, purchase date, and current earnings
- **CoD Operations:** Deposit amount, simple interest mode, admin override

**Right Panel:**
- **Earnings Calculator:** 3 projection modes showing interest projections

#### 4. **View Transactions Records** (700×650)
**Purpose:** Search, filter, and manage all transactions

**Filter Controls:**
- Period (All Periods, Period 1-4)
- Student (populated based on period)
- Type (All, Income, Spending, Deposit, Withdraw)
- Date range (From/To)

**Statistics Bar:**
- Total count, Income count, Spending count, Net amount

**Record Display:**
- Transaction entries with full equations showing balance changes
- Color-coded by transaction type
- Click-to-select with centralized undo for selected transactions

#### 5. **Banking Settings** (550×700)
**Purpose:** Configure economy parameters

**Sections:**
- **Interest Rates:** Simple interest percentage with save button
- **CoD Restrictions:** Lockout period, min/max deposit with save button
- **Important Dates:** Banking period, investments period, quarterly date ranges with save button
- **Transaction Logging:** Auto-logging, timestamp tracking, audit log toggles with save button
- **Maintenance:** 
  - Data Management (update timestamps, record daily data)
  - Historical Records (save to new tab, reset working tab)
  - Reset Operations (all money, by column) with save button
  - Destructive actions require confirmation modal verification

---

## Project Structure

```
B-Bucks-Economy-Scripts/
├── src/
│   ├── Code.ts                    # Main Apps Script file with menu & dialog functions
│   ├── AccountingManager.html      # Manual balance modifier UI
│   ├── PricingManager.html         # Income/Consumables manager UI
│   ├── InvestmentsManager.html     # Investment banking UI
│   ├── TransactionsRecords.html    # Transaction records viewer UI
│   ├── BankingSettings.html        # Settings & configuration UI
│   └── appsscript.json            # Apps Script manifest
├── package.json                   # Project metadata
├── tsconfig.json                  # TypeScript configuration
├── LICENSE                        # MIT License
└── README.md                      # This file
```

---

## Design System

### Color Palette (GitHub-Inspired)
- **Primary:** #0366d6 (Blue) - Actions, active states
- **Success:** #28a745 (Green) - Income, positive amounts
- **Danger:** #d73a49 (Red) - Spending, negative amounts
- **Secondary:** #6f42c1 (Purple) - Deposits, secondary actions
- **Background:** #fafbfc (Light gray)
- **Text:** #24292e (Dark gray)
- **Borders:** #e1e4e8 (Light border)

### Typography
- **Font Stack:** System fonts (-apple-system, Segoe UI, Roboto, sans-serif)
- **Labels:** 9-10px, uppercase, 600 weight, 0.3-0.5px letter-spacing
- **Body:** 11-12px
- **Headings:** 12-14px, bold

---

## UI Features & Implementation Notes

### Click-to-Select Pattern
- Click any transaction entry to select it (dark gray background)
- Hover shows light background preview
- Click again to deselect
- Multiple selections supported
- Undo button enables when items selected
- Click outside to deselect all

**Backend Integration:**
```typescript
const selectedItems = document.querySelectorAll('[data-selectable="true"].selected');
selectedItems.forEach(item => {
  // Access transaction data from item
});
```

### Navigation Header
All dialog pages include a navigation header for quick switching between dialogs:
- **Manual B-Bucks Modifier**
- **Income / Consumables**
- **Investment Banking**
- **Transaction Records**
- **Banking Settings**

Active page shows blue underline; hover highlights any link.

### Confirmation Modal
Destructive operations require three-field verification:
1. Student/individual name
2. Sheet name
3. Action name

Prevents accidental data loss through multiple confirmations.

---

## Backend Implementation Checklist

The UI structure is complete. To fully enable functionality, implement:

- [ ] Transaction data retrieval from spreadsheet
- [ ] Search/filter logic for activity logs
- [ ] Date range filtering
- [ ] Period → Student dropdown population
- [ ] Income/spending transaction submission
- [ ] Balance update calculations
- [ ] Undo transaction reversal logic
- [ ] CoD deposit/withdrawal processing
- [ ] Interest calculation and projections
- [ ] Historical record save/archive operations
- [ ] Column-specific reset operations
- [ ] Confirmation modal validation
- [ ] Timestamp updates and logging

All element IDs are preserved and ready for data binding.

---

## Keyboard Shortcuts & Tips

**In Manual B-Bucks Modifier:**
- Use quick buttons for common adjustments
- Custom amount field supports operation syntax
- Enable "Auto-submit" checkbox for hands-free entry
- Scroll activity log to see more transactions

**In Transaction Records:**
- Select period first to populate student dropdown
- Use multiple filters together for precise results
- Click entries to select batch operations

**In Banking Settings:**
- Each section has its own Save button
- Set important dates at start of school year
- Adjust restrictions to match economy goals

---

## Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## Troubleshooting

### Menu not appearing?
- Refresh the Google Sheet
- Check that Apps Script authorization completed
- Verify all files are created in the Apps Script editor

### Dialogs opening but blank?
- Clear browser cache
- Ensure all HTML files are properly formatted
- Check Console (Ctrl+Shift+J) for JavaScript errors

### Transactions not saving?
- Verify spreadsheet structure matches backend expectations
- Check Google Sheets API permissions
- Review Apps Script execution logs (Tools → Execution log)

---

## Future Enhancements

Potential features for future versions:
- Email notifications for large transactions
- Graphical analytics and reports
- Budget tracking and limits
- Student-facing portal for balance viewing
- Automated interest accrual
- Multi-teacher support with role-based access
- Mobile-friendly responsive design
- Transaction export to CSV/PDF
- Student achievement badges/milestones
- Classroom marketplace system

---

## License

This project is licensed under the **MIT License** - see [LICENSE](LICENSE) file for details.

---

## Support

For issues, questions, or suggestions:
- Open an issue on [GitHub Issues](https://github.com/enVId-tech/B-Bucks-Economy-Scripts/issues)
- Review existing documentation and FAQs
- Check the Apps Script console for error messages