# Gmail → QBO Invoice Auto-Forwarder
### Google Apps Script for automatic vendor invoice routing

Watches Gmail for vendor emails, matches subject lines against keywords you define, and forwards qualifying emails to your accounting software's email intake address — with a CC to whoever needs to see it. Built and battle-tested at [Longshore Contracting](https://longshorecontracting.com).

---

## The problem it solves

Most accounting software (QuickBooks, Xero, Wave) has an email address that auto-imports bills when you forward invoices to it. The manual version of this workflow — someone on your team forwarding vendor emails — breaks down two ways:

1. Someone forgets to forward
2. Someone forwards but forgets the CC, creating duplicates in QBO

Native Gmail filters can't CC a second address on a forward. Apps Script can.

---

## How it works

- Runs every 15 minutes via a time-based trigger
- Searches Gmail for emails from vendors listed in your Rules sheet
- Checks subject lines against include/exclude keywords you define
- Forwards matches to your accounting software's intake address + CC
- Labels processed threads so nothing gets forwarded twice
- Logs every forward with date, sender, subject, and message ID

---

## Setup (one-time, ~10 minutes)

### 1. Find your accounting software's email intake address

| Software | Where to find it |
|----------|-----------------|
| QuickBooks Online | Expenses & Bills → Bills → Look for banner "Anyone can autofill multiple receipts..." |
| Xero | Settings → Email to Bills |
| Wave | Purchases → Bills → Email Bills to Wave |

### 2. Create a Google Sheet

Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet. Copy the ID from the URL — it's the long string between `/d/` and `/edit`.

### 3. Set up the script

1. Go to [script.google.com](https://script.google.com) and create a new project
2. Paste the full script from `invoice-forwarder.gs`
3. Fill in your values at the top of the CONFIG block:

```javascript
const CONFIG = {
  FORWARD_TO: 'your-qbo-intake@assist.intuit.com',  // your accounting software's address
  CC: 'finance@yourcompany.com',                     // whoever needs a copy
  SPREADSHEET_ID: 'paste-your-sheet-id-here',
  START_DATE: 'after:2026/03/13'                     // ← SET THIS TO TODAY'S DATE
};
```

> ⚠️ **Set START_DATE before your first run.** Without it, the script will sweep your entire inbox history and forward every matching email ever. Ask me how I know.

### 4. Run setup functions

In the Apps Script editor, run these two functions in order:

1. `setupSheet()` — builds the Rules and Log tabs and pre-populates example vendor rules
2. `createTrigger()` — sets the 15-minute timer

Authorize Gmail and Sheets access when prompted.

---

## Managing rules (no code required)

All vendor rules live in the **Rules tab** of your Google Sheet. Your finance team can manage this directly.

| Column | What it does |
|--------|-------------|
| Sender Email | Exact email address to watch |
| Include Keywords | Subject must contain at least one (comma-separated) |
| Exclude Keywords | Subject must NOT contain any of these (comma-separated) |
| Active | TRUE to enable, FALSE to pause without deleting |

**Example:** To forward invoices from a supplier but not their monthly statements:
- Include: `invoice, bill`
- Exclude: `statement, reminder`

Add a new row for each vendor. Changes take effect on the next 15-minute run.

---

## Auditing

Every forward is logged in the **Log tab** with:
- Date forwarded
- Which rule matched
- Subject line
- Actual sender address
- Message ID (used internally to prevent duplicates)

Do not edit the Log tab manually.

---

## Requirements

- Google account with Gmail and Google Sheets
- Apps Script access (free, included with Google account)
- Accounting software with an email intake address

---

## License

MIT — use it, adapt it, share it.

Built by [Longshore Contracting](https://longshorecontracting.com) · Eastern NC
