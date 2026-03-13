// ============================================================
// Gmail → QBO Invoice Auto-Forwarder
// Built by Longshore Contracting — longshorecontracting.com
//
// ⚠️  BEFORE YOUR FIRST RUN: set START_DATE to today's date.
//     Without it, the script will forward your entire inbox
//     history. Every matching email. All at once. Don't ask.
// ============================================================

const CONFIG = {
  RULES_SHEET_NAME: 'Rules',
  LOG_SHEET_NAME:   'Log',
  FORWARD_TO:       'your-intake-address@accounting-software.com', // ← your QBO / Xero / Wave email
  CC:               'finance@yourcompany.com',                     // ← who gets a copy
  LABEL_NAME:       'QBO-Forwarded',
  SPREADSHEET_ID:   'REPLACE_WITH_YOUR_SHEET_ID',                  // ← paste Sheet ID here
  START_DATE:       'after:2026/03/13'                             // ← SET TO TODAY BEFORE FIRST RUN
};

// ============================================================
// MAIN — runs every 15 minutes via trigger
// ============================================================

function forwardInvoices() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const rulesSheet = ss.getSheetByName(CONFIG.RULES_SHEET_NAME);
  const logSheet   = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);

  const rules = getRules(rulesSheet);
  if (!rules.length) return;

  // Get or create the processed label
  let label = GmailApp.getUserLabelByName(CONFIG.LABEL_NAME);
  if (!label) label = GmailApp.createLabel(CONFIG.LABEL_NAME);

  // Load already-processed message IDs from log to prevent double-sends
  const processedIds = getProcessedIds(logSheet);

  rules.forEach(rule => {
    if (!rule.active) return;

    // START_DATE pins the search so no historical emails are swept up
    const threads = GmailApp.search(
      `from:${rule.sender} -label:${CONFIG.LABEL_NAME} ${CONFIG.START_DATE}`,
      0, 50
    );

    threads.forEach(thread => {
      thread.getMessages().forEach(message => {
        const msgId = message.getId();
        if (processedIds.has(msgId)) return;

        const subject = message.getSubject();
        if (!matchesRule(subject, rule)) return;

        try {
          message.forward(CONFIG.FORWARD_TO, { cc: CONFIG.CC });
          logForward(logSheet, message, rule.sender);
          processedIds.add(msgId);
          Logger.log(`Forwarded: "${subject}" from ${rule.sender}`);
        } catch (e) {
          Logger.log(`ERROR forwarding "${subject}": ${e.message}`);
        }
      });

      // Label the whole thread so it's excluded from future searches
      label.addToThread(thread);
    });
  });
}

// ============================================================
// HELPERS
// ============================================================

function getRules(sheet) {
  const data = sheet.getDataRange().getValues();
  const rules = [];
  for (let i = 1; i < data.length; i++) {
    const [sender, include, exclude, active] = data[i];
    if (!sender || String(sender).trim() === '') continue;
    rules.push({
      sender:  String(sender).trim().toLowerCase(),
      include: include ? String(include).split(',').map(s => s.trim().toLowerCase()).filter(Boolean) : [],
      exclude: exclude ? String(exclude).split(',').map(s => s.trim().toLowerCase()).filter(Boolean) : [],
      active:  active === true || String(active).toLowerCase() === 'true'
    });
  }
  return rules;
}

function matchesRule(subject, rule) {
  const s = subject.toLowerCase();
  if (rule.include.length > 0 && !rule.include.some(kw => s.includes(kw))) return false;
  if (rule.exclude.length > 0 && rule.exclude.some(kw => s.includes(kw))) return false;
  return true;
}

function getProcessedIds(logSheet) {
  const data = logSheet.getDataRange().getValues();
  const ids = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][4]) ids.add(String(data[i][4]));  // Column E = Message ID
  }
  return ids;
}

function logForward(logSheet, message, matchedRule) {
  logSheet.appendRow([
    new Date(),
    matchedRule,
    message.getSubject(),
    message.getFrom(),
    message.getId()
  ]);
}

// ============================================================
// SETUP — run once manually after creating the Sheet
// ============================================================

function setupSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // --- Rules tab ---
  let rulesSheet = ss.getSheetByName(CONFIG.RULES_SHEET_NAME);
  if (!rulesSheet) rulesSheet = ss.insertSheet(CONFIG.RULES_SHEET_NAME);
  rulesSheet.clearContents();

  const rulesHeaders = [
    'Sender Email',
    'Include Keywords (comma-separated)',
    'Exclude Keywords (comma-separated)',
    'Active'
  ];
  rulesSheet.getRange(1, 1, 1, 4).setValues([rulesHeaders])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  rulesSheet.setFrozenRows(1);
  rulesSheet.setColumnWidths(1, 1, 280);
  rulesSheet.setColumnWidths(2, 1, 340);
  rulesSheet.setColumnWidths(3, 1, 340);
  rulesSheet.setColumnWidths(4, 1, 70);

  // Example rules — replace with your actual vendors
  const defaultRules = [
    ['ar@pleasepayus.com',                   'invoice, bill',       'statement, reminder',              true],
    ['billing@wehavemet.com',                'your invoice',        'payment confirmation, schedule',   true],
    ['noreply@definitely-a-vendor.com',      'e-subscriptions',     '',                                 true],
    ['invoices@notastatement.com',           'daily invoices',      'monthly statement',                true],
    ['ar@justfriendlyreminder.com',          'new payment request', 'reminder, payment confirmation',   true],
    ['collections@yourproblemnowoursending.com', 'receipt',         '',                                 true],
  ];
  rulesSheet.getRange(2, 1, defaultRules.length, 4).setValues(defaultRules);

  for (let i = 2; i <= defaultRules.length + 1; i++) {
    rulesSheet.getRange(i, 1, 1, 4).setBackground(i % 2 === 0 ? '#f8f9fa' : '#ffffff');
  }

  // --- Log tab ---
  let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!logSheet) logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
  logSheet.clearContents();

  const logHeaders = ['Date Forwarded', 'Matched Rule (Sender)', 'Subject', 'From (Actual)', 'Message ID'];
  logSheet.getRange(1, 1, 1, 5).setValues([logHeaders])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  logSheet.setFrozenRows(1);
  logSheet.setColumnWidths(1, 1, 160);
  logSheet.setColumnWidths(2, 1, 260);
  logSheet.setColumnWidths(3, 1, 340);
  logSheet.setColumnWidths(4, 1, 260);
  logSheet.setColumnWidths(5, 1, 160);

  Logger.log('✅ Sheet setup complete. Update CONFIG values, then run createTrigger().');
}

function createTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'forwardInvoices') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('forwardInvoices')
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log('✅ Trigger created: forwardInvoices runs every 15 minutes.');
}
