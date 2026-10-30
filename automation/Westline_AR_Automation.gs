/**
 * Westline AR Operating Workbook — automation
 * ------------------------------------------------------------
 * Paste into Extensions → Apps Script of the workbook, save, reload the sheet.
 * A "Westline AR" menu appears. Run "Install daily triggers" once.
 *
 * What it does
 *   1. Builds the call-disposition Google Form and links it to the Call Log.
 *   2. Moves the as-of date to today each morning (6:00 AM PT) so ages,
 *      scores and the Work Queue recalculate before the team logs in.
 *   3. Saves a daily snapshot of the headline figures (trend history).
 *   4. Emails the morning summary: top 10 accounts, promises due, requests
 *      past service level.
 *   5. Flags requests that pass the 4-hour first-contact service level.
 *
 * Westline Auto Parts is fictional. No real customer data.
 */

const CFG = {
  TZ: 'America/Los_Angeles',
  SHEETS: {
    settings: 'Settings', accounts: 'Accounts', worklist: 'Worklist', calls: 'Call Log',
    requests: 'Requests', queue: 'Work Queue', lists: 'Lists', snapshot: 'Daily Snapshot'
  },
  FORM_TITLE: 'Westline – Call disposition',
  SUMMARY_TO: '',            // leave blank to send to the account running the trigger
  SLA_HOURS: 4
};

/* ---------------- menu ---------------- */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Westline AR')
    .addItem('Set as-of date to today', 'setAsOfToday')
    .addItem('Build call disposition form', 'setupDispositionForm')
    .addSeparator()
    .addItem('Save daily snapshot now', 'saveDailySnapshot')
    .addItem('Send morning summary now', 'sendMorningSummary')
    .addItem('Highlight requests past service level', 'flagRequestSla')
    .addSeparator()
    .addItem('Install daily triggers', 'installTriggers')
    .addItem('Remove triggers', 'removeTriggers')
    .addToUi();
}

/* ---------------- helpers ---------------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Missing sheet: ' + name);
  return s;
}
function named_(name) { return ss_().getRangeByName(name); }
function rowsAsObjects_(sheet) {
  const v = sheet.getDataRange().getValues();
  const head = v.shift();
  return v.filter(r => r[0] !== '').map(r => Object.fromEntries(head.map((h, i) => [h, r[i]])));
}
function money_(n) { return '$' + Math.round(n).toLocaleString('en-US'); }

/* ---------------- 1. as-of date ---------------- */
function setAsOfToday() {
  const today = new Date(Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd') + 'T00:00:00');
  named_('AsOf').setValue(today);
  SpreadsheetApp.flush();
}

/* ---------------- 2. disposition form ---------------- */
function setupDispositionForm() {
  const lists = sh_(CFG.SHEETS.lists);
  const col = header => {
    const head = lists.getRange(1, 1, 1, lists.getLastColumn()).getValues()[0];
    const c = head.indexOf(header) + 1;
    return lists.getRange(2, c, lists.getLastRow() - 1, 1).getValues().flat().filter(String);
  };
  const form = FormApp.create(CFG.FORM_TITLE)
    .setDescription('One response per call. Takes under 20 seconds. Promises need an amount and a date.')
    .setCollectEmail(false)
    .setProgressBar(false);

  form.addListItem().setTitle('Specialist').setChoiceValues(col('Specialists')).setRequired(true);
  form.addTextItem().setTitle('Account ID').setRequired(true)
      .setValidation(FormApp.createTextValidation().requireTextMatchesPattern('^C-\\d{5}$')
      .setHelpText('Format C-12345').build());
  form.addListItem().setTitle('Outcome').setChoiceValues(col('Outcomes')).setRequired(true);
  form.addTextItem().setTitle('Amount')
      .setHelpText('Promised or paid amount, numbers only')
      .setValidation(FormApp.createTextValidation().requireNumberGreaterThan(0).build());
  form.addDateItem().setTitle('Promise date');
  form.addParagraphTextItem().setTitle('Notes');

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss_().getId());
  SpreadsheetApp.flush();
  // The form writes to a new "Form Responses" tab with the same column order as Call Log.
  SpreadsheetApp.getUi().alert('Form created:\n' + form.getPublishedUrl() +
    '\n\nResponses land in a new "Form Responses" tab. Point the Call Log at it, or keep both and append nightly.');
}

/* ---------------- 3. daily snapshot ---------------- */
function saveDailySnapshot() {
  let snap = ss_().getSheetByName(CFG.SHEETS.snapshot);
  if (!snap) {
    snap = ss_().insertSheet(CFG.SHEETS.snapshot);
    snap.appendRow(['Date', 'Worklist balance', 'Past due', 'Older than 60 days %', 'Accounts past due',
                    'Promises due unpaid', 'Open requests', 'Requests past SLA']);
    snap.setFrozenRows(1);
  }
  const inv = rowsAsObjects_(sh_(CFG.SHEETS.worklist));
  const acc = rowsAsObjects_(sh_(CFG.SHEETS.accounts));
  const req = rowsAsObjects_(sh_(CFG.SHEETS.requests));
  const asOf = named_('AsOf').getValue();
  const bal = inv.reduce((s, r) => s + (Number(r['Open balance']) || 0), 0);
  const pd = inv.filter(r => r['Days past due'] > 0 && r['Flag'] !== 'Legacy').reduce((s, r) => s + (Number(r['Open balance']) || 0), 0);
  const old = inv.filter(r => r['Days since invoice'] > 60).reduce((s, r) => s + (Number(r['Open balance']) || 0), 0);
  const promises = acc.filter(a => a['Promise due'] instanceof Date && a['Promise due'] <= asOf && a['Past due'] > 0).length;
  const open = req.filter(r => r['Status'] !== 'Resolved');
  const late = open.filter(r => !(r['First customer contact'] instanceof Date) &&
                                (new Date() - r['Raised']) / 36e5 > CFG.SLA_HOURS).length;
  snap.appendRow([asOf, bal, pd, bal ? old / bal : 0, acc.filter(a => a['Past due'] > 0).length, promises, open.length, late]);
}

/* ---------------- 4. morning summary ---------------- */
function sendMorningSummary() {
  const q = sh_(CFG.SHEETS.queue).getRange('B8:L17').getValues().filter(r => r[0] !== '');
  const acc = rowsAsObjects_(sh_(CFG.SHEETS.accounts));
  const asOf = named_('AsOf').getValue();
  const due = acc.filter(a => a['Promise due'] instanceof Date && a['Promise due'] <= asOf && a['Past due'] > 0);
  const rows = q.map(r => `<tr><td>${r[0]}</td><td>${r[1]}</td><td>${r[4]}</td><td style="text-align:right">${money_(r[5])}</td><td style="text-align:right">${r[6]}</td><td>${r[9]}</td></tr>`).join('');
  const html = `
    <div style="font-family:Arial,sans-serif;font-size:13px;color:#0E1726">
      <h2 style="font-size:16px;margin:0 0 6px">Collections desk · ${Utilities.formatDate(asOf, CFG.TZ, 'EEE d MMM yyyy')}</h2>
      <p style="margin:0 0 12px;color:#435166">${due.length} promises are due and unpaid. Top 10 accounts to contact first:</p>
      <table cellpadding="6" style="border-collapse:collapse;font-size:12px">
        <tr style="background:#0E1726;color:#fff"><th>#</th><th>Account</th><th>Owner</th><th>Past due</th><th>Oldest</th><th>Next action</th></tr>
        ${rows}
      </table>
      <p style="color:#6F7B8E;font-size:11px">Sent by the Westline AR workbook.</p>
    </div>`;
  const to = CFG.SUMMARY_TO || Session.getActiveUser().getEmail();
  MailApp.sendEmail({ to, subject: 'Collections desk – ' + Utilities.formatDate(asOf, CFG.TZ, 'd MMM'), htmlBody: html });
}

/* ---------------- 5. request service level ---------------- */
function flagRequestSla() {
  const sheet = sh_(CFG.SHEETS.requests);
  const v = sheet.getDataRange().getValues();
  const head = v[0];
  const iRaised = head.indexOf('Raised'), iFirst = head.indexOf('First customer contact'), iStatus = head.indexOf('Status');
  const now = new Date();
  for (let r = 1; r < v.length; r++) {
    if (!v[r][0]) continue;
    const late = v[r][iStatus] !== 'Resolved' && !(v[r][iFirst] instanceof Date) &&
                 (now - v[r][iRaised]) / 36e5 > CFG.SLA_HOURS;
    sheet.getRange(r + 1, 1, 1, head.length).setBackground(late ? '#FBE9E7' : null);
  }
}

/* ---------------- triggers ---------------- */
function installTriggers() {
  removeTriggers();
  ScriptApp.newTrigger('morningRun').timeBased().everyDays(1).atHour(6).inTimezone(CFG.TZ).create();
  ScriptApp.newTrigger('flagRequestSla').timeBased().everyHours(1).create();
  SpreadsheetApp.getActive().toast('Daily 6 AM PT run and hourly SLA check installed.');
}
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}
function morningRun() {
  setAsOfToday();
  Utilities.sleep(4000);       // let formulas settle
  saveDailySnapshot();
  sendMorningSummary();
}
