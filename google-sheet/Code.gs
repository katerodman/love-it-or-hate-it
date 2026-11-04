/**
 * Love It or Hate It: receives finished games, saves them to this Google Sheet
 * and emails a copy of each response.
 *
 * Paste this whole file into Extensions > Apps Script in your Google Sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */

const EMAIL_TO = 'kate@violetstudios.org';
const RESPONSES_TAB = 'Responses';
const SUMMARY_TAB = 'Summary';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ratings = Array.isArray(data.ratings) ? data.ratings.slice(0, 100) : [];
    if (!ratings.length) return reply({ ok: false, error: 'No ratings' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = setUpSheets(ss);
    const submitted = new Date();
    const responseId = String(data.responseId || Utilities.getUuid()).slice(0, 40);
    const finishedAt = data.finishedAt ? new Date(data.finishedAt) : submitted;

    const rows = ratings.map(r => [
      submitted, responseId, clean(r.org),
      score(r.name), score(r.logo), score(r.work),
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    sendEmail(ratings, finishedAt, responseId, ss.getUrl());
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Lets you open the web app link in a browser to check it is running. */
function doGet() {
  return reply({ ok: true, message: 'Love It or Hate It is connected.' });
}

/** Run this once from the editor to create the tabs and approve permissions. */
function setUp() {
  setUpSheets(SpreadsheetApp.getActiveSpreadsheet());
  MailApp.getRemainingDailyQuota();
}

function setUpSheets(ss) {
  let sheet = ss.getSheetByName(RESPONSES_TAB);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    if (sheet.getLastRow() > 0) sheet = ss.insertSheet(RESPONSES_TAB);
    else sheet.setName(RESPONSES_TAB);
    sheet.appendRow(['Submitted', 'Response ID', 'Organization', 'Name', 'Logo', 'Work']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:F1').setFontWeight('bold');
  }
  if (!ss.getSheetByName(SUMMARY_TAB)) {
    const summary = ss.insertSheet(SUMMARY_TAB);
    summary.getRange('A1').setFormula(
      '=QUERY(' + RESPONSES_TAB + '!A:F, "select C, avg(D), avg(E), avg(F), count(D) ' +
      'where C is not null group by C order by avg(F) desc ' +
      'label C \'Organization\', avg(D) \'Avg name\', avg(E) \'Avg logo\', avg(F) \'Avg work\', count(D) \'Responses\'", 1)'
    );
    summary.getRange('F1').setValue('Not familiar with work');
    summary.getRange('F2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="", "", COUNTIFS(' + RESPONSES_TAB + '!C:C, A2:A, ' + RESPONSES_TAB + '!F:F, "Not familiar")))'
    );
    summary.getRange('A1:F1').setFontWeight('bold');
    summary.setFrozenRows(1);
  }
  return sheet;
}

function sendEmail(ratings, finishedAt, responseId, sheetUrl) {
  const sorted = ratings.slice().sort((a, b) => clean(a.org).localeCompare(clean(b.org)));
  const unfamiliar = sorted.filter(r => score(r.work) === 'Not familiar').length;
  const rowsHtml = sorted.map(r =>
    '<tr><td style="padding:4px 10px;border-bottom:1px solid #eee">' + html(clean(r.org)) + '</td>' +
    [r.name, r.logo, r.work].map(v => '<td style="padding:4px 10px;border-bottom:1px solid #eee;text-align:center">' + html(String(score(v))) + '</td>').join('') +
    '</tr>'
  ).join('');
  const body =
    '<p>New Love It or Hate It response, finished ' + html(finishedAt.toUTCString()) + '.</p>' +
    '<p>Not familiar with the work of ' + unfamiliar + ' of ' + sorted.length + ' orgs.</p>' +
    '<table style="border-collapse:collapse;font-family:sans-serif;font-size:14px">' +
    '<tr><th style="text-align:left;padding:4px 10px">Organization</th><th style="padding:4px 10px">Name</th><th style="padding:4px 10px">Logo</th><th style="padding:4px 10px">Work</th></tr>' +
    rowsHtml + '</table>' +
    '<p><a href="' + sheetUrl + '">Open the results sheet</a> · Response ' + html(responseId) + '</p>';
  const text = sorted.map(r => clean(r.org) + ': Name ' + score(r.name) + ' / Logo ' + score(r.logo) + ' / Work ' + score(r.work)).join('\n');
  MailApp.sendEmail({ to: EMAIL_TO, subject: 'Love It or Hate It: new response', body: text, htmlBody: body, name: 'Love It or Hate It' });
}

function score(v) {
  if (v === 'na' || v === 'Not familiar') return 'Not familiar';
  const n = Number(v);
  return n >= 1 && n <= 5 ? n : '';
}
function clean(s) { return String(s || '').slice(0, 120); }
function html(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
