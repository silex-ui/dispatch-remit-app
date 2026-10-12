/**
 * Shared backend for the Dispatch & Remit app.
 * This turns a Google Sheet into a tiny database the app's JavaScript talks to.
 *
 * SETUP (one time):
 * 1. Go to sheets.google.com and create a new blank spreadsheet.
 *    Name it anything, e.g. "Dispatch Remit DB".
 * 2. Extensions > Apps Script.
 * 3. Delete whatever default code is there and paste this whole file in.
 * 4. Click Deploy > New deployment.
 *    - Click the gear icon next to "Select type" and choose "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    - Click Deploy, then authorize it with your own Google account
 *      (you'll see an "unsafe app" warning because it's your own unverified
 *      script — click Advanced > Go to [project name] (unsafe) > Allow).
 * 5. Copy the Web app URL it gives you (ends in /exec) and send it back
 *    so it can be wired into the app.
 *
 * If you're updating an EXISTING deployment with a new version of this file,
 * use Deploy > Manage deployments > pencil icon > New version > Deploy,
 * so the same URL keeps serving the updated code.
 *
 * Do not change SECRET below — it must match the value baked into the app.
 *
 * NOTE: reads go out as JSONP (a <script> tag) and writes go out as a plain
 * HTML form POST, both read back via postMessage — not a normal fetch() —
 * because some browsers block a direct cross-origin fetch to Apps Script's
 * web app URLs. This file supports both styles.
 */

const SHEET_NAME = 'State';
const SECRET = 'bj6aF0zueMn6KtDGerbPDoG0alvwuwB3';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 3).setValues([['version', 'data', 'updatedAt']]);
    const seed = {
      settings: { ownerPin: '1903', adminPin: '1234', riders: [] },
      bookings: [], transactions: [], ledger: [], riderDays: {}
    };
    sh.getRange(2, 1, 1, 3).setValues([[1, JSON.stringify(seed), new Date().toISOString()]]);
  }
  return sh;
}

function readState_() {
  const sh = getSheet_();
  const version = sh.getRange(2, 1).getValue();
  const data = JSON.parse(sh.getRange(2, 2).getValue());
  return { version, data };
}

function writeState_(version, data) {
  const sh = getSheet_();
  sh.getRange(2, 1, 1, 3).setValues([[version, JSON.stringify(data), new Date().toISOString()]]);
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** JSONP: wraps the payload as `callbackName(payload);` and serves it as
 * JavaScript, so a <script src="...&callback=xyz"> tag can load it —
 * script tags aren't subject to the cross-origin fetch restriction. */
function jsonpOut_(callback, obj) {
  const safe = /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback || '') ? callback : 'callback';
  return ContentService
    .createTextOutput(safe + '(' + JSON.stringify(obj) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/** For writes submitted via a hidden HTML form (not fetch): respond with a
 * tiny page whose script posts the result back to the parent window. */
function postMessageOut_(reqId, obj) {
  const payload = JSON.stringify(Object.assign({ reqId: reqId || null }, obj))
    .replace(/</g, '\\u003c'); // defang in case any stored text contains "</script>"
  const html = '<!doctype html><html><body><script>' +
    'parent.postMessage(' + payload + ', "*");' +
    '</script></body></html>';
  return HtmlService.createHtmlOutput(html);
}

function doGet(e) {
  const key = (e.parameter && e.parameter.key) || '';
  const callback = e.parameter && e.parameter.callback;
  if (key !== SECRET) {
    const err = { ok: false, error: 'unauthorized' };
    return callback ? jsonpOut_(callback, err) : jsonOut_(err);
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const { version, data } = readState_();
    const out = { ok: true, version, data };
    return callback ? jsonpOut_(callback, out) : jsonOut_(out);
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  // Submitted either as a classic JSON POST body, or as HTML form fields
  // (from the app's hidden-form workaround) — handle both.
  let body, reqId = null;
  if (e.parameter && e.parameter.key) {
    reqId = e.parameter.reqId || null;
    try {
      body = {
        key: e.parameter.key,
        expectedVersion: e.parameter.expectedVersion,
        data: JSON.parse(e.parameter.data)
      };
    } catch (err) {
      return postMessageOut_(reqId, { ok: false, error: 'bad_json' });
    }
  } else {
    try {
      body = JSON.parse(e.postData.contents);
    } catch (err) {
      return jsonOut_({ ok: false, error: 'bad_json' });
    }
  }
  const respond = reqId !== null || (e.parameter && e.parameter.reqId !== undefined)
    ? (obj) => postMessageOut_(reqId, obj)
    : jsonOut_;

  if ((body.key || '') !== SECRET) return respond({ ok: false, error: 'unauthorized' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const current = readState_();
    if (Number(body.expectedVersion) !== Number(current.version)) {
      return respond({ ok: false, conflict: true, version: current.version, data: current.data });
    }
    const newVersion = current.version + 1;
    writeState_(newVersion, body.data);
    return respond({ ok: true, version: newVersion });
  } finally {
    lock.releaseLock();
  }
}
