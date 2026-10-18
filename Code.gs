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
 *
 * ROOT CAUSE FOUND 2026-10-18: a single Google Sheets cell can hold at most
 * 50,000 characters. The whole app's data (bookings + transactions + ledger
 * etc.) was being JSON-stringified into ONE cell — and once the real data
 * grew past that limit, every write containing the full, current dataset
 * started throwing an uncaught exception server-side ("Your input contains
 * more than the maximum of 50000 characters in a single cell"). Because the
 * exception happened before any response was ever sent back, the client's
 * formPost() just sat waiting until its own timeout fired — showing as a
 * generic "offline"/timeout on the device, with nothing in the UI pointing
 * at the real cause. Meanwhill, any write whose data happened to still be
 * under the limit (e.g. an older/smaller snapshot) could still succeed,
 * which is why the sheet kept settling back onto an incomplete copy no
 * matter how many times a full push was retried.
 *
 * FIX: the JSON is now split across several cells ("chunks") instead of one,
 * well under the 50,000-character limit each, and rejoined on read. See
 * writeChunks_ / readChunks_ below. migrateIfNeeded_ converts an existing
 * sheet from the old one-cell layout to this one automatically and safely,
 * the first time this code runs against it — no data is lost in the move.
 */

const SHEET_NAME = 'State';
const SECRET = 'bj6aF0zueMn6KtDGerbPDoG0alvwuwB3';
// Stay well under Google Sheets' real 50,000-character-per-cell ceiling.
const CHUNK_SIZE = 40000;

/** New layout, row 2 = the one and only state row:
 *   A version | B updatedAt | C writer | D chunk count | E.. the JSON, split
 *   into CHUNK_SIZE-character pieces, one piece per cell. */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 4).setValues([['version', 'updatedAt', 'writer', 'chunks']]);
    const seed = {
      settings: { ownerPin: '1903', adminPin: '1234', riders: [] },
      bookings: [], transactions: [], ledger: [], riderDays: {}
    };
    writeChunks_(sh, 1, JSON.stringify(seed), '', new Date().toISOString());
    return sh;
  }
  migrateIfNeeded_(sh);
  return sh;
}

/** A sheet built by the OLD code has the full JSON sitting in B2 (a string
 * starting with "{"). The new layout's B2 is always a timestamp instead —
 * that difference is how we tell old and new apart and convert safely,
 * without needing a separate "have we migrated yet" flag anywhere. */
function migrateIfNeeded_(sh) {
  const b2 = sh.getRange(2, 2).getValue();
  const looksOld = typeof b2 === 'string' && b2.trim().charAt(0) === '{';
  if (!looksOld) return;
  const oldVersion = sh.getRange(2, 1).getValue();
  const oldDataStr = String(b2);
  const oldUpdatedAtCell = sh.getRange(2, 3).getValue();
  const oldUpdatedAt = oldUpdatedAtCell instanceof Date
    ? oldUpdatedAtCell.toISOString()
    : String(oldUpdatedAtCell || new Date().toISOString());
  const oldWriter = sh.getRange(2, 4).getValue() || '';
  sh.getRange(1, 1, 1, 4).setValues([['version', 'updatedAt', 'writer', 'chunks']]);
  writeChunks_(sh, oldVersion, oldDataStr, oldWriter, oldUpdatedAt);
}

function writeChunks_(sh, version, jsonStr, writer, updatedAt) {
  const chunks = [];
  for (let i = 0; i < jsonStr.length; i += CHUNK_SIZE) chunks.push(jsonStr.slice(i, i + CHUNK_SIZE));
  if (chunks.length === 0) chunks.push('');
  // Clear out any leftover chunk cells from a previous, longer write so old
  // data never lingers past the end of the new (possibly shorter) one.
  const prevCount = Number(sh.getRange(2, 4).getValue()) || 0;
  const clearCount = Math.max(prevCount, chunks.length);
  if (clearCount > 0) sh.getRange(2, 5, 1, clearCount).clearContent();
  sh.getRange(2, 1, 1, 4).setValues([[version, updatedAt, writer || '', chunks.length]]);
  sh.getRange(2, 5, 1, chunks.length).setValues([chunks]);
}

function readChunks_(sh) {
  const version = sh.getRange(2, 1).getValue();
  const updatedAtCell = sh.getRange(2, 2).getValue();
  const updatedAt = updatedAtCell instanceof Date ? updatedAtCell.toISOString() : updatedAtCell;
  const writer = sh.getRange(2, 3).getValue() || '';
  const count = Number(sh.getRange(2, 4).getValue()) || 0;
  let jsonStr = '';
  if (count > 0) jsonStr = sh.getRange(2, 5, 1, count).getValues()[0].join('');
  return { version, data: JSON.parse(jsonStr || '{}'), updatedAt, writer };
}

function readState_() {
  return readChunks_(getSheet_());
}

/** Identifies which device/browser sent a write, so two devices racing each
 * other to overwrite the shared copy can actually be told apart afterward —
 * otherwise every successful write looks identical from the outside, even
 * when it came from a device holding stale/incomplete data. */
function writeState_(version, data, writer) {
  const sh = getSheet_();
  writeChunks_(sh, version, JSON.stringify(data), writer, new Date().toISOString());
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
  lock.waitLock(15000);
  try {
    const { version, data, updatedAt, writer } = readState_();
    const out = { ok: true, version, data, updatedAt, writer };
    return callback ? jsonpOut_(callback, out) : jsonOut_(out);
  } catch (err) {
    const out = { ok: false, error: 'server_error', message: String(err && err.message || err) };
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
        data: JSON.parse(e.parameter.data),
        writer: e.parameter.writer || ''
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

  Logger.log('doPost IN writer=[%s] expectedVersion=%s dataLen=%s',
    body.writer, body.expectedVersion, (body.data ? JSON.stringify(body.data).length : 0));

  if ((body.key || '') !== SECRET) {
    Logger.log('doPost REJECT unauthorized');
    return respond({ ok: false, error: 'unauthorized' });
  }
  // Every up-to-date copy of the app always identifies itself on a save
  // (the "writer" field). A copy old enough to predate that can silently
  // undo someone else's correct save with its own stale data and no one
  // would ever see which device did it — refusing it outright, instead of
  // writing it, forces that device to actually refresh before it can touch
  // the shared sheet again.
  if (!body.writer) {
    Logger.log('doPost REJECT stale_client (no writer field sent)');
    return respond({ ok: false, error: 'stale_client', message: 'This device is running an outdated copy of the app and was blocked from saving. Please fully close and reopen it to get the latest version.' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const current = readState_();
    if (Number(body.expectedVersion) !== Number(current.version)) {
      Logger.log('doPost REJECT conflict: client expected %s, sheet is at %s', body.expectedVersion, current.version);
      return respond({ ok: false, conflict: true, version: current.version, data: current.data });
    }
    const newVersion = current.version + 1;
    writeState_(newVersion, body.data, body.writer || '');
    Logger.log('doPost OK writer=[%s] newVersion=%s', body.writer, newVersion);
    return respond({ ok: true, version: newVersion });
  } catch (err) {
    // Whatever goes wrong here, always send SOMETHING back — an uncaught
    // exception used to mean the client's save request just hung until its
    // own timeout, with no clue anywhere about what actually happened.
    Logger.log('doPost ERROR %s', String(err && err.message || err));
    return respond({ ok: false, error: 'server_error', message: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}
