/**
 * PO WEBAPP V1.3
 * Backend for the multi-module PO webapp.
 *
 * V1.3 changes:
 *   - Generic row API: getRows / addRow / updateRow / deleteRow.
 *     The first argument is always the tab name (sheetName).
 *   - Multi-module support: one Apps Script project bound to one
 *     "PO WEBAPP BACKEND" spreadsheet that contains a tab per module:
 *       PURCHASED_ORDER, INCOMING_SHIPMENT, SAMPLES, NOTES.
 *   - listSheets() returns every tab in the active spreadsheet so the
 *     Settings panel can show them as suggestions.
 *
 * Setup:
 *   1. Open your "PO WEBAPP BACKEND" Google Sheet
 *      (rename the existing PURCHASED ORDERS sheet for clarity, optional)
 *   2. Add 4 tabs: PURCHASED_ORDER, INCOMING_SHIPMENT, SAMPLES, NOTES
 *      (or whatever names you want — set them in the in-app Settings later)
 *      Add column headers in row 1 of each tab.
 *   3. Extensions -> Apps Script
 *   4. Replace Code.gs with this file.
 *   5. Replace the Index HTML file with the V1.3 Index.html.
 *   6. Deploy -> Manage deployments -> Edit -> New version -> Deploy.
 */

// ===================== CONFIG =====================
// Default tab name used when the frontend does not pass one. The frontend
// stores its own per-module tab names in localStorage (Settings panel).
const SHEET_NAME = 'PURCHASED_ORDER';
// Server-side CacheService TTL for getRows() results.
// Reads within this window return the cached blob (~10x faster than re-reading
// the sheet). Writes (addRow / updateRow / deleteRow) invalidate the cache
// for that tab so users see their own edits immediately.
const CACHE_TTL_SECONDS = 60;
// ===================================================

/**
 * Serves the web app HTML (GAS-hosted mode) when called with no action param.
 * When called with ?action=<fn> from an external frontend (e.g. GitHub Pages),
 * routes to the appropriate read function and returns JSON.
 * ContentService responses automatically include Access-Control-Allow-Origin: *.
 */
function doGet(e) {
  if (!e || !e.parameter || !e.parameter.action) {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('PO WEBAPP V1.3')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  var p = e.parameter;

  // All read actions require a valid session token
  var session = _verifyToken_(p.token);
  if (!session) return jsonResponse_({ error: 'Auth required', authRequired: true });

  var result;
  switch (p.action) {
    case 'getRows':
      // Server-side module access check — cannot be bypassed by any client.
      // Falls back to sheet-name check (_sheets array) when old clients omit moduleId.
      if (session.r !== 'admin') {
        var _tabCfg = getTabConfig();
        if (_tabCfg.ok) {
          try {
            var _cfgObj = JSON.parse(_tabCfg.value || '{}');
            var _disabled = p.moduleId && !!_cfgObj[p.moduleId];
            if (!_disabled && Array.isArray(_cfgObj._sheets)) {
              _disabled = _cfgObj._sheets.indexOf(String(p.sheetName || '').trim()) !== -1;
            }
            if (_disabled) return jsonResponse_({ error: 'Module not available', disabled: true });
          } catch (e) {}
        }
      }
      result = getRows(p.sheetName);
      break;
    case 'getUsers':     result = getUsers();             break;
    case 'getColHidden': result = getColHidden();         break;
    case 'getColVis':    result = getColVis();            break;
    case 'getTabConfig': result = getTabConfig();         break;
    case 'getItemCodes': result = getItemCodes();         break;
    case 'listSheets':   result = listSheets();           break;
    case 'getImages':    result = getImages(p.sheetName, p.rowIndex);  break;
    default:             result = { error: 'Unknown action: ' + p.action };
  }
  return jsonResponse_(result);
}

/**
 * Routes write requests from an external frontend.
 * Body must be a JSON string with { action, ...params }.
 * Content-Type: text/plain avoids CORS preflight (simple request).
 */
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) {
    return jsonResponse_({ success: false, error: 'Invalid JSON body' });
  }
  // login is the only unauthenticated action
  if (body.action === 'login') {
    return jsonResponse_(login(body.username, body.password));
  }

  // All other write actions require a valid session token
  var session = _verifyToken_(body.token);
  if (!session) return jsonResponse_({ success: false, error: 'Auth required', authRequired: true });

  // Admin-only writes
  var ADMIN_ONLY = { saveUser: 1, deleteUser: 1, saveColVis: 1, saveColHidden: 1, saveTabConfig: 1 };
  if (ADMIN_ONLY[body.action] && session.r !== 'admin') {
    return jsonResponse_({ success: false, error: 'Admin role required', adminRequired: true });
  }

  var result;
  switch (body.action) {
    case 'addRow':         result = addRow(body.sheetName, body.data);                                        break;
    case 'updateRow':      result = updateRow(body.sheetName, body.rowIndex, body.data);                      break;
    case 'deleteRow':      result = deleteRow(body.sheetName, body.rowIndex);                                 break;
    case 'moveToReceived': result = moveToReceived(body.sourceTab, body.rowIndex, body.data, body.targetTab); break;
    case 'saveUser':       result = saveUser(body.username, body.password, body.role);                        break;
    case 'deleteUser':     result = deleteUser(body.username);                                                break;
    case 'saveColVis':     result = saveColVis(body.jsonStr);                                                 break;
    case 'saveColHidden':  result = saveColHidden(body.jsonStr);                                                 break;
    case 'saveTabConfig':  result = saveTabConfig(body.jsonStr);                                              break;
    case 'saveImage':      result = saveImage(body.sheetName, body.rowIndex, body.name, body.base64, body.mimeType); break;
    case 'getImageData':   result = getImageData(body.imageId);                                                      break;
    case 'deleteImage':    result = deleteImage(body.imageId);                                                       break;
    default:               result = { success: false, error: 'Unknown action: ' + body.action };
  }
  return jsonResponse_(result);
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Resolve a tab name to a Sheet object.
 * If `sheetName` is empty/undefined, falls back to SHEET_NAME.
 * If still not found, throws a descriptive error so the frontend can
 * surface it (instead of silently using the wrong tab).
 */
function getSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = (sheetName && String(sheetName).trim()) || SHEET_NAME;
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet tab "' + name + '" not found in "' + ss.getName() +
                    '". Available tabs: ' +
                    ss.getSheets().map(function (s) { return s.getName(); }).join(', '));
  }
  return sheet;
}

/**
 * Lists every tab in the active spreadsheet (used by the Settings panel).
 */
function listSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const names = ss.getSheets().map(function (s) { return s.getName(); });
    return { ok: true, sheets: names, defaultSheet: SHEET_NAME, spreadsheetName: ss.getName() };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

/**
 * Reads all rows from a tab and returns { headers, rows, sheetName }.
 * Each row object has a __rowIndex pointing to its 1-based sheet row,
 * which the frontend uses for edit / delete.
 *
 * Results are cached in CacheService for CACHE_TTL_SECONDS so subsequent
 * reads are very fast. The cache is invalidated by addRow/updateRow/deleteRow.
 */
function getRows(sheetName) {
  try {
    const sheet = getSheet_(sheetName);
    const actualName = sheet.getName();
    const cache = CacheService.getScriptCache();
    const cacheKey = 'rows_' + actualName;

    const cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) { /* fall through */ }
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    let result;
    if (lastRow < 1 || lastCol < 1) {
      result = { headers: [], rows: [], sheetName: actualName };
    } else {
      const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      const headers = values[0].map(function (h) { return String(h || '').trim(); });
      const tz = Session.getScriptTimeZone();
      const rows = values.slice(1).map(function (row, idx) {
        const obj = { __rowIndex: idx + 2 };
        headers.forEach(function (h, i) {
          let val = row[i];
          if (val instanceof Date) {
            val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
          }
          obj[h] = val;
        });
        return obj;
      });
      result = { headers: headers, rows: rows, sheetName: actualName };
    }

    // CacheService values are capped at 100KB — large sheets silently skip caching.
    try {
      cache.put(cacheKey, JSON.stringify(result), CACHE_TTL_SECONDS);
    } catch (e) { /* payload too large; serve uncached */ }

    return result;
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * Invalidates the CacheService entry for a tab so the next getRows() call
 * fetches fresh from the spreadsheet. Called after every write.
 */
function invalidateCache_(sheetName) {
  try {
    const sheet = getSheet_(sheetName);
    CacheService.getScriptCache().remove('rows_' + sheet.getName());
  } catch (e) { /* ignore */ }
}

/**
 * Appends a row to a tab. `data` is keyed by column header.
 */
function addRow(sheetName, data) {
  try {
    const sheet = getSheet_(sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(function (h) {
      const v = data[h];
      return (v === undefined || v === null) ? '' : v;
    });
    sheet.appendRow(newRow);
    invalidateCache_(sheetName);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Updates a row in place. `rowIndex` is the 1-based sheet row.
 */
function updateRow(sheetName, rowIndex, data) {
  try {
    const sheet = getSheet_(sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(function (h) {
      const v = data[h];
      return (v === undefined || v === null) ? '' : v;
    });
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([newRow]);
    invalidateCache_(sheetName);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Moves a row from sourceTab to targetTab (called when STATUS → "Received").
 * 1. Writes the updated data back to the source row.
 * 2. Appends a copy to targetTab, mapped by that sheet's column headers.
 * 3. Deletes the source row so it no longer appears in PURCHASED_ORDER.
 * Both caches are invalidated so the next getRows() call is fresh.
 */
function moveToReceived(sourceTab, rowIndex, data, targetTab) {
  try {
    const srcSheet = getSheet_(sourceTab);
    const tgtSheet = getSheet_(targetTab);

    // Write the updated row back to the source sheet first.
    const srcHeaders = srcSheet.getRange(1, 1, 1, srcSheet.getLastColumn()).getValues()[0];
    const srcRow = srcHeaders.map(function (h) {
      const v = data[h];
      return (v === undefined || v === null) ? '' : v;
    });
    srcSheet.getRange(rowIndex, 1, 1, srcHeaders.length).setValues([srcRow]);

    // Append a copy to the target (RECEIVED) sheet, mapped by its own headers.
    // If the target sheet has no headers yet, seed it with the source headers first.
    let tgtHeaders;
    const tgtLastCol = tgtSheet.getLastColumn();
    if (tgtLastCol < 1) {
      tgtHeaders = srcHeaders;
      tgtSheet.appendRow(tgtHeaders);
    } else {
      tgtHeaders = tgtSheet.getRange(1, 1, 1, tgtLastCol).getValues()[0];
    }
    const tgtRow = tgtHeaders.map(function (h) {
      const v = data[h];
      return (v === undefined || v === null) ? '' : v;
    });
    tgtSheet.appendRow(tgtRow);

    invalidateCache_(sourceTab);
    invalidateCache_(targetTab);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Deletes a row from a tab. `rowIndex` is the 1-based sheet row.
 */
function deleteRow(sheetName, rowIndex) {
  try {
    const sheet = getSheet_(sheetName);
    sheet.deleteRow(rowIndex);
    invalidateCache_(sheetName);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * ============ Password hashing ============
 * Passwords are stored as 'sha256:<base64-digest>' where the digest is
 * SHA-256(pepper + ':' + username + ':' + password). The pepper is a random
 * value generated once and stored in PropertiesService — it never leaves the
 * server and stays the same across reads/writes.
 *
 * Plaintext (legacy) passwords are still accepted on login and migrated to
 * hashed form automatically on the next successful login.
 */
function _getHashPepper_() {
  const props = PropertiesService.getScriptProperties();
  let pepper = props.getProperty('PW_PEPPER');
  if (!pepper) {
    const seed = String(Math.random()) + ':' + String(Date.now()) + ':' + Utilities.getUuid();
    pepper = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed)
    );
    props.setProperty('PW_PEPPER', pepper);
  }
  return pepper;
}

function _hashPassword_(username, password) {
  const pepper = _getHashPepper_();
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    pepper + ':' + String(username || '').toLowerCase() + ':' + String(password || '')
  );
  return 'sha256:' + Utilities.base64Encode(bytes);
}

function _verifyPassword_(username, password, stored) {
  if (!stored) return false;
  const s = String(stored);
  if (s.indexOf('sha256:') === 0) {
    return _hashPassword_(username, password) === s;
  }
  // Legacy plaintext compare
  return s === String(password);
}

/**
 * ============ Session tokens (HMAC-signed) ============
 * Token format: <base64url(payload)>.<base64url(hmac-sha256)>
 * Payload JSON: { u: username, r: role, exp: epoch_ms }
 * Secret is generated once and stored in PropertiesService.
 *
 * Tokens expire after TOKEN_TTL_MS. To force logout of all sessions, delete
 * the TOKEN_SECRET property (every existing token then becomes unverifiable).
 */
const TOKEN_TTL_MS = 12 * 60 * 60 * 1000; // 12 hours

function _getTokenSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('TOKEN_SECRET');
  if (!secret) {
    const seed = String(Math.random()) + ':' + String(Date.now()) + ':' + Utilities.getUuid();
    secret = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed)
    );
    props.setProperty('TOKEN_SECRET', secret);
  }
  return secret;
}

function _signToken_(username, role) {
  const payload = JSON.stringify({
    u: String(username || '').toLowerCase(),
    r: role === 'admin' ? 'admin' : 'user',
    exp: Date.now() + TOKEN_TTL_MS
  });
  const sig = Utilities.computeHmacSha256Signature(payload, _getTokenSecret_());
  const payloadEnc = Utilities.base64EncodeWebSafe(payload).replace(/=+$/, '');
  const sigEnc = Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '');
  return payloadEnc + '.' + sigEnc;
}

function _verifyToken_(token) {
  try {
    if (!token || typeof token !== 'string') return null;
    const parts = String(token).split('.');
    if (parts.length !== 2) return null;
    let payloadB64 = parts[0];
    while (payloadB64.length % 4 !== 0) payloadB64 += '=';
    const payloadBytes = Utilities.base64DecodeWebSafe(payloadB64);
    const payloadStr = Utilities.newBlob(payloadBytes).getDataAsString();
    const expectedSig = Utilities.computeHmacSha256Signature(payloadStr, _getTokenSecret_());
    const expectedSigB64 = Utilities.base64EncodeWebSafe(expectedSig).replace(/=+$/, '');
    if (expectedSigB64 !== parts[1]) return null;
    const payload = JSON.parse(payloadStr);
    if (!payload || typeof payload.exp !== 'number' || Date.now() > payload.exp) return null;
    return payload; // { u, r, exp }
  } catch (e) {
    return null;
  }
}

/**
 * Validates credentials server-side. Returns { ok: true, role } on success,
 * { ok: false, error } on failure. Plaintext passwords found in the USERS
 * sheet are upgraded to hashed form on successful login.
 *
 * Built-in defaults (admin / admin1234, user / user1234) work only when the
 * USERS sheet has no row for that username — so the first thing an operator
 * should do is set real passwords.
 */
function login(username, password) {
  try {
    username = String(username || '').trim().toLowerCase();
    password = String(password || '').trim();
    if (!username || !password) return { ok: false, error: 'Missing credentials' };

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USERS');
    let stored = null, role = null, foundRow = -1;
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
        for (let i = 0; i < data.length; i++) {
          if (String(data[i][0] || '').trim().toLowerCase() === username) {
            stored   = String(data[i][1] || '');
            role     = String(data[i][2] || '').trim().toLowerCase() === 'admin' ? 'admin' : 'user';
            foundRow = i + 2;
            break;
          }
        }
      }
    }

    // Fallback to built-in defaults only when USERS has no entry for this name
    if (!stored) {
      if (username === 'admin' && password === 'admin1234') {
        return { ok: true, role: 'admin', token: _signToken_(username, 'admin') };
      }
      if (username === 'user' && password === 'user1234') {
        return { ok: true, role: 'user', token: _signToken_(username, 'user') };
      }
      return { ok: false, error: 'Invalid credentials' };
    }

    if (!_verifyPassword_(username, password, stored)) {
      return { ok: false, error: 'Invalid credentials' };
    }

    // Migrate plaintext → hashed on successful login
    if (stored.indexOf('sha256:') !== 0 && sheet && foundRow > 0) {
      try { sheet.getRange(foundRow, 2).setValue(_hashPassword_(username, password)); } catch (e) {}
    }

    return { ok: true, role: role, token: _signToken_(username, role) };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

/**
 * Returns the list of accounts from the USERS tab — usernames and roles only.
 * Passwords are NEVER returned, even hashed. Used by the admin Settings panel
 * to populate the user list and to check name uniqueness.
 */
function getUsers() {
  try {
    const sheet = getSheet_('USERS');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, accounts: {} };
    const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const accounts = {};
    values.forEach(function (row) {
      const username = String(row[0] || '').trim().toLowerCase();
      const role     = String(row[2] || '').trim().toLowerCase();
      if (username) accounts[username] = { role: role === 'admin' ? 'admin' : 'user' };
    });
    return { ok: true, accounts: accounts };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

/**
 * Creates or updates a user row in the USERS tab. Passwords are hashed before
 * storage. If the tab does not exist it is created with headers automatically.
 */
function saveUser(username, password, role) {
  try {
    username = String(username || '').trim().toLowerCase();
    password = String(password || '').trim();
    role     = String(role     || '').trim().toLowerCase() === 'admin' ? 'admin' : 'user';
    if (!username || !password) return { success: false, error: 'Username and password are required' };

    const hashed = _hashPassword_(username, password);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('USERS');
    if (!sheet) {
      sheet = ss.insertSheet('USERS');
      sheet.appendRow(['username', 'password', 'role']);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const col = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < col.length; i++) {
        if (String(col[i][0] || '').trim().toLowerCase() === username) {
          sheet.getRange(i + 2, 1, 1, 3).setValues([[username, hashed, role]]);
          return { success: true };
        }
      }
    }
    sheet.appendRow([username, hashed, role]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Deletes a user row from the USERS tab by username.
 */
function deleteUser(username) {
  try {
    username = String(username || '').trim().toLowerCase();
    const sheet = getSheet_('USERS');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'User not found' };
    const col = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = col.length - 1; i >= 0; i--) {
      if (String(col[i][0] || '').trim().toLowerCase() === username) {
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: false, error: 'User "' + username + '" not found' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Reads the column-visibility config JSON from the APP_CONFIG tab.
 * Returns { ok: true, value: '{}' } if the tab or key doesn't exist yet.
 */
function getColVis() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('APP_CONFIG');
    if (!sheet) return { ok: true, value: '{}' };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, value: '{}' };
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'col_vis') {
        return { ok: true, value: String(data[i][1] || '{}') };
      }
    }
    return { ok: true, value: '{}' };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

/**
 * Persists the column-visibility config JSON to the APP_CONFIG tab.
 * Creates the tab with headers if it doesn't exist.
 */
function saveColVis(jsonStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('APP_CONFIG');
    if (!sheet) {
      sheet = ss.insertSheet('APP_CONFIG');
      sheet.appendRow(['key', 'value']);
    }
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (String(keys[i][0]).trim() === 'col_vis') {
          sheet.getRange(i + 2, 2).setValue(jsonStr);
          return { success: true };
        }
      }
    }
    sheet.appendRow(['col_vis', jsonStr]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Reads the column-hidden config JSON from APP_CONFIG (key: col_hidden).
 * Columns in this config are fully hidden from everyone, including admin.
 */
function getColHidden() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('APP_CONFIG');
    if (!sheet) return { ok: true, value: '{}' };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, value: '{}' };
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'col_hidden') {
        return { ok: true, value: String(data[i][1] || '{}') };
      }
    }
    return { ok: true, value: '{}' };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

/**
 * Persists the column-hidden config JSON to APP_CONFIG (key: col_hidden).
 */
function saveColHidden(jsonStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('APP_CONFIG');
    if (!sheet) {
      sheet = ss.insertSheet('APP_CONFIG');
      sheet.appendRow(['key', 'value']);
    }
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (String(keys[i][0]).trim() === 'col_hidden') {
          sheet.getRange(i + 2, 2).setValue(jsonStr);
          return { success: true };
        }
      }
    }
    sheet.appendRow(['col_hidden', jsonStr]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
 * Expects row 1 to be a header row; data starts at row 2.
 * Col A = Item Code, Col B = Description.
 * Result is cached in CacheService for CACHE_TTL_SECONDS.
 */
function getItemCodes() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'itemcodes_map';
    const cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) { /* fall through */ }
    }
    const sheet = getSheet_('ITEMCODES');
    const lastRow = sheet.getLastRow();
    const result = { ok: true, map: {} };
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      values.forEach(function (row) {
        const code = String(row[0] || '').trim();
        const desc = String(row[1] || '').trim();
        if (code) result.map[code.toUpperCase()] = desc;
      });
    }
    try { cache.put(cacheKey, JSON.stringify(result), CACHE_TTL_SECONDS); } catch (e) {}
    return result;
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

/**
 * Reads the tab-availability config JSON from the APP_CONFIG tab (key: tab_config).
 * Returns { ok: true, value: '{}' } if the tab or key doesn't exist yet.
 */
function getTabConfig() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('tab_config');
  if (cached) return { ok: true, value: cached };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('APP_CONFIG');
    if (!sheet) return { ok: true, value: '{}' };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, value: '{}' };
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'tab_config') {
        const val = String(data[i][1] || '{}');
        try { cache.put('tab_config', val, CACHE_TTL_SECONDS); } catch(e) {}
        return { ok: true, value: val };
      }
    }
    return { ok: true, value: '{}' };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

/**
 * Persists the tab-availability config JSON to the APP_CONFIG tab (key: tab_config).
 */
function saveTabConfig(jsonStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('APP_CONFIG');
    if (!sheet) {
      sheet = ss.insertSheet('APP_CONFIG');
      sheet.appendRow(['key', 'value']);
    }
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (String(keys[i][0]).trim() === 'tab_config') {
          sheet.getRange(i + 2, 2).setValue(jsonStr);
          try { CacheService.getScriptCache().remove('tab_config'); } catch(e) {}
          return { success: true };
        }
      }
    }
    sheet.appendRow(['tab_config', jsonStr]);
    try { CacheService.getScriptCache().remove('tab_config'); } catch(e) {}
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =====================================================================
// ============ Personal Notes — Image Attachments (Google Drive) ======
// =====================================================================
// Images are stored as files in a Google Drive folder "PO_WEBAPP_IMAGES".
// The sheet "PERSONAL_NOTES_IMAGES" acts as an index:
//   [imageId | sheetName | rowIndex | fileName | mimeType | driveFileId | createdAt]
//
// getImages    — returns image list (id, name, mimeType) for a row; no binary data
// getImageData — returns base64 of ONE image by imageId (called on demand when viewing)
// saveImage    — uploads base64 to Drive, records fileId in index sheet
// deleteImage  — trashes Drive file and removes index row
//
// IMPORTANT: Requires Drive authorization. Run authorizeDrive() once from the
// Apps Script editor to trigger the permission prompt, then redeploy.

const IMAGES_INDEX_SHEET  = 'PERSONAL_NOTES_IMAGES';
const IMAGES_DRIVE_FOLDER = 'PO_WEBAPP_IMAGES';

/** One-time authorization helper — run this from the editor once, then delete. */
function authorizeDrive() {
  DriveApp.getRootFolder();
  Logger.log('Drive authorized successfully.');
}

function _getImagesFolder_() {
  const folders = DriveApp.getFoldersByName(IMAGES_DRIVE_FOLDER);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(IMAGES_DRIVE_FOLDER);
}

function _getImagesIndexSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(IMAGES_INDEX_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(IMAGES_INDEX_SHEET);
    sheet.appendRow(['imageId', 'sheetName', 'rowIndex', 'fileName', 'mimeType', 'driveFileId', 'createdAt']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Returns image list for a row — lightweight, no binary data.
 * Response: { ok: true, images: [{ id, name, mimeType }] }
 */
function getImages(sheetName, rowIndex) {
  try {
    sheetName = String(sheetName || '').trim();
    rowIndex  = String(rowIndex  || '').trim();
    const sheet = _getImagesIndexSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, images: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const images = [];
    data.forEach(function (row) {
      if (String(row[1] || '').trim() === sheetName &&
          String(row[2] || '').trim() === rowIndex  &&
          String(row[0] || '').trim()) {
        images.push({
          id:       String(row[0]),
          name:     String(row[3] || ''),
          mimeType: String(row[4] || 'image/png'),
          fileId:   String(row[5] || ''),
        });
      }
    });
    return { ok: true, images: images };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

/**
 * Fetches ONE image from Drive and returns it as base64.
 * Called on-demand when the viewer opens a specific image.
 * Response: { ok: true, base64, mimeType, name }
 */
function getImageData(imageId) {
  try {
    imageId = String(imageId || '').trim();
    if (!imageId) return { ok: false, error: 'imageId required' };

    const sheet = _getImagesIndexSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: false, error: 'Index sheet is empty — no images recorded' };
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === imageId) {
        const fileId   = String(data[i][5] || '').trim();
        const name     = String(data[i][3] || '');
        const mimeType = String(data[i][4] || 'image/png');

        if (!fileId) {
          return { ok: false, error: 'No Drive file ID for this image — it was uploaded before Drive was authorized. Please delete and re-upload this image.' };
        }

        let file;
        try {
          file = DriveApp.getFileById(fileId);
        } catch (driveErr) {
          return { ok: false, error: 'Drive file not found (ID: ' + fileId + '). It may have been deleted from Drive. Error: ' + driveErr.toString() };
        }

        const bytes  = file.getBlob().getBytes();
        const base64 = Utilities.base64Encode(bytes);
        return { ok: true, base64: base64, mimeType: mimeType, name: name };
      }
    }
    return { ok: false, error: 'Image ID "' + imageId + '" not found in index sheet' };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

/**
 * Debug helper — run this from the Apps Script editor to inspect the image index.
 * Check the Logs (View > Logs) after running.
 */
function debugImageIndex() {
  const sheet = _getImagesIndexSheet_();
  const lastRow = sheet.getLastRow();
  Logger.log('PERSONAL_NOTES_IMAGES rows: ' + (lastRow - 1));
  if (lastRow < 2) { Logger.log('No images recorded.'); return; }
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  data.forEach(function (row, i) {
    Logger.log('Row ' + (i+2) + ': imageId=' + row[0] + ' | sheet=' + row[1] + ' | rowIndex=' + row[2] + ' | file=' + row[3] + ' | mime=' + row[4] + ' | driveFileId=' + (row[5] ? row[5] : '*** EMPTY ***'));
  });
}

/**
 * Uploads a base64-encoded image to Drive and records its fileId in the index sheet.
 * Response: { success: true, imageId }
 */
function saveImage(sheetName, rowIndex, fileName, base64Data, mimeType) {
  try {
    sheetName  = String(sheetName  || '').trim();
    rowIndex   = String(rowIndex   || '').trim();
    fileName   = String(fileName   || 'image').replace(/[^\w.\-]/g, '_');
    mimeType   = String(mimeType   || 'image/png').trim();
    if (!base64Data) return { success: false, error: 'No image data provided' };

    const folder    = _getImagesFolder_();
    const imageId   = Utilities.getUuid();
    const driveName = sheetName + '__row' + rowIndex + '__' + imageId + '__' + fileName;
    const blob      = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, driveName);
    const file      = folder.createFile(blob);
    const fileId    = file.getId();

    const indexSheet = _getImagesIndexSheet_();
    const createdAt  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    indexSheet.appendRow([imageId, sheetName, rowIndex, fileName, mimeType, fileId, createdAt]);

    return { success: true, imageId: imageId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Trashes the Drive file and removes the index row.
 * Response: { success: true }
 */
function deleteImage(imageId) {
  try {
    imageId = String(imageId || '').trim();
    if (!imageId) return { success: false, error: 'imageId required' };

    const indexSheet = _getImagesIndexSheet_();
    const lastRow = indexSheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Image not found' };

    const data = indexSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0] || '').trim() === imageId) {
        try {
          const fileId = String(data[i][5] || '').trim();
          if (fileId) DriveApp.getFileById(fileId).setTrashed(true);
        } catch (e2) { /* file already gone — ignore */ }
        indexSheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: false, error: 'Image "' + imageId + '" not found' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
