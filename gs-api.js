/**
 * gs-api.js  —  Google Sheets API layer for Hardware Shop POS
 * All apps include this via <script src="gs-api.js">
 *
 * Key design decisions:
 *  - read()   always fetches live from Sheets, caches result locally
 *  - append() writes to Sheets first, updates cache on success
 *  - update() writes to Sheets first, updates cache on success
 *  - deleteRow() uses batchUpdate deleteDimension — actually removes the row
 *  - All writes queue when offline and flush on reconnect
 *  - _heads[tab] always populated from the last read() — ensures correct column order
 */

var GS = (function () {
  'use strict';

  var _ctx    = null;   // { token, tokenExpires, spreadsheetId, shopName, userName, userEmail }
  var _heads  = {};     // { tabName: ['col1','col2',...] }      — from last read
  var _sheetIds = {};   // { tabName: sheetId }                  — from spreadsheet metadata
  var _cache  = {};     // { tabName: [rowObj, ...] }            — local copy
  var _queue  = [];     // offline write queue
  var _online = true;

  var QUEUE_KEY = 'hs_sync_queue';
  var CACHE_KEY = 'hs_gs_cache';

  // ──────────────────────────────────────────────────────────
  // PUBLIC — INIT
  // ──────────────────────────────────────────────────────────
  function init() {
    try { _ctx = JSON.parse(localStorage.getItem('hs_gcontext') || 'null'); }
    catch (e) { _ctx = null; }

    if (!_ctx || !_ctx.spreadsheetId) {
      console.warn('GS: no context — localStorage-only mode');
      return Promise.resolve(false);
    }

    try { _queue = JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]'); } catch (e) { _queue = []; }
    try { _cache = JSON.parse(localStorage.getItem(CACHE_KEY) || '{}'); } catch (e) { _cache = {}; }

    window.addEventListener('online',  function () { _online = true;  flushQueue(); });
    window.addEventListener('offline', function () { _online = false; });
    _online = navigator.onLine;

    // Load sheet IDs (needed for deleteRow)
    return _loadSheetIds().then(function () { return true; }).catch(function () { return false; });
  }

  function isOnline() {
    return _online && !!_ctx && !!_ctx.token && _ctx.tokenExpires > Date.now();
  }

  // ──────────────────────────────────────────────────────────
  // PUBLIC — READ
  // Always fetches live from Sheets, updates cache + headers
  // ──────────────────────────────────────────────────────────
  function read(tab) {
    if (!isOnline()) return Promise.resolve(_cache[tab] || []);

    return _gFetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + _ctx.spreadsheetId
      + '/values/' + encodeURIComponent(tab) + '?majorDimension=ROWS'
    ).then(function (data) {
      var rows = data.values || [];
      if (rows.length < 1) { _cache[tab] = []; _saveCache(); return []; }

      var headers = rows[0];
      _heads[tab] = headers;

      var objs = rows.slice(1).map(function (row) {
        var obj = {};
        headers.forEach(function (h, i) { obj[h] = row[i] !== undefined ? row[i] : ''; });
        return obj;
      });

      _cache[tab] = objs;
      _saveCache();
      return objs;
    }).catch(function (err) {
      console.warn('GS.read(' + tab + ') error, using cache:', err);
      return _cache[tab] || [];
    });
  }

  // ──────────────────────────────────────────────────────────
  // PUBLIC — APPEND
  // Writes to Sheets, then updates local cache
  // ──────────────────────────────────────────────────────────
  function append(tab, rowObj) {
    var op = { type: 'append', tab: tab, rowObj: rowObj, ts: Date.now() };

    if (!isOnline()) {
      _cache[tab] = _cache[tab] || [];
      _cache[tab].push(rowObj);
      _saveCache();
      _queue.push(op);
      _saveQueue();
      return Promise.resolve({ queued: true });
    }

    return _ensureHeaders(tab).then(function () {
      return _doAppend(tab, rowObj);
    }).then(function (result) {
      _cache[tab] = _cache[tab] || [];
      _cache[tab].push(rowObj);
      _saveCache();
      return result;
    }).catch(function (err) {
      console.warn('GS.append(' + tab + ') failed, queuing:', err);
      _cache[tab] = _cache[tab] || [];
      _cache[tab].push(rowObj);
      _saveCache();
      _queue.push(op);
      _saveQueue();
    });
  }

  // ──────────────────────────────────────────────────────────
  // PUBLIC — UPDATE  (rowIndex is 1-based, after header)
  // ──────────────────────────────────────────────────────────
  function update(tab, rowIndex, rowObj) {
    var op = { type: 'update', tab: tab, rowIndex: rowIndex, rowObj: rowObj, ts: Date.now() };

    // Update cache immediately
    if (_cache[tab]) _cache[tab][rowIndex - 1] = rowObj;
    _saveCache();

    if (!isOnline()) { _queue.push(op); _saveQueue(); return Promise.resolve({ queued: true }); }

    return _ensureHeaders(tab).then(function () {
      return _doUpdate(tab, rowIndex, rowObj);
    }).catch(function (err) {
      console.warn('GS.update(' + tab + ') failed, queuing:', err);
      _queue.push(op);
      _saveQueue();
    });
  }

  // ──────────────────────────────────────────────────────────
  // PUBLIC — DELETE ROW  (rowIndex 1-based after header)
  // Uses batchUpdate deleteDimension — actually removes the row
  // ──────────────────────────────────────────────────────────
  function deleteRow(tab, rowIndex) {
    // Remove from cache
    if (_cache[tab]) { _cache[tab].splice(rowIndex - 1, 1); _saveCache(); }

    if (!isOnline()) {
      _queue.push({ type: 'delete', tab: tab, rowIndex: rowIndex, ts: Date.now() });
      _saveQueue();
      return Promise.resolve({ queued: true });
    }

    var sheetId = _sheetIds[tab];
    if (sheetId === undefined) {
      return _loadSheetIds().then(function () { return _doDeleteRow(tab, rowIndex); });
    }
    return _doDeleteRow(tab, rowIndex);
  }

  // ──────────────────────────────────────────────────────────
  // PUBLIC — FIND ROW  (searches cache, re-reads if not found)
  // ──────────────────────────────────────────────────────────
  function findRow(tab, field, value) {
    var rows = _cache[tab] || [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][field]) === String(value)) return { rowIndex: i + 1, row: rows[i] };
    }
    return null;
  }

  // ──────────────────────────────────────────────────────────
  // PUBLIC — UPSERT  (find + update or append)
  // Reads live from sheet first to ensure cache is current
  // ──────────────────────────────────────────────────────────
  function upsert(tab, keyField, rowObj) {
    if (!isOnline()) {
      // Offline: just append (dedup on next sync)
      return append(tab, rowObj);
    }
    // Re-read to get fresh data + headers before upsert
    return read(tab).then(function () {
      var found = findRow(tab, keyField, rowObj[keyField]);
      if (found) return update(tab, found.rowIndex, rowObj);
      return append(tab, rowObj);
    });
  }

  // ──────────────────────────────────────────────────────────
  // PUBLIC — UPDATE FIELD  (single cell by key lookup)
  // ──────────────────────────────────────────────────────────
  function updateField(tab, keyField, keyValue, field, value) {
    return read(tab).then(function () {
      var found = findRow(tab, keyField, keyValue);
      if (!found) return null;
      var updated = Object.assign({}, found.row);
      updated[field] = value;
      return update(tab, found.rowIndex, updated);
    });
  }

  // ──────────────────────────────────────────────────────────
  // PUBLIC — SETTINGS HELPERS
  // ──────────────────────────────────────────────────────────
  function readSettings() {
    return read('Settings').then(function (rows) {
      var out = {};
      rows.forEach(function (r) { if (r.key) out[r.key] = r.value; });
      return out;
    });
  }

  function writeSetting(key, value) {
    return read('Settings').then(function () {
      var found = findRow('Settings', 'key', key);
      if (found) return update('Settings', found.rowIndex, { key: key, value: value });
      return append('Settings', { key: key, value: value });
    });
  }

  // ──────────────────────────────────────────────────────────
  // PUBLIC — FLUSH OFFLINE QUEUE
  // ──────────────────────────────────────────────────────────
  function flushQueue() {
    if (!isOnline() || _queue.length === 0) return Promise.resolve();
    var ops = _queue.slice();
    _queue = [];
    _saveQueue();

    var chain = Promise.resolve();
    ops.forEach(function (op) {
      chain = chain.then(function () {
        if (op.type === 'append') return _ensureHeaders(op.tab).then(function () { return _doAppend(op.tab, op.rowObj); });
        if (op.type === 'update') return _ensureHeaders(op.tab).then(function () { return _doUpdate(op.tab, op.rowIndex, op.rowObj); });
        if (op.type === 'delete') return _doDeleteRow(op.tab, op.rowIndex);
        return Promise.resolve();
      }).catch(function (err) {
        _queue.push(op);
        _saveQueue();
        console.warn('GS.flushQueue op failed, re-queued:', err);
      });
    });
    return chain;
  }

  // ──────────────────────────────────────────────────────────
  // PUBLIC — CONTEXT ACCESSORS
  // ──────────────────────────────────────────────────────────
  function getSpreadsheetId() { return _ctx ? _ctx.spreadsheetId : null; }
  function getShopName()      { return _ctx ? _ctx.shopName : 'Hardware Shop'; }
  function getUserName()      { return _ctx ? _ctx.userName : ''; }
  function getUserEmail()     { return _ctx ? _ctx.userEmail : ''; }
  function getToken()         { return _ctx ? _ctx.token : null; }

  // ──────────────────────────────────────────────────────────
  // PRIVATE — LOAD SHEET IDs (needed for deleteRow)
  // ──────────────────────────────────────────────────────────
  function _loadSheetIds() {
    if (!isOnline()) return Promise.resolve();
    return _gFetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + _ctx.spreadsheetId
      + '?fields=sheets.properties'
    ).then(function (data) {
      (data.sheets || []).forEach(function (s) {
        _sheetIds[s.properties.title] = s.properties.sheetId;
      });
    });
  }

  // ──────────────────────────────────────────────────────────
  // PRIVATE — ENSURE HEADERS LOADED FOR TAB
  // ──────────────────────────────────────────────────────────
  function _ensureHeaders(tab) {
    if (_heads[tab] && _heads[tab].length) return Promise.resolve();
    // Read the header row only
    return _gFetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + _ctx.spreadsheetId
      + '/values/' + encodeURIComponent(tab) + '!1:1'
    ).then(function (data) {
      var rows = data.values || [];
      if (rows.length) _heads[tab] = rows[0];
    }).catch(function () {});
  }

  // ──────────────────────────────────────────────────────────
  // PRIVATE — DO APPEND
  // ──────────────────────────────────────────────────────────
  function _doAppend(tab, rowObj) {
    var values = [_rowToArray(tab, rowObj)];
    return _gFetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + _ctx.spreadsheetId
      + '/values/' + encodeURIComponent(tab) + ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS',
      {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ values: values })
      }
    );
  }

  // ──────────────────────────────────────────────────────────
  // PRIVATE — DO UPDATE  (rowIndex 1-based after header → sheet row = rowIndex+1)
  // ──────────────────────────────────────────────────────────
  function _doUpdate(tab, rowIndex, rowObj) {
    var sheetRow = rowIndex + 1; // header is row 1, data starts row 2
    var colCount = (_heads[tab] || Object.keys(rowObj)).length;
    var endCol   = _colLetter(colCount);
    var range    = encodeURIComponent(tab) + '!A' + sheetRow + ':' + endCol + sheetRow;
    var values   = [_rowToArray(tab, rowObj)];
    return _gFetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + _ctx.spreadsheetId
      + '/values/' + range + '?valueInputOption=RAW',
      {
        method:  'PUT',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ values: values })
      }
    );
  }

  // ──────────────────────────────────────────────────────────
  // PRIVATE — DO DELETE ROW via batchUpdate
  // ──────────────────────────────────────────────────────────
  function _doDeleteRow(tab, rowIndex) {
    var sheetId = _sheetIds[tab];
    if (sheetId === undefined) return Promise.resolve();
    var startIndex = rowIndex; // 0-based: header=0, first data row=1
    return _gFetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + _ctx.spreadsheetId + ':batchUpdate',
      {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({
          requests: [{
            deleteDimension: {
              range: {
                sheetId:    sheetId,
                dimension:  'ROWS',
                startIndex: startIndex,     // row to delete (0-based)
                endIndex:   startIndex + 1
              }
            }
          }]
        })
      }
    );
  }

  // ──────────────────────────────────────────────────────────
  // PRIVATE — ROW OBJECT → ARRAY  (ordered by sheet headers)
  // ──────────────────────────────────────────────────────────
  function _rowToArray(tab, rowObj) {
    var headers = _heads[tab];
    if (!headers || !headers.length) {
      // No headers loaded — use object key order (fallback only)
      return Object.keys(rowObj).map(function (k) { return rowObj[k] !== undefined ? String(rowObj[k]) : ''; });
    }
    return headers.map(function (h) {
      var v = rowObj[h];
      return (v !== undefined && v !== null) ? String(v) : '';
    });
  }

  // ──────────────────────────────────────────────────────────
  // PRIVATE — COLUMN LETTER  (1→A, 26→Z, 27→AA …)
  // ──────────────────────────────────────────────────────────
  function _colLetter(n) {
    var s = '';
    while (n > 0) {
      var r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s || 'A';
  }

  // ──────────────────────────────────────────────────────────
  // PRIVATE — AUTHENTICATED FETCH
  // ──────────────────────────────────────────────────────────
  function _gFetch(url, options) {
    options = options || {};
    options.headers = Object.assign(
      { 'Authorization': 'Bearer ' + (_ctx ? _ctx.token : '') },
      options.headers || {}
    );
    return fetch(url, options).then(function (r) {
      if (!r.ok) return r.json().then(function (e) { throw new Error(JSON.stringify(e)); });
      return r.json();
    });
  }

  function _saveCache() { try { localStorage.setItem(CACHE_KEY, JSON.stringify(_cache)); } catch (e) {} }
  function _saveQueue() { try { localStorage.setItem(QUEUE_KEY, JSON.stringify(_queue)); } catch (e) {} }

  // ──────────────────────────────────────────────────────────
  // EXPOSE PUBLIC API
  // ──────────────────────────────────────────────────────────
  return {
    init:             init,
    isOnline:         isOnline,
    read:             read,
    append:           append,
    update:           update,
    deleteRow:        deleteRow,
    findRow:          findRow,
    upsert:           upsert,
    updateField:      updateField,
    flushQueue:       flushQueue,
    readSettings:     readSettings,
    writeSetting:     writeSetting,
    getSpreadsheetId: getSpreadsheetId,
    getShopName:      getShopName,
    getUserName:      getUserName,
    getUserEmail:     getUserEmail,
    getToken:         getToken,
    get syncQueue()   { return _queue; },
    get cache()       { return _cache; },
    get heads()       { return _heads; },
  };
})();
