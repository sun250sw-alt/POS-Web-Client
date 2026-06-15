/**
 * gs-api.js  —  Optimized Google Sheets API layer for Hardware Shop POS
 * All apps include this via <script src="gs-api.js">
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

  // Automatically refresh credential context from shared storage
  function _reloadContext() {
    try {
      var fresh = JSON.parse(localStorage.getItem('hs_gcontext') || 'null');
      if (fresh) { _ctx = fresh; }
    } catch (e) {}
  }

  // ──────────────────────────────────────────────────────────
  // PUBLIC — INIT
  // ──────────────────────────────────────────────────────────
  function init() {
    _reloadContext();

    if (!_ctx || !_ctx.spreadsheetId) {
      console.warn('GS: no context — localStorage-only mode');
      return Promise.resolve(false);
    }

    try { _queue = JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]'); } catch (e) { _queue = []; }
    try { _cache = JSON.parse(localStorage.getItem(CACHE_KEY) || '{}'); } catch (e) { _cache = {}; }

    window.addEventListener('online',  function () { _online = true;  flushQueue(); });
    window.addEventListener('offline', function () { _online = false; });
    _online = navigator.onLine;

    var tokenOk = !!_ctx.token && _ctx.tokenExpires > Date.now();
    if (!tokenOk) {
      console.warn('GS: token expired — need re-auth');
      return Promise.resolve(false);
    }

    return _loadSheetIds().then(function () { return true; }).catch(function () { return false; });
  }

  function clearCache() {
    _cache = {};
    _heads = {};
    try { localStorage.removeItem(CACHE_KEY); } catch(e) {}
    console.log('GS cache cleared');
  }

  function isOnline() {
    _reloadContext();
    return _online && !!_ctx && !!_ctx.spreadsheetId &&
           !!_ctx.token && _ctx.tokenExpires > Date.now();
  }

  function isNetworkUp() {
    return _online && !!_ctx && !!_ctx.spreadsheetId;
  }

  function isTokenExpired() {
    _reloadContext();
    return !!_ctx && !!_ctx.spreadsheetId && (!_ctx.token || _ctx.tokenExpires <= Date.now());
  }

  function hasSpreadsheet() {
    _reloadContext();
    return !!_ctx && !!_ctx.spreadsheetId;
  }

  // ──────────────────────────────────────────────────────────
  // PUBLIC — READ
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
  // PUBLIC — UPDATE
  // ──────────────────────────────────────────────────────────
  function update(tab, rowIndex, rowObj) {
    var op = { type: 'update', tab: tab, rowIndex: rowIndex, rowObj: rowObj, ts: Date.now() };

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
  // PUBLIC — BATCH UPDATE (Saves multiple rows in a single HTTP request)
  // ──────────────────────────────────────────────────────────
  function batchUpdateRows(tab, updates) {
    // updates: Array of { rowIndex: Number (1-based), rowObj: Object }
    if (_cache[tab]) {
      updates.forEach(function (u) {
        _cache[tab][u.rowIndex - 1] = u.rowObj;
      });
      _saveCache();
    }

    if (!isOnline()) {
      updates.forEach(function (u) {
        _queue.push({ type: 'update', tab: tab, rowIndex: u.rowIndex, rowObj: u.rowObj, ts: Date.now() });
      });
      _saveQueue();
      return Promise.resolve({ queued: true });
    }

    return _ensureHeaders(tab).then(function () {
      var data = updates.map(function (u) {
        var sheetRow = u.rowIndex + 1;
        var colCount = (_heads[tab] || Object.keys(u.rowObj)).length;
        var endCol   = _colLetter(colCount);
        var range    = encodeURIComponent(tab) + '!A' + sheetRow + ':' + endCol + sheetRow;
        return {
          range: range,
          values: [_rowToArray(tab, u.rowObj)]
        };
      });

      return _gFetch(
        'https://sheets.googleapis.com/v4/spreadsheets/' + _ctx.spreadsheetId + '/values:batchUpdate',
        {
          method:  'POST',
          headers: { 'Content-Type': 'application/json' },
          body:    JSON.stringify({
            valueInputOption: 'RAW',
            data: data
          })
        }
      );
    });
  }

  // ──────────────────────────────────────────────────────────
  // PUBLIC — DELETE ROW
  // ──────────────────────────────────────────────────────────
  function deleteRow(tab, rowIndex) {
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

  function findRow(tab, field, value) {
    var rows = _cache[tab] || [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][field]) === String(value)) return { rowIndex: i + 1, row: rows[i] };
    }
    return null;
  }

  function upsert(tab, keyField, rowObj) {
    if (!isOnline()) {
      return append(tab, rowObj);
    }
    return read(tab).then(function () {
      var found = findRow(tab, keyField, rowObj[keyField]);
      if (found) return update(tab, found.rowIndex, rowObj);
      return append(tab, rowObj);
    });
  }

  function updateField(tab, keyField, keyValue, field, value) {
    return read(tab).then(function () {
      var found = findRow(tab, keyField, keyValue);
      if (!found) return null;
      var updated = Object.assign({}, found.row);
      updated[field] = value;
      return update(tab, found.rowIndex, updated);
    });
  }

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

  // Preserves strict write-order sequentially, stopping on first error
  function flushQueue() {
    if (!isOnline() || _queue.length === 0) return Promise.resolve();

    var index = 0;
    function processNext() {
      if (index >= _queue.length) {
        _queue = [];
        _saveQueue();
        return Promise.resolve();
      }

      var op = _queue[index];
      return Promise.resolve().then(function () {
        if (op.type === 'append') return _ensureHeaders(op.tab).then(function () { return _doAppend(op.tab, op.rowObj); });
        if (op.type === 'update') return _ensureHeaders(op.tab).then(function () { return _doUpdate(op.tab, op.rowIndex, op.rowObj); });
        if (op.type === 'delete') return _doDeleteRow(op.tab, op.rowIndex);
        return Promise.resolve();
      }).then(function () {
        index++;
        return processNext();
      }).catch(function (err) {
        console.warn('Queue flush suspended at operation ' + index + ' due to error:', err);
        _queue = _queue.slice(index);
        _saveQueue();
        throw err;
      });
    }

    return processNext();
  }

  function getSpreadsheetId() { return _ctx ? _ctx.spreadsheetId : null; }
  function getShopName()      { return _ctx ? _ctx.shopName : 'Hardware Shop'; }
  function getUserName()      { return _ctx ? _ctx.userName : ''; }
  function getUserEmail()     { return _ctx ? _ctx.userEmail : ''; }
  function getToken()         { return _ctx ? _ctx.token : null; }

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

  function _ensureHeaders(tab) {
    if (_heads[tab] && _heads[tab].length) return Promise.resolve();
    return _gFetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + _ctx.spreadsheetId
      + '/values/' + encodeURIComponent(tab) + '!1:1'
    ).then(function (data) {
      var rows = data.values || [];
      if (rows.length) _heads[tab] = rows[0];
    }).catch(function () {});
  }

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

  function _doUpdate(tab, rowIndex, rowObj) {
    var sheetRow = rowIndex + 1;
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

  function _doDeleteRow(tab, rowIndex) {
    var sheetId = _sheetIds[tab];
    if (sheetId === undefined) return Promise.resolve();
    var startIndex = rowIndex;
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
                startIndex: startIndex,
                endIndex:   startIndex + 1
              }
            }
          }]
        })
      }
    );
  }

  function _rowToArray(tab, rowObj) {
    var headers = _heads[tab];
    if (!headers || !headers.length) {
      return Object.keys(rowObj).map(function (k) { return rowObj[k] !== undefined ? String(rowObj[k]) : ''; });
    }
    return headers.map(function (h) {
      var v = rowObj[h];
      return (v !== undefined && v !== null) ? String(v) : '';
    });
  }

  function _colLetter(n) {
    var s = '';
    while (n > 0) {
      var r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s || 'A';
  }

  function _gFetch(url, options) {
    _reloadContext();
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

  return {
    init:             init,
    isOnline:         isOnline,
    isNetworkUp:      isNetworkUp,
    isTokenExpired:   isTokenExpired,
    hasSpreadsheet:   hasSpreadsheet,
    clearCache:       clearCache,
    read:             read,
    append:           append,
    update:           update,
    batchUpdateRows:  batchUpdateRows,
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
