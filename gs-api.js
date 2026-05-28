/**
 * gs-api.js  —  Google Sheets API layer for Hardware Shop POS
 * Included by store-app, till-app, admin-app via <script src="gs-api.js">
 *
 * Provides:
 *   GS.init()          — call on app load, reads context from hub
 *   GS.read(tab)       — returns array of row objects for a sheet tab
 *   GS.append(tab, row)— appends one row object
 *   GS.update(tab, rowIndex, row) — updates a specific row (1-based, after header)
 *   GS.batchUpdate(requests) — raw batchUpdate for bulk ops
 *   GS.findRow(tab, field, value) — returns { rowIndex, row } or null
 *   GS.isOnline()      — returns true if token valid and network available
 *   GS.syncQueue       — array of pending writes for offline mode
 *   GS.flushQueue()    — flush pending writes when back online
 */

var GS = (function() {
  'use strict';

  var _ctx   = null;   // { token, tokenExpires, spreadsheetId, shopName, userName, userEmail }
  var _heads = {};     // { tabName: [col headers] }
  var _cache = {};     // { tabName: [ {row object}, ... ] }
  var _queue = [];     // offline write queue
  var _online = true;

  var QUEUE_KEY = 'hs_sync_queue';
  var CACHE_KEY = 'hs_gs_cache';

  // ── PUBLIC API ──────────────────────────────────────────

  function init() {
    // Load context written by index.html
    try {
      _ctx = JSON.parse(localStorage.getItem('hs_gcontext') || 'null');
    } catch(e) { _ctx = null; }

    if (!_ctx || !_ctx.spreadsheetId) {
      console.warn('GS: no context — running in localStorage-only mode');
      return Promise.resolve(false);
    }

    // Load offline queue
    try { _queue = JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]'); } catch(e) { _queue = []; }

    // Load local cache
    try { _cache = JSON.parse(localStorage.getItem(CACHE_KEY) || '{}'); } catch(e) { _cache = {}; }

    // Monitor online status
    window.addEventListener('online',  function() { _online = true;  flushQueue(); });
    window.addEventListener('offline', function() { _online = false; });
    _online = navigator.onLine;

    return Promise.resolve(true);
  }

  function isOnline() {
    return _online && _ctx && _ctx.token && _ctx.tokenExpires > Date.now();
  }

  // ── READ ─────────────────────────────────────────────────
  function read(tab) {
    if (!isOnline()) {
      // Return from cache
      return Promise.resolve(_cache[tab] || []);
    }

    return _gFetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + _ctx.spreadsheetId
      + '/values/' + encodeURIComponent(tab) + '?majorDimension=ROWS'
    ).then(function(data) {
      var rows = data.values || [];
      if (rows.length < 2) { _cache[tab] = []; _saveCache(); return []; }

      var headers = rows[0];
      _heads[tab] = headers;

      var objs = rows.slice(1).map(function(row) {
        var obj = {};
        headers.forEach(function(h, i) { obj[h] = row[i] !== undefined ? row[i] : ''; });
        return obj;
      });

      _cache[tab] = objs;
      _saveCache();
      return objs;
    }).catch(function(err) {
      console.warn('GS.read error, using cache:', err);
      return _cache[tab] || [];
    });
  }

  // ── APPEND ───────────────────────────────────────────────
  function append(tab, rowObj) {
    // Always update cache immediately
    _cache[tab] = _cache[tab] || [];
    _cache[tab].push(rowObj);
    _saveCache();

    var op = { type: 'append', tab, rowObj, ts: Date.now() };

    if (!isOnline()) {
      _queue.push(op);
      _saveQueue();
      return Promise.resolve({ queued: true });
    }

    return _doAppend(tab, rowObj).catch(function(err) {
      console.warn('GS.append failed, queuing:', err);
      _queue.push(op);
      _saveQueue();
    });
  }

  // ── UPDATE (by row index 1-based after header) ────────────
  function update(tab, rowIndex, rowObj) {
    // Update cache
    if (_cache[tab] && _cache[tab][rowIndex - 1]) {
      _cache[tab][rowIndex - 1] = rowObj;
      _saveCache();
    }

    var op = { type: 'update', tab, rowIndex, rowObj, ts: Date.now() };

    if (!isOnline()) {
      _queue.push(op);
      _saveQueue();
      return Promise.resolve({ queued: true });
    }

    return _doUpdate(tab, rowIndex, rowObj).catch(function(err) {
      console.warn('GS.update failed, queuing:', err);
      _queue.push(op);
      _saveQueue();
    });
  }

  // ── FIND ROW ─────────────────────────────────────────────
  function findRow(tab, field, value) {
    var rows = _cache[tab] || [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][field]) === String(value)) {
        return { rowIndex: i + 1, row: rows[i] };
      }
    }
    return null;
  }

  // ── UPSERT (find + update or append) ─────────────────────
  function upsert(tab, keyField, rowObj) {
    var found = findRow(tab, keyField, rowObj[keyField]);
    if (found) {
      return update(tab, found.rowIndex, rowObj);
    }
    return append(tab, rowObj);
  }

  // ── UPDATE FIELD (single cell) ────────────────────────────
  function updateField(tab, keyField, keyValue, field, value) {
    var found = findRow(tab, keyField, keyValue);
    if (!found) return Promise.resolve(null);
    var updated = Object.assign({}, found.row);
    updated[field] = value;
    return update(tab, found.rowIndex, updated);
  }

  // ── BATCH UPDATE (raw Sheets API) ─────────────────────────
  function batchUpdate(requests) {
    if (!isOnline()) return Promise.resolve({ queued: true });
    return _gFetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + _ctx.spreadsheetId + ':batchUpdate',
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ requests })
      }
    );
  }

  // ── FLUSH OFFLINE QUEUE ──────────────────────────────────
  function flushQueue() {
    if (!isOnline() || _queue.length === 0) return Promise.resolve();

    var ops   = _queue.slice();
    _queue    = [];
    _saveQueue();

    var chain = Promise.resolve();
    ops.forEach(function(op) {
      chain = chain.then(function() {
        if (op.type === 'append') return _doAppend(op.tab, op.rowObj);
        if (op.type === 'update') return _doUpdate(op.tab, op.rowIndex, op.rowObj);
        return Promise.resolve();
      }).catch(function(err) {
        // Re-queue failed ops
        _queue.push(op);
        _saveQueue();
        console.warn('GS.flushQueue op failed, re-queued:', err);
      });
    });

    return chain;
  }

  // ── CONTEXT HELPERS ──────────────────────────────────────
  function getSpreadsheetId() { return _ctx ? _ctx.spreadsheetId : null; }
  function getShopName()      { return _ctx ? _ctx.shopName      : 'Hardware Shop'; }
  function getUserName()      { return _ctx ? _ctx.userName      : ''; }
  function getUserEmail()     { return _ctx ? _ctx.userEmail     : ''; }
  function getToken()         { return _ctx ? _ctx.token         : null; }

  // ── SETTINGS HELPERS ─────────────────────────────────────
  async function readSettings() {
    var rows = await read('Settings');
    var out  = {};
    rows.forEach(function(r) { if (r.key) out[r.key] = r.value; });
    return out;
  }

  async function writeSetting(key, value) {
    var found = findRow('Settings', 'key', key);
    if (found) {
      return update('Settings', found.rowIndex, { key, value });
    }
    return append('Settings', { key, value });
  }

  // ── PRIVATE HELPERS ──────────────────────────────────────
  function _gFetch(url, options) {
    options = options || {};
    options.headers = Object.assign(
      { 'Authorization': 'Bearer ' + (_ctx ? _ctx.token : '') },
      options.headers || {}
    );
    return fetch(url, options).then(function(r) {
      if (!r.ok) throw new Error('HTTP ' + r.status + ' ' + url);
      return r.json();
    });
  }

  function _rowToArray(tab, rowObj) {
    var headers = _heads[tab];
    if (!headers || !headers.length) return Object.values(rowObj).map(String);
    return headers.map(function(h) { return rowObj[h] !== undefined ? String(rowObj[h]) : ''; });
  }

  function _doAppend(tab, rowObj) {
    var values = [_rowToArray(tab, rowObj)];
    return _gFetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + _ctx.spreadsheetId
      + '/values/' + encodeURIComponent(tab) + '!A1:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS',
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ values })
      }
    );
  }

  function _doUpdate(tab, rowIndex, rowObj) {
    var sheetRow = rowIndex + 1; // +1 for header row
    var values   = [_rowToArray(tab, rowObj)];
    return _gFetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + _ctx.spreadsheetId
      + '/values/' + encodeURIComponent(tab) + '!A' + sheetRow + '?valueInputOption=USER_ENTERED',
      {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ values })
      }
    );
  }

  function _saveCache() {
    try { localStorage.setItem(CACHE_KEY, JSON.stringify(_cache)); } catch(e) {}
  }

  function _saveQueue() {
    try { localStorage.setItem(QUEUE_KEY, JSON.stringify(_queue)); } catch(e) {}
  }

  // ── EXPOSE ───────────────────────────────────────────────
  return {
    init,
    isOnline,
    read,
    append,
    update,
    upsert,
    updateField,
    findRow,
    batchUpdate,
    flushQueue,
    readSettings,
    writeSetting,
    getSpreadsheetId,
    getShopName,
    getUserName,
    getUserEmail,
    getToken,
    get syncQueue() { return _queue; },
    get cache()     { return _cache; },
  };
})();
