// =========================
// LINE Ride - Apps Script
// 強化完整覆蓋版 v1.1
// =========================

const SHEET_ID = '1OyEUm52Yh7_B235AtKtcDan0XvfbpsKp5KTHRhSDXsA';
const APP_NAME = 'LINE Ride';
const TIMEZONE = Session.getScriptTimeZone() || 'Asia/Taipei';

var SHEET_CACHE_ = {};

// =========================
// 入口
// =========================
function doGet(e) {
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    const action = lower_(p.action);
    const page = clean_(p.page) || 'home';

    if (action === 'health') {
      return jsonOutput_({
        ok: true,
        app: APP_NAME,
        time: new Date().toISOString(),
        timezone: TIMEZONE,
        config: validateRequiredSettings_()
      });
    }

    if (action === 'init') {
      return jsonOutput_(initSystem());
    }

    if (action === 'config') {
      return jsonOutput_({
        ok: true,
        settings: getPublicConfig_()
      });
    }

    const htmlFile = resolveHtmlFile_(page);
    const t = HtmlService.createTemplateFromFile(htmlFile);

    t.boot = JSON.stringify({
      ok: true,
      appName: APP_NAME,
      page: page,
      orderId: clean_(p.orderId),
      role: clean_(p.role),
      liffId: getSetting_('liff_id'),
      appUrl: getSetting_('app_url'),
      settings: getPublicConfig_()
    });

    return t.evaluate()
      .setTitle(APP_NAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    logError_('doGet', err, {
      parameter: (e && e.parameter) ? e.parameter : null
    });

    return ContentService
      .createTextOutput('系統錯誤：' + safeErrorMessage_(err))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function doPost(e) {
  try {
    const body = parsePostBody_(e);

    if (body && body.action) {
      return handleApiAction_(body);
    }

    const events = Array.isArray(body.events) ? body.events : [];
    const cache = CacheService.getScriptCache();

    events.forEach(function(event) {
      try {
        const userId = getNested_(event, ['source', 'userId']) || 'common';
        const ts = clean_(event.timestamp) || String(Date.now());
        const cacheKey = 'ev_' + userId + '_' + ts;

        if (cache.get(cacheKey)) return;
        cache.put(cacheKey, '1', 30);

        handleLineEvent_(event);
      } catch (err) {
        logError_('EventProcessing', err, event);
      }
    });

    return textOutput_('ok');
  } catch (err) {
    logError_('doPost', err, {
      raw: (e && e.postData && e.postData.contents) ? e.postData.contents : ''
    });
    return textOutput_('ok');
  }
}

// =========================
// API Action Router
// =========================
function handleApiAction_(body) {
  try {
    const action = lower_(body.action);

    if (action === 'ping') {
      return jsonOutput_({
        ok: true,
        pong: true,
        time: new Date().toISOString()
      });
    }

    if (action === 'init') {
      return jsonOutput_(initSystem());
    }

    if (action === 'config') {
      return jsonOutput_({
        ok: true,
        settings: getPublicConfig_(),
        required: validateRequiredSettings_()
      });
    }

    if (action === 'create_order') {
      return jsonOutput_(createOrder_(body));
    }

    if (action === 'cancel_order') {
      return jsonOutput_(cancelOrder_(body));
    }

    if (action === 'list_orders') {
      return jsonOutput_(listOrders_(body));
    }

    if (action === 'save_favorite') {
      return jsonOutput_(saveFavorite_(body));
    }

    if (action === 'delete_favorite') {
      return jsonOutput_(deleteFavorite_(body));
    }

    if (action === 'list_favorites') {
      return jsonOutput_(listFavorites_(body));
    }

    if (action === 'blacklist_add') {
      return jsonOutput_(addBlacklist_(body));
    }

    if (action === 'blacklist_list') {
      return jsonOutput_(listBlacklist_(body));
    }

    return jsonOutput_({
      ok: false,
      message: '未知 action：' + clean_(body.action)
    });
  } catch (err) {
    logError_('handleApiAction', err, body);
    return jsonOutput_({
      ok: false,
      message: safeErrorMessage_(err)
    });
  }
}

// =========================
// 初始化
// =========================
function initSystem() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const tables = {
    Users: [
      'line_uid',
      'role',
      'display_name',
      'online',
      'last_lat',
      'last_lon',
      'last_update',
      'created_at',
      'updated_at'
    ],
    Orders: [
      'order_id',
      'customer_id',
      'customer_name',
      'driver_id',
      'driver_name',
      'status',
      'pickup_lat',
      'pickup_lon',
      'pickup_label',
      'drop_lat',
      'drop_lon',
      'drop_label',
      'quoted_distance_km',
      'quoted_eta_min',
      'quoted_fare',
      'search_expires_at',
      'created_at',
      'matched_at',
      'started_at',
      'completed_at',
      'cancelled_at',
      'customer_rating',
      'driver_rating'
    ],
    Settings: [
      'key',
      'value'
    ],
    Logs: [
      'time',
      'type',
      'payload'
    ],
    Blacklist: [
      'blacklist_id',
      'target_user_id',
      'reason',
      'created_by',
      'created_at'
    ],
    Favorites: [
      'favorite_id',
      'user_id',
      'label',
      'address',
      'lat',
      'lon',
      'created_at'
    ]
  };

  Object.keys(tables).forEach(function(name) {
    ensureSheet_(ss, name, tables[name]);
  });

  const settings = {
    line_channel_access_token: '',
    liff_id: '',
    app_url: '',
    search_radius_km: '10',
    fare_base: '80',
    fare_per_km: '20',
    eta_minutes_per_km: '3',
    order_search_expire_minutes: '5'
  };

  Object.keys(settings).forEach(function(key) {
    if (!hasSetting_(key)) {
      upsertSetting_(key, settings[key]);
    }
  });

  clearSheetCache_();

  return {
    ok: true,
    message: '初始化完成',
    sheets: Object.keys(tables),
    required: validateRequiredSettings_()
  };
}

// =========================
// HTML / Template
// =========================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function resolveHtmlFile_(page) {
  const candidates = [];
  const cleanPage = clean_(page);

  if (cleanPage) {
    candidates.push(cleanPage);
  }
  candidates.push('index');

  for (var i = 0; i < candidates.length; i++) {
    if (fileExists_(candidates[i])) return candidates[i];
  }

  throw new Error('找不到 HTML 檔案，已嘗試：' + candidates.join(', '));
}

function fileExists_(filename) {
  try {
    HtmlService.createHtmlOutputFromFile(filename);
    return true;
  } catch (err) {
    return false;
  }
}

// =========================
// 基本工具
// =========================
function clean_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function lower_(v) {
  return clean_(v).toLowerCase();
}

function toNumber_(v, fallback) {
  const n = Number(v);
  return isNaN(n) ? (fallback || 0) : n;
}

function toBool_(v) {
  const s = lower_(v);
  return s === '1' || s === 'true' || s === 'yes' || s === 'y';
}

function now_() {
  return new Date();
}

function uuid_() {
  return Utilities.getUuid();
}

function formatDateTime_(d) {
  if (!d) return '';
  return Utilities.formatDate(d, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function safeErrorMessage_(err) {
  if (!err) return 'Unknown error';
  return clean_(err.message) || String(err);
}

function getNested_(obj, path) {
  try {
    return path.reduce(function(acc, key) {
      return acc && acc[key];
    }, obj);
  } catch (e) {
    return '';
  }
}

function textOutput_(text) {
  return ContentService
    .createTextOutput(clean_(text))
    .setMimeType(ContentService.MimeType.TEXT);
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function parsePostBody_(e) {
  const contents = (e && e.postData && e.postData.contents)
    ? e.postData.contents
    : '{}';

  try {
    return JSON.parse(contents);
  } catch (err) {
    throw new Error('POST JSON 解析失敗：' + err.message);
  }
}

function clearSheetCache_() {
  SHEET_CACHE_ = {};
}

// =========================
// Settings
// =========================
function hasSetting_(key) {
  return getSettingRow_(key) !== null;
}

function getSettingRow_(key) {
  const rows = getObjects_('Settings');
  const target = clean_(key);

  for (var i = 0; i < rows.length; i++) {
    if (clean_(rows[i].key) === target) {
      return rows[i];
    }
  }
  return null;
}

function getSetting_(key) {
  const row = getSettingRow_(key);
  return row ? clean_(row.value) : '';
}

function getSettingNumber_(key, fallback) {
  const v = getSetting_(key);
  const n = Number(v);
  return isNaN(n) ? fallback : n;
}

function upsertSetting_(key, value) {
  const sh = getSheet_('Settings');
  const data = sh.getDataRange().getValues();
  let updated = false;

  for (var i = 1; i < data.length; i++) {
    if (clean_(data[i][0]) === clean_(key)) {
      sh.getRange(i + 1, 2).setValue(value);
      updated = true;
      break;
    }
  }

  if (!updated) {
    sh.appendRow([key, value]);
  }

  clearSheetCache_();
}

function validateRequiredSettings_() {
  return {
    line_channel_access_token: !!getSetting_('line_channel_access_token'),
    liff_id: !!getSetting_('liff_id'),
    app_url: !!getSetting_('app_url')
  };
}

function getPublicConfig_() {
  return {
    liffId: getSetting_('liff_id'),
    appUrl: getSetting_('app_url'),
    searchRadiusKm: getSettingNumber_('search_radius_km', 10),
    fareBase: getSettingNumber_('fare_base', 80),
    farePerKm: getSettingNumber_('fare_per_km', 20),
    etaMinutesPerKm: getSettingNumber_('eta_minutes_per_km', 3),
    orderSearchExpireMinutes: getSettingNumber_('order_search_expire_minutes', 5)
  };
}

// =========================
// Sheet Helpers
// =========================
function getSheet_(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(name);
  if (!sh) {
    throw new Error('找不到工作表：' + name);
  }
  return sh;
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
  }

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold');
    sh.setFrozenRows(1);
    return;
  }

  const currentLastColumn = Math.max(sh.getLastColumn(), headers.length);
  const firstRow = sh.getRange(1, 1, 1, currentLastColumn).getValues()[0];
  const existing = firstRow.map(function(h) { return clean_(h); });

  let changed = false;

  headers.forEach(function(header, idx) {
    if (existing[idx] !== clean_(header)) {
      sh.getRange(1, idx + 1).setValue(header).setFontWeight('bold');
      changed = true;
    }
  });

  if (changed) {
    sh.setFrozenRows(1);
  }
}

function getObjects_(sheetName) {
  if (SHEET_CACHE_[sheetName]) {
    return SHEET_CACHE_[sheetName];
  }

  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();

  if (!values || values.length < 2) {
    SHEET_CACHE_[sheetName] = [];
    return [];
  }

  const headers = values[0].map(function(h) {
    return clean_(h);
  });

  const rows = values.slice(1).map(function(row, index) {
    const obj = { _row: index + 2 };
    headers.forEach(function(h, j) {
      obj[h] = row[j];
    });
    return obj;
  });

  SHEET_CACHE_[sheetName] = rows;
  return rows;
}

function appendObject_(sheetName, obj, headerOrder) {
  const sh = getSheet_(sheetName);
  const headers = (headerOrder && headerOrder.length)
    ? headerOrder
    : sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const row = headers.map(function(h) {
    return obj.hasOwnProperty(h) ? obj[h] : '';
  });

  sh.appendRow(row);
  clearSheetCache_();
}

function updateRowByRowNumber_(sheetName, rowNumber, updates) {
  const sh = getSheet_(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  headers.forEach(function(header, i) {
    if (updates.hasOwnProperty(header)) {
      sh.getRange(rowNumber, i + 1).setValue(updates[header]);
    }
  });

  clearSheetCache_();
}

function findRowByField_(sheetName, field, value) {
  const rows = getObjects_(sheetName);
  const target = clean_(value);

  for (var i = 0; i < rows.length; i++) {
    if (clean_(rows[i][field]) === target) {
      return rows[i];
    }
  }
  return null;
}

// =========================
// Logs
// =========================
function logError_(type, err, payload) {
  try {
    getSheet_('Logs').appendRow([
      new Date(),
      clean_(type),
      JSON.stringify({
        message: safeErrorMessage_(err),
        stack: err && err.stack ? String(err.stack) : '',
        payload: payload || null
      })
    ]);
    clearSheetCache_();
  } catch (logErr) {
    console.error(logErr);
  }
}

function logInfo_(type, payload) {
  try {
    getSheet_('Logs').appendRow([
      new Date(),
      clean_(type),
      JSON.stringify(payload || {})
    ]);
    clearSheetCache_();
  } catch (err) {
    console.error(err);
  }
}

// =========================
// LINE Webhook
// =========================
function handleLineEvent_(event) {
  if (!event || !event.type) return;

  const replyToken = clean_(event.replyToken);
  const userId = getNested_(event, ['source', 'userId']) || '';
  const eventType = clean_(event.type);

  if (userId) {
    upsertUserFromLineEvent_(event);
  }

  if (eventType === 'follow') {
    if (replyToken) {
      replyText_(replyToken, '歡迎使用 ' + APP_NAME);
    }
    return;
  }

  if (eventType === 'unfollow') {
    setUserOnline_(userId, false);
    return;
  }

  if (eventType === 'message' && getNested_(event, ['message', 'type']) === 'text') {
    const userMsg = clean_(getNested_(event, ['message', 'text']));
    const msg = lower_(userMsg);

    if (msg === 'ping') {
      replyText_(replyToken, 'pong');
      return;
    }

    if (msg === 'id') {
      replyText_(replyToken, userId ? ('你的 LINE UID：' + userId) : '讀取不到 userId');
      return;
    }

    if (msg === 'help' || msg === '幫助') {
      replyText_(replyToken, [
        '可用指令：',
        'ping',
        'id',
        'help',
        '我的最愛',
        '我的訂單'
      ].join('\n'));
      return;
    }

    if (msg === '我的最愛') {
      replyText_(replyToken, formatFavoritesText_(userId));
      return;
    }

    if (msg === '我的訂單') {
      replyText_(replyToken, formatMyOrdersText_(userId));
      return;
    }

    replyText_(replyToken, '收到：' + userMsg);
    return;
  }

  if (eventType === 'postback') {
    if (replyToken) {
      replyText_(replyToken, '已收到操作');
    }
  }
}

function upsertUserFromLineEvent_(event) {
  const userId = getNested_(event, ['source', 'userId']);
  if (!userId) return;

  const users = getObjects_('Users');
  const now = new Date();
  let found = null;

  for (var i = 0; i < users.length; i++) {
    if (clean_(users[i].line_uid) === clean_(userId)) {
      found = users[i];
      break;
    }
  }

  if (found) {
    updateRowByRowNumber_('Users', found._row, {
      updated_at: now,
      last_update: now,
      online: '1'
    });
  } else {
    appendObject_('Users', {
      line_uid: userId,
      role: 'user',
      display_name: '',
      online: '1',
      last_lat: '',
      last_lon: '',
      last_update: now,
      created_at: now,
      updated_at: now
    }, [
      'line_uid',
      'role',
      'display_name',
      'online',
      'last_lat',
      'last_lon',
      'last_update',
      'created_at',
      'updated_at'
    ]);
  }
}

function setUserOnline_(userId, online) {
  if (!userId) return;
  const row = findRowByField_('Users', 'line_uid', userId);
  if (!row) return;

  updateRowByRowNumber_('Users', row._row, {
    online: online ? '1' : '0',
    updated_at: new Date(),
    last_update: new Date()
  });
}

function replyText_(replyToken, text) {
  const token = getSetting_('line_channel_access_token');

  if (!token) {
    throw new Error('尚未設定 line_channel_access_token');
  }

  if (!replyToken) {
    throw new Error('replyToken 為空');
  }

  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{
        type: 'text',
        text: clean_(text) || ' '
      }]
    }),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('LINE reply 失敗：' + code + ' / ' + res.getContentText());
  }
}

function pushText_(userId, text) {
  const token = getSetting_('line_channel_access_token');
  if (!token) throw new Error('尚未設定 line_channel_access_token');
  if (!userId) throw new Error('userId 為空');

  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token
    },
    payload: JSON.stringify({
      to: userId,
      messages: [{
        type: 'text',
        text: clean_(text) || ' '
      }]
    }),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('LINE push 失敗：' + code + ' / ' + res.getContentText());
  }
}

// =========================
// Orders
// =========================
function createOrder_(body) {
  const customerId = clean_(body.customer_id || body.customerId);
  const customerName = clean_(body.customer_name || body.customerName);
  const pickupLat = toNumber_(body.pickup_lat || body.pickupLat, 0);
  const pickupLon = toNumber_(body.pickup_lon || body.pickupLon, 0);
  const pickupLabel = clean_(body.pickup_label || body.pickupLabel);
  const dropLat = toNumber_(body.drop_lat || body.dropLat, 0);
  const dropLon = toNumber_(body.drop_lon || body.dropLon, 0);
  const dropLabel = clean_(body.drop_label || body.dropLabel);

  if (!customerId) throw new Error('缺少 customer_id');
  if (!pickupLabel) throw new Error('缺少 pickup_label');
  if (!dropLabel) throw new Error('缺少 drop_label');

  const distanceKm = estimateDistanceKm_(pickupLat, pickupLon, dropLat, dropLon);
  const fare = estimateFare_(distanceKm);
  const etaMin = estimateEtaMin_(distanceKm);
  const now = new Date();
  const expireMinutes = getSettingNumber_('order_search_expire_minutes', 5);
  const expiresAt = new Date(now.getTime() + expireMinutes * 60 * 1000);
  const orderId = 'ORD-' + Utilities.formatDate(now, TIMEZONE, 'yyyyMMddHHmmss') + '-' + uuid_().slice(0, 8);

  appendObject_('Orders', {
    order_id: orderId,
    customer_id: customerId,
    customer_name: customerName,
    driver_id: '',
    driver_name: '',
    status: 'searching',
    pickup_lat: pickupLat,
    pickup_lon: pickupLon,
    pickup_label: pickupLabel,
    drop_lat: dropLat,
    drop_lon: dropLon,
    drop_label: dropLabel,
    quoted_distance_km: distanceKm,
    quoted_eta_min: etaMin,
    quoted_fare: fare,
    search_expires_at: expiresAt,
    created_at: now,
    matched_at: '',
    started_at: '',
    completed_at: '',
    cancelled_at: '',
    customer_rating: '',
    driver_rating: ''
  }, [
    'order_id',
    'customer_id',
    'customer_name',
    'driver_id',
    'driver_name',
    'status',
    'pickup_lat',
    'pickup_lon',
    'pickup_label',
    'drop_lat',
    'drop_lon',
    'drop_label',
    'quoted_distance_km',
    'quoted_eta_min',
    'quoted_fare',
    'search_expires_at',
    'created_at',
    'matched_at',
    'started_at',
    'completed_at',
    'cancelled_at',
    'customer_rating',
    'driver_rating'
  ]);

  return {
    ok: true,
    message: '訂單建立成功',
    order_id: orderId,
    quoted_distance_km: distanceKm,
    quoted_eta_min: etaMin,
    quoted_fare: fare,
    search_expires_at: formatDateTime_(expiresAt)
  };
}

function cancelOrder_(body) {
  const orderId = clean_(body.order_id || body.orderId);
  if (!orderId) throw new Error('缺少 order_id');

  const row = findRowByField_('Orders', 'order_id', orderId);
  if (!row) throw new Error('找不到訂單：' + orderId);

  if (clean_(row.status) === 'completed') {
    throw new Error('已完成訂單不可取消');
  }

  updateRowByRowNumber_('Orders', row._row, {
    status: 'cancelled',
    cancelled_at: new Date()
  });

  return {
    ok: true,
    message: '訂單已取消',
    order_id: orderId
  };
}

function listOrders_(body) {
  const customerId = clean_(body.customer_id || body.customerId);
  const status = lower_(body.status);
  let rows = getObjects_('Orders');

  if (customerId) {
    rows = rows.filter(function(r) {
      return clean_(r.customer_id) === customerId;
    });
  }

  if (status) {
    rows = rows.filter(function(r) {
      return lower_(r.status) === status;
    });
  }

  rows = rows
    .sort(function(a, b) {
      return new Date(b.created_at) - new Date(a.created_at);
    })
    .slice(0, 50)
    .map(formatOrderObject_);

  return {
    ok: true,
    count: rows.length,
    orders: rows
  };
}

function formatOrderObject_(r) {
  return {
    order_id: clean_(r.order_id),
    customer_id: clean_(r.customer_id),
    customer_name: clean_(r.customer_name),
    driver_id: clean_(r.driver_id),
    driver_name: clean_(r.driver_name),
    status: clean_(r.status),
    pickup_label: clean_(r.pickup_label),
    drop_label: clean_(r.drop_label),
    quoted_distance_km: toNumber_(r.quoted_distance_km, 0),
    quoted_eta_min: toNumber_(r.quoted_eta_min, 0),
    quoted_fare: toNumber_(r.quoted_fare, 0),
    created_at: formatDateTime_(r.created_at ? new Date(r.created_at) : ''),
    search_expires_at: formatDateTime_(r.search_expires_at ? new Date(r.search_expires_at) : '')
  };
}

function formatMyOrdersText_(customerId) {
  const result = listOrders_({ customer_id: customerId });
  if (!result.orders.length) return '目前沒有訂單';

  return result.orders.slice(0, 5).map(function(o, i) {
    return [
      (i + 1) + '. ' + o.order_id,
      '狀態：' + o.status,
      '上車：' + o.pickup_label,
      '下車：' + o.drop_label,
      '車資：' + o.quoted_fare
    ].join('\n');
  }).join('\n\n');
}

// =========================
// Favorites
// =========================
function saveFavorite_(body) {
  const userId = clean_(body.user_id || body.userId);
  const label = clean_(body.label);
  const address = clean_(body.address);
  const lat = toNumber_(body.lat, 0);
  const lon = toNumber_(body.lon, 0);

  if (!userId) throw new Error('缺少 user_id');
  if (!label) throw new Error('缺少 label');
  if (!address) throw new Error('缺少 address');

  appendObject_('Favorites', {
    favorite_id: 'FAV-' + uuid_().slice(0, 8),
    user_id: userId,
    label: label,
    address: address,
    lat: lat,
    lon: lon,
    created_at: new Date()
  }, [
    'favorite_id',
    'user_id',
    'label',
    'address',
    'lat',
    'lon',
    'created_at'
  ]);

  return {
    ok: true,
    message: '已儲存常用地點'
  };
}

function deleteFavorite_(body) {
  const favoriteId = clean_(body.favorite_id || body.favoriteId);
  if (!favoriteId) throw new Error('缺少 favorite_id');

  const sh = getSheet_('Favorites');
  const rows = getObjects_('Favorites');
  const row = rows.find(function(r) {
    return clean_(r.favorite_id) === favoriteId;
  });

  if (!row) throw new Error('找不到 favorite_id：' + favoriteId);

  sh.deleteRow(row._row);
  clearSheetCache_();

  return {
    ok: true,
    message: '已刪除常用地點'
  };
}

function listFavorites_(body) {
  const userId = clean_(body.user_id || body.userId);
  if (!userId) throw new Error('缺少 user_id');

  const rows = getObjects_('Favorites')
    .filter(function(r) {
      return clean_(r.user_id) === userId;
    })
    .sort(function(a, b) {
      return new Date(b.created_at) - new Date(a.created_at);
    })
    .map(function(r) {
      return {
        favorite_id: clean_(r.favorite_id),
        label: clean_(r.label),
        address: clean_(r.address),
        lat: toNumber_(r.lat, 0),
        lon: toNumber_(r.lon, 0),
        created_at: formatDateTime_(r.created_at ? new Date(r.created_at) : '')
      };
    });

  return {
    ok: true,
    count: rows.length,
    favorites: rows
  };
}

function formatFavoritesText_(userId) {
  const result = listFavorites_({ user_id: userId });
  if (!result.favorites.length) return '目前沒有常用地點';

  return result.favorites.slice(0, 10).map(function(item, i) {
    return (i + 1) + '. ' + item.label + '｜' + item.address;
  }).join('\n');
}

// =========================
// Blacklist
// =========================
function addBlacklist_(body) {
  const targetUserId = clean_(body.target_user_id || body.targetUserId);
  const reason = clean_(body.reason);
  const createdBy = clean_(body.created_by || body.createdBy);

  if (!targetUserId) throw new Error('缺少 target_user_id');

  appendObject_('Blacklist', {
    blacklist_id: 'BL-' + uuid_().slice(0, 8),
    target_user_id: targetUserId,
    reason: reason,
    created_by: createdBy,
    created_at: new Date()
  }, [
    'blacklist_id',
    'target_user_id',
    'reason',
    'created_by',
    'created_at'
  ]);

  return {
    ok: true,
    message: '已加入黑名單'
  };
}

function listBlacklist_(body) {
  const rows = getObjects_('Blacklist')
    .sort(function(a, b) {
      return new Date(b.created_at) - new Date(a.created_at);
    })
    .slice(0, body && body.limit ? Number(body.limit) : 100)
    .map(function(r) {
      return {
        blacklist_id: clean_(r.blacklist_id),
        target_user_id: clean_(r.target_user_id),
        reason: clean_(r.reason),
        created_by: clean_(r.created_by),
        created_at: formatDateTime_(r.created_at ? new Date(r.created_at) : '')
      };
    });

  return {
    ok: true,
    count: rows.length,
    rows: rows
  };
}

// =========================
// Fare / ETA / Geo
// =========================
function estimateDistanceKm_(lat1, lon1, lat2, lon2) {
  if (!lat1 || !lon1 || !lat2 || !lon2) return 0;

  const R = 6371;
  const dLat = deg2rad_(lat2 - lat1);
  const dLon = deg2rad_(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(deg2rad_(lat1)) * Math.cos(deg2rad_(lat2)) *
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return Math.round(R * c * 100) / 100;
}

function deg2rad_(deg) {
  return deg * (Math.PI / 180);
}

function estimateFare_(distanceKm) {
  const base = getSettingNumber_('fare_base', 80);
  const perKm = getSettingNumber_('fare_per_km', 20);
  const fare = base + Math.ceil(Math.max(distanceKm - 1, 0)) * perKm;
  return Math.max(base, fare);
}

function estimateEtaMin_(distanceKm) {
  const rate = getSettingNumber_('eta_minutes_per_km', 3);
  return Math.max(3, Math.ceil(distanceKm * rate));
}

// =========================
// 測試工具
// =========================
function testInitSystem() {
  Logger.log(JSON.stringify(initSystem()));
}

function testHealth() {
  const res = doGet({ parameter: { action: 'health' } });
  Logger.log(res.getContent());
}

function testCreateOrder() {
  const res = createOrder_({
    customer_id: 'U123456',
    customer_name: '測試乘客',
    pickup_lat: 24.1477,
    pickup_lon: 120.6736,
    pickup_label: '台中火車站',
    drop_lat: 24.1374,
    drop_lon: 120.6869,
    drop_label: '審計新村'
  });
  Logger.log(JSON.stringify(res));
}

function testListOrders() {
  Logger.log(JSON.stringify(listOrders_({ customer_id: 'U123456' })));
}

function testSaveFavorite() {
  Logger.log(JSON.stringify(saveFavorite_({
    user_id: 'U123456',
    label: '公司',
    address: '台中市西區範例路100號',
    lat: 24.145,
    lon: 120.67
  })));
}

function testListFavorites() {
  Logger.log(JSON.stringify(listFavorites_({ user_id: 'U123456' })));
}

function testPushText() {
  pushText_('請替換成真正的 userId', '測試推播');
}
