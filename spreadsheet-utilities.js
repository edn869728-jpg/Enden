/**
 * spreadsheet-utilities.js
 * Google Apps Script backend for the HR system.
 *
 * Column layout for 申請資料 sheet (RECEIVE_COLS):
 *   A (0): reqId
 *   B (1): type
 *   C (2): empId
 *   D (3): name
 *   E (4): title
 *   F (5): note        ← supplementary notes ONLY
 *   G (6): amount      ← 代墊金額 (expense amount)
 *   H (7): item        ← 代墊商品 (expense item)
 *   I (8): location    ← 位置資訊 (location info)
 *   J (9): attachment
 *
 * Column layout for 排班預選 sheet (PRESELECT_COLS):
 *   A (0): empId
 *   B (1): date        ← "YYYY/MM/DD"
 *   C (2): isOff       ← TRUE = vacation/day-off, FALSE = working
 */

// ── Constants ──────────────────────────────────────────────────────────────────

var RECEIVE_COLS = {
  REQ_ID:     0,
  TYPE:       1,
  EMP_ID:     2,
  NAME:       3,
  TITLE:      4,
  NOTE:       5,
  AMOUNT:     6,
  ITEM:       7,
  LOCATION:   8,
  ATTACHMENT: 9
};

var PRESELECT_COLS = {
  EMP_ID: 0,
  DATE:   1,
  IS_OFF: 2   // TRUE = vacation/off, FALSE = working
};

var SALARY_SHEET      = '薪資紀錄';
var SALARY_DETAIL_SHT = '薪資明細';
var PRESELECT_SHEET   = '排班預選';
var UPLOAD_SHEET      = '申請資料';
var CLOCK_FIX_SHEET   = '補打卡';
var LEAVE_SHEET       = '請假申請';

// ── Web App Entry Point ────────────────────────────────────────────────────────

function doGet(e) {
  return HtmlService.createTemplateFromFile('!DOCTYPE')
    .evaluate()
    .setTitle('HR 系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Client-Callable API Dispatch ───────────────────────────────────────────────

/**
 * Unified entry point for all client-side google.script.run calls.
 * @param {string} action
 * @param {Object} payload
 * @returns {Object} { success, data?, message? }
 */
function processRequest(action, payload) {
  try {
    switch (action) {
      case 'submitPreselect':        return submitPreselect(payload);
      case 'submitEmployeeUpload':   return submitEmployeeUpload(payload);
      case 'submitEmployeeClockFix': return submitEmployeeClockFix(payload);
      case 'getSalaryHistory':       return getSalaryHistory(payload);
      case 'getSalaryDetail':        return getSalaryDetail(payload);
      case 'getReviewCards':         return getReviewCards(payload);
      case 'getEmployeeInfo':        return getEmployeeInfo(payload);
      case 'getHomeStats':           return getHomeStats(payload);
      case 'getNotifications':       return getNotifications(payload);
      default:
        return { success: false, message: '未知的 action：' + action };
    }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ── Schedule / Preselect ───────────────────────────────────────────────────────

/**
 * Employee preselects their desired vacation days.
 * Each date the employee selected is treated as a REST/VACATION day.
 *
 * @param {Object} payload
 * @param {string}   payload.empId
 * @param {string[]} payload.vacationDates  Dates the employee wants OFF (YYYY/MM/DD)
 * @param {string}   payload.periodStart    Start of selectable period
 * @param {string}   payload.periodEnd      End of selectable period
 */
function submitPreselect(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PRESELECT_SHEET) || ss.insertSheet(PRESELECT_SHEET);

  writePreselectSchedule_(sheet, payload.empId, payload.vacationDates || [],
                          payload.periodStart, payload.periodEnd);
  return { success: true, message: '休假申請已送出' };
}

/**
 * Writes preselect schedule to the sheet.
 *
 * Convention:  IS_OFF = TRUE  → vacation / day off
 *              IS_OFF = FALSE → working day
 *
 * BUG FIX: Previously selected dates were stored as TRUE meaning "working",
 * which is the opposite of the employee's intent.  Now:
 *   - Selected dates (vacationDates[])  → isOff = TRUE  (employee wants rest)
 *   - All other dates in the period     → isOff = FALSE (working as normal)
 *
 * @param {Sheet}    sheet
 * @param {string}   empId
 * @param {string[]} vacationDates  Dates requested as vacation
 * @param {string}   periodStart
 * @param {string}   periodEnd
 */
function writePreselectSchedule_(sheet, empId, vacationDates, periodStart, periodEnd) {
  // Build a Set of normalized vacation date strings for O(1) lookup.
  var vacSet = {};
  for (var i = 0; i < vacationDates.length; i++) {
    vacSet[normalizeDate_(vacationDates[i])] = true;
  }

  var start = new Date(periodStart);
  var end   = new Date(periodEnd);
  var rows  = [];

  for (var d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    var dateStr = formatDateStr_(d);
    // TRUE = vacation/off (employee selected this date as a rest day)
    // FALSE = working day (not selected)
    var isOff = vacSet[dateStr] === true;
    rows.push([empId, dateStr, isOff]);
  }

  // Remove existing entries for this employee within the period.
  var data = sheet.getDataRange().getValues();
  var normStart = normalizeDate_(periodStart);
  var normEnd   = normalizeDate_(periodEnd);

  // Iterate backwards to avoid index shifting on deletion.
  for (var r = data.length - 1; r >= 1; r--) {
    var row = data[r];
    if (String(row[PRESELECT_COLS.EMP_ID]) === String(empId)) {
      var rowDate = normalizeDate_(row[PRESELECT_COLS.DATE]);
      if (rowDate >= normStart && rowDate <= normEnd) {
        sheet.deleteRow(r + 1); // sheet rows are 1-indexed
      }
    }
  }

  if (rows.length > 0) {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

// ── Employee Upload / Request ──────────────────────────────────────────────────

/**
 * Submits an employee data/expense request.
 *
 * BUG FIX: Previously amount, item, and location were all concatenated
 * into the `note` field.  They are now stored in their own dedicated columns
 * (G, H, I respectively).  The `note` field (F) now holds supplementary
 * text only.
 *
 * @param {Object} payload
 * @param {string}  payload.type
 * @param {string}  payload.empId
 * @param {string}  payload.name
 * @param {string}  payload.title
 * @param {string}  [payload.note]       Supplementary notes only
 * @param {number}  [payload.amount]     代墊金額
 * @param {string}  [payload.item]       代墊商品
 * @param {string}  [payload.location]   位置資訊
 * @param {string}  [payload.attachment] Attachment URL
 */
function submitEmployeeUpload(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(UPLOAD_SHEET) || ss.insertSheet(UPLOAD_SHEET);

  var reqId = generateReqId_();
  var row = [
    reqId,
    payload.type       || '',
    payload.empId      || '',
    payload.name       || '',
    payload.title      || '',
    payload.note       || '',   // F: supplementary note only
    payload.amount     || '',   // G: 代墊金額
    payload.item       || '',   // H: 代墊商品
    payload.location   || '',   // I: 位置資訊
    payload.attachment || ''    // J: attachment
  ];

  sheet.appendRow(row);
  return { success: true, reqId: reqId, message: '申請已送出' };
}

// ── Clock-In Fix ───────────────────────────────────────────────────────────────

/**
 * Submits a clock-in/out correction request.
 *
 * BUG FIX: Previously only `payload.time` (e.g. "14:00") was stored.
 * Google Sheets interprets a bare time string as a fraction of a day relative
 * to the serial epoch (1899/12/30), displaying "1899/12/30 14:00:00".
 * Fix: store the full datetime string "YYYY/MM/DD HH:MM" in the time column.
 *
 * @param {Object} payload
 * @param {string} payload.empId
 * @param {string} payload.date    e.g. "2026/04/17"
 * @param {string} payload.time    e.g. "14:00"
 * @param {string} payload.type    "in" or "out"
 * @param {string} payload.reason
 */
function submitEmployeeClockFix(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CLOCK_FIX_SHEET) || ss.insertSheet(CLOCK_FIX_SHEET);

  // Combine date + time to prevent Google Sheets from interpreting a bare
  // time value against the 1899 serial epoch.
  var fullDatetime = (payload.date || '') + ' ' + (payload.time || '');

  sheet.appendRow([
    payload.empId  || '',
    payload.date   || '',
    fullDatetime,            // full "YYYY/MM/DD HH:MM" – no 1899 bug
    payload.type   || '',
    payload.reason || '',
    new Date()               // submission timestamp
  ]);

  return { success: true, message: '補打卡申請已送出' };
}

// ── Salary ─────────────────────────────────────────────────────────────────────

/**
 * Returns issued-salary records for an employee.
 * Supports both monthly ("2026-04") and weekly ("2026-W15") pay periods.
 *
 * Sheet columns: empId | periodType | periodKey | amount | hours | issued
 *
 * @param {Object} payload
 * @param {string} payload.empId
 * @returns {{ success: boolean, data: Array }}
 */
function getSalaryHistory(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SALARY_SHEET);
  if (!sheet) return { success: true, data: [] };

  var data    = sheet.getDataRange().getValues();
  var results = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[0]) !== String(payload.empId)) continue;

    var periodType = String(row[1] || 'monthly');  // 'monthly' | 'weekly'
    var periodKey  = String(row[2] || '');
    var label      = formatPeriodLabel_(periodType, periodKey);

    results.push({
      key:    periodKey,
      type:   periodType,
      label:  label,
      amount: row[3] || 0,
      hours:  row[4] || 0
    });
  }

  // Newest first
  results.reverse();
  return { success: true, data: results };
}

/**
 * Returns payslip detail for a given period.
 *
 * Sheet columns: empId | periodKey | periodType | basePay | overtime |
 *                bonus | deduction | total | hours
 *
 * @param {Object} payload
 * @param {string} payload.empId
 * @param {string} payload.periodKey
 */
function getSalaryDetail(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SALARY_DETAIL_SHT);
  if (!sheet) return { success: true, data: null };

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[0]) === String(payload.empId) &&
        String(row[1]) === String(payload.periodKey)) {
      return {
        success: true,
        data: {
          empId:      clean_(row[0]),
          periodKey:  clean_(row[1]),
          periodType: clean_(row[2]),
          basePay:    row[3] || 0,
          overtime:   row[4] || 0,
          bonus:      row[5] || 0,
          deduction:  row[6] || 0,
          total:      row[7] || 0,
          hours:      row[8] || 0
        }
      };
    }
  }

  return { success: true, data: null };
}

// ── Review Cards ───────────────────────────────────────────────────────────────

/**
 * Returns pending review cards for manager/admin.
 *
 * BUG FIX: Previously all fields were passed through formatDateTimeMaybe_()
 * which returned an empty string for non-Date values.  Fields like 假別 (leave
 * type) and 事由 (reason) are plain strings and must use clean_() instead.
 *
 * @param {Object} payload
 * @param {string} [payload.status]  Filter by status (default "待審核")
 */
function getReviewCards(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LEAVE_SHEET);
  if (!sheet) return { success: true, data: [] };

  var targetStatus = (payload && payload.status) || '待審核';
  var data         = sheet.getDataRange().getValues();
  var cards        = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // status is expected in column H (index 7) – adjust to your actual layout
    if (clean_(row[7]) !== targetStatus) continue;

    cards.push({
      reqId:     clean_(row[0]),
      leaveType: clean_(row[1]),          // 假別 – plain string, NOT a date
      empId:     clean_(row[2]),
      name:      clean_(row[3]),
      startDate: formatDateMaybe_(row[4]),
      endDate:   formatDateMaybe_(row[5]),
      reason:    clean_(row[6]),          // 事由 – plain string, NOT a date
      status:    clean_(row[7])
    });
  }

  return { success: true, data: cards };
}

// ── Employee Info & Home Stats ─────────────────────────────────────────────────

/**
 * Returns basic employee profile info.
 * @param {Object} payload
 * @param {string} payload.empId
 */
function getEmployeeInfo(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('員工資料');
  if (!sheet) return { success: true, data: null };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[0]) === String(payload.empId)) {
      return {
        success: true,
        data: {
          empId:    clean_(row[0]),
          name:     clean_(row[1]),
          dept:     clean_(row[2]),
          role:     clean_(row[3]),
          shift:    clean_(row[4]),
          joinDate: formatDateMaybe_(row[5])
        }
      };
    }
  }
  return { success: true, data: null };
}

/**
 * Returns home-page statistics for an employee (weekly/monthly hours & pay).
 * @param {Object} payload
 * @param {string} payload.empId
 */
function getHomeStats(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var salSheet = ss.getSheetByName(SALARY_SHEET);
  if (!salSheet) return { success: true, data: {} };

  var today      = new Date();
  var yearMonth  = today.getFullYear() + '-' +
                   String(today.getMonth() + 1).padStart(2, '0');
  var weekKey    = getISOWeekKey_(today);

  var monthHours = 0, monthPay = 0, weekHours = 0, weekPay = 0;
  var data = salSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[0]) !== String(payload.empId)) continue;
    var pType = String(row[1] || 'monthly');
    var pKey  = String(row[2] || '');

    if (pType === 'monthly' && pKey === yearMonth) {
      monthPay   += Number(row[3]) || 0;
      monthHours += Number(row[4]) || 0;
    }
    if (pType === 'weekly' && pKey === weekKey) {
      weekPay   += Number(row[3]) || 0;
      weekHours += Number(row[4]) || 0;
    }
  }

  return {
    success: true,
    data: {
      monthHours: monthHours,
      monthPay:   monthPay,
      weekHours:  weekHours,
      weekPay:    weekPay
    }
  };
}

/**
 * Returns notifications/reminders for an employee.
 * @param {Object} payload
 * @param {string} payload.empId
 */
function getNotifications(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('通知');
  if (!sheet) return { success: true, data: [] };

  var data   = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var target = clean_(row[3]);
    // Show if addressed to this employee, all employees ('all'), or is public
    if (target === 'all' || target === '' ||
        String(target) === String(payload.empId)) {
      result.push({
        id:      clean_(row[0]),
        type:    clean_(row[1]),   // 'reminder' | 'notice' | 'message'
        title:   clean_(row[2]),
        target:  target,
        content: clean_(row[4]),
        date:    formatDateMaybe_(row[5]),
        from:    clean_(row[6])
      });
    }
  }

  result.reverse(); // newest first
  return { success: true, data: result };
}

// ── Utility Helpers ────────────────────────────────────────────────────────────

function generateReqId_() {
  return 'REQ-' + new Date().getTime();
}

function normalizeDate_(d) {
  if (!d && d !== 0) return '';
  var dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt.getTime())) return String(d);
  return formatDateStr_(dt);
}

function formatDateStr_(d) {
  var y   = d.getFullYear();
  var m   = String(d.getMonth() + 1).padStart(2, '0');
  var day = String(d.getDate()).padStart(2, '0');
  return y + '/' + m + '/' + day;
}

/**
 * Formats a pay-period key into a human-readable label.
 * monthly: "2026-04"  → "2026年04月"
 * weekly:  "2026-W15" → "2026 第15週"
 *
 * @param {string} type  'monthly' | 'weekly'
 * @param {string} key
 * @returns {string}
 */
function formatPeriodLabel_(type, key) {
  if (type === 'weekly') {
    var parts = key.split('-W');
    return (parts[0] || '') + ' 第' + (parts[1] || '') + '週';
  }
  // monthly (default)
  var parts = key.split('-');
  return (parts[0] || '') + '年' + (parts[1] || '') + '月';
}

/**
 * Returns a raw string value; never interprets the value as a date.
 * Use for fields that contain text (leave type, reason, etc.).
 */
function clean_(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

/**
 * Formats a cell value as a date string ONLY if it is actually a Date object.
 * Falls back to clean_() for everything else (prevents "" for string fields).
 */
function formatDateMaybe_(val) {
  if (val instanceof Date && !isNaN(val.getTime())) {
    return formatDateStr_(val);
  }
  return clean_(val);
}

/**
 * Returns an ISO week key for a given date, e.g. "2026-W15".
 */
function getISOWeekKey_(date) {
  var d  = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  var day = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - day);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  var week = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return d.getUTCFullYear() + '-W' + String(week).padStart(2, '0');
}
