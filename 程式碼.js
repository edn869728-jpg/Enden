// ============================================================
// 程式碼.js  —  Google Apps Script back-end
// Employee management system utilities
// ============================================================

// ── Constants ────────────────────────────────────────────────
var SHEET_CLOCK     = '打卡紀錄';
var SHEET_CLOCK_FIX = '補打卡申請';
var SHEET_LEAVE     = '請假申請';
var SHEET_SALARY    = '薪資明細';
var SHEET_EMPLOYEE  = '員工資料';
var SHEET_PRESELECT = '選休排班';
var SHEET_UPLOAD    = '上傳記錄';
var SHEET_SCHEDULE  = '班表';

var RECEIVE_COLS = ['申請編號', '類型', '員編', '姓名', '職稱', '補充說明', '代墊金額', '代墊商品', '位置資訊', '附件', '申請時間', '審核狀態', '審核人員', '審核時間', '審核意見', '審核人清單', '目前審核人', '已審核人員'];

// ── Generic helpers ──────────────────────────────────────────

/**
 * Zero-pad a number to at least 2 digits.
 */
function pad_(n) { return n < 10 ? '0' + n : String(n); }

/**
 * Return the active spreadsheet's sheet by name, or null.
 */
function getSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

/**
 * Trim and stringify a cell value; return '' for null/undefined.
 */
function clean_(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

/**
 * If val looks like a Date or ISO date string, format it as
 * 'YYYY/MM/DD HH:mm'; otherwise return ''.
 */
function formatDateTimeMaybe_(val) {
  if (!val) return '';
  var d;
  if (val instanceof Date) {
    d = val;
  } else {
    var s = String(val).trim();
    if (!/^\d{4}[-\/]\d{2}[-\/]\d{2}/.test(s)) return '';
    d = new Date(s.replace(/\//g, '-'));
  }
  if (isNaN(d.getTime())) return '';
  return d.getFullYear() + '/' +
         pad_(d.getMonth() + 1) + '/' +
         pad_(d.getDate()) + ' ' +
         pad_(d.getHours()) + ':' +
         pad_(d.getMinutes());
}

/**
 * Return today's date as 'YYYY/MM/DD'.
 */
function todayStr_() {
  var now = new Date();
  return now.getFullYear() + '/' + pad_(now.getMonth() + 1) + '/' + pad_(now.getDate());
}

/**
 * Generate a unique application ID: prefix + timestamp.
 */
function generateId_(prefix) {
  return prefix + new Date().getTime();
}

// ── Employee lookup ───────────────────────────────────────────

/**
 * Find the row for an employee by their ID.
 * Returns a header-keyed object or null.
 */
function getEmployee_(empId) {
  var sheet = getSheet_(SHEET_EMPLOYEE);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  var headers = data[0].map(function(h) { return clean_(h); });
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var idIdx = headers.indexOf('員編');
    if (idIdx < 0) idIdx = 0;
    if (clean_(row[idIdx]) === clean_(empId)) {
      var obj = {};
      headers.forEach(function(h, j) { obj[h] = clean_(row[j]); });
      return obj;
    }
  }
  return null;
}

// ── Clock fix (補打卡) ────────────────────────────────────────

/**
 * Submit a clock-fix request.
 *
 * Fix 5: Store date (col 4) and time (col 5) in SEPARATE columns so that
 * Google Sheets never mis-parses a bare time string as 1899/12/30.
 *
 * @param {string} empId    Employee ID
 * @param {string} timeVal  ISO datetime string, e.g. '2026-04-17T14:00'
 * @param {string} type     '上班' or '下班'
 * @param {string} reason   Reason text
 * @returns {Object} {success, message}
 */
function submitEmployeeClockFix(empId, timeVal, type, reason) {
  try {
    var sheet = getSheet_(SHEET_CLOCK_FIX);
    if (!sheet) return { success: false, message: '找不到補打卡申請工作表' };

    var emp = getEmployee_(empId);
    var empName = emp ? emp['姓名'] || emp['名字'] || '' : '';

    // Parse timeVal into separate date and time parts
    var dt = String(timeVal).trim();
    var datePart, timePart;
    if (dt.indexOf('T') !== -1) {
      // ISO 8601: '2026-04-17T14:00' or '2026-04-17T14:00:00'
      datePart = dt.slice(0, 10).replace(/-/g, '/');  // '2026/04/17'
      timePart = dt.slice(11, 16);                    // '14:00'
    } else if (/^\d{4}[-\/]\d{2}[-\/]\d{2}\s/.test(dt)) {
      // Already 'YYYY-MM-DD HH:mm' or 'YYYY/MM/DD HH:mm'
      var normalized = dt.replace(/-/g, '/');
      datePart = normalized.slice(0, 10);
      timePart = normalized.slice(11, 16);
    } else {
      datePart = dt;
      timePart = '';
    }

    var appId = generateId_('CF');

    // Columns: [申請編號, 員編, 姓名, 日期, 時間, 類型, 事由, 申請日期, 狀態]
    sheet.appendRow([
      appId,
      empId,
      empName,
      datePart,   // e.g. '2026/04/17' - date column
      timePart,   // e.g. '14:00'      - time column (separate)
      type,
      reason,
      todayStr_(),
      '待審核'
    ]);

    return { success: true, message: '補打卡申請已送出，申請編號：' + appId };
  } catch (e) {
    return { success: false, message: '錯誤：' + e.message };
  }
}

// ── Leave request (請假) ──────────────────────────────────────

/**
 * Submit a leave request.
 */
function submitLeaveRequest(empId, leaveType, startDate, days, reason) {
  try {
    var sheet = getSheet_(SHEET_LEAVE);
    if (!sheet) return { success: false, message: '找不到請假申請工作表' };

    var emp = getEmployee_(empId);
    var empName = emp ? emp['姓名'] || emp['名字'] || '' : '';
    var appId = generateId_('LV');

    sheet.appendRow([
      appId,
      empId,
      empName,
      leaveType,
      startDate,
      days,
      reason,
      todayStr_(),
      '待審核'
    ]);

    return { success: true, message: '請假申請已送出，申請編號：' + appId };
  } catch (e) {
    return { success: false, message: '錯誤：' + e.message };
  }
}

// ── Upload / reimbursement request (上傳申請) ─────────────────

/**
 * Submit an employee upload/reimbursement request.
 * Columns follow RECEIVE_COLS order:
 * [reqId, type, empId, name, title, note, amount, item, location, attachment, now, '待審核', '', '', '', approvers, approvers[0], '[]']
 *
 * @param {string} empId      Employee ID
 * @param {string} type       Request type
 * @param {string} note       Supplementary note
 * @param {string} amount     Reimbursement amount
 * @param {string} item       Reimbursement item
 * @param {string} location   Location info
 * @param {string} fileName   Attachment file name
 * @param {string} fileMime   Attachment MIME type
 * @param {string} fileData   Attachment base64 data
 * @returns {Object} {success, message}
 */
function submitEmployeeUpload(empId, type, note, amount, item, location, fileName, fileMime, fileData) {
  try {
    var sheet = getSheet_(SHEET_UPLOAD);
    if (!sheet) return { success: false, message: '找不到上傳記錄工作表' };

    var emp = getEmployee_(empId);
    var empName  = emp ? (emp['姓名'] || emp['名字'] || '') : '';
    var empTitle = emp ? (emp['職稱'] || emp['職位'] || '') : '';

    // Build approver list (use manager from employee record, or default)
    var approvers = emp && emp['主管員編'] ? [emp['主管員編']] : ['MANAGER'];

    var reqId = generateId_('UP');
    var now   = new Date();
    var nowStr = now.getFullYear() + '/' + pad_(now.getMonth()+1) + '/' + pad_(now.getDate()) + ' ' +
                 pad_(now.getHours()) + ':' + pad_(now.getMinutes());

    // Handle file attachment
    var attachmentStr = '';
    if (fileData && fileName) {
      // Store filename; optionally save to Drive and store the URL
      attachmentStr = fileName;
    }

    // appendRow 按照 RECEIVE_COLS 定義的順序寫入
    sheet.appendRow([
      reqId,               // 申請編號
      type,                // 類型
      empId,               // 員編
      empName,             // 姓名
      empTitle,            // 職稱
      note,                // 補充說明
      amount || '',        // 代墊金額
      item || '',          // 代墊商品
      location || '',      // 位置資訊
      attachmentStr,       // 附件
      nowStr,              // 申請時間
      '待審核',             // 審核狀態
      '',                  // 審核人員
      '',                  // 審核時間
      '',                  // 審核意見
      approvers.join(','), // 審核人清單
      approvers[0] || '',  // 目前審核人
      '[]'                 // 已審核人員
    ]);

    return { success: true, message: '上傳申請已送出，編號：' + reqId };
  } catch (e) {
    return { success: false, message: '錯誤：' + e.message };
  }
}

// ── Review cards ──────────────────────────────────────────────

/**
 * Build review card objects from a sheet.
 *
 * Uses clean_() for all fields; formatDateTimeMaybe_() only for known
 * date/time columns so that fields like 假別 and 事由 are never empty.
 *
 * @param {string}   sheetName      Name of the sheet to read
 * @param {string[]} includeHeaders Column headers to include in desc
 * @param {string}   titlePrefix    Prefix for the card title
 * @param {string}   statusHeader   Header name of the status column
 * @param {string}   targetStatus   Filter to only rows with this status
 * @returns {Object[]} Array of card objects
 */
function getReviewCards_(sheetName, includeHeaders, titlePrefix, statusHeader, targetStatus) {
  var sheet = getSheet_(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0].map(function(h) { return clean_(h); });
  var statusIdx = headers.indexOf(clean_(statusHeader));

  var dateTimeHeaders = ['申請日期', '開始日期', '補打卡時間', '打卡時間', '結束日期', '日期'];

  var cards = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (statusIdx >= 0 && clean_(row[statusIdx]) !== targetStatus) continue;

    var descParts = [];
    includeHeaders.forEach(function(header) {
      var idx = headers.indexOf(clean_(header));
      if (idx < 0) return;
      var rawVal = row[idx];
      var displayVal;
      if (dateTimeHeaders.indexOf(clean_(header)) !== -1) {
        displayVal = formatDateTimeMaybe_(rawVal) || clean_(rawVal);
      } else {
        displayVal = clean_(rawVal);
      }
      if (displayVal !== '') {
        descParts.push(header + '：' + displayVal);
      }
    });

    var appIdIdx = headers.indexOf('申請編號');
    var nameIdx  = headers.indexOf('姓名');
    var empIdIdx = headers.indexOf('員編');
    var typeIdx  = headers.indexOf('假別');
    var titleParts = [titlePrefix];
    if (appIdIdx >= 0 && clean_(row[appIdIdx])) titleParts.push(clean_(row[appIdIdx]));
    if (nameIdx >= 0 && clean_(row[nameIdx]))   titleParts.push(clean_(row[nameIdx]));
    if (empIdIdx >= 0 && clean_(row[empIdIdx])) titleParts.push('(' + clean_(row[empIdIdx]) + ')');
    if (typeIdx >= 0 && clean_(row[typeIdx]))   titleParts.push('—' + clean_(row[typeIdx]));

    cards.push({
      rowIndex: i + 1,
      title: titleParts.join(' '),
      desc: descParts.join('\n'),
      status: statusIdx >= 0 ? clean_(row[statusIdx]) : ''
    });
  }
  return cards;
}

/**
 * Get review cards for a specific role.
 *
 * @param {string} role    Role name (e.g. '主管', '人資')
 * @param {string} empDept Department filter ('' = all)
 * @returns {Object} { leaveCards, clockFixCards }
 */
function getReviewCardsForRole_(role, empDept) {
  var leaveInclude = ['員編', '姓名', '假別', '開始日期', '天數', '事由', '申請日期'];
  var leaveCards = getReviewCards_(
    SHEET_LEAVE,
    leaveInclude,
    '請假申請',
    '狀態',
    '待審核'
  );

  var clockFixInclude = ['員編', '姓名', '補打卡時間', '類型', '事由', '申請日期'];
  var clockFixCards = getReviewCards_(
    SHEET_CLOCK_FIX,
    clockFixInclude,
    '補打卡申請',
    '狀態',
    '待審核'
  );

  if (empDept) {
    var filterByDept = function(cards) {
      return cards.filter(function(c) {
        return c.desc.indexOf(empDept) !== -1;
      });
    };
    leaveCards    = filterByDept(leaveCards);
    clockFixCards = filterByDept(clockFixCards);
  }

  return { leaveCards: leaveCards, clockFixCards: clockFixCards };
}

/**
 * Approve or reject a leave/clock-fix request.
 *
 * @param {string} sheetName  Sheet to update
 * @param {number} rowIndex   1-based row index
 * @param {string} decision   '核准' or '拒絕'
 * @param {string} comment    Reviewer comment
 * @returns {Object} {success, message}
 */
function reviewRequest_(sheetName, rowIndex, decision, comment) {
  try {
    var sheet = getSheet_(sheetName);
    if (!sheet) return { success: false, message: '找不到工作表：' + sheetName };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(function(h) { return clean_(h); });

    var statusIdx     = headers.indexOf('狀態');
    var commentIdx    = headers.indexOf('審核意見');
    var reviewDateIdx = headers.indexOf('審核日期');

    if (statusIdx >= 0)     sheet.getRange(rowIndex, statusIdx + 1).setValue(decision);
    if (commentIdx >= 0)    sheet.getRange(rowIndex, commentIdx + 1).setValue(comment || '');
    if (reviewDateIdx >= 0) sheet.getRange(rowIndex, reviewDateIdx + 1).setValue(todayStr_());

    return { success: true, message: '審核完成' };
  } catch (e) {
    return { success: false, message: '錯誤：' + e.message };
  }
}

// ── Pre-select (選休) ─────────────────────────────────────────

/**
 * Submit pre-selected vacation days.
 *
 * @param {string}   empId        Employee ID
 * @param {string}   yearMonth    'YYYY-MM'
 * @param {string[]} vacationDays Array of date strings ('YYYY-MM-DD') to take off
 * @returns {Object} {success, message}
 */
function submitPreselect(empId, yearMonth, vacationDays) {
  try {
    var sheet = getSheet_(SHEET_PRESELECT);
    if (!sheet) return { success: false, message: '找不到選休排班工作表' };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(function(h) { return clean_(h); });

    var empIdIdx  = headers.indexOf('員編');
    var monthIdx  = headers.indexOf('月份');
    var dateIdx   = headers.indexOf('日期');
    var isWorkIdx = headers.indexOf('上班');

    // Remove existing rows for this employee+month
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (clean_(data[i][empIdIdx]) === clean_(empId) &&
          clean_(data[i][monthIdx]) === yearMonth) {
        sheet.deleteRow(i + 1);
      }
    }

    // Write one row per day in the month
    var parts = yearMonth.split('-');
    var year  = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10);
    var daysInMonth = new Date(year, month, 0).getDate();

    for (var d = 1; d <= daysInMonth; d++) {
      var dateStr = year + '-' + pad_(month) + '-' + pad_(d);
      var isWork = vacationDays.indexOf(dateStr) === -1;
      var row = [];
      for (var c = 0; c < headers.length; c++) {
        if (c === empIdIdx)   row.push(empId);
        else if (c === monthIdx)  row.push(yearMonth);
        else if (c === dateIdx)   row.push(dateStr);
        else if (c === isWorkIdx) row.push(isWork);
        else row.push('');
      }
      sheet.appendRow(row);
    }

    return { success: true, message: '選休設定已儲存' };
  } catch (e) {
    return { success: false, message: '錯誤：' + e.message };
  }
}

// ── Schedule (班表) ───────────────────────────────────────────

/**
 * Write pre-selected schedule for a week/period.
 * New schedule defaults to ALL TRUE (working days).
 * Then sets the employee's selected weekdays (vacation days) to FALSE.
 *
 * @param {string}   empId     Employee ID
 * @param {string}   weekStart Start date of the week 'YYYY-MM-DD'
 * @param {number[]} weekdays  Array of weekday indexes (0=Sun..6=Sat) to set as vacation (FALSE)
 * @returns {Object} {success, message}
 */
function writePreselectSchedule_(empId, weekStart, weekdays) {
  try {
    var sheet = getSheet_(SHEET_SCHEDULE);
    if (!sheet) return { success: false, message: '找不到班表工作表' };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(function(h) { return clean_(h); });

    var empIdIdx    = headers.indexOf('員編');
    var weekIdx     = headers.indexOf('週起始');
    var dayColStart = headers.indexOf('週日'); // first day column

    if (dayColStart < 0) {
      // Fallback: try numeric day column '0'
      dayColStart = headers.indexOf('0');
    }

    // Find existing row for this emp+week
    var data = sheet.getDataRange().getValues();
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (clean_(data[i][empIdIdx]) === clean_(empId) &&
          clean_(data[i][weekIdx]) === weekStart) {
        targetRow = i + 1; // 1-based
        break;
      }
    }

    // If not found, append a new row
    if (targetRow < 0) {
      var newRow = [];
      for (var c = 0; c < headers.length; c++) {
        if (c === empIdIdx)      newRow.push(empId);
        else if (c === weekIdx)  newRow.push(weekStart);
        else newRow.push('');
      }
      sheet.appendRow(newRow);
      targetRow = sheet.getLastRow();
    }

    // Step 1: Set all 7 day columns to TRUE (上班)
    if (dayColStart >= 0) {
      for (var d = 0; d < 7; d++) {
        sheet.getRange(targetRow, dayColStart + 1 + d).setValue(true);
      }
      // Step 2: Set employee's selected vacation days to FALSE (休假)
      weekdays.forEach(function(dayIdx) {
        if (dayIdx >= 0 && dayIdx < 7) {
          sheet.getRange(targetRow, dayColStart + 1 + dayIdx).setValue(false);
        }
      });
    }

    return { success: true, message: '班表已更新' };
  } catch (e) {
    return { success: false, message: '錯誤：' + e.message };
  }
}

// ── Salary (薪資) ─────────────────────────────────────────────

/**
 * Get salary records for an employee.
 *
 * @param {string} empId  Employee ID
 * @returns {Object[]} Array of salary record objects
 */
function getSalaryRecords(empId) {
  var sheet = getSheet_(SHEET_SALARY);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h) { return clean_(h); });
  var empIdIdx = headers.indexOf('員編');
  var records = [];
  for (var i = 1; i < data.length; i++) {
    if (empIdIdx >= 0 && clean_(data[i][empIdIdx]) !== clean_(empId)) continue;
    var obj = {};
    headers.forEach(function(h, j) { obj[h] = clean_(data[i][j]); });
    records.push(obj);
  }
  return records;
}

/**
 * Get available salary months for an employee (for the download selector).
 *
 * @param {string} empId
 * @returns {string[]} Sorted list of 'YYYY/MM' month strings
 */
function getSalaryMonths(empId) {
  var records = getSalaryRecords(empId);
  var months = records
    .map(function(r) { return r['月份'] || ''; })
    .filter(function(m) { return m !== ''; });
  var unique = months.filter(function(v, i, a) { return a.indexOf(v) === i; });
  unique.sort(function(a, b) { return b > a ? 1 : -1; });
  return unique;
}

// ── Admin homepage (管理員首頁) ───────────────────────────────

/**
 * Get pending review counts for admin homepage.
 * @returns {Object} {leave, clockFix, upload, message}
 */
function getPendingCounts() {
  try {
    var countSheet = function(sheetName) {
      var s = getSheet_(sheetName);
      if (!s) return 0;
      var data = s.getDataRange().getValues();
      if (data.length < 2) return 0;
      var headers = data[0].map(function(h) { return clean_(h); });
      var statusIdx = headers.indexOf('狀態') >= 0 ? headers.indexOf('狀態') : headers.indexOf('審核狀態');
      if (statusIdx < 0) return 0;
      var count = 0;
      for (var i = 1; i < data.length; i++) {
        if (clean_(data[i][statusIdx]) === '待審核') count++;
      }
      return count;
    };
    return {
      leave:    countSheet(SHEET_LEAVE),
      clockFix: countSheet(SHEET_CLOCK_FIX),
      upload:   countSheet(SHEET_UPLOAD),
      message:  0  // placeholder for message review
    };
  } catch (e) {
    return { leave: 0, clockFix: 0, upload: 0, message: 0 };
  }
}

// ── Web-app entry points ──────────────────────────────────────

/**
 * doGet: Serve the employee HTML page.
 */
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('employee')
    .setTitle('員工自助系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * doPost: Handle AJAX calls from the front-end.
 */
function doPost(e) {
  var params = JSON.parse(e.postData.contents);
  var action = params.action;
  var result;

  switch (action) {
    case 'submitClockFix':
      result = submitEmployeeClockFix(
        params.empId, params.timeVal, params.type, params.reason);
      break;
    case 'submitLeave':
      result = submitLeaveRequest(
        params.empId, params.leaveType, params.startDate, params.days, params.reason);
      break;
    case 'submitUpload':
      result = submitEmployeeUpload(
        params.empId, params.type, params.note, params.amount, params.item,
        params.location, params.fileName, params.fileMime, params.fileData);
      break;
    case 'getReviewCards':
      result = getReviewCardsForRole_(params.role, params.dept || '');
      break;
    case 'reviewLeave':
      result = reviewRequest_(SHEET_LEAVE, params.rowIndex, params.decision, params.comment);
      break;
    case 'reviewClockFix':
      result = reviewRequest_(SHEET_CLOCK_FIX, params.rowIndex, params.decision, params.comment);
      break;
    case 'submitPreselect':
      result = submitPreselect(params.empId, params.yearMonth, params.vacationDays);
      break;
    case 'writePreselectSchedule':
      result = writePreselectSchedule_(params.empId, params.weekStart, params.weekdays || []);
      break;
    case 'getSalaryMonths':
      result = { months: getSalaryMonths(params.empId) };
      break;
    case 'getPendingCounts':
      result = getPendingCounts();
      break;
    default:
      result = { success: false, message: '未知的 action：' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
