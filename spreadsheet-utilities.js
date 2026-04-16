// ============================================================
// spreadsheet-utilities.js  –  Google Apps Script backend
// ============================================================

// ---------- 欄位索引定義 ----------
var RECEIVE_COLS = {
  ID:                 0,   // A 資料編號
  TYPE:               1,   // B 類型
  EMP_ID:             2,   // C 員編
  NAME:               3,   // D 姓名
  TITLE:              4,   // E 標題
  NOTE:               5,   // F 申請內容（純補充說明）
  AMOUNT:             6,   // G 代墊金額
  ITEM:               7,   // H 代墊商品
  LOCATION:           8,   // I 位置資訊
  ATTACHMENT:         9,   // J 附件
  SUBMIT_TIME:        10,  // K 提交時間
  STATUS:             11,  // L 狀態
  REVIEWER:           12,  // M 審核者
  REVIEW_TIME:        13,  // N 審核時間
  REVIEW_NOTE:        14,  // O 審核備註
  REQUIRED_APPROVERS: 15,  // P requiredApprovers
  CURRENT_STEP:       16,  // Q currentStep
  APPROVED_LIST:      17   // R approvedList
};

// ---------- 工具函式 ----------

/** 移除首尾空白；null/undefined 轉成空字串 */
function clean_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

/** 取得試算表（依名稱） */
function getSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

/** 產生唯一 ID */
function generateId_() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMddHHmmss') +
         Math.random().toString(36).slice(2, 6).toUpperCase();
}

/** 取得目前登入員工資訊 */
function getCurrentUser_() {
  var email = Session.getActiveUser().getEmail();
  var sh = getSheet_('員工資料');
  if (!sh) return null;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === email) { // 假設第 3 欄為 email
      return { id: data[i][0], name: data[i][1], email: email, row: i + 1 };
    }
  }
  return { id: email, name: email, email: email, row: -1 };
}

/** 取得審核者清單 */
function getApprovers_(type) {
  var sh = getSheet_('審核設定');
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === type) {
      return String(data[i][1]).split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    }
  }
  return [];
}

// ============================================================
// 選休（Preselect）相關
// ============================================================

/**
 * 將選休結果寫入排班試算表。
 *
 * 邏輯（Fix 1）：
 *   - 先將當月所有日期對應儲存格設為 TRUE（= 上班）
 *   - 再將員工選中的日期設為 FALSE（= 休假）
 *
 * @param {string} empId   員工編號
 * @param {string} yearMonth  格式 "YYYY-MM"
 * @param {number[]} selectedDates  員工選擇休假的日期陣列（1-based）
 */
function writePreselectSchedule_(empId, yearMonth, selectedDates) {
  var sh = getSheet_('排班');
  if (!sh) throw new Error('找不到「排班」工作表');

  var parts = yearMonth.split('-');
  var year  = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var daysInMonth = new Date(year, month, 0).getDate();

  // 找到員工所在欄
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var empCol = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]) === empId) { empCol = c + 1; break; }
  }
  if (empCol === -1) throw new Error('找不到員工：' + empId);

  // 找到該月第一個日期所在列
  var dateCol = 1; // 假設第 A 欄為日期
  var allDates = sh.getRange(2, dateCol, sh.getLastRow() - 1, 1).getValues();
  var startRow = -1;
  for (var r = 0; r < allDates.length; r++) {
    var d = new Date(allDates[r][0]);
    if (d.getFullYear() === year && d.getMonth() + 1 === month && d.getDate() === 1) {
      startRow = r + 2; break;
    }
  }
  if (startRow === -1) throw new Error('找不到該月份起始列');

  // Fix 1：先全部設 TRUE（= 上班）
  for (var day = 0; day < daysInMonth; day++) {
    sh.getRange(startRow + day, empCol).setValue(true);
  }

  // Fix 1：選中的設 FALSE（= 休假）
  var selectedSet = {};
  (selectedDates || []).forEach(function(d) { selectedSet[d] = true; });
  for (var day2 = 1; day2 <= daysInMonth; day2++) {
    if (selectedSet[day2]) {
      sh.getRange(startRow + day2 - 1, empCol).setValue(false);
    }
  }
}

/**
 * 員工提交選休。
 * Fix 1：selectedDates 裡的日期代表「休假」，寫入 FALSE。
 *
 * @param {Object} payload  { yearMonth: "YYYY-MM", selectedDates: [1,2,...] }
 */
function submitPreselect(payload) {
  try {
    var user = getCurrentUser_();
    if (!user) return { ok: false, msg: '無法識別使用者' };

    var yearMonth    = clean_(payload.yearMonth);
    var selectedDates = payload.selectedDates || [];   // Fix 1：選中 = 休假

    writePreselectSchedule_(user.id, yearMonth, selectedDates);

    // 記錄選休紀錄
    var logSh = getSheet_('選休紀錄');
    if (logSh) {
      logSh.appendRow([
        user.id,
        user.name,
        yearMonth,
        selectedDates.join(','),   // 休假日期清單
        new Date()
      ]);
    }

    return { ok: true, msg: '選休提交成功' };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

/**
 * 取得目前使用者的選休紀錄。
 * Fix 1：FALSE = 休假 = 員工選中的日期。
 *
 * @param {string} yearMonth  "YYYY-MM"
 * @returns {{ ok: boolean, selected: number[] }}  selected = 休假日期（1-based）
 */
function getMyPreselect(yearMonth) {
  try {
    var user = getCurrentUser_();
    if (!user) return { ok: false, msg: '無法識別使用者', selected: [] };
    return getPreselectByUser_(user.id, yearMonth);
  } catch (e) {
    return { ok: false, msg: e.message, selected: [] };
  }
}

/**
 * 依員工編號取得選休紀錄。
 * Fix 1：從排班表讀取時，FALSE = 休假 = 員工選中的日期，回傳這些日期。
 *
 * @param {string} empId
 * @param {string} yearMonth  "YYYY-MM"
 * @returns {{ ok: boolean, selected: number[] }}
 */
function getPreselectByUser_(empId, yearMonth) {
  var sh = getSheet_('排班');
  if (!sh) return { ok: false, msg: '找不到排班工作表', selected: [] };

  var parts = yearMonth.split('-');
  var year  = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var daysInMonth = new Date(year, month, 0).getDate();

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var empCol = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]) === empId) { empCol = c + 1; break; }
  }
  if (empCol === -1) return { ok: true, selected: [] };

  var dateCol = 1;
  var allDates = sh.getRange(2, dateCol, sh.getLastRow() - 1, 1).getValues();
  var startRow = -1;
  for (var r = 0; r < allDates.length; r++) {
    var d = new Date(allDates[r][0]);
    if (d.getFullYear() === year && d.getMonth() + 1 === month && d.getDate() === 1) {
      startRow = r + 2; break;
    }
  }
  if (startRow === -1) return { ok: true, selected: [] };

  var selected = [];
  for (var day = 1; day <= daysInMonth; day++) {
    // Fix 1：FALSE = 休假 = 員工選中的日期
    var val = sh.getRange(startRow + day - 1, empCol).getValue();
    if (val === false || val === 'FALSE' || val === 0) {
      selected.push(day);
    }
  }

  return { ok: true, selected: selected };
}

// ============================================================
// 資料上傳（Employee Upload）
// ============================================================

/**
 * 員工提交資料上傳（請假、外勤回報、費用請款等）。
 *
 * Fix 3：每個欄位獨立寫入對應欄，不再全部塞進 note。
 *
 * @param {Object} payload
 *   {
 *     type:     string,   // 類型
 *     title:    string,   // 標題
 *     note:     string,   // 申請內容（補充說明）
 *     amount:   string,   // 代墊金額
 *     item:     string,   // 代墊商品
 *     location: string,   // 位置資訊
 *     attachment: string  // 附件 Base64 或 URL（可選）
 *   }
 */
function submitEmployeeUpload(payload) {
  try {
    var user = getCurrentUser_();
    if (!user) return { ok: false, msg: '無法識別使用者' };

    var sh = getSheet_('申請紀錄');
    if (!sh) return { ok: false, msg: '找不到「申請紀錄」工作表' };

    // 處理附件
    var attachmentUrl = '';
    if (payload.attachment) {
      try {
        var blob = Utilities.newBlob(
          Utilities.base64Decode(payload.attachment.split(',')[1] || payload.attachment),
          'application/octet-stream',
          '附件_' + generateId_()
        );
        var folderIter = DriveApp.getFoldersByName('申請附件');
        var folder = folderIter.hasNext() ? folderIter.next() : DriveApp.createFolder('申請附件');
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        attachmentUrl = file.getUrl();
      } catch (e) {
        attachmentUrl = '';
      }
    }

    var reqId    = generateId_();
    var approvers = getApprovers_(clean_(payload.type));

    // Fix 3：每個欄位獨立寫入對應欄
    sh.appendRow([
      reqId,                                        // A 資料編號
      clean_(payload.type),                         // B 類型
      user.id,                                      // C 員編
      user.name,                                    // D 姓名
      clean_(payload.title || payload.type),        // E 標題
      clean_(payload.note),                         // F 申請內容（純補充說明）
      clean_(payload.amount || ''),                 // G 代墊金額（獨立）
      clean_(payload.item   || ''),                 // H 代墊商品（獨立）
      clean_(payload.location || ''),               // I 位置資訊（獨立）
      attachmentUrl,                                // J 附件
      new Date(),                                   // K 提交時間
      '待審核',                                     // L 狀態
      '', '', '',                                   // M~O 審核欄位
      approvers.join(','),                          // P requiredApprovers
      approvers[0] || '',                           // Q currentStep
      '[]'                                          // R approvedList
    ]);

    return { ok: true, msg: '提交成功', id: reqId };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// ============================================================
// 薪資相關
// ============================================================

/** 取得員工薪資紀錄清單 */
function getSalaryHistory() {
  try {
    var user = getCurrentUser_();
    if (!user) return { ok: false, msg: '無法識別使用者', records: [] };

    var sh = getSheet_('薪資紀錄');
    if (!sh) return { ok: true, records: [] };

    var data = sh.getDataRange().getValues();
    var records = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(user.id)) {
        records.push({
          month:  data[i][1],   // YYYY-MM
          amount: data[i][2],
          status: data[i][3]
        });
      }
    }
    return { ok: true, records: records };
  } catch (e) {
    return { ok: false, msg: e.message, records: [] };
  }
}

/** 下載薪資單（回傳 PDF base64） */
function getSalarySlipData(yearMonth) {
  try {
    var user = getCurrentUser_();
    if (!user) return { ok: false, msg: '無法識別使用者' };

    var sh = getSheet_('薪資單');
    if (!sh) return { ok: false, msg: '找不到薪資單工作表' };

    // 實際實作依試算表結構調整
    return { ok: true, month: yearMonth, url: '' };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// ============================================================
// 打卡相關
// ============================================================

/** 打卡 */
function clockIn(payload) {
  try {
    var user = getCurrentUser_();
    if (!user) return { ok: false, msg: '無法識別使用者' };

    var sh = getSheet_('出勤紀錄');
    if (!sh) return { ok: false, msg: '找不到出勤紀錄工作表' };

    sh.appendRow([
      user.id,
      user.name,
      payload.type || '上班',
      new Date(),
      clean_(payload.location || '')
    ]);
    return { ok: true, msg: payload.type === '下班' ? '下班打卡成功' : '上班打卡成功' };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

/** 取得今日打卡紀錄 */
function getTodayClockRecord() {
  try {
    var user = getCurrentUser_();
    if (!user) return { ok: false, msg: '無法識別使用者', records: [] };

    var sh = getSheet_('出勤紀錄');
    if (!sh) return { ok: true, records: [] };

    var today = new Date();
    var todayStr = Utilities.formatDate(today, 'Asia/Taipei', 'yyyy-MM-dd');
    var data = sh.getDataRange().getValues();
    var records = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(user.id)) {
        var rowDate = Utilities.formatDate(new Date(data[i][3]), 'Asia/Taipei', 'yyyy-MM-dd');
        if (rowDate === todayStr) {
          records.push({ type: data[i][2], time: data[i][3] });
        }
      }
    }
    return { ok: true, records: records };
  } catch (e) {
    return { ok: false, msg: e.message, records: [] };
  }
}

// ============================================================
// 申請紀錄查詢
// ============================================================

/** 取得員工自己的申請紀錄 */
function getMyUploadRecords() {
  try {
    var user = getCurrentUser_();
    if (!user) return { ok: false, msg: '無法識別使用者', records: [] };

    var sh = getSheet_('申請紀錄');
    if (!sh) return { ok: true, records: [] };

    var data = sh.getDataRange().getValues();
    var records = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][RECEIVE_COLS.EMP_ID]) === String(user.id)) {
        records.push({
          id:         data[i][RECEIVE_COLS.ID],
          type:       data[i][RECEIVE_COLS.TYPE],
          title:      data[i][RECEIVE_COLS.TITLE],
          note:       data[i][RECEIVE_COLS.NOTE],
          amount:     data[i][RECEIVE_COLS.AMOUNT],
          item:       data[i][RECEIVE_COLS.ITEM],
          location:   data[i][RECEIVE_COLS.LOCATION],
          attachment: data[i][RECEIVE_COLS.ATTACHMENT],
          submitTime: data[i][RECEIVE_COLS.SUBMIT_TIME],
          status:     data[i][RECEIVE_COLS.STATUS]
        });
      }
    }
    return { ok: true, records: records };
  } catch (e) {
    return { ok: false, msg: e.message, records: [] };
  }
}
