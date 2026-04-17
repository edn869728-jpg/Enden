const SPREADSHEET_ID = '1NZxOjKWa6OmzeYlDklbN_GcDzsFlLkC7kJGzi-1XDpU';
const UPLOAD_FOLDER_ID = '10ED2Bd72agzl6ZyuQYFAHdl0de02zgv0';
const TZ = Session.getScriptTimeZone() || 'Asia/Taipei';

const SHEET_USERS = '人員資料';
const SHEET_SYSTEM = '系統設定';
const SHEET_CLOCK = '打卡原始記錄';
const SHEET_LEAVE = '請假申請';
const SHEET_UPLOAD = '資料接收審核';
const SHEET_NOTICE = '通知發布';
const SHEET_MESSAGE = '留言審核';
const SHEET_CLOCK_FIX = '補打卡申請';
const SHEET_SALARY = '薪資審核';
const SHEET_PUBLISH = '排班發布';
const SHEET_DEADLINE = '班表提報期限';
const SHEET_MODIFY_LOG = '資料修改紀錄';
const SHEET_PERMISSION = '權限管理';
const SHEET_PRESELECT = '選休表';

const DUP_MINUTES = 15;

/* =========================  
 * 基本工具
 * ========================= */
function openSS_() {
  const id = String(SPREADSHEET_ID || '').trim();
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) {
    throw new Error('找不到綁定試算表。請確認此 GAS 專案已綁定到目標試算表。');
  }
  return ss;
}

function getSheet_(name) {
  return openSS_().getSheetByName(name);
}

function getOrCreateSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function clean_(v) {
  return String(v == null ? '' : v).trim();
}

function normalizeEnabled_(v) {
  const s = clean_(v).toLowerCase();
  return ['是', 'y', 'yes', 'true', '1', 'on', '啟用'].indexOf(s) > -1;
}

function formatDate_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy/MM/dd');
}

function formatTime_(d) {
  return Utilities.formatDate(d, TZ, 'HH:mm:ss');
}

function formatDateTime_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy/MM/dd HH:mm:ss');
}

function formatYmdMaybe_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy/MM/dd');
  return clean_(v);
}

function formatDateTimeMaybe_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy/MM/dd HH:mm:ss');
  return clean_(v);
}

function formatInputDateTime_(v) {
  if (!(v instanceof Date)) return clean_(v);
  return Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm");
}

function now_() {
  return new Date();
}

function minutesLater_(m) {
  return new Date(Date.now() + m * 60 * 1000);
}

function generateKey_(len) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < len; i++) {
    s += chars[Math.floor(Math.random() * chars.length)];
  }
  return s;
}

function generateToken_() {
  return Utilities.getUuid().replace(/-/g, '') + String(Date.now());
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseRequest_(e) {
  let data = {};

  if (e && e.postData && e.postData.contents) {
    const raw = String(e.postData.contents || '');
    const type = String(e.postData.type || '').toLowerCase();

    if (
      type.indexOf('application/json') > -1 ||
      type.indexOf('text/plain') > -1
    ) {
      try {
        data = JSON.parse(raw || '{}');
      } catch (err) {
        data = {};
      }
    } else {
      data = Object.fromEntries(
        Object.entries(e.parameter || {}).map(([k, v]) => [k, v])
      );
    }
  } else if (e && e.parameter) {
    data = Object.fromEntries(
      Object.entries(e.parameter).map(([k, v]) => [k, v])
    );
  }

  return data || {};
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* =========================
 * 班表優先工時計算
 * ========================= */
function parseHmToMinutes_(hm) {
  const s = clean_(hm);
  if (!s) return null;

  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;

  const hh = Number(m[1]);
  const mm = Number(m[2]);

  if (isNaN(hh) || isNaN(mm)) return null;
  if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;

  return hh * 60 + mm;
}

function minutesToHm_(mins) {
  const n = Number(mins);
  if (!isFinite(n)) return '';

  const hh = Math.floor(n / 60);
  const mm = n % 60;

  return ('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2);
}

function clampRoundUnitMinutes_(v) {
  const n = Number(v || 0);
  if ([5, 10, 15, 30].indexOf(n) > -1) return n;
  return 15;
}

function clampGraceMinutes_(v) {
  const n = Number(v || 0);
  if (!isFinite(n)) return 15;
  if (n < 0) return 0;
  if (n > 30) return 30;
  return n;
}

function getAttendanceSettings_() {
  const system = getSystemSettingsMap_();

  return {
    schedule_first_enabled: String(system.schedule_first_enabled || 'true').toLowerCase() !== 'false',
    checkin_early_allow_minutes: clampGraceMinutes_(system.checkin_early_allow_minutes || 30),
    checkin_grace_minutes: clampGraceMinutes_(system.checkin_grace_minutes || 15),
    checkout_grace_minutes: clampGraceMinutes_(system.checkout_grace_minutes || 15),
    round_unit_minutes: clampRoundUnitMinutes_(system.round_unit_minutes || 15)
  };
}

function calcCheckInPayMinutes_(rawInMinutes, shiftStartMinutes, settings) {
  const earlyAllow = Number(settings.checkin_early_allow_minutes || 30);
  const grace = Number(settings.checkin_grace_minutes || 15);
  const unit = Number(settings.round_unit_minutes || 15);

  const earliestNormal = shiftStartMinutes - earlyAllow;
  const latestNormal = shiftStartMinutes + grace;

  if (rawInMinutes >= earliestNormal && rawInMinutes <= latestNormal) {
    return {
      payMinutes: shiftStartMinutes,
      lateFlag: rawInMinutes > shiftStartMinutes,
      lateMinutes: rawInMinutes > shiftStartMinutes ? (rawInMinutes - shiftStartMinutes) : 0,
      note: rawInMinutes < shiftStartMinutes ? '提早到班，計薪以上班班表時間計' : (rawInMinutes > shiftStartMinutes ? '寬限內遲到，計薪以上班班表時間計' : '正常上班')
    };
  }

  if (rawInMinutes < earliestNormal) {
    return {
      payMinutes: shiftStartMinutes,
      lateFlag: false,
      lateMinutes: 0,
      note: '提早到班超過可接受區間，仍以上班班表時間計'
    };
  }

  const over = rawInMinutes - latestNormal;
  const steps = Math.floor(over / unit) + 1;
  const payMinutes = shiftStartMinutes + steps * unit;

  return {
    payMinutes: payMinutes,
    lateFlag: true,
    lateMinutes: Math.max(0, rawInMinutes - shiftStartMinutes),
    note: '超過上班寬限，往後推到下一個單位班別'
  };
}

function calcCheckOutPayMinutes_(rawOutMinutes, shiftEndMinutes, settings) {
  const grace = Number(settings.checkout_grace_minutes || 15);
  const unit = Number(settings.round_unit_minutes || 15);

  const earliestNormal = shiftEndMinutes - grace;

  if (rawOutMinutes >= earliestNormal) {
    return {
      payMinutes: shiftEndMinutes,
      earlyLeaveFlag: rawOutMinutes < shiftEndMinutes,
      earlyLeaveMinutes: rawOutMinutes < shiftEndMinutes ? (shiftEndMinutes - rawOutMinutes) : 0,
      note: rawOutMinutes < shiftEndMinutes ? '寬限內早退，計薪以下班班表時間計' : '正常下班或延後下班'
    };
  }

  const diff = earliestNormal - rawOutMinutes;
  const steps = Math.floor(diff / unit) + 1;
  const payMinutes = shiftEndMinutes - steps * unit;

  return {
    payMinutes: payMinutes,
    earlyLeaveFlag: true,
    earlyLeaveMinutes: Math.max(0, shiftEndMinutes - rawOutMinutes),
    note: '超過下班寬限，往前退到上一個單位班別'
  };
}

function calcAttendanceByShift_(rawInTime, rawOutTime, shiftStart, shiftEnd, settings) {
  const cfg = settings || getAttendanceSettings_();

  const rawInMinutes = parseHmToMinutes_(rawInTime);
  const rawOutMinutes = parseHmToMinutes_(rawOutTime);
  const shiftStartMinutes = parseHmToMinutes_(shiftStart);
  const shiftEndMinutes = parseHmToMinutes_(shiftEnd);

  if (rawInMinutes == null || rawOutMinutes == null || shiftStartMinutes == null || shiftEndMinutes == null) {
    return {
      ok: false,
      rawIn: clean_(rawInTime),
      rawOut: clean_(rawOutTime),
      payIn: '',
      payOut: '',
      shiftStart: clean_(shiftStart),
      shiftEnd: clean_(shiftEnd),
      hours: 0,
      lateFlag: false,
      earlyLeaveFlag: false,
      lateMinutes: 0,
      earlyLeaveMinutes: 0,
      checkInNote: '時間格式錯誤',
      checkOutNote: '時間格式錯誤'
    };
  }

  const checkIn = calcCheckInPayMinutes_(rawInMinutes, shiftStartMinutes, cfg);
  const checkOut = calcCheckOutPayMinutes_(rawOutMinutes, shiftEndMinutes, cfg);

  let payInMinutes = checkIn.payMinutes;
  let payOutMinutes = checkOut.payMinutes;

  if (payOutMinutes < payInMinutes) {
    payOutMinutes = payInMinutes;
  }

  const mins = payOutMinutes - payInMinutes;
  const hours = Math.round((mins / 60) * 100) / 100;

  return {
    ok: true,
    rawIn: minutesToHm_(rawInMinutes),
    rawOut: minutesToHm_(rawOutMinutes),
    payIn: minutesToHm_(payInMinutes),
    payOut: minutesToHm_(payOutMinutes),
    shiftStart: minutesToHm_(shiftStartMinutes),
    shiftEnd: minutesToHm_(shiftEndMinutes),
    hours: hours,
    lateFlag: !!checkIn.lateFlag,
    earlyLeaveFlag: !!checkOut.earlyLeaveFlag,
    lateMinutes: Number(checkIn.lateMinutes || 0),
    earlyLeaveMinutes: Number(checkOut.earlyLeaveMinutes || 0),
    checkInNote: checkIn.note || '',
    checkOutNote: checkOut.note || ''
  };
}

function calcHours_(inTime, outTime) {
  if (!inTime || !outTime) return 0;

  const inMinutes = parseHmToMinutes_(inTime);
  const outMinutes = parseHmToMinutes_(outTime);

  if (inMinutes == null || outMinutes == null) return 0;
  if (outMinutes <= inMinutes) return 0;

  return Math.round(((outMinutes - inMinutes) / 60) * 100) / 100;
}

function applyShiftAttendance_(rawInTime, rawOutTime, shiftStart, shiftEnd) {
  return calcAttendanceByShift_(
    rawInTime,
    rawOutTime,
    shiftStart,
    shiftEnd,
    getAttendanceSettings_()
  );
}

function roundMoney_(n) {
  return Math.round(Number(n || 0));
}

function safeLastRow_(sh) {
  return Math.max(0, sh.getLastRow());
}

function ensureSheetHeaders_(sh, headers, rows) {
  sh.clear();

  if (headers && headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#ffe6d5')
      .setHorizontalAlignment('center');
    sh.setFrozenRows(1);
  }

  if (rows && rows.length) {
    sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  sh.autoResizeColumns(1, Math.max(headers.length, 1));
}

function ensureClockHeaders_(sh) {
  if (safeLastRow_(sh) === 0) {
    sh.getRange(1, 1, 1, 8).setValues([[
      '姓名',
      '日期時間',
      '員編',
      '動作',
      '來源',
      '時間戳記',
      'UUID',
      '備註'
    ]]);
  }
}

function ensureSessionsHeaders_(sh) {
  if (safeLastRow_(sh) === 0) {
    sh.appendRow(['key', 'id', 'page', 'createdAt', 'expireAt', 'used']);
  }
}

function ensureSheetHasRows_(sh, rowCount) {
  const need = rowCount - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
}

/* =========================
 * 初始化
 * ========================= */
function 初始化試算表() {
  const ss = openSS_();

  const SHEETS = [
    {
      name: SHEET_USERS,
      headers: [
      '暱稱',
      '姓名',
      '員編',
      '職稱',
      '部門',
      '生日',
      '身分證',
      '班別',
      '職級',
      '顏色',
      '啟用',
      '建立時間',
      'specialdeviceid',
      'token',
      '密碼',
      '換機密碼'
      ]
    },
    {
      name: SHEET_SYSTEM,
      headers: [
        '設定鍵',
        '設定值',
        '說明',
        '更新時間',
        '更新者'
      ],
      rows: [
      ['company_name', 'ANG.lo', '公司名稱', new Date(), 'system'],
      ['company_subtitle', 'Humanized system technology', '公司副名稱', new Date(), 'system'],
      ['system_title', 'ANG.lo Engine', '系統主標題', new Date(), 'system'],
      ['logo_url', '', 'Logo 圖片網址或 base64', new Date(), 'system'],
      ['fallback_text', 'ANG', '無 Logo 時顯示文字', new Date(), 'system'],
      ['week_open_at', '', '下週班表開放提報時間', new Date(), 'system'],
      ['week_close_at', '', '下週班表截止提報時間', new Date(), 'system'],
      ['month_open_at', '', '月班表開放提報時間', new Date(), 'system'],
      ['month_close_at', '', '月班表截止提報時間', new Date(), 'system'],
      ['schedule_rule', '每週休2天', '排班規則', new Date(), 'system'],
      ['publish_schedule', '每週三 20:00 發布排班', '發布日時程', new Date(), 'system']
]
    },
    {
      name: SHEET_CLOCK,
      headers: [
        '姓名',
        '日期時間',
        '員編',
        '動作',
        '來源',
        '時間戳記',
        'UUID',
        '備註'
      ]
    },
    {
      name: SHEET_LEAVE,
      headers: [
        '申請編號',
        '員編',
        '姓名',
        '假別',
        '開始日期',
        '結束日期',
        '天數',
        '事由',
        '申請時間',
        '狀態',
        '初審者',
        '初審時間',
        '最終審核者',
        '最終審核時間',
        '審核備註'
      ]
    },
    {
      name: SHEET_UPLOAD,
      headers: [
        '資料編號',
        '類型',
        '員編',
        '姓名',
        '標題',
        '內容',
        '附件',
        '提交時間',
        '狀態',
        '審核者',
        '審核時間',
        '審核備註'
      ]
    },
    {
      name: SHEET_NOTICE,
      headers: [
        '發布編號',
        '類型',
        '標題',
        '內容',
        '狀態',
        '發布時間',
        '建立者',
        '更新時間',
        '備註'
      ]
    },
    {
      name: SHEET_MESSAGE,
      headers: [
        '留言編號',
        '員編',
        '姓名',
        '留言內容',
        '留言時間',
        '狀態',
        '審核者',
        '審核時間',
        '審核備註'
      ]
    },
    {
      name: SHEET_CLOCK_FIX,
      headers: [
        '申請編號',
        '員編',
        '姓名',
        '補打卡日期',
        '補打卡時間',
        '補打卡動作',
        '申請事由',
        '申請時間',
        '狀態',
        '審核者',
        '審核時間',
        '審核備註'
      ]
    },
    {
      name: SHEET_SALARY,
      headers: [
        '薪資編號',
        '員編',
        '姓名',
        '月份',
        '工時',
        '加班時數',
        '應發金額',
        '狀態',
        '審核者',
        '審核時間',
        '備註'
      ]
    },
    {
      name: SHEET_PUBLISH,
      headers: [
        '發布編號',
        '類型',
        '期間',
        '標題',
        '內容',
        '狀態',
        '發布時間',
        '發布者',
        '備註'
      ]
    },
    {
      name: SHEET_DEADLINE,
      headers: [
        '設定編號',
        '類型',
        '開放時間',
        '截止時間',
        '提醒說明',
        '更新時間',
        '更新者'
      ],
      rows: [
        ['DL001', '下週班表', '', '', '逾時後不可再自行提報，需由主管協助處理。', new Date(), 'system'],
        ['DL002', '月班表', '', '', '月班表截止後，僅能由管理端調整。', new Date(), 'system']
      ]
    },
    {
      name: SHEET_MODIFY_LOG,
      headers: [
        '修改編號',
        '資料表',
        '資料列主鍵',
        '修改欄位',
        '修改前',
        '修改後',
        '修改者',
        '修改時間',
        '備註'
      ]
    },
        
    {
      name: SHEET_PERMISSION,
      headers: [
        '角色',
        '功能名稱',
        '啟用',
        '更新時間',
        '更新者'
      ],
      rows: [
        // Manager 預設權限
        ['Manager', '排班發布中心', 'Y', new Date(), 'system'],
        ['Manager', '請假審核中心', 'Y', new Date(), 'system'],
        ['Manager', '資料接收審核中心', 'Y', new Date(), 'system'],
        ['Manager', '通知發布中心', 'Y', new Date(), 'system'],
        ['Manager', '留言審核中心', 'Y', new Date(), 'system'],

        // Admin 預設權限
        ['Admin', '打卡管理中心', 'Y', new Date(), 'system'],
        ['Admin', '補打卡審核中心', 'Y', new Date(), 'system'],
        ['Admin', '資料修改中心', 'Y', new Date(), 'system'],
        ['Admin', '薪資審核中��', 'Y', new Date(), 'system'],
        ['Admin', '下週／月班表提報期限設定', 'Y', new Date(), 'system'],
        ['Admin', 'Logo and Title 設定上傳', 'Y', new Date(), 'system'],

        // Creator 完整權限
        ['Creator', '排班發布中心', 'Y', new Date(), 'system'],
        ['Creator', '請假審核中心', 'Y', new Date(), 'system'],
        ['Creator', '資料接收審核中心', 'Y', new Date(), 'system'],
        ['Creator', '通知發布中心', 'Y', new Date(), 'system'],
        ['Creator', '留言審核中心', 'Y', new Date(), 'system'],
        ['Creator', '打卡管理中心', 'Y', new Date(), 'system'],
        ['Creator', '補打卡審核中心', 'Y', new Date(), 'system'],
        ['Creator', '資料修改中心', 'Y', new Date(), 'system'],
        ['Creator', '薪資審核中心', 'Y', new Date(), 'system'],
        ['Creator', '排班規則設定中心', 'Y', new Date(), 'system'],
        ['Creator', '發布日時程設定中心', 'Y', new Date(), 'system'],
        ['Creator', '下週／月班表提報期限設定', 'Y', new Date(), 'system'],
        ['Creator', 'Logo and Title 設定上傳', 'Y', new Date(), 'system'],
        ['Creator', '權限管理', 'Y', new Date(), 'system'],
        ['Creator', '系統設定', 'Y', new Date(), 'system']
      ]
    }
  ];

  SHEETS.forEach(cfg => {
    const sh = getOrCreateSheet_(ss, cfg.name);
    ensureSheetHeaders_(sh, cfg.headers, cfg.rows || []);
  });

  建立範例人員資料_();

  SpreadsheetApp.flush();
  Logger.log('初始化完成：' + ss.getUrl());
}

function initSheets_() {
  const ss = openSS_();

  const shUsers = getOrCreateSheet_(ss, SHEET_USERS);
  const shClock = getOrCreateSheet_(ss, SHEET_CLOCK);

  if (safeLastRow_(shUsers) === 0) {
    shUsers.getRange(1, 1, 1, 16).setValues([[
      '暱稱',
      '姓名',
      '員編',
      '職稱',
      '部門',
      '生日',
      '身分證',
      '班別',
      '職級',
      '顏色',
      '啟用',
      '建立時間',
      'specialdeviceid',
      'token',
      '密碼',
      '換機密碼'
    ]]);
  }

  ensureClockHeaders_(shClock);
}

function 建立範例人員資料_() {
  const ss = openSS_();
  const sh = ss.getSheetByName(SHEET_USERS);
  if (!sh) return;

  const sample = [
    ['悠悠', '林芷儀', 'ANG0601', '美編', '廣告', '', '', 'B', 'Manager', '#9b59b6', '是', new Date(), '', '9d8c7b6a504132219d8c7b6a504132219d8c7b6a50413', '', ''],
    ['程程', '張程淮', 'ANG0602', '媒體', '廣告', '', '', 'B', 'Admin', '#27ae60', '是', new Date(), '', 'bc9a8d7f6e5d4c3b2a109876543210febc9a8d7f6e5d4', '', ''],
    ['米米', '米馳浩', 'ANG0603', '資訊', '系統', '1997/1/28', 'M122800149', 'F', 'Creator', '#ff6a00', '是', new Date(), '926Enden 的 iPhone428', '0273d013afe64538aeb3f52d18d289ab1775139447000', '', ''],
    ['翰翰', '吳勇翰', 'ANG0604', '銷售', '門市', '', '', 'B', 'Employee', '#3498db', '是', new Date(), '', 'f8e9d7c6b5a41234f8e9d7c6b5a41234f8e9d7c6b5a41', '', ''],
    ['楊楊', '楊福濡', 'ANG0605', '銷售', '門市', '1979/2/25', '', 'F', 'Employee', '#e84393', '是', new Date(), '', 'a1b2c3d4e5f67890a1b2c3d4e5f67890a1b2c3d4e5f67', '', '']
  ];

  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, sample.length, sample[0].length).setValues(sample);
  }
}

/* =========================
 * 驗證 / 使用者
 * ========================= */
function getUserById_(id) {
  const sh = getSheet_(SHEET_USERS);

  if (!sh) {
    return { ok: false, message: '找不到 人員資料 工作表' };
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return { ok: false, message: '人員資料沒有任何員工' };
  }

  const values = sh.getRange(2, 1, lastRow - 1, 16).getValues();
  const targetId = clean_(id).toUpperCase();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const empId = clean_(row[2]).toUpperCase();
    if (empId !== targetId) continue;

    return {
      ok: true,
      rowIndex: i + 2,
      nickname: clean_(row[0]),
      name: clean_(row[1]) || clean_(row[0]) || empId,
      id: empId,
      jobTitle: clean_(row[3]),
      dept: clean_(row[4]),
      role: clean_(row[8]),
      color: clean_(row[9]),
      enabled: normalizeEnabled_(row[10]),
      specialdeviceid: clean_(row[12]),
      token: clean_(row[13])
    };
  }

  return { ok: false, message: '查無此員編' };
}

function getUserByIdAndToken_(id, token) {
  const user = getUserById_(id);
  if (!user.ok) return user;

  if (!user.enabled) {
    return { ok: false, message: '此帳號未啟用' };
  }

  if (clean_(user.token) !== clean_(token)) {
    return { ok: false, message: '員編或 token 不正確' };
  }

  return {
    ok: true,
    nickname: user.nickname,
    name: user.name,
    id: user.id,
    token: user.token,
    role: user.role,
    color: user.color,
    specialdeviceid: user.specialdeviceid
  };
}

function authUser_(id, token) {
  const user = getUserByIdAndToken_(id, token);
  if (!user.ok) return { ok: false, message: user.message || '驗證失敗' };
  return { ok: true, user: user };
}

/* =========================
 * 打卡 + 設備綁定
 * ========================= */
function ensureUserTokenByIdAndDevice_(id, deviceId) {
  const sh = getSheet_(SHEET_USERS);

  if (!sh) {
    return {
      ok: false,
      message: '找不到 人員資料 工作表'
    };
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return {
      ok: false,
      message: '人員資料沒有任何員工'
    };
  }

  const values = sh.getRange(2, 1, lastRow - 1, 16).getValues();
  const targetId = clean_(id).toUpperCase();
  const targetDevice = clean_(deviceId);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const empId = clean_(row[2]).toUpperCase();
    const enabled = normalizeEnabled_(row[10]);
    const specialdeviceid = clean_(row[12]);
    let token = clean_(row[13]);

    if (empId !== targetId) continue;

    if (!enabled) {
      return { ok: false, message: '此帳號未啟用' };
    }

    // ✅ 修正：如果沒有 deviceId，允許打卡（用員編 + Token）
    if (!targetDevice) {
      if (!token) {
        token = generateToken_();
        sh.getRange(i + 2, 14).setValue(token);
        SpreadsheetApp.flush();
      }
      return { ok: true, token: token };
    }

    // ✅ 如果有 deviceId，檢查是否已綁定
    if (!specialdeviceid) {
      // 首次 NFC 打卡：自動綁定設備
      sh.getRange(i + 2, 13).setValue(targetDevice);
      if (!token) {
        token = generateToken_();
        sh.getRange(i + 2, 14).setValue(token);
      }
      SpreadsheetApp.flush();
      console.log('✅ 首次打卡，自動綁定設備:', targetDevice);
      return { ok: true, token: token };
    }

    // ✅ 如果已綁定，驗證設備是否相符
    if (specialdeviceid !== targetDevice) {
      return { ok: false, message: '裝置不符（已綁定為:' + specialdeviceid + '）' };
    }

    // ✅ 設備相符，回傳 Token
    if (!token) {
      token = generateToken_();
      sh.getRange(i + 2, 14).setValue(token);
      SpreadsheetApp.flush();
    }

    return { ok: true, token: token };
  }

  return { ok: false, message: '查無此員編' };
}

function checkRecentDuplicate_(sh, id, now, limitMinutes) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: false };

  const values = sh.getRange(2, 1, lastRow - 1, 8).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const rowId = clean_(row[2]).toUpperCase();
    if (rowId !== id) continue;

    const dt = row[1];
    const lastAction = clean_(row[3]);
    if (!(dt instanceof Date)) continue;

    const diffMs = now.getTime() - dt.getTime();
    const diffMin = diffMs / 1000 / 60;

    if (diffMin >= 0 && diffMin < limitMinutes) {
      return {
        ok: true,
        lastTime: dt,
        lastTimeText: formatDateTime_(dt),
        lastAction: lastAction
      };
    }

    return { ok: false };
  }

  return { ok: false };
}

function getTodayClockCount_(sh, id, now) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  const start = new Date(now);
  start.setHours(0, 0, 0, 0);

  const end = new Date(now);
  end.setHours(23, 59, 59, 999);

  const values = sh.getRange(2, 1, lastRow - 1, 8).getValues();

  let count = 0;
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowId = clean_(row[2]).toUpperCase();
    const dt = row[1];

    if (rowId !== id) continue;
    if (!(dt instanceof Date)) continue;

    if (dt >= start && dt <= end) count++;
  }

  return count;
}

// ✅ 改進的打卡函數
function handleClock_(data) {
  const id = clean_(data.id).toUpperCase();
  let token = clean_(data.token);
  const source = clean_(data.source) || 'nfc';
  const device = clean_(data.device) || '';
  const note = clean_(data.note) || '';

  if (!id) {
    return jsonOutput_({
      ok: false,
      message: '缺少員編 id'
    });
  }

  // ✅ 自動綁定或驗證設備
  const tokenResult = ensureUserTokenByIdAndDevice_(id, device);
  if (!tokenResult.ok) {
    return jsonOutput_({
      ok: false,
      message: tokenResult.message
    });
  }
  token = tokenResult.token;

  const user = getUserByIdAndToken_(id, token);
  if (!user.ok) {
    return jsonOutput_({
      ok: false,
      message: user.message || '身分驗證失敗'
    });
  }

  if (device && user.specialdeviceid && device !== user.specialdeviceid) {
    return jsonOutput_({
      ok: false,
      message: '裝置不符'
    });
  }

  const now = new Date();
  const ss = openSS_();
  const sh = getOrCreateSheet_(ss, SHEET_CLOCK);

  ensureClockHeaders_(sh);

  const duplicate = checkRecentDuplicate_(sh, id, now, DUP_MINUTES);
  if (duplicate.ok) {
    return jsonOutput_({
      ok: false,
      message: '重複打卡，上一筆打卡時間',
      lastClockTime: duplicate.lastTimeText,
      lastAction: duplicate.lastAction,
      name: user.name,
      nickname: user.nickname,
      id: user.id,
      token: token
    });
  }

  const todayCount = getTodayClockCount_(sh, id, now);
  const actionType = (todayCount % 2 === 0) ? '上班' : '下班';

  const uuid = Utilities.getUuid();
  const timestamp = Date.now();

  const remarkParts = [];
  if (source) remarkParts.push('source=' + source);
  if (device) remarkParts.push('device=' + device);
  if (note) remarkParts.push('note=' + note);

  const remark = remarkParts.join(' | ');

  sh.appendRow([
    user.name,
    now,
    user.id,
    actionType,
    source,
    timestamp,
    uuid,
    remark
  ]);

  return jsonOutput_({
    ok: true,
    message: actionType + '打卡成功',
    actionType: actionType,
    name: user.name,
    nickname: user.nickname,
    id: user.id,
    token: token,
    date: formatDate_(now),
    time: formatTime_(now),
    dateTime: formatDateTime_(now),
    source: source
  });
}

/* =========================
 * 註冊
 * ========================= */
function toBirth6_(value) {
  if (value === null || value === undefined || value === '') return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyMMdd');
  }

  if (typeof value === 'number' && !isNaN(value)) {
    const d = new Date(Math.round((value - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, TZ, 'yyMMdd');
    }
  }

  const s = String(value).trim().replace(/^'+/, '');
  const digits = s.replace(/\D/g, '');

  if (digits.length === 8) return digits.slice(2, 8);
  if (digits.length === 7) return ('0' + digits).slice(-8).slice(2, 8);
  if (digits.length === 6) return digits;

  const m1 = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (m1) {
    const yy = m1[1].slice(-2);
    const mm = ('0' + m1[2]).slice(-2);
    const dd = ('0' + m1[3]).slice(-2);
    return yy + mm + dd;
  }

  return '';
}

function normalizeBirth6Input_(value) {
  const s = clean_(value).replace(/\D/g, '');
  if (s.length === 8) return s.slice(2, 8);
  if (s.length === 6) return s;
  return '';
}

function handleRegister_(data) {
  const id = clean_(data.id).toUpperCase();
  const birth6 = normalizeBirth6Input_(data.birth6);
  const device = clean_(data.device);

  if (!id) {
    return jsonOutput_({ ok: false, message: '缺少員編 id' });
  }

  if (!birth6) {
    return jsonOutput_({ ok: false, message: '缺少 birth6' });
  }

  if (!device) {
    return jsonOutput_({ ok: false, message: '缺少 device' });
  }

  const ss = openSS_();
  const sh = ss.getSheetByName(SHEET_USERS);

  if (!sh) {
    return jsonOutput_({ ok: false, message: '找不到 人員資料 工作表' });
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return jsonOutput_({ ok: false, message: '人員資料沒有任何員工' });
  }

  const values = sh.getRange(2, 1, lastRow - 1, 16).getValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const nickname = clean_(row[0]);
    const name = clean_(row[1]);
    const empId = clean_(row[2]).toUpperCase();
    const role = clean_(row[8]);
    const color = clean_(row[9]);
    const enabled = normalizeEnabled_(row[10]);
    const birthValue = row[5];
    let specialdeviceid = clean_(row[12]);
    let token = clean_(row[13]);

    if (empId !== id) continue;

    if (!enabled) {
      return jsonOutput_({ ok: false, message: '此帳號未啟用' });
    }

    const sheetBirth6 = toBirth6_(birthValue);

    if (!sheetBirth6) {
      return jsonOutput_({
        ok: false,
        message: '人員資料表內的生日格式無法辨識，請檢查 F 欄'
      });
    }

    if (sheetBirth6 !== birth6) {
      return jsonOutput_({
        ok: false,
        message: '生日驗證失敗',
        debug_id: empId,
        debug_birth6_input: birth6,
        debug_birth6_sheet: sheetBirth6
      });
    }

    if (!specialdeviceid) {
      specialdeviceid = device;
      sh.getRange(i + 2, 13).setValue(specialdeviceid);
    } else if (specialdeviceid !== device) {
      return jsonOutput_({
        ok: false,
        message: '此帳號已綁定其他裝置'
      });
    }

    if (!token) {
      token = generateToken_();
      sh.getRange(i + 2, 14).setValue(token);
    }

    SpreadsheetApp.flush();

    return jsonOutput_({
      ok: true,
      message: '註冊成功，請開始感應',
      id: empId,
      name: name || nickname || empId,
      nickname: nickname,
      role: role,
      color: color,
      token: token,
      device: specialdeviceid
    });
  }

  return jsonOutput_({ ok: false, message: '查無此員編' });
}

/* =========================
 * 系統設定 / 權限 / 公用查詢
 * ========================= */
function getSystemSettingsMap_() {
  const sh = getOrCreateSheet_(openSS_(), SHEET_SYSTEM);
  const out = {};

  if (sh.getLastRow() < 2) return out;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  values.forEach(r => {
    out[clean_(r[0])] = r[1];
  });

  return out;
}


function getEmployeeHeaderData() {
  const sh = getOrCreateSheet_(openSS_(), SHEET_SYSTEM);
  return {
    ok: true,
    brandTitleMain: clean_(sh.getRange('B2').getDisplayValue()),
    brandTitleSub: clean_(sh.getRange('B3').getDisplayValue()),
    logoValue: clean_(sh.getRange('B4').getDisplayValue())
  };
}

function saveAttachmentToDrive_(payload, user) {
  const fileData = clean_(payload.fileData || '');
  const fileName = clean_(payload.fileName || 'upload.bin');
  const fileMime = clean_(payload.fileMime || 'application/octet-stream');
  if (!fileData) return '';
  try {
    const bytes = Utilities.base64Decode(fileData);
    const blob = Utilities.newBlob(bytes, fileMime, fileName);
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    const file = folder.createFile(blob);
    file.setDescription('uploaded by ' + (user && user.id ? user.id : 'unknown'));
    return file.getUrl();
  } catch (err) {
    console.error('saveAttachmentToDrive_ error', err);
    return '';
  }
}

/* ============================================================
 * 選休表 寫入（配合實際水平展開結構）
 *
 * 表格結構（第1列=標頭）：
 *   A=員編, B=暱稱,
 *   C=下週標籤, D~J=下週週一~日 (7欄),
 *   K=下月標籤, L~?=下月每天 (最多31欄),
 *   接著=本週標籤, 本週週一~日 (7欄),
 *   接著=本月標籤, 本月每天 (最多31欄),
 *   最後一欄=備註代號
 *
 * employee.html 送來的 payload:
 *   weekOffset: 0=本週, 1=下週
 *   monthOffset: 0=本月, 1=下月 (選月用)
 *   weekdays: [0,1,2,...6] 選中的日期索引（0=週一）
 *   monthDays: [0,1,2,...30] 選中的日期索引（0=第1天）
 *   mode: 'week' 或 'month'
 * ============================================================ */

function writePreselectSchedule_(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;
  var user = auth.user;

  var sh = getOrCreateSheet_(openSS_(), SHEET_PRESELECT);
  var allValues = sh.getDataRange().getValues();

  if (allValues.length < 2) return { ok: false, message: '選休表格式不正確（至少需要標頭列+資料列）' };

  var headerRow = allValues[0]; // 第1列標頭
  var mode = clean_(payload.mode || 'week');
  var weekOffset = Number(payload.weekOffset || 0); // 0=本週, 1=下週
  var monthOffset = Number(payload.monthOffset || 0); // 0=本月, 1=下月
  var weekdays = Array.isArray(payload.weekdays) ? payload.weekdays.map(Number) : [];
  var monthDays = Array.isArray(payload.monthDays) ? payload.monthDays.map(Number) : [];

  // === 找到員工所在的列（第2列開始） ===
  var targetRow = -1;
  for (var r = 1; r < allValues.length; r++) {
    var rowId = clean_(allValues[r][0]).toUpperCase();
    var rowNickname = clean_(allValues[r][1]);
    if (rowId === clean_(user.id).toUpperCase() ||
        rowNickname === clean_(user.nickname) ||
        rowNickname === clean_(user.name)) {
      targetRow = r + 1; // Sheet 的列號（1-based）
      break;
    }
  }

  if (targetRow === -1) {
    return { ok: false, message: '在選休表找不到你的資料列（' + (user.nickname || user.name || user.id) + '）' };
  }

  // === 找到各區塊在標頭列的起始位置 ===
  // 掃描標頭列，找出包含「週」或「月」關鍵字的區塊分隔欄
  var blocks = findPreselectBlocks_(headerRow);

  if (mode === 'week') {
    // 選週：weekOffset=1 → 下週, weekOffset=0 → 本週
    var targetBlock = weekOffset === 1 ? blocks.nextWeek : blocks.thisWeek;

    if (!targetBlock) {
      return { ok: false, message: '找不到' + (weekOffset === 1 ? '下週' : '本週') + '區塊' };
    }

    // 先清除該區塊所有日期（設為 FALSE）
    for (var d = 0; d < 7; d++) {
      var col = targetBlock.dataStartCol + d;
      if (col <= headerRow.length) {
        sh.getRange(targetRow, col).setValue(false);
      }
    }

    // 寫入選中的日期（設為 TRUE）
    weekdays.forEach(function(idx) {
      if (idx >= 0 && idx < 7) {
        var col = targetBlock.dataStartCol + idx;
        if (col <= headerRow.length) {
          sh.getRange(targetRow, col).setValue(true);
        }
      }
    });

  } else if (mode === 'month') {
    // 選月：monthOffset=1 → 下月, monthOffset=0 → 本月
    var targetBlock = monthOffset === 1 ? blocks.nextMonth : blocks.thisMonth;

    if (!targetBlock) {
      return { ok: false, message: '找不到' + (monthOffset === 1 ? '下月' : '本月') + '區塊' };
    }

    // 先清除該區塊所有日期
    for (var d = 0; d < targetBlock.dayCount; d++) {
      var col = targetBlock.dataStartCol + d;
      if (col <= headerRow.length) {
        sh.getRange(targetRow, col).setValue(false);
      }
    }

    // 寫入選中的日期
    monthDays.forEach(function(idx) {
      if (idx >= 0 && idx < targetBlock.dayCount) {
        var col = targetBlock.dataStartCol + idx;
        if (col <= headerRow.length) {
          sh.getRange(targetRow, col).setValue(true);
        }
      }
    });
  }

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: (mode === 'week' ? '週' : '月') + '選休已更新',
    mode: mode,
    weekOffset: weekOffset,
    monthOffset: monthOffset,
    weekdays: weekdays,
    monthDays: monthDays
  };
}

/* ============================================================
 * 掃描標頭列，找出四個區塊的位置
 * 順序（從左到右）：下週 → 下月 → 本週 → 本月
 * ============================================================ */
function findPreselectBlocks_(headerRow) {
  var result = {
    nextWeek: null,   // 下週
    nextMonth: null,  // 下月
    thisWeek: null,   // 本週
    thisMonth: null    // 本月
  };

  // 找出所有「區塊分隔欄」（包含「週」或「月」關鍵字的欄位）
  var separators = [];
  for (var c = 0; c < headerRow.length; c++) {
    var val = clean_(headerRow[c]);
    if (!val) continue;

    // 判斷是否為區塊分隔欄（含「第X週」或「X月份」）
    var isWeekSep = /第\s*\d+\s*週/.test(val) || /週$/.test(val);
    var isMonthSep = /\d+\s*月/.test(val) || /月份$/.test(val);

    if (isWeekSep || isMonthSep) {
      separators.push({
        col: c,                    // 0-based 在 headerRow 中的位置
        sheetCol: c + 1,          // 1-based Sheet 欄號
        type: isWeekSep ? 'week' : 'month',
        label: val
      });
    }
  }

  // 根據你的表格順序：下週 → 下月 → 本週 → 本月
  // 第1個 week = 下週, 第2個 week = 本週
  // 第1個 month = 下月, 第2個 month = 本月
  var weekBlocks = separators.filter(function(s) { return s.type === 'week'; });
  var monthBlocks = separators.filter(function(s) { return s.type === 'month'; });

  // 下週（第1個 week 分隔欄）
  if (weekBlocks.length >= 1) {
    var sep = weekBlocks[0];
    result.nextWeek = {
      separatorCol: sep.sheetCol,
      dataStartCol: sep.sheetCol + 1, // 分隔欄的下一欄開始是資料
      dayCount: 7,
      label: sep.label
    };
  }

  // 本週（第2個 week 分隔欄）
  if (weekBlocks.length >= 2) {
    var sep = weekBlocks[1];
    result.thisWeek = {
      separatorCol: sep.sheetCol,
      dataStartCol: sep.sheetCol + 1,
      dayCount: 7,
      label: sep.label
    };
  }

  // 下月（第1個 month 分隔欄）
  if (monthBlocks.length >= 1) {
    var sep = monthBlocks[0];
    // 下月的天數 = 從下月分隔欄到下一個分隔欄之間的欄數
    var nextSepCol = headerRow.length; // 預設到最後
    for (var i = 0; i < separators.length; i++) {
      if (separators[i].col > sep.col) {
        nextSepCol = separators[i].col;
        break;
      }
    }
    var dayCount = nextSepCol - sep.col - 1;
    result.nextMonth = {
      separatorCol: sep.sheetCol,
      dataStartCol: sep.sheetCol + 1,
      dayCount: Math.min(dayCount, 31),
      label: sep.label
    };
  }

  // 本月（第2個 month 分隔欄）
  if (monthBlocks.length >= 2) {
    var sep = monthBlocks[1];
    var nextSepCol = headerRow.length;
    for (var i = 0; i < separators.length; i++) {
      if (separators[i].col > sep.col) {
        nextSepCol = separators[i].col;
        break;
      }
    }
    var dayCount = nextSepCol - sep.col - 1;
    result.thisMonth = {
      separatorCol: sep.sheetCol,
      dataStartCol: sep.sheetCol + 1,
      dayCount: Math.min(dayCount, 31),
      label: sep.label
    };
  }

  return result;
}

/* ============================================================
 * 讀取選休表（給 employee.html 顯示目前選休狀態）
 * ============================================================ */
function getPreselectByUser_(userId) {
  var sh = getOrCreateSheet_(openSS_(), SHEET_PRESELECT);
  var allValues = sh.getDataRange().getValues();

  if (allValues.length < 2) return { nextWeek: [], nextMonth: [], thisWeek: [], thisMonth: [] };

  var headerRow = allValues[0];
  var blocks = findPreselectBlocks_(headerRow);

  // 找員工列
  var rowData = null;
  for (var r = 1; r < allValues.length; r++) {
    var rowId = clean_(allValues[r][0]).toUpperCase();
    if (rowId === clean_(userId).toUpperCase()) {
      rowData = allValues[r];
      break;
    }
  }

  if (!rowData) return { nextWeek: [], nextMonth: [], thisWeek: [], thisMonth: [] };

  function readBlock(block) {
    if (!block) return [];
    var arr = [];
    for (var d = 0; d < block.dayCount; d++) {
      var colIdx = block.dataStartCol - 1 + d;
      if (colIdx < rowData.length) {
        var val = rowData[colIdx];
        arr.push(val === true || val === 'TRUE' || val === 'true');
      } else {
        arr.push(false);
      }
    }
    return arr;
  }

  function readDateHeaders(block) {
    if (!block) return [];
    var arr = [];
    for (var d = 0; d < block.dayCount; d++) {
      var colIdx = block.dataStartCol - 1 + d;
      if (colIdx < headerRow.length) {
        arr.push(clean_(headerRow[colIdx]));
      }
    }
    return arr;
  }

  return {
    nextWeek: readBlock(blocks.nextWeek),
    nextWeekHeaders: readDateHeaders(blocks.nextWeek),
    nextMonth: readBlock(blocks.nextMonth),
    nextMonthHeaders: readDateHeaders(blocks.nextMonth),
    thisWeek: readBlock(blocks.thisWeek),
    thisWeekHeaders: readDateHeaders(blocks.thisWeek),
    thisMonth: readBlock(blocks.thisMonth),
    thisMonthHeaders: readDateHeaders(blocks.thisMonth)
  };
}

/* ============================================================
 * 管理端：讀取所有員工選休狀態（給 admin.html 用）
 * ============================================================ */
function getPreselectWeekCards_() {
  var sh = getOrCreateSheet_(openSS_(), SHEET_PRESELECT);
  var allValues = sh.getDataRange().getValues();

  if (allValues.length < 2) return [];

  var headerRow = allValues[0];
  var blocks = findPreselectBlocks_(headerRow);
  var out = [];

  for (var r = 1; r < allValues.length; r++) {
    var row = allValues[r];
    var empId = clean_(row[0]).toUpperCase();
    var nickname = clean_(row[1]);
    if (!empId) continue;

    // 下週選休摘要
    var nextWeekPicks = [];
    if (blocks.nextWeek) {
      for (var d = 0; d < 7; d++) {
        var colIdx = blocks.nextWeek.dataStartCol - 1 + d;
        if (colIdx < row.length && (row[colIdx] === true || row[colIdx] === 'TRUE' || row[colIdx] === 'true')) {
          var header = colIdx < headerRow.length ? clean_(headerRow[colIdx]) : ('第' + (d + 1) + '天');
          nextWeekPicks.push(header);
        }
      }
    }

    // 下月選休摘要
    var nextMonthPicks = [];
    if (blocks.nextMonth) {
      for (var d = 0; d < blocks.nextMonth.dayCount; d++) {
        var colIdx = blocks.nextMonth.dataStartCol - 1 + d;
        if (colIdx < row.length && (row[colIdx] === true || row[colIdx] === 'TRUE' || row[colIdx] === 'true')) {
          var header = colIdx < headerRow.length ? clean_(headerRow[colIdx]) : ('第' + (d + 1) + '天');
          nextMonthPicks.push(header);
        }
      }
    }

    var desc = '';
    if (nextWeekPicks.length) desc += '下週休：' + nextWeekPicks.join('、');
    if (nextMonthPicks.length) desc += (desc ? '\n' : '') + '下月休：' + nextMonthPicks.length + '天';

    out.push({
      rowKey: empId,
      applicant: nickname + '｜' + empId,
      title: nickname + ' 的選休',
      desc: desc || '尚未選擇',
      status: (nextWeekPicks.length || nextMonthPicks.length) ? 'pending' : 'draft',
      sheetName: SHEET_PRESELECT
    });
  }

  return out;
}

function setSystemSettings_(map, updater) {
  const sh = getOrCreateSheet_(openSS_(), SHEET_SYSTEM);
  const values = sh.getDataRange().getValues();

  Object.keys(map).forEach(key => {
    let rowIndex = -1;

    for (let i = 1; i < values.length; i++) {
      if (clean_(values[i][0]) === key) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      sh.appendRow([key, map[key], '', new Date(), updater || 'system']);
    } else {
      sh.getRange(rowIndex, 2, 1, 4).setValues([[
        map[key],
        '',
        new Date(),
        updater || 'system'
      ]]);
    }
  });
}

function getPermissionsByRole_(role) {
  const sh = getOrCreateSheet_(openSS_(), SHEET_PERMISSION);
  if (sh.getLastRow() < 2) return [];

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();

  return values
    .filter(r => clean_(r[0]) === clean_(role) && normalizeEnabled_(r[2]))
    .map(r => clean_(r[1]));
}

function isAdminOrCreator_(role) {
  return ['Admin', 'Creator'].indexOf(clean_(role)) > -1;
}

function canReviewSheet_(role, sheetName) {
  const r = clean_(role);
  if (r === 'Creator') return true;
  if (r === 'Manager') return [SHEET_LEAVE, SHEET_UPLOAD, SHEET_MESSAGE].indexOf(sheetName) > -1;
  if (r === 'Admin') return [SHEET_CLOCK_FIX, SHEET_SALARY].indexOf(sheetName) > -1;
  return false;
}

function getRowsByUser_(sheetName, userId, limit) {
  const sh = getOrCreateSheet_(openSS_(), sheetName);
  if (sh.getLastRow() < 2) return [];

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  return values
    .filter(r => clean_(r[1]).toUpperCase() === clean_(userId).toUpperCase())
    .slice(-limit)
    .reverse();
}

function normalizePublishStatus_(v) {
  const s = clean_(v).toLowerCase();
  if (s === 'published' || clean_(v) === '已發布') return 'published';
  if (s === 'scheduled' || clean_(v) === '已排程') return 'scheduled';
  return 'draft';
}

function normalizeReviewStatus_(v) {
  const s = clean_(v).toLowerCase();
  if (s === 'approved' || clean_(v) === '已通過') return 'approved';
  if (s === 'rejected' || clean_(v) === '已退回') return 'rejected';
  return 'pending';
}

/* =========================
 * 員工頁資料
 * ========================= */

function normalizeShiftCode_(v) {
  return clean_(v).toLowerCase().replace(/\s+/g, '');
}

function getShiftRuleMap_() {
  const system = getSystemSettingsMap_();

  return {
    default: {
      start: clean_(system.default_shift_start || system.shift_start || '07:00'),
      end: clean_(system.default_shift_end || system.shift_end || '16:00')
    },
    a: {
      start: clean_(system.a_shift_start || system.default_shift_start || system.shift_start || '07:00'),
      end: clean_(system.a_shift_end || system.default_shift_end || system.shift_end || '16:00')
    },
    b: {
      start: clean_(system.b_shift_start || system.default_shift_start || system.shift_start || '07:00'),
      end: clean_(system.b_shift_end || system.default_shift_end || system.shift_end || '16:00')
    },
    b2: {
      start: clean_(system.b2_shift_start || system.default_shift_start || system.shift_start || '07:00'),
      end: clean_(system.b2_shift_end || system.default_shift_end || system.shift_end || '16:00')
    },
    c: {
      start: clean_(system.c_shift_start || system.default_shift_start || system.shift_start || '07:00'),
      end: clean_(system.c_shift_end || system.default_shift_end || system.shift_end || '16:00')
    }
  };
}

function getShiftForDateByUser_(userId, workDate) {
  const system = getSystemSettingsMap_();
  const user = getUserById_(userId);
  const ymd = clean_(workDate).replace(/\D/g, '');
  const userKey = clean_(userId).toLowerCase();

  const rules = getShiftRuleMap_();

  const dateStartKey = `shift_start_${userKey}_${ymd}`;
  const dateEndKey = `shift_end_${userKey}_${ymd}`;
  const dateCodeKey = `shift_code_${userKey}_${ymd}`;

  if (clean_(system[dateStartKey]) && clean_(system[dateEndKey])) {
    return {
      shiftCode: normalizeShiftCode_(system[dateCodeKey] || 'date'),
      shiftStart: clean_(system[dateStartKey]),
      shiftEnd: clean_(system[dateEndKey]),
      source: 'system_date_override'
    };
  }

  const userStartKey = `${userKey}_shift_start`;
  const userEndKey = `${userKey}_shift_end`;
  const userCodeKey = `${userKey}_shift_code`;

  if (clean_(system[userStartKey]) && clean_(system[userEndKey])) {
    return {
      shiftCode: normalizeShiftCode_(system[userCodeKey] || 'user'),
      shiftStart: clean_(system[userStartKey]),
      shiftEnd: clean_(system[userEndKey]),
      source: 'system_user_override'
    };
  }

  let code = '';
  if (user && user.ok) {
    code = normalizeShiftCode_(user.jobTitle || '');
  }

  if (rules[code] && clean_(rules[code].start) && clean_(rules[code].end)) {
    return {
      shiftCode: code,
      shiftStart: rules[code].start,
      shiftEnd: rules[code].end,
      source: 'shift_code_rule'
    };
  }

  return {
    shiftCode: 'default',
    shiftStart: rules.default.start,
    shiftEnd: rules.default.end,
    source: 'default_rule'
  };
}

function getClockRecordsByUser_(userId, limit) {
  const sh = getOrCreateSheet_(openSS_(), SHEET_CLOCK);
  if (sh.getLastRow() < 2) return [];

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  const map = {};

  values.forEach(r => {
    if (clean_(r[2]).toUpperCase() !== clean_(userId).toUpperCase()) return;

    const dt = r[1] instanceof Date ? r[1] : new Date(r[1]);
    if (isNaN(dt.getTime())) return;

    const date = Utilities.formatDate(dt, TZ, 'yyyy-MM-dd');
    const weekday = ['週日', '週一', '週二', '週三', '週四', '週五', '週六'][dt.getDay()];
    const action = clean_(r[3]);
    const timeText = Utilities.formatDate(dt, TZ, 'HH:mm');

    if (!map[date]) {
      map[date] = {
        date: date,
        weekday: weekday,
        inTime: '',
        outTime: '',
        rawIn: '',
        rawOut: '',
        payIn: '',
        payOut: '',
        shiftCode: '',
        shiftStart: '',
        shiftEnd: '',
        hours: 0,
        lateFlag: false,
        earlyLeaveFlag: false,
        lateMinutes: 0,
        earlyLeaveMinutes: 0,
        checkInNote: '',
        checkOutNote: '',
        note: clean_(r[7])
      };
    }

    if (action === '上班') {
      map[date].inTime = timeText;
      map[date].rawIn = timeText;
    }

    if (action === '下班') {
      map[date].outTime = timeText;
      map[date].rawOut = timeText;
    }

    map[date].note = clean_(r[7]) || map[date].note;
  });

  return Object.keys(map)
    .sort()
    .reverse()
    .slice(0, limit || 60)
    .map(k => {
      const rec = map[k];
      const shift = getShiftForDateByUser_(userId, rec.date);

      rec.shiftCode = shift.shiftCode || 'default';
      rec.shiftStart = shift.shiftStart || '';
      rec.shiftEnd = shift.shiftEnd || '';

      if (rec.rawIn && rec.rawOut && rec.shiftStart && rec.shiftEnd) {
        const att = applyShiftAttendance_(rec.rawIn, rec.rawOut, rec.shiftStart, rec.shiftEnd);

        rec.payIn = att.payIn || '';
        rec.payOut = att.payOut || '';
        rec.hours = Number(att.hours || 0);
        rec.lateFlag = !!att.lateFlag;
        rec.earlyLeaveFlag = !!att.earlyLeaveFlag;
        rec.lateMinutes = Number(att.lateMinutes || 0);
        rec.earlyLeaveMinutes = Number(att.earlyLeaveMinutes || 0);
        rec.checkInNote = att.checkInNote || '';
        rec.checkOutNote = att.checkOutNote || '';
      } else {
        rec.payIn = rec.rawIn || '';
        rec.payOut = rec.rawOut || '';
        rec.hours = calcHours_(rec.rawIn, rec.rawOut);
        rec.lateFlag = false;
        rec.earlyLeaveFlag = false;
        rec.lateMinutes = 0;
        rec.earlyLeaveMinutes = 0;
        rec.checkInNote = '';
        rec.checkOutNote = '';
      }

      return rec;
    })
    .reverse();
}

function getLeaveByUser_(userId, limit) {
  return getRowsByUser_(SHEET_LEAVE, userId, limit).map(r => ({
    id: clean_(r[0]),
    type: clean_(r[3]),
    startDate: formatYmdMaybe_(r[4]),
    endDate: formatYmdMaybe_(r[5]),
    days: Number(r[6] || 0),
    reason: clean_(r[7]),
    status: clean_(r[9]) || '待審核'
  }));
}

function getUploadsByUser_(userId, limit) {
  return getRowsByUser_(SHEET_UPLOAD, userId, limit).map(r => ({
    id: clean_(r[0]),
    type: clean_(r[1]),
    title: clean_(r[4]),
    content: clean_(r[5]),
    attachment: clean_(r[6]),
    status: clean_(r[8]) || '待審核'
  }));
}

function getMessagesByUser_(userId, limit) {
  return getRowsByUser_(SHEET_MESSAGE, userId, limit).map(r => ({
    author: clean_(r[2]) || '我',
    text: clean_(r[3]),
    status: clean_(r[5]) || '待主管審核'
  }));
}

function getClockFixByUser_(userId, limit) {
  return getRowsByUser_(SHEET_CLOCK_FIX, userId, limit).map(r => ({
    id: clean_(r[0]),
    date: clean_(r[3]),
    time: clean_(r[4]),
    action: clean_(r[5]),
    reason: clean_(r[6]),
    status: clean_(r[8]) || '待審核'
  }));
}

function getSalaryByUser_(userId, limit) {
  return getRowsByUser_(SHEET_SALARY, userId, limit).map(r => ({
    key: clean_(r[0]),
    label: clean_(r[3]),
    amount: Number(r[6] || 0),
    issueDate: formatDateTimeMaybe_(r[9]),
    detail: clean_(r[10]),
    hours: Number(r[4] || 0),
    overtime: Number(r[5] || 0),
    status: clean_(r[7]) || ''
  }));
}

function getApprovedNotices_(limit) {
  const sh = getOrCreateSheet_(openSS_(), SHEET_NOTICE);
  if (sh.getLastRow() < 2) return [];

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  return values
    .filter(r => ['published', '已發布', 'Y'].indexOf(clean_(r[4])) > -1)
    .slice(-limit)
    .reverse()
    .map(r => ({
      type: clean_(r[1]),
      title: clean_(r[2]),
      content: clean_(r[3]),
      status: clean_(r[4]),
      publishTime: formatDateTimeMaybe_(r[5])
    }));
}

function getPublishRows_(limit) {
  const sh = getOrCreateSheet_(openSS_(), SHEET_PUBLISH);
  if (sh.getLastRow() < 2) return [];

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  return values.slice(-limit).reverse().map(r => ({
    id: clean_(r[0]),
    type: clean_(r[1]),
    period: clean_(r[2]),
    title: clean_(r[3]),
    desc: clean_(r[4]),
    status: clean_(r[5]),
    publishTime: formatDateTimeMaybe_(r[6]),
    publisher: clean_(r[7])
  }));
}

function getDeadlineMap_() {
  const sh = getOrCreateSheet_(openSS_(), SHEET_DEADLINE);
  const out = {
    weekOpenAt: '',
    weekCloseAt: '',
    weekDeadlineNote: '',
    monthOpenAt: '',
    monthCloseAt: '',
    monthDeadlineNote: ''
  };

  if (sh.getLastRow() < 2) return out;

  const values = sh.getRange(2, 1, Math.min(2, sh.getLastRow() - 1), 7).getValues();
  values.forEach(r => {
    const type = clean_(r[1]);
    if (type === '下週班表') {
      out.weekOpenAt = formatInputDateTime_(r[2]);
      out.weekCloseAt = formatInputDateTime_(r[3]);
      out.weekDeadlineNote = clean_(r[4]);
    }
    if (type === '月班表') {
      out.monthOpenAt = formatInputDateTime_(r[2]);
      out.monthCloseAt = formatInputDateTime_(r[3]);
      out.monthDeadlineNote = clean_(r[4]);
    }
  });

  return out;
}

function parseYmdLocal_(ymd) {
  const s = clean_(ymd);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0, 0);
}

function startOfWeekMonday_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  const day = x.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  x.setDate(x.getDate() + diff);
  return x;
}

function endOfWeekMonday_(d) {
  const x = startOfWeekMonday_(d);
  x.setDate(x.getDate() + 6);
  x.setHours(23, 59, 59, 999);
  return x;
}

function startOfMonth_(d) {
  return new Date(d.getFullYear(), d.getMonth(), 1, 0, 0, 0, 0);
}

function endOfMonth_(d) {
  return new Date(d.getFullYear(), d.getMonth() + 1, 0, 23, 59, 59, 999);
}

function hourlyRateFromSystem_(system) {
  const raw = Number(system && system.hourly_rate || 0);
  return raw > 0 ? raw : 200;
}

function sumHoursBetween_(records, start, end) {
  return records.reduce((sum, rec) => {
    const d = parseYmdLocal_(rec.date);
    if (!d) return sum;
    if (d >= start && d <= end) {
      return sum + Number(rec.hours || 0);
    }
    return sum;
  }, 0);
}

function getRecordByDate_(records, dateObj) {
  const ymd = Utilities.formatDate(dateObj, TZ, 'yyyy-MM-dd');
  for (let i = 0; i < records.length; i++) {
    if (clean_(records[i].date) === ymd) return records[i];
  }
  return null;
}

function getDayStatusText_(rec) {
  if (!rec) return '休息';
  if (rec.inTime && rec.outTime) return '已完成出勤';
  if (rec.inTime && !rec.outTime) return '上班中';
  return '休息';
}

function buildHomeData_(user, system, workRecords, leaveRows, uploads, clockFix, notices, messages) {
  const now = new Date();
  const todayRec = getRecordByDate_(workRecords, now);

  const tomorrow = new Date(now);
  tomorrow.setDate(tomorrow.getDate() + 1);

  const afterTomorrow = new Date(now);
  afterTomorrow.setDate(afterTomorrow.getDate() + 2);

  const tomorrowRec = getRecordByDate_(workRecords, tomorrow);
  const afterTomorrowRec = getRecordByDate_(workRecords, afterTomorrow);

  const weekHours = sumHoursBetween_(workRecords, startOfWeekMonday_(now), endOfWeekMonday_(now));
  const monthHours = sumHoursBetween_(workRecords, startOfMonth_(now), endOfMonth_(now));

  const hourlyRate = hourlyRateFromSystem_(system);
  const weekSalary = roundMoney_(weekHours * hourlyRate);
  const monthSalary = roundMoney_(monthHours * hourlyRate);

  const pendingCount =
    leaveRows.filter(x => clean_(x.status) === '待審核').length +
    uploads.filter(x => clean_(x.status) === '待審核').length +
    clockFix.filter(x => clean_(x.status) === '待審核').length;

  const alertItems = [];
  clockFix.filter(x => clean_(x.status) === '待審核').slice(0, 2).forEach(x => {
    alertItems.push({
      title: (x.date || '') + ' ' + (x.action || '補打卡') + '待審核',
      meta: x.reason || '補打卡申請已送出，等待審核。',
      level: 'warn'
    });
  });

  uploads.filter(x => clean_(x.status) === '待審核').slice(0, 2).forEach(x => {
    alertItems.push({
      title: (x.title || x.type || '資料上傳') + '待審核',
      meta: x.content || '資料已送出，等待審核。',
      level: 'danger'
    });
  });

  const timelineItems = workRecords.slice().reverse().slice(0, 5).map(rec => {
    let title = rec.date + ' ';
    if (rec.inTime && rec.outTime) {
      title += rec.inTime + ' - ' + rec.outTime + ' 已完成出勤';
    } else if (rec.inTime) {
      title += rec.inTime + ' 已完成上班打卡';
    } else {
      title += '無完整出勤紀錄';
    }

    let meta = rec.hours ? ('工時 ' + rec.hours + 'h') : '無工時';
    if (rec.note) meta += '｜' + rec.note;

    return {
      title: title,
      meta: meta
    };
  });

  const noticeItems = (notices || []).slice(0, 5).map(n => ({
    title: n.title || n.type || '主管通知',
    meta: n.content || ''
  }));

  return {
    todayStatus: getDayStatusText_(todayRec),
    nextShift: '明天 ' + getDayStatusText_(tomorrowRec),
    weekSalary: weekSalary,
    weekHours: Math.round(weekHours * 100) / 100,
    monthSalary: monthSalary,
    monthHours: Math.round(monthHours * 100) / 100,
    tomorrowStatus: getDayStatusText_(tomorrowRec),
    afterTomorrowStatus: getDayStatusText_(afterTomorrowRec),
    pendingCount: pendingCount,
    alerts: alertItems,
    notices: noticeItems,
    timeline: timelineItems,
    latestMessages: (messages || []).slice(0, 10)
  };
}

// ✅ 員工資料初始化
function getEmployeeBootstrapData(id, token) {
  console.log('📨 getEmployeeBootstrapData 收到請求:', { id });
  
  const auth = authUser_(id, token);
  if (!auth.ok) {
    console.log('❌ 身分驗證失敗:', auth.message);
    return auth;
  }

  const user = auth.user;
  console.log('✅ 身分驗證成功:', user.id);

  const system = getSystemSettingsMap_();
  const workRecords = getClockRecordsByUser_(user.id, 60);
  const leaveRows = getLeaveByUser_(user.id, 30);
  const uploads = getUploadsByUser_(user.id, 20);
  const clockFix = getClockFixByUser_(user.id, 20);
  const salaryHistory = getSalaryByUser_(user.id, 20);
  const notices = getApprovedNotices_(5);
  const messages = getMessagesByUser_(user.id, 20);
  const publishRows = getPublishRows_(12);
  const home = buildHomeData_(user, system, workRecords, leaveRows, uploads, clockFix, notices, messages);

  return {
    ok: true,
    profile: {
      id: user.id,
      name: user.name,
      nickname: user.nickname,
      role: user.role || '員工',
      color: user.color || '#ff87e0',
      chip: user.role || '標準版'
    },
    system: system,
    home: home,
    notices: notices,
    messages: messages,
    workRecords: workRecords,
    leaveRows: leaveRows,
    uploads: uploads,
    clockFix: clockFix,
    salaryHistory: salaryHistory,
    publishRows: publishRows,
    deadlineMap: getDeadlineMap_(),
    adminSalaryConfig: {
      mode: 'month',
      weekStart: '週一',
      monthStart: '1號'
    }
  };
}

function submitEmployeeLeave(payload) {
  const auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;

  const user = auth.user;
  const sh = getOrCreateSheet_(openSS_(), SHEET_LEAVE);
  const items = Array.isArray(payload.items) ? payload.items : [];

  if (!items.length) return { ok: false, message: '沒有請假項目' };

  const startDate = clean_(payload.startDate);
  const reason = clean_(payload.reason);
  const now = new Date();

  items.forEach((item, idx) => {
    const reqId = 'LV' + Utilities.formatDate(now, TZ, 'yyyyMMddHHmmss') + String(idx + 1).padStart(2, '0');
    sh.appendRow([
      reqId,
      user.id,
      user.name,
      clean_(item.raw || item.display),
      startDate,
      startDate,
      Number(item.days || 0),
      reason,
      now,
      '待審核',
      '',
      '',
      '',
      '',
      ''
    ]);
  });

  return { ok: true, message: '請假申請已送出' };
}

function submitEmployeeClockFix(payload) {
  const auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;

  const user = auth.user;
  const sh = getOrCreateSheet_(openSS_(), SHEET_CLOCK_FIX);
  const reqId = 'CF' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss');
  const dt = clean_(payload.timeVal);
  const date = dt ? dt.slice(0, 10) : '';
  const time = dt ? dt.slice(11, 16) : '';

  sh.appendRow([
    reqId,
    user.id,
    user.name,
    date,
    time,
    clean_(payload.type),
    clean_(payload.note),
    new Date(),
    '待審核',
    '',
    '',
    ''
  ]);

  return { ok: true, message: '補打卡申請已送出' };
}

function submitEmployeeUpload(payload) {
  const auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;

  const user = auth.user;
  const sh = getOrCreateSheet_(openSS_(), SHEET_UPLOAD);
  const reqId = 'UP' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss');
  const attachmentUrl = saveAttachmentToDrive_(payload, user) || clean_(payload.attachment || '');

  sh.appendRow([
    reqId,
    clean_(payload.type),
    user.id,
    user.name,
    clean_(payload.title || payload.type),
    clean_(payload.note),
    attachmentUrl,
    new Date(),
    '待審核',
    '',
    '',
    ''
  ]);

  return { ok: true, message: '資料已送出', attachmentUrl: attachmentUrl };
}

function submitEmployeeMessage(payload) {
  const auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;

  const user = auth.user;
  const text = clean_(payload.text);
  if (!text) return { ok: false, message: '留言不可空白' };

  const sh = getOrCreateSheet_(openSS_(), SHEET_MESSAGE);
  const reqId = 'MS' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss');

  sh.appendRow([
    reqId,
    user.id,
    user.name,
    text,
    new Date(),
    '待審核',
    '',
    '',
    ''
  ]);

  return { ok: true, message: '留言已送出，待主管審核' };
}

/* =========================
 * 管理頁資料
 * ========================= */
function getNoticePublishItems_(limit) {
  const sh = getOrCreateSheet_(openSS_(), SHEET_NOTICE);
  if (sh.getLastRow() < 2) return [];

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  return values.slice(-limit).reverse().map(r => ({
    id: clean_(r[0]),
    title: clean_(r[2]),
    desc: clean_(r[3]),
    status: normalizePublishStatus_(r[4]),
    publisher: clean_(r[6])
  }));
}

function getPublishByType_(type, limit) {
  return getPublishRows_(limit)
    .filter(r => !type || r.type === type)
    .map(r => ({
      id: r.id,
      title: r.title,
      desc: r.desc,
      status: normalizePublishStatus_(r.status),
      period: r.period
    }));
}

function getReviewCards_(sheetName, includeHeaders) {
  const sh = getOrCreateSheet_(openSS_(), sheetName);
  if (sh.getLastRow() < 2) return [];

  const values = sh.getDataRange().getValues();
  const headers = values[0];

  return values.slice(1).reverse().map(r => {
    const obj = {};
    headers.forEach((h, idx) => {
      obj[clean_(h)] = r[idx];
    });

    const applicant = `${clean_(obj['姓名']) || ''}｜${clean_(obj['員編']) || ''}`.replace(/^｜|｜$/g, '');
    const parts = (includeHeaders || [])
      .map(h => `${h}：${formatDateTimeMaybe_(obj[h])}`)
      .filter(x => !x.endsWith('：'));

    return {
      rowKey: clean_(r[0]),
      applicant: applicant,
      title: clean_(r[0]) || clean_(obj['標題']) || sheetName,
      desc: parts.join('\n'),
      status: normalizeReviewStatus_(obj['狀態']),
      sheetName: sheetName
    };
  });
}

// ✅ 管理資料初始化

function getPreselectWeekCards_() {
  const sh = getOrCreateSheet_(openSS_(), SHEET_PRESELECT);
  if (sh.getLastRow() < 3) return [];
  const values = sh.getDataRange().getValues();
  const out = [];
  ['當週(備用)','下週(主要)'].forEach(function(sectionLabel){
    let sectionRow = -1;
    for (let i = 0; i < values.length; i++) { if (clean_(values[i][0]) === sectionLabel) { sectionRow = i; break; } }
    if (sectionRow < 0) return;
    const dateRow = values[sectionRow + 1] || [];
    for (let r = sectionRow + 2; r < values.length; r++) {
      const row = values[r];
      const nm = clean_(row[0]);
      if (!nm) break;
      if (nm === '當週(備用)' || nm === '下週(主要)') break;
      const picks = [];
      for (let c = 1; c <= 7; c++) { if (row[c] === true) picks.push(clean_(dateRow[c]) || ('週' + c)); }
      out.push({ rowKey: sectionLabel + '_' + nm, applicant: nm, title: sectionLabel, desc: picks.join('、') || '未選擇', status: 'pending', sheetName: SHEET_PRESELECT });
    }
  });
  return out;
}

function getAdminBootstrapData(id, token) {
  console.log('📨 getAdminBootstrapData 收到請求:', { id });
  
  const auth = authUser_(id, token);
  if (!auth.ok) {
    console.log('❌ 身分驗證失敗:', auth.message);
    return auth;
  }

  const user = auth.user;
  
  // ✅ 只允許 Creator, Admin, Manager 進入後台
  const allowedRoles = ['Creator', 'Admin', 'Manager'];
  if (allowedRoles.indexOf(user.role) === -1) {
    console.log('❌ 無管理權限:', { id, role: user.role });
    return { ok: false, message: '無管理權限，僅限 Manager / Admin / Creator' };
  }

  console.log('✅ 身分驗證成功，角色:', user.role);

  const system = getSystemSettingsMap_();

  return {
    ok: true,
    profile: {
      id: user.id,
      name: user.name,
      nickname: user.nickname,
      role: user.role || 'Employee',
      color: user.color || ''
    },
    system: system,
    permissions: getPermissionsByRole_(user.role || 'Employee'),
    deadlineMap: getDeadlineMap_(),
    schedulePublishItems: getPublishByType_('排班發布', 20),
    noticePublishItems: getNoticePublishItems_(20),
    leaveReviews: getReviewCards_(SHEET_LEAVE, ['假別', '開始日期', '天數', '事由']),
    receiveReviews: getReviewCards_(SHEET_UPLOAD, ['類型', '標題', '內容']),
    messageReviews: getReviewCards_(SHEET_MESSAGE, ['留言內容']),
    clockFixReviews: getReviewCards_(SHEET_CLOCK_FIX, ['補打卡日期', '補打卡時間', '補打卡動作', '申請事由']),
    salaryReviews: getReviewCards_(SHEET_SALARY, ['月份', '工時', '加班時數', '應發金額', '備註']),
    preselectWeekCards: getPreselectWeekCards_()
  };
}

function saveDeadlineSettings(payload) {
  const auth = authUser_(payload.userId, payload.token || payload.userToken);
  if (!auth.ok) return auth;

  const user = auth.user;
  if (!isAdminOrCreator_(user.role)) return { ok: false, message: '無權限' };

  const sh = getOrCreateSheet_(openSS_(), SHEET_DEADLINE);
  ensureSheetHasRows_(sh, 3);

  sh.getRange(2, 3, 1, 5).setValues([[
    payload.weekOpenAt || '',
    payload.weekCloseAt || '',
    payload.weekDeadlineNote || '',
    new Date(),
    user.id
  ]]);

  sh.getRange(3, 3, 1, 5).setValues([[
    payload.monthOpenAt || '',
    payload.monthCloseAt || '',
    payload.monthDeadlineNote || '',
    new Date(),
    user.id
  ]]);

  return { ok: true, message: '期限設定已儲存' };
}

function saveBrandSettings(payload) {
  const auth = authUser_(payload.userId, payload.token || payload.userToken);
  if (!auth.ok) return auth;

  const user = auth.user;
  if (!isAdminOrCreator_(user.role)) return { ok: false, message: '無權限' };

  setSystemSettings_({
    main_title: payload.mainTitle || 'ANG.lo Engine',
    system_name: payload.systemName || 'HR 系統',
    fallback_text: payload.fallbackText || 'ANG',
    logo_url: payload.logoDataUrl || ''
  }, user.id);

  return { ok: true, message: 'Logo & Title 設定已儲存' };
}

function adminSetReviewStatus(payload) {
  const auth = authUser_(payload.userId, payload.token || payload.userToken);
  if (!auth.ok) return auth;

  const user = auth.user;
  const sheetName = clean_(payload.sheetName);
  const rowKey = clean_(payload.rowKey);
  const status = clean_(payload.status);

  if (!sheetName || !rowKey || !status) return { ok: false, message: '缺少必要欄位' };
  if (!canReviewSheet_(user.role, sheetName)) return { ok: false, message: '無權限' };

  const sh = getOrCreateSheet_(openSS_(), sheetName);
  const values = sh.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (clean_(values[i][0]) === rowKey) {
      const headers = values[0];
      const statusCol = headers.indexOf('狀態') + 1;
      const reviewerCol = headers.indexOf('審核者') + 1;
      const reviewTimeCol = headers.indexOf('審核時間') + 1;

      if (statusCol) sh.getRange(i + 1, statusCol).setValue(status);
      if (reviewerCol) sh.getRange(i + 1, reviewerCol).setValue(user.id);
      if (reviewTimeCol) sh.getRange(i + 1, reviewTimeCol).setValue(new Date());

      return { ok: true, message: '已更新狀態' };
    }
  }

  return { ok: false, message: '找不到資料' };
}

/* =========================
 * GET / POST 入口
 * ========================= */
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};

    const pageParam = clean_(params.page || params.PAGE || '').toLowerCase();
    const idParam = clean_(params.id || params.ID || '').toUpperCase();
    const tokenParam = clean_(params.token || params.TOKEN || '');

    console.log('📨 收到 GET 請求:', { page: pageParam, id: idParam });

    const templateMap = {
      employee: 'employee',
      admin: 'admin',
      clock: 'clock',
      mysalary: 'mysalary',
      index: 'index',
      upload: 'upload',
      settle: 'settle',
      creator: 'creator',
      boss: 'boss'
    };

    const page = templateMap[pageParam] ? pageParam : 'employee';
    let user = null;

    // ✅ 驗證身分
    if (page !== 'boss' && idParam && tokenParam) {
      const authUser = getUserByIdAndToken_(idParam, tokenParam);
      if (authUser && authUser.ok) {
        user = {
          id: authUser.id || '',
          token: authUser.token || '',
          page: page,
          role: authUser.role || '',
          color: authUser.color || '',
          name: authUser.name || '',
          nickname: authUser.nickname || '',
          isGuest: false
        };
      }
    }

    // ✅ 驗證失敗則為訪客
    if (!user) {
      user = {
        id: idParam || '',
        token: tokenParam || '',
        page: page,
        role: page === 'boss' ? 'public' : 'guest',
        color: '#9e9e9e',
        name: page === 'boss' ? '公開看板' : '訪客',
        nickname: page === 'boss' ? '公開看板' : '訪客',
        isGuest: page === 'boss' ? false : true
      };
    }

    const t = HtmlService.createTemplateFromFile(templateMap[page]);
    t.id = user.id;
    t.token = user.token;
    t.page = user.page;
    t.key = '';
    t.role = user.role;
    t.color = user.color;
    t.name = user.name;
    t.nickname = user.nickname;
    t.isGuest = user.isGuest;
    t.webAppUrl = ScriptApp.getService().getUrl();

    return t.evaluate()
      .setTitle('ANG HR Engine')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0, user-scalable=no, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    console.error('❌ doGet 錯誤:', err);
    return HtmlService.createHtmlOutput(
      '<h2>doGet error</h2><pre>' + String(err && err.stack ? err.stack : err) + '</pre>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ✅ 完整的 doPost 路由
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action ? String(data.action).trim() : '';

    // === 員工功能 ===
    if (action === 'employeeBootstrap') return jsonOutput_(getEmployeeBootstrapData(data.id, data.token));
    if (action === 'employeeHeaderData') return jsonOutput_(getEmployeeHeaderData());
    if (action === 'employeePreselect') return jsonOutput_(writePreselectSchedule_(data));
    if (action === 'employeeLeave') return jsonOutput_(submitEmployeeLeave(data));
    if (action === 'employeeClockFix') return jsonOutput_(submitEmployeeClockFix(data));
    if (action === 'employeeUpload') return jsonOutput_(submitEmployeeUpload(data));
    if (action === 'employeeMessage') return jsonOutput_(submitEmployeeMessage(data));

    // === 分層管理端 ===
    if (action === 'getManagerBootstrapData') return jsonOutput_(getManagerBootstrapData(data.id || '', data.token || ''));
    if (action === 'getAdminBootstrapData') return jsonOutput_(getAdminBootstrapData(data.id || '', data.token || ''));
    if (action === 'getCreatorBootstrapData') return jsonOutput_(getCreatorBootstrapData(data.id || '', data.token || ''));

    // === 審核 ===
    if (action === 'adminSetReviewStatus') return jsonOutput_(adminSetReviewStatus(data));

    // === Creator 設定 ===
    if (action === 'saveApproverSettings') return jsonOutput_(saveApproverSettings(data));
    if (action === 'saveDeadlineSettings') return jsonOutput_(saveDeadlineSettings(data));
    if (action === 'saveBrandSettings') return jsonOutput_(saveBrandSettings(data));

    if (action === 'generateSalaryDraft') return jsonOutput_(generateSalaryDraft(data));
    if (action === 'saveSalaryReview') return jsonOutput_(saveSalaryReview(data));
    if (action === 'downloadSalarySlip') return jsonOutput_(downloadSalarySlip(data));
    if (action === 'savePreselectSettings') return jsonOutput_(savePreselectSettings(data));
    if (action === 'isPreselectOpen') return jsonOutput_(isPreselectOpen(data));
    if (action === 'submitPreselect') return jsonOutput_(submitPreselect(data));
    if (action === 'getMyPreselect') return jsonOutput_(getMyPreselect(data));
    if (action === 'getTodayStatus') return jsonOutput_(getTodayStatus(data));
    if (action === 'getRecentActivities') return jsonOutput_(getRecentActivities(data));
    if (action === 'getNoticesForEmployee') return jsonOutput_(getNoticesForEmployee(data));
    if (action === 'getSettingsHash') return jsonOutput_(getSettingsHash(data));
    if (action === 'getShiftTypes') return jsonOutput_(getShiftTypes(data));
    if (action === 'saveShiftTypes') return jsonOutput_(saveShiftTypes(data));
    
    // === 歸檔 ===
    if (action === 'archiveOldRecords') return jsonOutput_(archiveOldRecords(data));
    if (action === 'exportArchivedToDrive') return jsonOutput_(exportArchivedToDrive(data));
    
    if (action === 'generateSalaryDraft') return jsonOutput_(generateSalaryDraft(data));
    if (action === 'saveSalaryReview') return jsonOutput_(saveSalaryReview(data));
    if (action === 'downloadSalarySlip') return jsonOutput_(downloadSalarySlip(data));
    if (action === 'savePreselectSettings') return jsonOutput_(savePreselectSettings(data));

    // === 打卡 / 註冊 ===
    if (action === 'handleClock' || action === 'clock') return handleClock_(data);
    if (action === 'handleRegister' || action === 'register') return handleRegister_(data);

    return jsonOutput_({ ok: false, message: '未知 action: ' + action });

  } catch (err) {
    return jsonOutput_({ ok: false, message: '後端錯誤：' + err.toString() });
  }
}

/* ============================================================
 * 多級審核系統 - 追加功能
 * 貼在現有程式碼.js 最下方，不要刪除原有程式碼
 * ============================================================ */

// ========== 1. 取得審核階層設定 ==========

function getApproversForType_(type) {
  var sys = getSystemSettingsMap_();
  var key = type + '_approvers';
  var raw = String(sys[key] || 'Manager').replace(/\s+/g, '');
  return raw.split(',').filter(function(x) { return x.length > 0; });
}

function isMultiLevelEnabled_() {
  var sys = getSystemSettingsMap_();
  var val = String(sys['multi_level_review_enabled'] || 'true').toLowerCase().trim();
  return val !== 'false' && val !== '0' && val !== 'no';
}

// ========== 2. 覆寫送出函數（加入審核流程欄位） ==========

// --- 請假申請 ---
function submitEmployeeLeave(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;

  var user = auth.user;
  var sh = getOrCreateSheet_(openSS_(), SHEET_LEAVE);
  var items = Array.isArray(payload.items) ? payload.items : [];
  if (!items.length) return { ok: false, message: '沒有請假項目' };

  var approvers = isMultiLevelEnabled_() ? getApproversForType_('leave') : ['Manager'];
  var startDate = clean_(payload.startDate);
  var endDate = clean_(payload.endDate || payload.startDate);
  var reason = clean_(payload.reason);
  var now = new Date();

  items.forEach(function(item, idx) {
    var reqId = 'LV' + Utilities.formatDate(now, TZ, 'yyyyMMddHHmmss') + String(idx + 1).padStart(2, '0');
    sh.appendRow([
      reqId,
      user.id,
      user.name,
      clean_(item.raw || item.display),
      startDate,
      endDate,
      Number(item.days || 0),
      reason,
      now,
      '待審核',
      '', '', '', '', '',
      approvers.join(','),
      approvers[0] || '',
      '[]'
    ]);
  });

  return { ok: true, message: '請假申請已送出' };
}

// --- 補打卡申請 ---
function submitEmployeeClockFix(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;

  var user = auth.user;
  var sh = getOrCreateSheet_(openSS_(), SHEET_CLOCK_FIX);
  var approvers = isMultiLevelEnabled_() ? getApproversForType_('clockfix') : ['Manager'];
  var reqId = 'CF' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss');
  var dt = clean_(payload.timeVal);
  var date = dt ? dt.slice(0, 10) : '';
  var time = dt ? dt.slice(11, 16) : '';

  sh.appendRow([
    reqId,
    user.id,
    user.name,
    date,
    time,
    clean_(payload.type),
    clean_(payload.note),
    new Date(),
    '待審核',
    '', '', '',
    approvers.join(','),
    approvers[0] || '',
    '[]'
  ]);

  return { ok: true, message: '補打卡申請已送出' };
}

// --- 資料上傳 ---
function submitEmployeeUpload(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;

  var user = auth.user;
  var sh = getOrCreateSheet_(openSS_(), SHEET_UPLOAD);
  var approvers = isMultiLevelEnabled_() ? getApproversForType_('upload') : ['Manager'];
  var reqId = 'UP' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss');
  var attachmentUrl = saveAttachmentToDrive_(payload, user) || clean_(payload.attachment || '');

  sh.appendRow([
    reqId,
    clean_(payload.type),
    user.id,
    user.name,
    clean_(payload.title || payload.type),
    clean_(payload.note),
    attachmentUrl,
    new Date(),
    '待審核',
    '', '', '',
    approvers.join(','),
    approvers[0] || '',
    '[]'
  ]);

  return { ok: true, message: '資料已送出', attachmentUrl: attachmentUrl };
}

// --- 留言 ---
function submitEmployeeMessage(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;

  var user = auth.user;
  var text = clean_(payload.text);
  if (!text) return { ok: false, message: '留言不可空白' };

  var sh = getOrCreateSheet_(openSS_(), SHEET_MESSAGE);
  var approvers = isMultiLevelEnabled_() ? getApproversForType_('message') : ['Manager'];
  var reqId = 'MS' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss');

  sh.appendRow([
    reqId,
    user.id,
    user.name,
    text,
    new Date(),
    '待審核',
    '', '', '',
    approvers.join(','),
    approvers[0] || '',
    '[]'
  ]);

  return { ok: true, message: '留言已送出，待審核' };
}

// ========== 3. 多級審核核心邏輯（覆寫） ==========

function adminSetReviewStatus(payload) {
  var auth = authUser_(payload.userId, payload.token || payload.userToken);
  if (!auth.ok) return auth;

  var user = auth.user;
  var sheetName = clean_(payload.sheetName);
  var rowKey = clean_(payload.rowKey);
  var status = clean_(payload.status);
  var reviewNote = clean_(payload.reviewNote || '');

  if (!sheetName || !rowKey || !status) return { ok: false, message: '缺少必要欄位' };
  if (!canReviewSheet_(user.role, sheetName)) return { ok: false, message: '無權限' };

  var sh = getOrCreateSheet_(openSS_(), sheetName);
  var values = sh.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) {
    if (clean_(values[i][0]) !== rowKey) continue;

    var headers = values[0];
    var statusCol = headers.indexOf('狀態') + 1;
    var reviewerCol = headers.indexOf('審核者') + 1;
    var reviewTimeCol = headers.indexOf('審核時間') + 1;
    var reviewNoteCol = headers.indexOf('審核備註') + 1;
    var firstReviewerCol = headers.indexOf('初審者') + 1;
    var firstReviewTimeCol = headers.indexOf('初審時間') + 1;
    var finalReviewerCol = headers.indexOf('最終審核者') + 1;
    var finalReviewTimeCol = headers.indexOf('最終審核時間') + 1;
    var requiredCol = headers.indexOf('requiredApprovers') + 1;
    var currentCol = headers.indexOf('currentStep') + 1;
    var approveCol = headers.indexOf('approvedList') + 1;

    var requiredApprovers = [];
    var approvedList = [];

    if (requiredCol > 0) {
      requiredApprovers = clean_(sh.getRange(i + 1, requiredCol).getValue())
        .replace(/\s+/g, '').split(',').filter(function(x) { return x.length > 0; });
    }

    if (approveCol > 0) {
      try {
        approvedList = JSON.parse(sh.getRange(i + 1, approveCol).getValue() || '[]');
      } catch (e) { approvedList = []; }
    }

    if (!Array.isArray(approvedList)) approvedList = [];

    var now = new Date();

    if (status === 'approved' || status === '已通過') {

      approvedList.push({
        role: user.role,
        user: user.id,
        name: user.name,
        date: formatDateTime_(now)
      });

      if (approveCol > 0) {
        sh.getRange(i + 1, approveCol).setValue(JSON.stringify(approvedList));
      }

      // 判斷是否有多級流程
      if (requiredApprovers.length > 0 && isMultiLevelEnabled_()) {
        var currentStepRole = currentCol > 0 ? clean_(sh.getRange(i + 1, currentCol).getValue()) : '';
        var stepIndex = requiredApprovers.indexOf(currentStepRole);
        if (stepIndex < 0) stepIndex = requiredApprovers.indexOf(user.role);

        var nextIdx = stepIndex + 1;

        if (nextIdx >= requiredApprovers.length) {
          // 最後一層 → 全部通過
          if (statusCol > 0) sh.getRange(i + 1, statusCol).setValue('已通過');
          if (currentCol > 0) sh.getRange(i + 1, currentCol).setValue('');
          if (finalReviewerCol > 0) sh.getRange(i + 1, finalReviewerCol).setValue(user.id);
          if (finalReviewTimeCol > 0) sh.getRange(i + 1, finalReviewTimeCol).setValue(now);
        } else {
          // 還有下一層
          if (currentCol > 0) sh.getRange(i + 1, currentCol).setValue(requiredApprovers[nextIdx]);
          if (statusCol > 0) sh.getRange(i + 1, statusCol).setValue('待審核');
        }

        // 初審紀錄
        if (stepIndex === 0) {
          if (firstReviewerCol > 0) sh.getRange(i + 1, firstReviewerCol).setValue(user.id);
          if (firstReviewTimeCol > 0) sh.getRange(i + 1, firstReviewTimeCol).setValue(now);
        }

      } else {
        // 單層審核 → 直接通過
        if (statusCol > 0) sh.getRange(i + 1, statusCol).setValue('已通過');
        if (currentCol > 0) sh.getRange(i + 1, currentCol).setValue('');
        if (finalReviewerCol > 0) sh.getRange(i + 1, finalReviewerCol).setValue(user.id);
        if (finalReviewTimeCol > 0) sh.getRange(i + 1, finalReviewTimeCol).setValue(now);
      }

    } else if (status === 'rejected' || status === '已退回') {

      approvedList.push({
        role: user.role,
        user: user.id,
        name: user.name,
        date: formatDateTime_(now),
        rejected: true
      });

      if (statusCol > 0) sh.getRange(i + 1, statusCol).setValue('已退回');
      if (currentCol > 0) sh.getRange(i + 1, currentCol).setValue('');
      if (approveCol > 0) sh.getRange(i + 1, approveCol).setValue(JSON.stringify(approvedList));
    }

    // 通用審核欄位
    if (reviewerCol > 0) sh.getRange(i + 1, reviewerCol).setValue(user.id);
    if (reviewTimeCol > 0) sh.getRange(i + 1, reviewTimeCol).setValue(now);
    if (reviewNoteCol > 0 && reviewNote) sh.getRange(i + 1, reviewNoteCol).setValue(reviewNote);

    return { ok: true, message: '審核狀態已更新' };
  }

  return { ok: false, message: '找不到資料' };
}

// ========== 4. 審核權限擴展 ==========

function canReviewSheet_(role, sheetName) {
  var r = clean_(role);
  if (r === 'Creator') return true;
  if (r === 'Admin') {
    return [SHEET_LEAVE, SHEET_UPLOAD, SHEET_MESSAGE, SHEET_CLOCK_FIX, SHEET_SALARY, SHEET_NOTICE].indexOf(sheetName) > -1;
  }
  if (r === 'Manager') {
    return [SHEET_LEAVE, SHEET_UPLOAD, SHEET_MESSAGE, SHEET_NOTICE].indexOf(sheetName) > -1;
  }
  return false;
}

// ========== 5. 審核清單查詢（按角色過濾 currentStep） ==========

function getReviewCardsForRole_(sheetName, role, includeHeaders) {
  var sh = getOrCreateSheet_(openSS_(), sheetName);
  if (sh.getLastRow() < 2) return [];

  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var currentStepIdx = headers.indexOf('currentStep');
  var statusIdx = headers.indexOf('狀態');

  return values.slice(1).reverse().map(function(r) {
    var obj = {};
    headers.forEach(function(h, idx) {
      obj[clean_(h)] = r[idx];
    });

    var applicant = (clean_(obj['姓名']) || '') + '｜' + (clean_(obj['員編']) || '');
    applicant = applicant.replace(/^｜|｜$/g, '');

    var parts = (includeHeaders || [])
      .map(function(h) { return h + '：' + formatDateTimeMaybe_(obj[h]); })
      .filter(function(x) { return !x.endsWith('：'); });

    var currentStep = currentStepIdx >= 0 ? clean_(r[currentStepIdx]) : '';
    var rowStatus = statusIdx >= 0 ? clean_(r[statusIdx]) : '';

    return {
      rowKey: clean_(r[0]),
      applicant: applicant,
      title: clean_(r[0]) || clean_(obj['標題']) || sheetName,
      desc: parts.join('\n'),
      status: normalizeReviewStatus_(rowStatus),
      currentStep: currentStep,
      sheetName: sheetName
    };
  }).filter(function(item) {
    // 只顯示目前該角色需要審核的 + 已結案的
    if (item.status === 'pending' && item.currentStep && item.currentStep !== role) {
      // 待審核但不是輪到這個角色 → 不顯示
      // 除非角色是 Creator（Creator 看全部）
      if (role === 'Creator') return true;
      return false;
    }
    return true;
  });
}

// ========== 6. 管理端資料（分角色） ==========

function getManagerBootstrapData(id, token) {
  var auth = authUser_(id, token);
  if (!auth.ok) return auth;

  var user = auth.user;
  if (['Creator', 'Admin', 'Manager'].indexOf(user.role) === -1) {
    return { ok: false, message: '無管理權限' };
  }

  var system = getSystemSettingsMap_();

  return {
    ok: true,
    profile: {
      id: user.id,
      name: user.name,
      nickname: user.nickname,
      role: user.role,
      color: user.color
    },
    system: system,
    leaveReviews: getReviewCardsForRole_(SHEET_LEAVE, user.role, ['假別', '開始日期', '天數', '事由']),
    receiveReviews: getReviewCardsForRole_(SHEET_UPLOAD, user.role, ['類型', '標題', '內容']),
    messageReviews: getReviewCardsForRole_(SHEET_MESSAGE, user.role, ['留言內容']),
    noticePublishItems: getNoticePublishItems_(20)
  };
}

function getAdminBootstrapData(id, token) {
  var auth = authUser_(id, token);
  if (!auth.ok) return auth;

  var user = auth.user;
  if (['Creator', 'Admin'].indexOf(user.role) === -1) {
    return { ok: false, message: '無管理權限，僅限 Admin / Creator' };
  }

  var system = getSystemSettingsMap_();

  return {
    ok: true,
    profile: {
      id: user.id,
      name: user.name,
      nickname: user.nickname,
      role: user.role,
      color: user.color
    },
    system: system,
    permissions: getPermissionsByRole_(user.role),
    deadlineMap: getDeadlineMap_(),
    leaveReviews: getReviewCardsForRole_(SHEET_LEAVE, user.role, ['假別', '開始日期', '天數', '事由']),
    receiveReviews: getReviewCardsForRole_(SHEET_UPLOAD, user.role, ['類型', '標題', '內容']),
    messageReviews: getReviewCardsForRole_(SHEET_MESSAGE, user.role, ['留言內容']),
    clockFixReviews: getReviewCardsForRole_(SHEET_CLOCK_FIX, user.role, ['補打卡日期', '補打卡時間', '補打卡動作', '申請事由']),
    salaryReviews: getReviewCardsForRole_(SHEET_SALARY, user.role, ['月份', '工時', '加班時數', '應發金額', '備註']),
    schedulePublishItems: getPublishByType_('排班發布', 20),
    noticePublishItems: getNoticePublishItems_(20),
    preselectWeekCards: getPreselectWeekCards_()
  };
}

function getCreatorBootstrapData(id, token) {
  var auth = authUser_(id, token);
  if (!auth.ok) return auth;

  var user = auth.user;
  if (user.role !== 'Creator') {
    return { ok: false, message: '無權限，僅限 Creator' };
  }

  var system = getSystemSettingsMap_();

  return {
    ok: true,
    profile: {
      id: user.id,
      name: user.name,
      nickname: user.nickname,
      role: user.role,
      color: user.color
    },
    system: system,
    permissions: getPermissionsByRole_('Creator'),
    deadlineMap: getDeadlineMap_(),
    leaveReviews: getReviewCardsForRole_(SHEET_LEAVE, 'Creator', ['假別', '開始日期', '天數', '事由']),
    receiveReviews: getReviewCardsForRole_(SHEET_UPLOAD, 'Creator', ['類型', '標題', '內容']),
    messageReviews: getReviewCardsForRole_(SHEET_MESSAGE, 'Creator', ['留言內容']),
    clockFixReviews: getReviewCardsForRole_(SHEET_CLOCK_FIX, 'Creator', ['補打卡日期', '補打卡時間', '補打卡動作', '申請事由']),
    salaryReviews: getReviewCardsForRole_(SHEET_SALARY, 'Creator', ['月份', '工時', '加班時數', '應發金額', '備註']),
    schedulePublishItems: getPublishByType_('排班發布', 20),
    noticePublishItems: getNoticePublishItems_(20),
    preselectWeekCards: getPreselectWeekCards_(),
    approverSettings: {
      multi_level_review_enabled: String(system['multi_level_review_enabled'] || 'true'),
      leave_approvers: String(system['leave_approvers'] || 'Manager'),
      clockfix_approvers: String(system['clockfix_approvers'] || 'Manager'),
      upload_approvers: String(system['upload_approvers'] || 'Manager'),
      message_approvers: String(system['message_approvers'] || 'Manager'),
      salary_approvers: String(system['salary_approvers'] || 'Admin')
    }
  };
}

// ========== 7. Creator 審核流程設定儲存 ==========

function saveApproverSettings(payload) {
  var auth = authUser_(payload.userId, payload.token || payload.userToken);
  if (!auth.ok) return auth;

  var user = auth.user;
  if (user.role !== 'Creator') return { ok: false, message: '僅 Creator 可修改審核流程設定' };

  var map = {};
  if (payload.multi_level_review_enabled !== undefined) {
    map['multi_level_review_enabled'] = String(payload.multi_level_review_enabled);
  }
  if (payload.leave_approvers) map['leave_approvers'] = clean_(payload.leave_approvers);
  if (payload.clockfix_approvers) map['clockfix_approvers'] = clean_(payload.clockfix_approvers);
  if (payload.upload_approvers) map['upload_approvers'] = clean_(payload.upload_approvers);
  if (payload.message_approvers) map['message_approvers'] = clean_(payload.message_approvers);
  if (payload.salary_approvers) map['salary_approvers'] = clean_(payload.salary_approvers);

  setSystemSettings_(map, user.id);

  return { ok: true, message: '審核流程設定已儲存' };
}

// ========== 8. 歸檔功能 ==========

function archiveOldRecords(payload) {
  var auth = authUser_(payload.userId, payload.token || payload.userToken);
  if (!auth.ok) return auth;

  var user = auth.user;
  if (!isAdminOrCreator_(user.role)) return { ok: false, message: '無權限' };

  var sheetName = clean_(payload.sheetName);
  if (!sheetName) return { ok: false, message: '缺少 sheetName' };

  var sh = getOrCreateSheet_(openSS_(), sheetName);
  if (sh.getLastRow() < 2) return { ok: false, message: '無資料可歸檔' };

  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var statusIdx = headers.indexOf('狀態');
  var archFlagIdx = headers.indexOf('archivedFlag');
  var archDateIdx = headers.indexOf('archivedDate');
  var archByIdx = headers.indexOf('archivedBy');

  if (archFlagIdx < 0) return { ok: false, message: '此表無歸檔欄位' };

  var now = new Date();
  var count = 0;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var status = statusIdx >= 0 ? clean_(row[statusIdx]) : '';
    var archived = archFlagIdx >= 0 ? clean_(row[archFlagIdx]) : '';

    // 只歸檔已結案（已通過/已退回）且未歸檔的
    if ((status === '已通過' || status === '已退回') && archived !== 'TRUE' && archived !== 'true') {
      sh.getRange(i + 1, archFlagIdx + 1).setValue('TRUE');
      if (archDateIdx >= 0) sh.getRange(i + 1, archDateIdx + 1).setValue(now);
      if (archByIdx >= 0) sh.getRange(i + 1, archByIdx + 1).setValue(user.id);
      count++;
    }
  }

  return { ok: true, message: '已歸檔 ' + count + ' 筆資料' };
}

// ========== 9. 匯出歸檔資料到 Drive ==========

function exportArchivedToDrive(payload) {
  var auth = authUser_(payload.userId, payload.token || payload.userToken);
  if (!auth.ok) return auth;

  var user = auth.user;
  if (!isAdminOrCreator_(user.role)) return { ok: false, message: '無權限' };

  var sheetName = clean_(payload.sheetName);
  if (!sheetName) return { ok: false, message: '缺少 sheetName' };

  var sh = getOrCreateSheet_(openSS_(), sheetName);
  if (sh.getLastRow() < 2) return { ok: false, message: '無資料' };

  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var archFlagIdx = headers.indexOf('archivedFlag');

  if (archFlagIdx < 0) return { ok: false, message: '此表無歸檔欄位' };

  // 建新 Spreadsheet 匯出
  var folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
  var now = new Date();
  var fileName = sheetName + '_歸檔_' + Utilities.formatDate(now, TZ, 'yyyyMMdd_HHmmss');
  var newSS = SpreadsheetApp.create(fileName);
  var newSh = newSS.getActiveSheet();
  newSh.setName(sheetName + '_歸檔');

  // 寫標題
  newSh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 寫歸檔資料
  var exportRows = [];
  for (var i = 1; i < values.length; i++) {
    if (clean_(values[i][archFlagIdx]) === 'TRUE' || clean_(values[i][archFlagIdx]) === 'true') {
      exportRows.push(values[i]);
    }
  }

  if (exportRows.length > 0) {
    newSh.getRange(2, 1, exportRows.length, headers.length).setValues(exportRows);
  }

  // 搬到指定資料夾
  var file = DriveApp.getFileById(newSS.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  return {
    ok: true,
    message: '已匯出 ' + exportRows.length + ' 筆歸檔資料',
    fileUrl: newSS.getUrl(),
    fileName: fileName
  };
}

/* ============================================================
 * 完整追加包：薪資系統 + 選休設定 + 發佈規則
 * 貼在程式碼.js 最下方
 * ============================================================ */

// ==================== 薪資審核系統 ====================

function generateSalaryDraft(payload) {
  var auth = authUser_(payload.userId || payload.id, payload.token || payload.userToken);
  if (!auth.ok) return auth;
  var user = auth.user;
  if (!isAdminOrCreator_(user.role)) return { ok: false, message: '無權限' };

  var targetEmpId = clean_(payload.empId).toUpperCase();
  var targetMonth = clean_(payload.month);
  if (!targetEmpId || !targetMonth) return { ok: false, message: '缺少員編或月份' };

  var empSh = getOrCreateSheet_(openSS_(), SHEET_STAFF);
  var empData = empSh.getDataRange().getValues();
  var empHeaders = empData[0];
  var empRow = null;
  for (var i = 1; i < empData.length; i++) {
    var eid = clean_(empData[i][empHeaders.indexOf('員編')] || empData[i][0]).toUpperCase();
    if (eid === targetEmpId) {
      empRow = {};
      empHeaders.forEach(function(h, idx) { empRow[clean_(h)] = empData[i][idx]; });
      break;
    }
  }
  if (!empRow) return { ok: false, message: '找不到員工：' + targetEmpId };

  var clockSh = getOrCreateSheet_(openSS_(), SHEET_CLOCK);
  var clockData = clockSh.getDataRange().getValues();
  var clockHeaders = clockData[0];
  var attendance = calcAttendance_(clockData, clockHeaders, targetEmpId, targetMonth);
  var leaveSh = getOrCreateSheet_(openSS_(), SHEET_LEAVE);
  var leaveData = leaveSh.getDataRange().getValues();
  var leaveHeaders = leaveData[0];
  var leaveDed = calcLeaveDeductions_(leaveData, leaveHeaders, targetEmpId, targetMonth);
  var lateDed = calcLateDeductions_(clockData, clockHeaders, targetEmpId, targetMonth);

  return {
    ok: true,
    draft: {
      empId: targetEmpId,
      empName: clean_(empRow['姓名'] || empRow['暱稱'] || ''),
      month: targetMonth,
      baseSalary: 0,
      workDays: attendance.workDays,
      workHours: attendance.totalHours,
      overtimeHours: attendance.overtimeHours,
      overtimePay: 0,
      lateMinutes: lateDed.totalMinutes,
      lateDeduction: lateDed.amount,
      leaveDeduction: leaveDed.amount,
      leaveDays: leaveDed.days,
      leaveDetail: leaveDed.detail,
      mealAllowance: attendance.workDays * 60,
      extras: [],
      attendanceDetail: attendance.detail
    }
  };
}

function calcAttendance_(data, headers, empId, month) {
  var empIdx = Math.max(headers.indexOf('員編'), 2);
  var dateIdx = Math.max(headers.indexOf('日期時��'), 1);
  var actionIdx = Math.max(headers.indexOf('動作'), 3);
  var dayMap = {};

  for (var i = 1; i < data.length; i++) {
    if (clean_(data[i][empIdx]).toUpperCase() !== empId) continue;
    var dt = data[i][dateIdx];
    if (!dt) continue;
    var d = dt instanceof Date ? dt : new Date(dt);
    if (isNaN(d.getTime())) continue;
    var ds = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    if (!ds.startsWith(month)) continue;
    var action = clean_(data[i][actionIdx]);
    var ts = Utilities.formatDate(d, TZ, 'HH:mm');
    if (!dayMap[ds]) dayMap[ds] = { ci: null, co: null };
    if (/上班|簽到|clockIn/i.test(action)) { if (!dayMap[ds].ci || ts < dayMap[ds].ci) dayMap[ds].ci = ts; }
    if (/下班|簽退|clockOut/i.test(action)) { if (!dayMap[ds].co || ts > dayMap[ds].co) dayMap[ds].co = ts; }
  }

  var totalH = 0, otH = 0, wd = 0, detail = [];
  Object.keys(dayMap).sort().forEach(function(ds) {
    var r = dayMap[ds], h = 0;
    if (r.ci && r.co) {
      var ci = r.ci.split(':'), co = r.co.split(':');
      h = Math.max(0, (co[0]*60+co[1]*1 - ci[0]*60-ci[1]*1)/60);
      h = Math.round(h*100)/100;
    }
    var ot = h > 8 ? Math.round((h-8)*100)/100 : 0;
    totalH += h; otH += ot; if (h > 0) wd++;
    detail.push({ date: ds, clockIn: r.ci||'-', clockOut: r.co||'-', hours: h, overtime: ot });
  });

  return { workDays: wd, totalHours: Math.round(totalH*100)/100, overtimeHours: Math.round(otH*100)/100, detail: detail };
}

function calcLeaveDeductions_(data, headers, empId, month) {
  var ei = headers.indexOf('員編'), di = headers.indexOf('天數'), ti = headers.indexOf('假別');
  var si = headers.indexOf('開始日期'), sti = headers.indexOf('狀態');
  var days = 0, amt = 0, detail = [];

  for (var i = 1; i < data.length; i++) {
    if (clean_(data[i][ei]).toUpperCase() !== empId) continue;
    if (clean_(data[i][sti]) !== '已通過') continue;
    var sd = data[i][si];
    if (!sd) continue;
    var d = sd instanceof Date ? sd : new Date(sd);
    if (isNaN(d.getTime())) continue;
    if (Utilities.formatDate(d, TZ, 'yyyy-MM') !== month) continue;
    var dd = Number(data[i][di]) || 0;
    var lt = clean_(data[i][ti]);
    var ded = 0;
    if (lt === '事假') ded = dd * 500;
    if (lt === '病假') ded = dd * 250;
    days += dd; amt += ded;
    detail.push({ type: lt, days: dd, deduction: ded });
  }
  return { days: days, amount: amt, detail: detail };
}

function calcLateDeductions_(data, headers, empId, month) {
  var ei = Math.max(headers.indexOf('員編'), 2);
  var di = Math.max(headers.indexOf('日期時間'), 1);
  var ai = Math.max(headers.indexOf('動作'), 3);
  var sys = getSystemSettingsMap_();
  var wst = clean_(sys['work_start_time'] || '09:00').split(':');
  var sm = wst[0]*60 + (wst[1]||0)*1;
  var rate = Number(sys['late_deduction_per_minute'] || 10);
  var total = 0, checked = {};

  for (var i = 1; i < data.length; i++) {
    if (clean_(data[i][ei]).toUpperCase() !== empId) continue;
    if (!/上班|簽到|clockIn/i.test(clean_(data[i][ai]))) continue;
    var dt = data[i][di];
    if (!dt) continue;
    var d = dt instanceof Date ? dt : new Date(dt);
    if (isNaN(d.getTime())) continue;
    var ds = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    if (!ds.startsWith(month) || checked[ds]) continue;
    checked[ds] = true;
    var ts = Utilities.formatDate(d, TZ, 'HH:mm').split(':');
    var cm = ts[0]*60 + ts[1]*1;
    if (cm > sm) total += (cm - sm);
  }
  return { totalMinutes: total, amount: total * rate };
}

function saveSalaryReview(payload) {
  var auth = authUser_(payload.userId || payload.id, payload.token || payload.userToken);
  if (!auth.ok) return auth;
  var user = auth.user;
  if (!isAdminOrCreator_(user.role)) return { ok: false, message: '無權限' };

  var sh = getOrCreateSheet_(openSS_(), SHEET_SALARY);
  var now = new Date();
  var reqId = 'SL' + Utilities.formatDate(now, TZ, 'yyyyMMddHHmmss');
  var approvers = isMultiLevelEnabled_() ? getApproversForType_('salary') : ['Admin'];

  var extras = Array.isArray(payload.extras) ? payload.extras : [];
  var extrasTotal = 0;
  extras.forEach(function(e) { extrasTotal += Number(e.amount || 0); });

  var base = Number(payload.baseSalary || 0);
  var otPay = Number(payload.overtimePay || 0);
  var meal = Number(payload.mealAllowance || 0);
  var lateDed = Number(payload.lateDeduction || 0);
  var leaveDed = Number(payload.leaveDeduction || 0);
  var total = base + otPay + meal + extrasTotal - lateDed - leaveDed;

  sh.appendRow([
    reqId, clean_(payload.empId), clean_(payload.empName), clean_(payload.month),
    Number(payload.workHours || 0), Number(payload.overtimeHours || 0), total,
    '待審核', '', '', '',
    approvers.join(','), approvers[0] || '', '[]'
  ]);

  var detailJson = JSON.stringify({
    baseSalary: base, overtimePay: otPay, mealAllowance: meal,
    lateDeduction: lateDed, leaveDeduction: leaveDed,
    extras: extras, total: total
  });

  var lastRow = sh.getLastRow();
  var hs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var noteCol = hs.indexOf('備註') + 1;
  if (noteCol > 0) sh.getRange(lastRow, noteCol).setValue(detailJson);

  return { ok: true, message: '薪資審核單已建立', reqId: reqId, total: total };
}

function downloadSalarySlip(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;
  var user = auth.user;
  var targetMonth = clean_(payload.month);
  if (!targetMonth) return { ok: false, message: '請選擇月份' };

  var sh = getOrCreateSheet_(openSS_(), SHEET_SALARY);
  var data = sh.getDataRange().getValues();
  var hs = data[0];
  var ei = hs.indexOf('員編'), mi = hs.indexOf('月份');
  var si = hs.indexOf('狀態'), ni = hs.indexOf('備註'), ai = hs.indexOf('應發金額');

  for (var i = 1; i < data.length; i++) {
    if (clean_(data[i][ei]).toUpperCase() !== clean_(user.id).toUpperCase()) continue;
    if (clean_(data[i][mi]) !== targetMonth) continue;
    if (clean_(data[i][si]) !== '已通過') continue;
    var detail = {};
    if (ni >= 0) { try { detail = JSON.parse(data[i][ni]); } catch(e) {} }
    return {
      ok: true,
      slip: {
        month: targetMonth, empId: user.id, empName: user.name,
        totalPay: data[i][ai] || 0, detail: detail, slipId: clean_(data[i][0])
      }
    };
  }
  return { ok: false, message: targetMonth + ' 薪資尚未發布或未通過審核' };
}

// ==================== 選休設定 ====================

function getPreselectSettings_() {
  var sys = getSystemSettingsMap_();
  return {
    mode: clean_(sys['preselect_mode'] || 'week'),
    weekOpen: clean_(sys['preselect_week_open'] || '三 20:00'),
    weekClose: clean_(sys['preselect_week_close'] || '五 18:00'),
    weekPublish: clean_(sys['preselect_week_publish'] || '六 10:00'),
    monthOpen: clean_(sys['preselect_month_open'] || '20 09:00'),
    monthClose: clean_(sys['preselect_month_close'] || '25 18:00'),
    monthPublish: clean_(sys['preselect_month_publish'] || '27 10:00')
  };
}

function savePreselectSettings(payload) {
  var auth = authUser_(payload.userId || payload.id, payload.token || payload.userToken);
  if (!auth.ok) return auth;
  if (!isAdminOrCreator_(auth.user.role)) return { ok: false, message: '無權限' };
  var map = {};
  var mode = clean_(payload.mode || 'week').toLowerCase();
  if (mode !== 'week' && mode !== 'month') mode = 'week';
  map['preselect_mode'] = mode;
  if (payload.weekOpen !== undefined) map['preselect_week_open'] = clean_(payload.weekOpen);
  if (payload.weekClose !== undefined) map['preselect_week_close'] = clean_(payload.weekClose);
  if (payload.weekPublish !== undefined) map['preselect_week_publish'] = clean_(payload.weekPublish);
  if (payload.monthOpen !== undefined) map['preselect_month_open'] = clean_(payload.monthOpen);
  if (payload.monthClose !== undefined) map['preselect_month_close'] = clean_(payload.monthClose);
  if (payload.monthPublish !== undefined) map['preselect_month_publish'] = clean_(payload.monthPublish);
  setSystemSettings_(map, auth.user.id);
  return { ok: true, message: '選休規則已儲存，自動套用到往後每一期' };
}

function isPreselectOpen_() {
  var s = getPreselectSettings_();
  var now = new Date();
  return s.mode === 'week'
    ? checkWeekWindow_(now, s.weekOpen, s.weekClose)
    : checkMonthWindow_(now, s.monthOpen, s.monthClose);
}

function checkWeekWindow_(now, openR, closeR) {
  var o = parseWeekR_(openR), c = parseWeekR_(closeR);
  if (!o || !c) return { isOpen: true, reason: '規則格式錯誤，預設開放' };
  var cur = (now.getDay()||7)*1440 + now.getHours()*60 + now.getMinutes();
  var oT = o.d*1440+o.m, cT = c.d*1440+c.m;
  var isO = oT <= cT ? (cur >= oT && cur <= cT) : (cur >= oT || cur <= cT);
  return { isOpen: isO, reason: isO ? '選休開放中' : '選休未開放（' + openR + ' ～ ' + closeR + '）' };
}

function checkMonthWindow_(now, openR, closeR) {
  var o = parseMonthR_(openR), c = parseMonthR_(closeR);
  if (!o || !c) return { isOpen: true, reason: '規則格式錯誤，預設開放' };
  var cur = now.getDate()*1440 + now.getHours()*60 + now.getMinutes();
  var oT = o.d*1440+o.m, cT = c.d*1440+c.m;
  var isO = oT <= cT ? (cur >= oT && cur <= cT) : (cur >= oT || cur <= cT);
  return { isOpen: isO, reason: isO ? '選休開放中' : '選休未開放（每月 ' + openR + ' ～ ' + closeR + '）' };
}

function parseWeekR_(r) {
  var map = {'日':7,'一':1,'二':2,'三':3,'四':4,'五':5,'六':6};
  var m = clean_(r).match(/^([日一二三四五六])\s*(\d{1,2}):(\d{2})$/);
  return m ? { d: map[m[1]], m: m[2]*60+m[3]*1 } : null;
}

function parseMonthR_(r) {
  var m = clean_(r).match(/^(\d{1,2})\s*(\d{1,2}):(\d{2})$/);
  return m ? { d: m[1]*1, m: m[2]*60+m[3]*1 } : null;
}

/* ============================================================
 *  完整追加包 — 貼在程式碼.js 最下方
 *  內容：選休系統 + 薪資系統 + 班別系統 + 輪詢 + 空行過濾
 * ============================================================ */

// ==================== 空白行過濾（全域） ====================

function isEmptyRow_(row) {
  if (!row || !Array.isArray(row)) return true;
  for (var i = 0; i < Math.min(row.length, 5); i++) {
    if (clean_(row[i]) !== '') return false;
  }
  return true;
}

// ==================== 選休設定 CRUD ====================

function getPreselectSettings_(forceRefresh) {
  var sys = getSystemSettingsMap_();
  return {
    mode: clean_(sys['preselect_mode'] || 'week'),
    weekOpen: clean_(sys['preselect_week_open'] || '三 20:00'),
    weekClose: clean_(sys['preselect_week_close'] || '五 18:00'),
    weekPublish: clean_(sys['preselect_week_publish'] || '六 10:00'),
    monthOpen: clean_(sys['preselect_month_open'] || '20 09:00'),
    monthClose: clean_(sys['preselect_month_close'] || '25 18:00'),
    monthPublish: clean_(sys['preselect_month_publish'] || '27 10:00'),
    lastModified: clean_(sys['settings_last_modified'] || '')
  };
}

function savePreselectSettings(payload) {
  var auth = authUser_(payload.userId || payload.id, payload.token || payload.userToken);
  if (!auth.ok) return auth;
  if (!isAdminOrCreator_(auth.user.role)) return { ok: false, message: '無權限' };

  var mode = clean_(payload.mode || 'week').toLowerCase();
  if (mode !== 'week' && mode !== 'month') mode = 'week';

  var map = { preselect_mode: mode };
  if (payload.weekOpen !== undefined) map['preselect_week_open'] = clean_(payload.weekOpen);
  if (payload.weekClose !== undefined) map['preselect_week_close'] = clean_(payload.weekClose);
  if (payload.weekPublish !== undefined) map['preselect_week_publish'] = clean_(payload.weekPublish);
  if (payload.monthOpen !== undefined) map['preselect_month_open'] = clean_(payload.monthOpen);
  if (payload.monthClose !== undefined) map['preselect_month_close'] = clean_(payload.monthClose);
  if (payload.monthPublish !== undefined) map['preselect_month_publish'] = clean_(payload.monthPublish);
  map['settings_last_modified'] = new Date().getTime().toString();

  setSystemSettings_(map, auth.user.id);
  return { ok: true, message: '選休規則已儲存，立即生效', mode: mode };
}

// ==================== 選休開放狀態判斷 ====================

function isPreselectOpen(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;

  var s = getPreselectSettings_();
  var now = new Date();
  var result;

  if (s.mode === 'week') {
    result = checkWeekWindow_(now, s.weekOpen, s.weekClose);
  } else {
    result = checkMonthWindow_(now, s.monthOpen, s.monthClose);
  }

  result.mode = s.mode;
  result.ok = true;
  return result;
}

function checkWeekWindow_(now, openR, closeR) {
  var o = parseWeekR_(openR), c = parseWeekR_(closeR);
  if (!o || !c) return { isOpen: true, reason: '規則格式錯誤，預設開放' };
  var dayOfWeek = now.getDay(); // 0=日
  var curDay = dayOfWeek === 0 ? 7 : dayOfWeek; // 轉成 1=一 7=日
  var cur = curDay * 1440 + now.getHours() * 60 + now.getMinutes();
  var oT = o.d * 1440 + o.m, cT = c.d * 1440 + c.m;
  var isO = oT <= cT ? (cur >= oT && cur <= cT) : (cur >= oT || cur <= cT);
  return { isOpen: isO, reason: isO ? '✅ 選休開放中' : '🔒 選休已截止（' + openR + ' ～ ' + closeR + '）' };
}

function checkMonthWindow_(now, openR, closeR) {
  var o = parseMonthR_(openR), c = parseMonthR_(closeR);
  if (!o || !c) return { isOpen: true, reason: '規則格式錯誤，預設開放' };
  var cur = now.getDate() * 1440 + now.getHours() * 60 + now.getMinutes();
  var oT = o.d * 1440 + o.m, cT = c.d * 1440 + c.m;
  var isO = oT <= cT ? (cur >= oT && cur <= cT) : (cur >= oT || cur <= cT);
  return { isOpen: isO, reason: isO ? '✅ 選休開放中' : '🔒 選休已截止（每月 ' + openR + ' ～ ' + closeR + '）' };
}

function parseWeekR_(r) {
  var map = { '日': 7, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6 };
  var m = clean_(r).match(/^([日一二三四五六])\s*(\d{1,2}):(\d{2})$/);
  return m ? { d: map[m[1]], m: m[2] * 60 + m[3] * 1 } : null;
}

function parseMonthR_(r) {
  var m = clean_(r).match(/^(\d{1,2})\s*(\d{1,2}):(\d{2})$/);
  return m ? { d: m[1] * 1, m: m[2] * 60 + m[3] * 1 } : null;
}

// ==================== 選休寫入（反轉邏輯） ====================

function submitPreselect(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;

  // 先檢查是否開放
  var s = getPreselectSettings_();
  var now = new Date();
  var windowCheck = s.mode === 'week'
    ? checkWeekWindow_(now, s.weekOpen, s.weekClose)
    : checkMonthWindow_(now, s.monthOpen, s.monthClose);

  if (!windowCheck.isOpen) {
    return { ok: false, message: windowCheck.reason };
  }

  var sh = getOrCreateSheet_(openSS_(), SHEET_PRESELECT);
  var data = sh.getDataRange().getValues();
  var headerRow = data[0];

  var empId = clean_(auth.user.id).toUpperCase();
  var targetRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (clean_(data[i][0]).toUpperCase() === empId) { targetRow = i + 1; break; }
  }
  if (targetRow < 0) return { ok: false, message: '找不到你的選休行' };

  var selected = Array.isArray(payload.selected) ? payload.selected : [];
  var blockInfo = findCurrentBlock_(headerRow, s.mode);
  if (!blockInfo) return { ok: false, message: '找不到對應的選休���塊' };

  // 先全部設 TRUE（上班）
  for (var d = 0; d < blockInfo.dayCount; d++) {
    var col = blockInfo.dataStartCol + d;
    if (col <= headerRow.length) {
      sh.getRange(targetRow, col).setValue(true);
    }
  }

  // 員工選中的設 FALSE（休假）
  selected.forEach(function(idx) {
    if (idx >= 0 && idx < blockInfo.dayCount) {
      var col = blockInfo.dataStartCol + idx;
      if (col <= headerRow.length) {
        sh.getRange(targetRow, col).setValue(false);
      }
    }
  });

  return { ok: true, message: '選休已送出' };
}

function findCurrentBlock_(headerRow, mode) {
  // 根據模式找到正確的區塊
  var keyword = mode === 'month' ? '下月' : '下週';
  for (var c = 0; c < headerRow.length; c++) {
    var h = clean_(headerRow[c]);
    if (h.indexOf(keyword) > -1) {
      // 找到起始，算天數
      var dayCount = 0;
      for (var d = c + 1; d < headerRow.length; d++) {
        var hd = clean_(headerRow[d]);
        if (hd === '' || hd.indexOf('週') > -1 || hd.indexOf('月') > -1) break;
        dayCount++;
      }
      if (dayCount === 0) dayCount = mode === 'month' ? 31 : 7;
      return { dataStartCol: c + 2, dayCount: dayCount };
    }
  }
  // 預設
  return { dataStartCol: 3, dayCount: mode === 'month' ? 31 : 7 };
}

// ==================== 選休讀取（反轉邏輯） ====================

function getMyPreselect(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;

  var s = getPreselectSettings_();
  var now = new Date();
  var windowCheck = s.mode === 'week'
    ? checkWeekWindow_(now, s.weekOpen, s.weekClose)
    : checkMonthWindow_(now, s.monthOpen, s.monthClose);

  var sh = getOrCreateSheet_(openSS_(), SHEET_PRESELECT);
  var data = sh.getDataRange().getValues();
  var headerRow = data[0];
  var empId = clean_(auth.user.id).toUpperCase();

  var selected = [];
  var labels = [];

  for (var i = 1; i < data.length; i++) {
    if (clean_(data[i][0]).toUpperCase() !== empId) continue;

    var blockInfo = findCurrentBlock_(headerRow, s.mode);
    if (!blockInfo) break;

    for (var d = 0; d < blockInfo.dayCount; d++) {
      var colIdx = blockInfo.dataStartCol - 1 + d;
      var headerLabel = colIdx < headerRow.length ? clean_(headerRow[colIdx]) : ('第' + (d + 1) + '天');
      labels.push(headerLabel);

      // FALSE = 休假 = 員工選中
      var val = colIdx < data[i].length ? data[i][colIdx] : true;
      selected.push(val === false || val === 'FALSE' || val === 'false');
    }
    break;
  }

  return {
    ok: true,
    mode: s.mode,
    isOpen: windowCheck.isOpen,
    reason: windowCheck.reason,
    selected: selected,
    labels: labels,
    dayCount: labels.length
  };
}

// ==================== 打卡寬限 ====================

function getClockGrace_() {
  var sys = getSystemSettingsMap_();
  return {
    beforeMinutes: Number(sys['clock_grace_before'] || 15),
    afterMinutes: Number(sys['clock_grace_after'] || 15)
  };
}

// ==================== 班別系統（Creator 設定） ====================

var SHEET_SHIFTS = '班別設定';

function getShiftTypes(payload) {
  var auth = authUser_(payload.id || payload.userId, payload.token || payload.userToken);
  if (!auth.ok) return auth;

  var sh = getOrCreateSheet_(openSS_(), SHEET_SHIFTS);
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return { ok: true, shifts: [], mixedMode: false };

  var headers = data[0];
  var shifts = [];
  for (var i = 1; i < data.length; i++) {
    if (isEmptyRow_(data[i])) continue;
    shifts.push({
      name: clean_(data[i][0]),
      startTime: clean_(data[i][1]),
      endTime: clean_(data[i][2]),
      breakMinutes: Number(data[i][3] || 60),
      hours: Number(data[i][4] || 8)
    });
  }

  var sys = getSystemSettingsMap_();
  var mixedMode = clean_(sys['shift_mixed_mode'] || 'false') === 'true';

  return { ok: true, shifts: shifts, mixedMode: mixedMode };
}

function saveShiftTypes(payload) {
  var auth = authUser_(payload.userId || payload.id, payload.token || payload.userToken);
  if (!auth.ok) return auth;
  if (!isAdminOrCreator_(auth.user.role)) return { ok: false, message: '無權限' };

  var sh = getOrCreateSheet_(openSS_(), SHEET_SHIFTS);
  sh.clear();
  sh.appendRow(['班別名稱', '上班時間', '下班時間', '休息(分鐘)', '實際工時']);

  var shifts = Array.isArray(payload.shifts) ? payload.shifts : [];
  shifts.forEach(function(s) {
    var startParts = clean_(s.startTime || '08:00').split(':');
    var endParts = clean_(s.endTime || '17:00').split(':');
    var startMin = startParts[0] * 60 + (startParts[1] || 0) * 1;
    var endMin = endParts[0] * 60 + (endParts[1] || 0) * 1;
    var breakMin = Number(s.breakMinutes || 60);
    var totalMin = endMin >= startMin ? (endMin - startMin - breakMin) : (1440 - startMin + endMin - breakMin);
    var hours = Math.round(totalMin / 60 * 100) / 100;

    sh.appendRow([
      clean_(s.name),
      clean_(s.startTime || '08:00'),
      clean_(s.endTime || '17:00'),
      breakMin,
      hours
    ]);
  });

  // 混班模式
  var mixedMode = payload.mixedMode === true || payload.mixedMode === 'true';
  setSystemSettings_({ shift_mixed_mode: mixedMode ? 'true' : 'false' }, auth.user.id);

  return { ok: true, message: '班別設定已儲存（共 ' + shifts.length + ' 個班別）' };
}

// ==================== 今日狀態 + 下一班提醒 ====================

function getTodayStatus(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;
  var user = auth.user;
  var empId = clean_(user.id).toUpperCase();
  var now = new Date();
  var todayStr = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  var tomorrowDate = new Date(now.getTime() + 86400000);
  var tomorrowStr = Utilities.formatDate(tomorrowDate, TZ, 'yyyy-MM-dd');

  // 1. 檢查今天是否有已通過的請假
  var leaveToday = checkLeaveOnDate_(empId, todayStr);
  var leaveTomorrow = checkLeaveOnDate_(empId, tomorrowStr);

  // 2. 從班表看今天/明天是否上班
  var scheduleToday = getScheduleForDate_(empId, todayStr);
  var scheduleTomorrow = getScheduleForDate_(empId, tomorrowStr);

  // 3. 今日狀態文字
  var todayText = '';
  var todayIcon = '';
  var todayUrgency = 'normal'; // normal, warning, danger

  if (leaveToday.hasLeave) {
    todayText = '今天' + leaveToday.leaveType + '假';
    todayIcon = '😴';
  } else if (!scheduleToday.isWork) {
    todayText = '今天休假';
    todayIcon = '😴';
  } else {
    // 今天要上班，算時間差
    var workStart = scheduleToday.startTime || '09:00';
    var startParts = workStart.split(':');
    var workStartMin = startParts[0] * 60 + (startParts[1] || 0) * 1;
    var nowMin = now.getHours() * 60 + now.getMinutes();
    var grace = getClockGrace_();
    var diff = workStartMin - nowMin;

    if (diff > 60) {
      todayText = '今天 ' + workStart + ' 上班';
      todayIcon = '☀️';
    } else if (diff > 30) {
      todayText = '準備上班囉';
      todayIcon = '🔔';
    } else if (diff > 15) {
      todayText = '快遲到了！';
      todayIcon = '⚡';
      todayUrgency = 'warning';
    } else if (diff > 0) {
      todayText = '完了完了要來不及了！';
      todayIcon = '🚨';
      todayUrgency = 'danger';
    } else if (diff > -grace.afterMinutes) {
      todayText = '你遲到了！';
      todayIcon = '😱';
      todayUrgency = 'danger';
    } else {
      todayText = '要被扣錢了！';
      todayIcon = '💸';
      todayUrgency = 'danger';
    }

    // 已經打過卡就改成正常
    var clocked = checkClockedToday_(empId, todayStr);
    if (clocked) {
      todayText = '已打卡上班中';
      todayIcon = '✅';
      todayUrgency = 'normal';
    }
  }

  // 4. 明天班別
  var tomorrowText = '';
  if (leaveTomorrow.hasLeave) {
    tomorrowText = '明天' + leaveTomorrow.leaveType + '假';
  } else if (!scheduleTomorrow.isWork) {
    tomorrowText = '明天休假';
  } else {
    tomorrowText = '明天 ' + (scheduleTomorrow.shiftName || '') + ' ' + (scheduleTomorrow.startTime || '') + ' 上班';
  }

  return {
    ok: true,
    today: { text: todayText, icon: todayIcon, urgency: todayUrgency },
    tomorrow: { text: tomorrowText },
    schedule: scheduleToday
  };
}

function checkLeaveOnDate_(empId, dateStr) {
  var sh = getOrCreateSheet_(openSS_(), SHEET_LEAVE);
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var eiCol = headers.indexOf('員編');
  var stCol = headers.indexOf('狀態');
  var startCol = headers.indexOf('開始日期');
  var endCol = headers.indexOf('結束日期');
  var typeCol = headers.indexOf('假別');

  for (var i = 1; i < data.length; i++) {
    if (isEmptyRow_(data[i])) continue;
    if (clean_(data[i][eiCol]).toUpperCase() !== empId) continue;
    if (clean_(data[i][stCol]) !== '已通過') continue;

    var sd = data[i][startCol];
    var ed = data[i][endCol];
    if (!sd) continue;
    var startD = sd instanceof Date ? sd : new Date(sd);
    var endD = ed ? (ed instanceof Date ? ed : new Date(ed)) : startD;
    var checkD = new Date(dateStr + 'T00:00:00');

    if (checkD >= new Date(Utilities.formatDate(startD, TZ, 'yyyy-MM-dd') + 'T00:00:00') &&
        checkD <= new Date(Utilities.formatDate(endD, TZ, 'yyyy-MM-dd') + 'T00:00:00')) {
      return { hasLeave: true, leaveType: clean_(data[i][typeCol]) || '休' };
    }
  }
  return { hasLeave: false };
}

function getScheduleForDate_(empId, dateStr) {
  // 嘗試從班表讀取
  try {
    var sh = getOrCreateSheet_(openSS_(), SHEET_PRESELECT);
    var data = sh.getDataRange().getValues();
    var headers = data[0];

    for (var i = 1; i < data.length; i++) {
      if (clean_(data[i][0]).toUpperCase() !== empId) continue;

      // 找到日期對應的欄
      for (var c = 1; c < headers.length; c++) {
        var h = clean_(headers[c]);
        if (h === dateStr || h.indexOf(dateStr) > -1) {
          var val = data[i][c];
          var isWork = !(val === false || val === 'FALSE' || val === 'false' || val === '休');

          // 讀取班別
          var shiftName = '';
          if (typeof val === 'string' && val !== 'TRUE' && val !== 'FALSE' && val !== '休') {
            shiftName = val; // 如 "早班"
          }

          // 讀班別的時間
          var startTime = '09:00';
          if (shiftName) {
            var shiftInfo = getShiftTime_(shiftName);
            if (shiftInfo) startTime = shiftInfo.startTime;
          } else {
            var sys = getSystemSettingsMap_();
            startTime = clean_(sys['work_start_time'] || '09:00');
          }

          return { isWork: isWork, shiftName: shiftName, startTime: startTime };
        }
      }
      break;
    }
  } catch (e) {}

  // 預設
  var sys2 = getSystemSettingsMap_();
  return { isWork: true, shiftName: '', startTime: clean_(sys2['work_start_time'] || '09:00') };
}

function getShiftTime_(shiftName) {
  try {
    var sh = getOrCreateSheet_(openSS_(), SHEET_SHIFTS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (clean_(data[i][0]) === shiftName) {
        return { startTime: clean_(data[i][1]), endTime: clean_(data[i][2]), breakMin: Number(data[i][3] || 60) };
      }
    }
  } catch (e) {}
  return null;
}

function checkClockedToday_(empId, todayStr) {
  try {
    var sh = getOrCreateSheet_(openSS_(), SHEET_CLOCK);
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var eiIdx = Math.max(headers.indexOf('員編'), 2);
    var dtIdx = Math.max(headers.indexOf('日期時間'), 1);

    for (var i = data.length - 1; i >= 1; i--) {
      if (clean_(data[i][eiIdx]).toUpperCase() !== empId) continue;
      var dt = data[i][dtIdx];
      if (!dt) continue;
      var d = dt instanceof Date ? dt : new Date(dt);
      if (isNaN(d.getTime())) continue;
      if (Utilities.formatDate(d, TZ, 'yyyy-MM-dd') === todayStr) return true;
    }
  } catch (e) {}
  return false;
}

// ==================== 近期動態（真實資料） ====================

function getRecentActivities(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;
  var empId = clean_(auth.user.id).toUpperCase();
  var activities = [];

  // 1. 打卡紀錄（最近5筆）
  try {
    var clockSh = getOrCreateSheet_(openSS_(), SHEET_CLOCK);
    var clockData = clockSh.getDataRange().getValues();
    var clockH = clockData[0];
    var ci = Math.max(clockH.indexOf('員編'), 2);
    var di = Math.max(clockH.indexOf('日期時間'), 1);
    var ai = Math.max(clockH.indexOf('動作'), 3);
    var count = 0;
    for (var i = clockData.length - 1; i >= 1 && count < 5; i--) {
      if (isEmptyRow_(clockData[i])) continue;
      if (clean_(clockData[i][ci]).toUpperCase() !== empId) continue;
      var dt = clockData[i][di];
      var d = dt instanceof Date ? dt : new Date(dt);
      var timeStr = !isNaN(d.getTime()) ? Utilities.formatDate(d, TZ, 'MM/dd HH:mm') : '';
      activities.push({
        type: 'clock',
        icon: '⏰',
        title: clean_(clockData[i][ai]) || '打卡',
        time: timeStr
      });
      count++;
    }
  } catch (e) {}

  // 2. 請假紀錄（最近5筆）
  try {
    var leaveSh = getOrCreateSheet_(openSS_(), SHEET_LEAVE);
    var leaveData = leaveSh.getDataRange().getValues();
    var lh = leaveData[0];
    var lei = lh.indexOf('員編');
    var lti = lh.indexOf('假別');
    var lsi = lh.indexOf('狀態');
    var ldi = lh.indexOf('提交時間');
    var lcount = 0;
    for (var i = leaveData.length - 1; i >= 1 && lcount < 5; i--) {
      if (isEmptyRow_(leaveData[i])) continue;
      if (clean_(leaveData[i][lei]).toUpperCase() !== empId) continue;
      var status = clean_(leaveData[i][lsi]);
      var statusText = status === '已通過' ? '✅' : status === '已退回' ? '❌' : '⏳';
      var dt2 = leaveData[i][ldi];
      var d2 = dt2 instanceof Date ? dt2 : new Date(dt2);
      var timeStr2 = !isNaN(d2.getTime()) ? Utilities.formatDate(d2, TZ, 'MM/dd HH:mm') : '';
      activities.push({
        type: 'leave',
        icon: '📋',
        title: clean_(leaveData[i][lti]) + ' ' + statusText,
        time: timeStr2
      });
      lcount++;
    }
  } catch (e) {}

  // 3. 補打卡紀錄（最近3筆）
  try {
    var cfSh = getOrCreateSheet_(openSS_(), SHEET_CLOCKFIX);
    var cfData = cfSh.getDataRange().getValues();
    var cfH = cfData[0];
    var cfei = cfH.indexOf('員編');
    var cfsi = cfH.indexOf('狀態');
    var cfdi = cfH.indexOf('提交時間');
    var cfcount = 0;
    for (var i = cfData.length - 1; i >= 1 && cfcount < 3; i--) {
      if (isEmptyRow_(cfData[i])) continue;
      if (clean_(cfData[i][cfei]).toUpperCase() !== empId) continue;
      var st = clean_(cfData[i][cfsi]);
      var stT = st === '已通過' ? '✅' : st === '已退回' ? '❌' : '⏳';
      var dt3 = cfData[i][cfdi];
      var d3 = dt3 instanceof Date ? dt3 : new Date(dt3);
      var ts3 = !isNaN(d3.getTime()) ? Utilities.formatDate(d3, TZ, 'MM/dd HH:mm') : '';
      activities.push({
        type: 'clockfix',
        icon: '🔧',
        title: '補打卡 ' + stT,
        time: ts3
      });
      cfcount++;
    }
  } catch (e) {}

  // 依時間排序（最新在前）
  activities.sort(function(a, b) { return (b.time || '').localeCompare(a.time || ''); });

  return { ok: true, activities: activities.slice(0, 10) };
}

// ==================== 主管通知（真實資料） ====================

function getNoticesForEmployee(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;

  var sh = getOrCreateSheet_(openSS_(), SHEET_NOTICE);
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var notices = [];

  var titleIdx = headers.indexOf('標題');
  var contentIdx = headers.indexOf('內容');
  var dateIdx = headers.indexOf('發布時間');
  var authorIdx = headers.indexOf('發布者');

  if (titleIdx < 0) titleIdx = 0;
  if (contentIdx < 0) contentIdx = 1;

  for (var i = data.length - 1; i >= 1; i--) {
    if (isEmptyRow_(data[i])) continue;
    var dt = dateIdx >= 0 ? data[i][dateIdx] : '';
    var d = dt instanceof Date ? dt : new Date(dt);
    var timeStr = !isNaN(d.getTime()) ? Utilities.formatDate(d, TZ, 'MM/dd HH:mm') : '';
    notices.push({
      title: clean_(data[i][titleIdx]) || '通知',
      content: clean_(data[i][contentIdx]) || '',
      time: timeStr,
      author: authorIdx >= 0 ? clean_(data[i][authorIdx]) : ''
    });
    if (notices.length >= 10) break;
  }

  return { ok: true, notices: notices };
}

// ==================== 輪詢（設定變動偵測） ====================

function getSettingsHash(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;
  var sys = getSystemSettingsMap_();
  return {
    ok: true,
    settings: {
      mode: clean_(sys['preselect_mode'] || 'week'),
      lastModified: clean_(sys['settings_last_modified'] || '')
    }
  };
}

// ==================== 薪資審核系統 ====================

function generateSalaryDraft(payload) {
  var auth = authUser_(payload.userId || payload.id, payload.token || payload.userToken);
  if (!auth.ok) return auth;
  if (!isAdminOrCreator_(auth.user.role)) return { ok: false, message: '無權限' };

  var targetEmpId = clean_(payload.empId).toUpperCase();
  var targetMonth = clean_(payload.month);
  if (!targetEmpId || !targetMonth) return { ok: false, message: '缺少員編或月份' };

  var empSh = getOrCreateSheet_(openSS_(), SHEET_STAFF);
  var empData = empSh.getDataRange().getValues();
  var empHeaders = empData[0];
  var empRow = null;
  for (var i = 1; i < empData.length; i++) {
    if (isEmptyRow_(empData[i])) continue;
    var eid = clean_(empData[i][empHeaders.indexOf('員編')] || empData[i][0]).toUpperCase();
    if (eid === targetEmpId) {
      empRow = {};
      empHeaders.forEach(function(h, idx) { empRow[clean_(h)] = empData[i][idx]; });
      break;
    }
  }
  if (!empRow) return { ok: false, message: '找不到員工：' + targetEmpId };

  var clockSh = getOrCreateSheet_(openSS_(), SHEET_CLOCK);
  var clockData = clockSh.getDataRange().getValues();
  var clockHeaders = clockData[0];
  var att = calcAttendance_(clockData, clockHeaders, targetEmpId, targetMonth);
  var leaveSh = getOrCreateSheet_(openSS_(), SHEET_LEAVE);
  var leaveData = leaveSh.getDataRange().getValues();
  var leaveHeaders = leaveData[0];
  var ld = calcLeaveDeductions_(leaveData, leaveHeaders, targetEmpId, targetMonth);
  var late = calcLateDeductions_(clockData, clockHeaders, targetEmpId, targetMonth);

  return {
    ok: true,
    draft: {
      empId: targetEmpId, empName: clean_(empRow['姓名'] || empRow['暱稱'] || ''),
      month: targetMonth, baseSalary: 0,
      workDays: att.workDays, workHours: att.totalHours,
      overtimeHours: att.overtimeHours, overtimePay: 0,
      lateMinutes: late.totalMinutes, lateDeduction: late.amount,
      leaveDeduction: ld.amount, leaveDays: ld.days, leaveDetail: ld.detail,
      mealAllowance: att.workDays * 60, extras: [], attendanceDetail: att.detail
    }
  };
}

function calcAttendance_(data, headers, empId, month) {
  var ei = Math.max(headers.indexOf('員編'), 2);
  var di = Math.max(headers.indexOf('日期時間'), 1);
  var ai = Math.max(headers.indexOf('動作'), 3);
  var dayMap = {};

  for (var i = 1; i < data.length; i++) {
    if (isEmptyRow_(data[i])) continue;
    if (clean_(data[i][ei]).toUpperCase() !== empId) continue;
    var dt = data[i][di]; if (!dt) continue;
    var d = dt instanceof Date ? dt : new Date(dt);
    if (isNaN(d.getTime())) continue;
    var ds = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    if (!ds.startsWith(month)) continue;
    var action = clean_(data[i][ai]);
    var ts = Utilities.formatDate(d, TZ, 'HH:mm');
    if (!dayMap[ds]) dayMap[ds] = { ci: null, co: null };
    if (/上班|簽到|clockIn/i.test(action)) { if (!dayMap[ds].ci || ts < dayMap[ds].ci) dayMap[ds].ci = ts; }
    if (/下班|簽退|clockOut/i.test(action)) { if (!dayMap[ds].co || ts > dayMap[ds].co) dayMap[ds].co = ts; }
  }

  var totalH = 0, otH = 0, wd = 0, detail = [];
  Object.keys(dayMap).sort().forEach(function(ds) {
    var r = dayMap[ds], h = 0;
    if (r.ci && r.co) {
      var c1 = r.ci.split(':'), c2 = r.co.split(':');
      h = Math.max(0, (c2[0]*60+c2[1]*1 - c1[0]*60-c1[1]*1) / 60);
      h = Math.round(h * 100) / 100;
    }
    var ot = h > 8 ? Math.round((h - 8) * 100) / 100 : 0;
    totalH += h; otH += ot; if (h > 0) wd++;
    detail.push({ date: ds, clockIn: r.ci || '-', clockOut: r.co || '-', hours: h, overtime: ot });
  });

  return { workDays: wd, totalHours: Math.round(totalH * 100) / 100, overtimeHours: Math.round(otH * 100) / 100, detail: detail };
}

function calcLeaveDeductions_(data, headers, empId, month) {
  var ei = headers.indexOf('員編'), di = headers.indexOf('天數'), ti = headers.indexOf('假別');
  var si = headers.indexOf('開始日期'), sti = headers.indexOf('狀態');
  var days = 0, amt = 0, detail = [];

  for (var i = 1; i < data.length; i++) {
    if (isEmptyRow_(data[i])) continue;
    if (clean_(data[i][ei]).toUpperCase() !== empId) continue;
    if (clean_(data[i][sti]) !== '已通過') continue;
    var sd = data[i][si]; if (!sd) continue;
    var d = sd instanceof Date ? sd : new Date(sd);
    if (isNaN(d.getTime())) continue;
    if (Utilities.formatDate(d, TZ, 'yyyy-MM') !== month) continue;
    var dd = Number(data[i][di]) || 0;
    var lt = clean_(data[i][ti]);
    var ded = 0;
    if (lt === '事假') ded = dd * 500;
    if (lt === '病假') ded = dd * 250;
    days += dd; amt += ded;
    detail.push({ type: lt, days: dd, deduction: ded });
  }
  return { days: days, amount: amt, detail: detail };
}

function calcLateDeductions_(data, headers, empId, month) {
  var ei = Math.max(headers.indexOf('員編'), 2);
  var di = Math.max(headers.indexOf('日期時間'), 1);
  var ai = Math.max(headers.indexOf('動作'), 3);
  var sys = getSystemSettingsMap_();
  var grace = getClockGrace_();
  var wst = clean_(sys['work_start_time'] || '09:00').split(':');
  var sm = wst[0] * 60 + (wst[1] || 0) * 1 + grace.beforeMinutes; // 加上寬限
  var rate = Number(sys['late_deduction_per_minute'] || 10);
  var total = 0, checked = {};

  for (var i = 1; i < data.length; i++) {
    if (isEmptyRow_(data[i])) continue;
    if (clean_(data[i][ei]).toUpperCase() !== empId) continue;
    if (!/上班|簽到|clockIn/i.test(clean_(data[i][ai]))) continue;
    var dt = data[i][di]; if (!dt) continue;
    var d = dt instanceof Date ? dt : new Date(dt);
    if (isNaN(d.getTime())) continue;
    var ds = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    if (!ds.startsWith(month) || checked[ds]) continue;
    checked[ds] = true;
    var ts = Utilities.formatDate(d, TZ, 'HH:mm').split(':');
    var cm = ts[0] * 60 + ts[1] * 1;
    if (cm > sm) total += (cm - sm);
  }
  return { totalMinutes: total, amount: total * rate };
}

function saveSalaryReview(payload) {
  var auth = authUser_(payload.userId || payload.id, payload.token || payload.userToken);
  if (!auth.ok) return auth;
  if (!isAdminOrCreator_(auth.user.role)) return { ok: false, message: '無權限' };

  var sh = getOrCreateSheet_(openSS_(), SHEET_SALARY);
  var now = new Date();
  var reqId = 'SL' + Utilities.formatDate(now, TZ, 'yyyyMMddHHmmss');
  var extras = Array.isArray(payload.extras) ? payload.extras : [];
  var extrasTotal = 0;
  extras.forEach(function(e) { extrasTotal += Number(e.amount || 0); });
  var base = Number(payload.baseSalary || 0);
  var otPay = Number(payload.overtimePay || 0);
  var meal = Number(payload.mealAllowance || 0);
  var lateDed = Number(payload.lateDeduction || 0);
  var leaveDed = Number(payload.leaveDeduction || 0);
  var total = base + otPay + meal + extrasTotal - lateDed - leaveDed;

  sh.appendRow([reqId, clean_(payload.empId), clean_(payload.empName), clean_(payload.month),
    Number(payload.workHours || 0), Number(payload.overtimeHours || 0), total, '待審核', '', '', '',
    '', '', '[]']);

  var detailJson = JSON.stringify({ baseSalary: base, overtimePay: otPay, mealAllowance: meal,
    lateDeduction: lateDed, leaveDeduction: leaveDed, extras: extras, total: total });
  var lastRow = sh.getLastRow();
  var hs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var noteCol = hs.indexOf('備註') + 1;
  if (noteCol > 0) sh.getRange(lastRow, noteCol).setValue(detailJson);

  return { ok: true, message: '薪資審核單已建立', reqId: reqId, total: total };
}

function downloadSalarySlip(payload) {
  var auth = authUser_(payload.id, payload.token);
  if (!auth.ok) return auth;
  var user = auth.user;
  var targetMonth = clean_(payload.month);
  if (!targetMonth) return { ok: false, message: '請選擇月份' };

  var sh = getOrCreateSheet_(openSS_(), SHEET_SALARY);
  var data = sh.getDataRange().getValues();
  var hs = data[0];
  var ei = hs.indexOf('員編'), mi = hs.indexOf('月份'), si = hs.indexOf('狀態'), ni = hs.indexOf('備註'), ai = hs.indexOf('應發金額');

  for (var i = 1; i < data.length; i++) {
    if (isEmptyRow_(data[i])) continue;
    if (clean_(data[i][ei]).toUpperCase() !== clean_(user.id).toUpperCase()) continue;
    if (clean_(data[i][mi]) !== targetMonth) continue;
    if (clean_(data[i][si]) !== '已通過') continue;
    var detail = {};
    if (ni >= 0) { try { detail = JSON.parse(data[i][ni]); } catch (e) {} }
    return {
      ok: true,
      slip: { month: targetMonth, empId: user.id, empName: user.name,
        totalPay: data[i][ai] || 0, detail: detail, slipId: clean_(data[i][0]) }
    };
  }
  return { ok: false, message: targetMonth + ' 薪資尚未發布或未通過審核' };
}

// ==================== 資料接收審核欄位映射 ====================
// 新欄位：A=資料編號 B=類型 C=員編 D=姓名 E=標題 F=申請內容
//         G=代墊金額 H=代墊商品 I=位置資訊 J=附件 K=提交時間
//         L=狀態 M=審核者 N=審核時間 O=審核備註
//         P=requiredApprovers Q=currentStep R=approvedList
//         S=archivedFlag T=archivedDate U=archivedBy

var RECEIVE_COLS = {
  id: 0, type: 1, empId: 2, name: 3, title: 4, content: 5,
  amount: 6, product: 7, location: 8, attachment: 9, submitTime: 10,
  status: 11, reviewer: 12, reviewTime: 13, reviewNote: 14,
  requiredApprovers: 15, currentStep: 16, approvedList: 17,
  archivedFlag: 18, archivedDate: 19, archivedBy: 20
};