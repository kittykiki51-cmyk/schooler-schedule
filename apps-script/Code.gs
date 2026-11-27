/**
 * Schooler｜課程錄製排程 — Google Apps Script 後端
 * 綁定試算表：講師回覆錄製日與檢核日
 *
 * 寫入規則（依現有資料比對）：
 *  - 課程總表：每次送出新增一列，課程狀態預設「待開始」、來源「網站」
 *  - 堂次排程：每堂一列（堂次從 0 起算，共「總時數 + 1」堂），檢核狀態預設「尚未檢核」
 *  - 工作檢核日：交付日的下一個工作日；交付日為週五時，取第 3 個工作日
 *    （跳過週六、週日與「國定假日」分頁中「是否停班 = 是」的日期）
 */

const SPREADSHEET_ID = '12M5HYmr2F1CBEAh1U-mhSU2XLSBrRuofZKRg5dAQWlo';
const TZ = 'Asia/Taipei';
const SHEET_COURSES = '課程總表';
const SHEET_LESSONS = '堂次排程';
const SHEET_HOLIDAYS = '國定假日';

const SITE_URL = 'https://kittykiki51-cmyk.github.io/schooler-schedule/';

/** 舊網址：引導到 GitHub Pages 網站 */
function doGet() {
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;max-width:480px;margin:80px auto;padding:0 20px;line-height:1.7;color:#1c211d">' +
    '<h2 style="margin:0 0 12px">Schooler 課程錄製排程已搬家</h2>' +
    '<p style="margin:0 0 20px;color:#687069">請改用新網址填寫排程：</p>' +
    '<a href="' + SITE_URL + '" target="_top" style="display:inline-block;background:#245947;color:#fff;padding:12px 20px;border-radius:12px;text-decoration:none;font-weight:700">前往填寫 →</a>' +
    '<p style="margin-top:20px;font-size:13px;color:#687069;word-break:break-all">' + SITE_URL + '</p></div>')
    .setTitle('Schooler｜課程錄製排程')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** GitHub Pages 網站送出的資料（JSON，以 text/plain 傳送避免 CORS 預檢） */
function doPost(e) {
  let payload = {};
  try { payload = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_(fail_('資料格式錯誤。')); }
  return json_(submitSchedule(payload));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 由網頁呼叫：建立排程並寫入試算表 */
function submitSchedule(p) {
  try {
    p = p || {};
    // 機器人防護：隱藏欄位有值就假裝成功、不寫入
    if (p.website) return { ok: true, scheduleId: 'SC-' + today_() + '-000000' };

    const teacherName = String(p.teacherName || '').trim().slice(0, 40);
    const courseName = String(p.courseName || '').trim().slice(0, 80);
    const notes = String(p.notes || '').trim().slice(0, 500);
    const totalHours = Number(p.totalHours);
    const dates = Array.isArray(p.dates) ? p.dates.map(String) : [];

    if (!teacherName || !courseName) return fail_('請完成所有必填欄位，再送出排程。');
    if (!Number.isInteger(totalHours) || totalHours < 1 || totalHours > 60) return fail_('課程總時數需為 1–60 的整數。');
    const lessonCount = totalHours + 1; // 第 0 堂起算
    if (dates.length !== lessonCount) return fail_('交付日期數量與堂數不符，請重新整理後再試。');
    for (const d of dates) if (!isValidDate_(d)) return fail_('請選擇有效日期。');

    // 防止重複送出（同一個 submissionKey 6 小時內只寫入一次）
    const cache = CacheService.getScriptCache();
    const key = p.submissionKey ? 'sub_' + String(p.submissionKey).slice(0, 80) : '';
    if (key) {
      const prev = cache.get(key);
      if (prev) return { ok: true, scheduleId: prev };
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (key) {
        const prev = cache.get(key);
        if (prev) return { ok: true, scheduleId: prev };
      }
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const courses = ss.getSheetByName(SHEET_COURSES);
      const lessons = ss.getSheetByName(SHEET_LESSONS);
      const holidays = loadHolidays_(ss.getSheetByName(SHEET_HOLIDAYS));

      const now = new Date();
      const scheduleId = 'SC-' + today_() + '-' +
        Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
      const dateObjs = dates.map(toNoon_);
      const sorted = dates.slice().sort();

      // 課程總表
      const courseRow = [
        scheduleId, now, safeText_(teacherName), safeText_(courseName), totalHours, lessonCount,
        toNoon_(sorted[0]), toNoon_(sorted[sorted.length - 1]),
        '待開始', safeText_(notes), '網站', now
      ];
      const cRow = lastDataRow_(courses) + 1;
      courses.getRange(cRow, 1, 1, courseRow.length).setValues([courseRow]);

      // 堂次排程
      const lessonRows = dates.map((d, i) => [
        scheduleId, safeText_(teacherName), safeText_(courseName), i,
        dateObjs[i], toNoon_(checkDate_(d, holidays)),
        '', '尚未檢核', '', '', now
      ]);
      const lRow = lastDataRow_(lessons) + 1;
      lessons.getRange(lRow, 1, lessonRows.length, lessonRows[0].length).setValues(lessonRows);
      lessons.getRange(lRow, 5, lessonRows.length, 2).setNumberFormat('yyyy/mm/dd');
      lessons.getRange(lRow, 11, lessonRows.length, 1).setNumberFormat('yyyy/m/d');

      SpreadsheetApp.flush();
      if (key) cache.put(key, scheduleId, 21600);
      return { ok: true, scheduleId: scheduleId };
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    console.error(e);
    return fail_('資料暫時無法送出，請稍後再試。');
  }
}

/* ---------- 工具函式 ---------- */

/** 避免使用者輸入被試算表當成公式執行 */
function safeText_(v) { const t = String(v || ''); return /^[=+\-@]/.test(t) ? "'" + t : t; }

function fail_(message) { return { ok: false, message: message }; }

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyyMMdd'); }

function isValidDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const [y, m, d] = s.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
}

/** 'yyyy-MM-dd' → 臺灣時間當日中午 12:00（與現有資料一致） */
function toNoon_(s) { return Utilities.parseDate(s + ' 12:00', TZ, 'yyyy-MM-dd HH:mm'); }

/** 讀取國定假日（是否停班 = 是），回傳 Set('yyyy-MM-dd') */
function loadHolidays_(sheet) {
  const set = new Set();
  if (!sheet) return set;
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const [date, , , off] = values[i];
    if (!date || String(off).trim() !== '是') continue;
    const key = date instanceof Date ? Utilities.formatDate(date, TZ, 'yyyy-MM-dd')
      : String(date).trim().replace(/\//g, '-').replace(/-(\d)(?=-|$)/g, '-0$1');
    set.add(key);
  }
  return set;
}

/** 純日期運算（以 UTC 表示日曆日，避免時區誤差） */
function addDays_(s, n) {
  const [y, m, d] = s.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d + n));
  return dt.toISOString().slice(0, 10);
}
function weekday_(s) {
  const [y, m, d] = s.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d)).getUTCDay(); // 0=日 … 5=五 6=六
}
function isWorkday_(s, holidays) {
  const w = weekday_(s);
  return w !== 0 && w !== 6 && !holidays.has(s);
}

/** 工作檢核日：下一個工作日；週五交付則取第 3 個工作日 */
function checkDate_(s, holidays) {
  let need = weekday_(s) === 5 ? 3 : 1;
  let cur = s;
  while (need > 0) {
    cur = addDays_(cur, 1);
    if (isWorkday_(cur, holidays)) need--;
  }
  return cur;
}

/** 以 A 欄（排程編號）判斷最後一筆資料列，避開已設定下拉選單的空白列 */
function lastDataRow_(sheet) {
  const n = sheet.getLastRow();
  if (n < 2) return 1;
  const col = sheet.getRange(1, 1, n, 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) if (String(col[i][0]).trim() !== '') return i + 1;
  return 1;
}

/** 手動測試工作檢核日規則（在編輯器中執行，查看記錄） */
function testCheckDates() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const h = loadHolidays_(ss.getSheetByName(SHEET_HOLIDAYS));
  ['2026-08-07', '2026-09-25', '2026-09-26', '2026-11-27', '2026-12-24'].forEach(d =>
    console.log(d, '→', checkDate_(d, h)));
}
