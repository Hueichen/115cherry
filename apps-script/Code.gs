/**
 * 115櫻桃班 公告欄：資料中介程式（Google Apps Script）
 *
 * 用途：試算表維持「限制」（不公開），由這支程式讀取指定分頁，
 *       只回傳網站需要的欄位，給 GitHub Pages 上的網站使用。
 *
 * 安裝：試算表 → 擴充功能 → Apps Script → 把這整份貼上 → 部署成「網頁應用程式」
 * 注意：修改程式後，要到「部署 → 管理部署作業 → 編輯 → 版本：新版本」才會生效。
 */

// 分頁名稱 → 網站使用的代號，以及「允許公開」的欄位（白名單）
// 沒列在這裡的分頁、欄位，一律不會被送出去（可放心在試算表加內部備註欄）
const SHEETS = {
  weekly:   { name: '週報',   fields: ['週次', '起', '週期', '迄', '櫻桃成長記', '櫻桃小須知'] },
  songs:    { name: '兒歌',   fields: ['週次', '名稱', '內容', '影片網址'] },
  rhymes:   { name: '手指謠', fields: ['週次', '名稱', '內容', '影片網址'] },
  gallery:  { name: '畫廊',   fields: ['週次', '日期', '標題', '圖片網址', '說明'] },
  books:    { name: '繪本',   fields: ['週次', '書名', '作者', '封面網址', '介紹'] },
  calendar: { name: '行事曆', fields: ['日期', '活動', '備註'] },
  notices:  { name: '須知',   fields: ['類別', '標題', '內容'] }
};

// 快取秒數：改了試算表後，最多等這麼久網站就會更新（0 = 不快取）
const CACHE_SECONDS = 60;

function doGet() {
  const cache = CacheService.getScriptCache();
  const KEY = 'cherry-data-v1';
  try {
    let body = CACHE_SECONDS > 0 ? cache.get(KEY) : null;
    if (!body) {
      body = JSON.stringify({ ok: true, updated: new Date().toISOString(), data: readAll_() });
      // 單一快取項目上限約 100KB，超過就不快取，直接回傳
      if (CACHE_SECONDS > 0 && body.length < 95000) cache.put(KEY, body, CACHE_SECONDS);
    }
    return json_(body);
  } catch (err) {
    return json_(JSON.stringify({ ok: false, error: String(err && err.message || err) }));
  }
}

function readAll_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = {};
  Object.keys(SHEETS).forEach(key => {
    const conf = SHEETS[key];
    const sheet = ss.getSheetByName(conf.name);
    if (!sheet) { out[key] = []; return; }               // 分頁不存在就回傳空的
    const values = sheet.getDataRange().getDisplayValues(); // 用「畫面上看到的文字」，日期不會跑時區
    if (values.length < 2) { out[key] = []; return; }
    const head = values[0].map(h => String(h).trim());
    const cols = conf.fields
      .map(f => ({ f, i: head.indexOf(f) }))
      .filter(c => c.i >= 0);
    out[key] = values.slice(1)
      .filter(row => row.some(v => String(v).trim() !== ''))
      .map(row => {
        const o = {};
        cols.forEach(c => { o[c.f] = String(row[c.i]).trim(); });
        return o;
      });
  });
  return out;
}

function json_(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

/** 在編輯器選這個函式按「執行」，可以先檢查讀到的資料（結果看「執行記錄」） */
function testRead() {
  Logger.log(JSON.stringify(readAll_(), null, 2));
}
