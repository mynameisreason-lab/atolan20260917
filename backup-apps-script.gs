/**
 * J人秘書 雲端備份 — Google Apps Script
 *
 * 用途：接收網站送來的商品／客戶行程單資料，寫入這個試算表。
 *
 * 設定步驟：
 * 1. 在 Google Drive 新增一個「Google 試算表」，命名例如「都蘭國資料備份」。
 * 2. 打開這個試算表 → 上方選單「擴充功能」→「Apps Script」。
 * 3. 把這個檔案的全部內容貼進去（取代預設的 myFunction 範例），儲存。
 * 4. 右上角「部署」→「新增部署作業」：
 *    - 類型選「網頁應用程式」
 *    - 執行身份：我
 *    - 存取權：任何人
 *    - 按「部署」，過程中會要求你授權，選你自己的帳號同意即可。
 * 5. 部署完成會給你一個網址（結尾是 /exec），把它複製起來。
 * 6. 回到 J人秘書網站 → 管理後台 →「雲端備份」分頁，貼上這個網址並儲存。
 *
 * 之後每次點「立即備份」，或客人送出新的行程需求時，資料就會自動寫進這個試算表裡，
 * 分成「Products」和「Orders」兩個工作表，每次備份都會用最新資料整份覆蓋（等於一份即時鏡像）。
 */

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (body.type === 'products') {
      writeSheet(ss, 'Products', body.data, [
        'id', 'name', 'category', 'duration', 'priceAdult', 'priceChild',
        'icon', 'tags', 'desc', 'active', 'image'
      ]);
    } else if (body.type === 'orders') {
      writeSheet(ss, 'Orders', body.data, [
        'id', 'refCode', 'createdAt', 'visitDate', 'days', 'mode',
        'adults', 'children', 'total', 'status', 'contactName', 'contactPhone', 'note', 'items'
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function writeSheet(ss, name, rows, cols) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear();
  sheet.appendRow(cols);

  rows.forEach(function (row) {
    const line = cols.map(function (c) {
      const v = row[c];
      if (Array.isArray(v)) return JSON.stringify(v);
      if (v === undefined || v === null) return '';
      return v;
    });
    sheet.appendRow(line);
  });

  sheet.appendRow([]);
  sheet.appendRow(['最後備份時間', new Date().toLocaleString('zh-TW')]);
}
