/**
 * BaF 圖紙索引同步 - Google Apps Script 端
 * =========================================================
 * 用途：接收 Revit(pyRevit) 傳來的圖紙索引，寫入指定頁籤，
 *       並自動套用「狀態」核取方塊與「圖紙類別 / 繪圖員」下拉選單。
 *
 * 設定步驟（每個試算表做一次）：
 *  1. 打開你的 Google Sheet → 上方「擴充功能」→「Apps Script」
 *  2. 把本檔內容全部貼上、覆蓋原本的 Code.gs，存檔
 *  3. （選用）SECRET 留空字串 '' = 不檢查密碼，靠 URL 保密即可；要加保險才填密碼
 *  4. 右上「部署」→「新增部署作業」→ 齒輪選「網頁應用程式」
 *       - 說明：隨意
 *       - 執行身分：我
 *       - 誰可以存取：任何人
 *  5. 按「部署」，第一次會要求授權 → 允許
 *  6. 複製「網頁應用程式 URL」(以 /exec 結尾) → 連同「目標頁籤名稱」交給 Revit 端
 *
 * 安全性：URL 形同鑰匙，請勿外流。若要多一層保險可填 SECRET。
 */

const SECRET = '';  // 留空 = 不檢查密碼

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (SECRET && body.secret !== SECRET) return json({ ok: false, error: '密鑰錯誤' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tabName = body.tab || '圖紙索引';
    let sh = ss.getSheetByName(tabName);
    if (!sh) sh = ss.insertSheet(tabName);

    const rows = body.rows || [];   // 二維陣列：第1列=更新時間, 第2列=表頭, 第3列起=資料
    const received = rows.length;
    const nCols = 8;                // A..H
    if (received < 2) return json({ ok: false, error: '沒有資料(received=' + received + ')' });

    // 0) 移除篩選器（會干擾寫入/顯示）
    try { if (sh.getFilter()) sh.getFilter().remove(); } catch (e0) {}

    // 確保列數足夠
    if (sh.getMaxRows() < received) {
      sh.insertRowsAfter(sh.getMaxRows(), received - sh.getMaxRows());
    }

    const writeRange = sh.getRange(1, 1, received, nCols);

    // 1) 拆開寫入範圍內的合併儲存格（避免 setValues 卡住）
    try { writeRange.breakApart(); } catch (e1) {}

    // 2) 清掉舊值與舊資料驗證
    writeRange.clearContent();
    writeRange.clearDataValidations();

    // 3) 寫值（補齊每列到 8 欄）
    const fixed = rows.map(function (r) {
      const rr = r.slice(0, nCols);
      while (rr.length < nCols) rr.push('');
      return rr;
    });
    writeRange.setValues(fixed);
    const wrote = received;

    const dataStart = 3;
    const dataCount = received - 2;
    let note = '';
    if (dataCount > 0) {
      // C 欄(狀態) → 核取方塊（TRUE=勾=真實圖紙）
      try {
        const cRange = sh.getRange(dataStart, 3, dataCount, 1);
        cRange.insertCheckboxes();
        cRange.setValues(rows.slice(2).map(function (r) {
          return [String(r[2]).toUpperCase() === 'TRUE'];
        }));
      } catch (e2) { note += '核取方塊失敗:' + e2 + '; '; }

      // D 欄(圖紙類別) / G 欄(繪圖員) → 下拉選單
      try { applyDropdown(sh, dataStart, 4, rows.slice(2), 3); } catch (e3) { note += '類別下拉失敗:' + e3 + '; '; }
      try { applyDropdown(sh, dataStart, 7, rows.slice(2), 6); } catch (e4) { note += '繪圖員下拉失敗:' + e4 + '; '; }
    }

    return json({ ok: true, tab: tabName, received: received, wrote: wrote, note: note });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function applyDropdown(sh, startRow, col, dataRows, idx) {
  const seen = {};
  const opts = [];
  dataRows.forEach(function (r) {
    const v = (r[idx] == null ? '' : String(r[idx])).trim();
    if (v && !seen[v]) { seen[v] = true; opts.push(v); }
  });
  if (opts.length === 0) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(opts, true)   // true = 顯示下拉箭頭
    .setAllowInvalid(true)            // 允許不在清單內的值（不擋輸入，只是不會有警告）
    .build();
  sh.getRange(startRow, col, dataRows.length, 1).setDataValidation(rule);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
