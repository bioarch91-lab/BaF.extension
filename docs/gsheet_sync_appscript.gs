/**
 * BaF 圖紙索引同步 - Google Apps Script 端 (v3)
 * =========================================================
 * 特色：
 *   - 認「表頭名稱」寫入：哪一欄的標題叫「圖紙類別」就寫到那欄，可自由調換欄位順序。
 *   - 寫入時自動鎖定表頭列（受保護範圍）：只有試算表擁有者(管理者)能改。
 *   - 狀態欄自動套核取方塊；圖紙類別/繪圖員自動套下拉選單。
 *   - 寫入前自動拆合併格、移除篩選器、補足列數。
 *
 * 設定步驟（每個試算表做一次）：
 *  1. Google Sheet →「擴充功能」→「Apps Script」
 *  2. 把本檔內容全部貼上、覆蓋舊內容，存檔
 *  3. （選用）要密碼才填 SECRET，否則留空
 *  4. 「部署」→「管理部署作業」→ 編輯(鉛筆) → 版本選「新版本」→ 部署（URL 不變）
 *     ※ 第一次用到「保護範圍」可能會再要求授權一次，按允許即可。
 *
 * 收到的 JSON：
 *   { secret, tab, headerLabels:[...], records:[{欄名:值,...}], updateDate, lockHeader }
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

    const labels = body.headerLabels || [];
    const records = body.records || [];
    if (!labels.length) return json({ ok: false, error: '缺 headerLabels' });

    // 0) 移除篩選器
    try { if (sh.getFilter()) sh.getFilter().remove(); } catch (e0) {}

    // 1) 找表頭列：掃前 10 列，找包含最多 label 的那一列
    const scan = Math.min(10, sh.getMaxRows());
    const lastCol = Math.max(sh.getLastColumn(), 8);
    let headerRow = 0, bestHit = 0, headerVals = null;
    for (let r = 1; r <= scan; r++) {
      const vals = sh.getRange(r, 1, 1, lastCol).getValues()[0];
      let hit = 0;
      for (let c = 0; c < vals.length; c++) {
        if (labels.indexOf(String(vals[c]).trim()) >= 0) hit++;
      }
      if (hit > bestHit) { bestHit = hit; headerRow = r; headerVals = vals; }
    }
    if (!headerRow) return json({ ok: false, error: '找不到表頭列(請確認頁籤有 UID / 圖紙號碼 等標題)' });

    // 2) 表頭文字 -> 欄索引(1-based)
    const colOf = {};
    for (let c = 0; c < headerVals.length; c++) {
      const t = String(headerVals[c]).trim();
      if (t && colOf[t] === undefined && labels.indexOf(t) >= 0) colOf[t] = c + 1;
    }
    const missing = labels.filter(function (l) { return colOf[l] === undefined; });

    const dataStart = headerRow + 1;
    const nData = records.length;

    // 確保列數足夠
    const needRows = dataStart - 1 + nData;
    if (sh.getMaxRows() < needRows) sh.insertRowsAfter(sh.getMaxRows(), needRows - sh.getMaxRows());

    // 3) 清掉舊資料(只清有對應到的欄)、拆合併、清舊驗證
    const usedRows = sh.getMaxRows() - dataStart + 1;
    if (usedRows > 0) {
      labels.forEach(function (l) {
        const c = colOf[l];
        if (!c) return;
        const rng = sh.getRange(dataStart, c, usedRows, 1);
        try { rng.breakApart(); } catch (e1) {}
        rng.clearContent();
        rng.clearDataValidations();
      });
    }

    // 4) 逐欄整批寫入資料
    if (nData > 0) {
      labels.forEach(function (l) {
        const c = colOf[l];
        if (!c) return;
        const colVals = records.map(function (rec) {
          let v = rec[l];
          if (v === undefined || v === null) v = '';
          return [v];
        });
        sh.getRange(dataStart, c, nData, 1).setValues(colVals);
      });
    }

    // 5) 狀態 -> 核取方塊；圖紙類別/繪圖員 -> 下拉選單
    let note = '';
    if (missing.length) note += '找不到表頭: ' + missing.join('/') + '; ';
    if (nData > 0) {
      const cStatus = colOf['狀態'];
      if (cStatus) {
        try { sh.getRange(dataStart, cStatus, nData, 1).insertCheckboxes(); }
        catch (e2) { note += '核取方塊失敗:' + e2 + '; '; }
      }
      ['圖紙類別', '繪圖員'].forEach(function (l) {
        const c = colOf[l];
        if (!c) return;
        try {
          const opts = distinct(records.map(function (rec) { return (rec[l] || '').toString().trim(); }));
          if (opts.length) {
            const rule = SpreadsheetApp.newDataValidation()
              .requireValueInList(opts, true).setAllowInvalid(true).build();
            sh.getRange(dataStart, c, nData, 1).setDataValidation(rule);
          }
        } catch (e3) { note += l + '下拉失敗:' + e3 + '; '; }
      });
    }

    // 6) 更新時間（找「更新時間」字樣，日期寫到它右邊第 3 欄，沿用 B→E 排版）
    if (body.updateDate) {
      try {
        const loc = findCell(sh, '更新時間', scan, lastCol);
        if (loc) sh.getRange(loc.row, Math.min(loc.col + 3, lastCol)).setValue(body.updateDate);
      } catch (e4) {}
    }

    // 7) 鎖定表頭列（只有擁有者=管理者可改）
    let locked = false;
    if (body.lockHeader) {
      try { locked = lockHeaderRows(sh, headerRow, lastCol); }
      catch (e5) { note += '鎖表頭失敗:' + e5 + '; '; }
    }

    return json({ ok: true, tab: tabName, headerRow: headerRow, wrote: nData, locked: locked, note: note });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function distinct(arr) {
  const seen = {}, out = [];
  arr.forEach(function (v) { if (v && !seen[v]) { seen[v] = true; out.push(v); } });
  return out;
}

function findCell(sh, text, scan, lastCol) {
  const rmax = Math.min(scan, sh.getMaxRows());
  for (let r = 1; r <= rmax; r++) {
    const vals = sh.getRange(r, 1, 1, lastCol).getValues()[0];
    for (let c = 0; c < vals.length; c++) {
      if (String(vals[c]).indexOf(text) >= 0) return { row: r, col: c + 1 };
    }
  }
  return null;
}

function lockHeaderRows(sh, headerRow, lastCol) {
  const desc = 'BaF表頭鎖定';
  // 移除舊的同名保護，避免重複
  const prots = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (let i = 0; i < prots.length; i++) {
    if (prots[i].getDescription() === desc) prots[i].remove();
  }
  // 鎖第 1 列(更新時間) 到 表頭列。protect() 預設只保留「目前的編輯者」，
  // 之後新加入的協作者預設不能編輯此範圍 → 達到「只有管理者能改表頭」。
  const rng = sh.getRange(1, 1, headerRow, lastCol);
  const p = rng.protect().setDescription(desc);
  // 進一步移除其他編輯者，只留擁有者；若沒權限讀 email 就略過(保護仍生效)。
  try {
    const emails = p.getEditors().map(function (u) { return u.getEmail(); });
    if (emails.length) p.removeEditors(emails);
  } catch (ee) {}
  try { if (p.canDomainEdit && p.canDomainEdit()) p.setDomainEdit(false); } catch (ee2) {}
  return true;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
