/**
 * BaF 圖紙索引同步 - Google Apps Script 端
 * （實際版本以程式內的 apiVersion 為準，會顯示在 Revit 工具的匯出訊息）
 * =========================================================
 * 特色：
 *   - 認「表頭名稱」寫入：哪一欄叫「圖紙類別」就寫到那欄，欄位順序可自由調換。
 *   - 空白頁籤自動建立表頭（只寫『值』，不套任何顏色；格式交給專案建築師）。
 *   - 狀態欄自動套核取方塊；指定欄自動套下拉選單；表頭自動鎖定（只有擁有者能改）。
 *   - 寫入前自動拆合併格、移除篩選器、補列數。
 *   - 可選擇依「圖紙類別」拆成多個工作表（action='split'）。
 *
 * 設定：擴充功能→Apps Script→貼上覆蓋→存檔→部署(管理部署作業→編輯→新版本)。
 */

const SECRET = '';  // 留空 = 不檢查密碼

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (SECRET && body.secret !== SECRET) return json({ ok: false, error: '密鑰錯誤' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tabName = body.tab || '圖紙索引';
    const labels = body.headerLabels || [];
    const records = body.records || [];
    const checkboxLabel = body.checkboxLabel || '是否為Revit出圖';
    const dropdownLabels = body.dropdownLabels || ['繪圖員'];

    // 讀取（匯入用）
    if (body.action === 'read') {
      const rsh = ss.getSheetByName(tabName);
      if (!rsh) return json({ ok: false, error: '找不到頁籤: ' + tabName });
      return json(readSheet(rsh, labels));
    }

    if (!labels.length) return json({ ok: false, error: '缺 headerLabels' });

    // 依圖紙類別拆分
    if (body.action === 'split') {
      const made = splitByCategory(ss, tabName, records, labels, checkboxLabel,
                                   dropdownLabels, body.updateDate, body.lockHeader);
      return json({ ok: true, apiVersion: 'v11', splitTabs: made });
    }

    // 一般匯出：寫主頁籤（並標記為「總表」）
    let sh = ss.getSheetByName(tabName);
    if (!sh) sh = ss.insertSheet(tabName);
    const res = writeIndexToSheet(sh, records, labels, checkboxLabel,
                                  dropdownLabels, body.updateDate, body.lockHeader);
    setRole(sh, 'main');
    return json({
      ok: true, apiVersion: 'v11', tab: tabName, headerRow: res.headerRow,
      wrote: records.length, checkboxCol: res.checkboxCol,
      checkboxLabel: checkboxLabel, locked: res.locked, note: res.note
    });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// 把一批 records 寫進指定工作表（值＋核取方塊＋下拉＋鎖表頭；不套顏色）
function writeIndexToSheet(sh, records, labels, checkboxLabel, dropdownLabels, updateDate, lockHeader, defaultHeaderRow) {
  const nCols = 8;
  let note = '';

  try { if (sh.getFilter()) sh.getFilter().remove(); } catch (e0) {}
  try { trimLeadingBlankCols(sh, labels); } catch (eT) {}  // 自動移除表頭左側空白欄

  // 找表頭列
  const scan = Math.min(10, sh.getMaxRows());
  const lastCol = Math.max(sh.getLastColumn(), nCols);
  let headerRow = 0, bestHit = 0, headerVals = null;
  for (let r = 1; r <= scan; r++) {
    const vals = sh.getRange(r, 1, 1, lastCol).getValues()[0];
    let hit = 0;
    for (let c = 0; c < vals.length; c++) if (labels.indexOf(String(vals[c]).trim()) >= 0) hit++;
    if (hit > bestHit) { bestHit = hit; headerRow = r; headerVals = vals; }
  }
  if (!headerRow) {
    // 自動建立表頭（從 A 欄起、第1列放更新時間）— 只寫值不套色
    headerRow = defaultHeaderRow || 2;
    const startCol = 1;
    sh.getRange(headerRow, startCol, 1, labels.length).setValues([labels]);
    sh.getRange(1, startCol).setValue('更新時間：');
    headerVals = sh.getRange(headerRow, 1, 1,
      Math.max(lastCol, startCol + labels.length - 1)).getValues()[0];
  }

  // label -> 欄
  const colOf = {};
  for (let c = 0; c < headerVals.length; c++) {
    const t = String(headerVals[c]).trim();
    if (t && colOf[t] === undefined && labels.indexOf(t) >= 0) colOf[t] = c + 1;
  }
  const missing = labels.filter(function (l) { return colOf[l] === undefined; });
  if (missing.length) note += '找不到表頭: ' + missing.join('/') + '; ';

  const dataStart = headerRow + 1;
  const nData = records.length;

  const needRows = dataStart - 1 + nData;
  if (sh.getMaxRows() < needRows) sh.insertRowsAfter(sh.getMaxRows(), needRows - sh.getMaxRows());

  // 清舊資料(只清有對應的欄)
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

  if (nData > 0) {
    // 寫值
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

    // 核取方塊
    const cStatus = colOf[checkboxLabel];
    if (cStatus) {
      try {
        const cRange = sh.getRange(dataStart, cStatus, nData, 1);
        cRange.insertCheckboxes();
        cRange.setValues(records.map(function (rec) {
          const v = rec[checkboxLabel];
          return [v === true || String(v).toUpperCase() === 'TRUE'];
        }));
      } catch (e2) { note += '核取方塊失敗:' + e2 + '; '; }
    }

    // 下拉
    dropdownLabels.forEach(function (l) {
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

  // 更新時間
  if (updateDate) {
    try {
      const lc = Math.max(sh.getLastColumn(), nCols);
      const loc = findCell(sh, '更新時間', scan, lc);
      if (loc) sh.getRange(loc.row, Math.min(loc.col + 3, lc)).setValue(updateDate);
    } catch (e4) {}
  }

  // 凍結表頭列（往下捲不會跑掉）
  try { sh.setFrozenRows(headerRow); } catch (eFz) {}

  // 鎖表頭
  let locked = false;
  if (lockHeader) {
    try { locked = lockHeaderRows(sh, headerRow, Math.max(sh.getLastColumn(), nCols)); }
    catch (e5) { note += '鎖表頭失敗:' + e5 + '; '; }
  }

  return { headerRow: headerRow, checkboxCol: (colOf[checkboxLabel] || 0), locked: locked, note: note };
}

// 依「圖紙類別」拆成多個工作表，每類一個分頁、以類別命名
function splitByCategory(ss, mainTabName, records, labels, checkboxLabel, dropdownLabels, updateDate, lockHeader) {
  const groups = {};
  const order = [];
  records.forEach(function (rec) {
    const cat = String(rec['圖紙類別'] || '').trim();
    if (!cat) return;            // 沒有類別的不拆
    if (cat === mainTabName) return;  // 不要覆蓋到總表
    if (!groups[cat]) { groups[cat] = []; order.push(cat); }
    groups[cat].push(rec);
  });
  const made = [];
  order.forEach(function (cat) {
    try {
      let sh = ss.getSheetByName(cat);
      if (!sh) sh = ss.insertSheet(cat);
      // 取該類別的批次參數(取第一筆，假設同類別一致)
      const bp = {};
      BAF_BATCH_LABELS.forEach(function (lbl) {
        bp[lbl] = (groups[cat][0] && groups[cat][0][lbl] != null) ? groups[cat][0][lbl] : '';
      });
      writeSubSheet(sh, groups[cat], labels, checkboxLabel, dropdownLabels,
                    updateDate, lockHeader, bp);
      setRole(sh, 'sub');         // 標記為「附表」
      made.push(cat);
    } catch (eSplit) { /* 分頁名稱不合法等→略過該類 */ }
  });
  return made;
}

// 分表：整頁乾淨重建。版面(從 A 欄起)：
//   第1列 更新時間；第2~6列 批次參數(label在A、值在B，可編輯)；空一列；欄位表頭；資料。
//   表頭(含批次參數)全部凍結；只鎖「欄位表頭那一列」，批次參數值仍可編輯。
function writeSubSheet(sh, records, labels, checkboxLabel, dropdownLabels, updateDate, lockHeader, batchParams) {
  // 整頁清乾淨：值＋資料驗證(核取方塊/下拉)都清掉，避免舊版面殘留
  const maxR = sh.getMaxRows(), maxC = sh.getMaxColumns();
  sh.getRange(1, 1, maxR, maxC).clearContent().clearDataValidations();
  try { if (sh.getFilter()) sh.getFilter().remove(); } catch (e0) {}
  const prots = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (let i = 0; i < prots.length; i++) {
    if (prots[i].getDescription() === 'BaF表頭鎖定') prots[i].remove();
  }

  // 第1列：更新時間
  sh.getRange(1, 1).setValue('更新時間：');
  if (updateDate) sh.getRange(1, 2).setValue(updateDate);

  // 批次參數區（label 在 A、值在 B）
  let r = 2;
  BAF_BATCH_LABELS.forEach(function (lbl) {
    sh.getRange(r, 1).setValue(lbl + '：');
    sh.getRange(r, 2).setValue((batchParams && batchParams[lbl] != null) ? batchParams[lbl] : '');
    r++;
  });
  const headerRow = r + 1;   // 空一列後放欄位表頭

  // 欄位表頭（從 A 欄起）
  sh.getRange(headerRow, 1, 1, labels.length).setValues([labels]);
  const colOf = {};
  for (let c = 0; c < labels.length; c++) colOf[labels[c]] = c + 1;

  const dataStart = headerRow + 1;
  const nData = records.length;
  if (nData > 0) {
    labels.forEach(function (l) {
      const c = colOf[l];
      const colVals = records.map(function (rec) {
        let v = rec[l]; if (v === undefined || v === null) v = ''; return [v];
      });
      sh.getRange(dataStart, c, nData, 1).setValues(colVals);
    });
    const cStatus = colOf[checkboxLabel];
    if (cStatus) {
      try {
        const cRange = sh.getRange(dataStart, cStatus, nData, 1);
        cRange.insertCheckboxes();
        cRange.setValues(records.map(function (rec) {
          const v = rec[checkboxLabel];
          return [v === true || String(v).toUpperCase() === 'TRUE'];
        }));
      } catch (e2) {}
    }
    dropdownLabels.forEach(function (l) {
      const c = colOf[l];
      if (!c) return;
      try {
        const opts = distinct(records.map(function (rec) { return (rec[l] || '').toString().trim(); }));
        if (opts.length) {
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(opts, true).setAllowInvalid(true).build();
          sh.getRange(dataStart, c, nData, 1).setDataValidation(rule);
        }
      } catch (e3) {}
    });
  }

  // 凍結到欄位表頭列（更新時間＋批次參數＋欄位表頭都算表頭）
  try { sh.setFrozenRows(headerRow); } catch (eFz) {}

  // 只鎖欄位表頭那一列（批次參數值仍可編輯）
  if (lockHeader) {
    try {
      const p = sh.getRange(headerRow, 1, 1, labels.length)
        .protect().setDescription('BaF表頭鎖定');
      const emails = p.getEditors().map(function (u) { return u.getEmail(); });
      if (emails.length) p.removeEditors(emails);
      if (p.canDomainEdit && p.canDomainEdit()) p.setDomainEdit(false);
    } catch (eL) {}
  }
  return headerRow;
}

function readSheet(sh, labels) {
  const scan = Math.min(10, sh.getMaxRows());
  const lastCol = Math.max(sh.getLastColumn(), 8);
  let headerRow = 0, bestHit = 0, headerVals = null;
  for (let r = 1; r <= scan; r++) {
    const vals = sh.getRange(r, 1, 1, lastCol).getValues()[0];
    let hit = 0;
    for (let c = 0; c < vals.length; c++) if (labels.indexOf(String(vals[c]).trim()) >= 0) hit++;
    if (hit > bestHit) { bestHit = hit; headerRow = r; headerVals = vals; }
  }
  if (!headerRow) return { ok: false, error: '找不到表頭列' };
  const colOf = {};
  for (let c = 0; c < headerVals.length; c++) {
    const t = String(headerVals[c]).trim();
    if (t && colOf[t] === undefined && labels.indexOf(t) >= 0) colOf[t] = c + 1;
  }
  const dataStart = headerRow + 1;
  const lastRow = sh.getLastRow();
  const records = [];
  const n = lastRow - dataStart + 1;
  if (n > 0) {
    const block = sh.getRange(dataStart, 1, n, lastCol).getValues();
    block.forEach(function (row) {
      const rec = {};
      labels.forEach(function (l) {
        const c = colOf[l];
        rec[l] = c ? row[c - 1] : '';
      });
      const u = String(rec['UID'] || '').trim();
      const num = String(rec['圖紙號碼'] || '').trim();
      const nm = String(rec['圖紙名稱'] || '').trim();
      if (u || num || nm) records.push(rec);
    });
  }
  const role = getRole(sh);
  const category = (role === 'sub') ? sh.getName() : '';
  const batchParams = {};
  if (role === 'sub') {
    BAF_BATCH_LABELS.forEach(function (lbl) {
      const loc = findCell(sh, lbl, 10, lastCol);
      batchParams[lbl] = loc ? sh.getRange(loc.row, loc.col + 1).getValue() : '';
    });
  }
  return {
    ok: true, headerRow: headerRow, records: records,
    role: role, category: category, batchParams: batchParams,
    mainTab: findMainTabName(sh.getParent())
  };
}

// 移除表頭左側「整欄空白」的欄，讓表頭從 A 欄開始（配色會跟著左移、不掉）
function trimLeadingBlankCols(sh, labels) {
  for (let guard = 0; guard < 5; guard++) {
    if (sh.getMaxColumns() < 2) break;
    const scan = Math.min(10, sh.getMaxRows());
    const lastCol = Math.max(sh.getLastColumn(), 8);
    let minLabelCol = 0;
    for (let r = 1; r <= scan && !minLabelCol; r++) {
      const vals = sh.getRange(r, 1, 1, lastCol).getValues()[0];
      for (let c = 0; c < vals.length; c++) {
        if (labels.indexOf(String(vals[c]).trim()) >= 0) { minLabelCol = c + 1; break; }
      }
    }
    if (minLabelCol <= 1) break;           // 表頭已在 A 欄(或找不到表頭)
    // 第1欄整欄(資料範圍)皆空才刪
    const colA = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues();
    let empty = true;
    for (let i = 0; i < colA.length; i++) {
      if (String(colA[i][0]).trim() !== '') { empty = false; break; }
    }
    if (!empty) break;
    sh.deleteColumn(1);
  }
}

function findMainTabName(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (getRole(sheets[i]) === 'main') return sheets[i].getName();
  }
  return '';
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
  const prots = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (let i = 0; i < prots.length; i++) {
    if (prots[i].getDescription() === desc) prots[i].remove();
  }
  const rng = sh.getRange(headerRow, 1, 1, lastCol);  // 只鎖欄位表頭那一列
  const p = rng.protect().setDescription(desc);
  try {
    const emails = p.getEditors().map(function (u) { return u.getEmail(); });
    if (emails.length) p.removeEditors(emails);
  } catch (ee) {}
  try { if (p.canDomainEdit && p.canDomainEdit()) p.setDomainEdit(false); } catch (ee2) {}
  return true;
}

// ---- 角色標記(總表/附表)：用隱形的 DeveloperMetadata ----
const BAF_ALL_LABELS = ['UID', '是否為Revit出圖', '圖紙類別', '圖紙號碼',
                        '圖紙名稱', '繪圖員', '修正備註'];
// 分表標頭的「整批共用」參數（每個分表填一次，匯入時套用到該類別所有圖紙）
const BAF_BATCH_LABELS = ['審圖員', '設計者', '批准者', '圖紙發布日期', '校核2'];

function setRole(sh, role) {
  try {
    const md = sh.getDeveloperMetadata();
    for (let i = 0; i < md.length; i++) {
      if (md[i].getKey() === 'BAF_ROLE') md[i].remove();
    }
    sh.addDeveloperMetadata('BAF_ROLE', role);
  } catch (e) {}
}

function getRole(sh) {
  try {
    const md = sh.getDeveloperMetadata();
    for (let i = 0; i < md.length; i++) {
      if (md[i].getKey() === 'BAF_ROLE') return md[i].getValue();
    }
  } catch (e) {}
  return '';
}

function headerInfo(sh) {
  const scan = Math.min(10, sh.getMaxRows());
  const lastCol = Math.max(sh.getLastColumn(), 8);
  for (let r = 1; r <= scan; r++) {
    const vals = sh.getRange(r, 1, 1, lastCol).getValues()[0];
    const colOf = {}, labelByCol = {};
    let hit = 0;
    for (let c = 0; c < vals.length; c++) {
      const t = String(vals[c]).trim();
      if (BAF_ALL_LABELS.indexOf(t) >= 0 && colOf[t] === undefined) {
        colOf[t] = c + 1; labelByCol[c + 1] = t; hit++;
      }
    }
    if (hit >= 3) return { headerRow: r, colOf: colOf, labelByCol: labelByCol };
  }
  return null;
}

// 編輯「附表」時，自動把同一張圖(以 UID 對應)的值寫回「總表」
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (getRole(sh) !== 'sub') return;          // 只處理附表的編輯
    const ss = e.source;
    let main = null;
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (getRole(sheets[i]) === 'main') { main = sheets[i]; break; }
    }
    if (!main) return;
    const subInfo = headerInfo(sh);
    const mainInfo = headerInfo(main);
    if (!subInfo || !mainInfo) return;
    const subUidCol = subInfo.colOf['UID'];
    if (!subUidCol || !mainInfo.colOf['UID']) return;

    // 主表 UID -> 列
    const mMap = {};
    const mLast = main.getLastRow();
    const mStart = mainInfo.headerRow + 1;
    if (mLast >= mStart) {
      const us = main.getRange(mStart, mainInfo.colOf['UID'], mLast - mStart + 1, 1).getValues();
      for (let i = 0; i < us.length; i++) {
        const u = String(us[i][0] || '').trim();
        if (u) mMap[u] = mStart + i;
      }
    }

    const r0 = e.range.getRow(), c0 = e.range.getColumn();
    const nR = e.range.getNumRows(), nC = e.range.getNumColumns();
    for (let rr = 0; rr < nR; rr++) {
      const row = r0 + rr;
      if (row <= subInfo.headerRow) continue;
      const uid = String(sh.getRange(row, subUidCol).getValue() || '').trim();
      if (!uid) continue;
      const mrow = mMap[uid];
      if (!mrow) continue;
      for (let cc = 0; cc < nC; cc++) {
        const col = c0 + cc;
        const label = subInfo.labelByCol[col];
        if (!label || label === 'UID') continue;   // UID 不改
        const mcol = mainInfo.colOf[label];
        if (!mcol) continue;
        main.getRange(mrow, mcol).setValue(sh.getRange(row, col).getValue());
      }
    }
  } catch (err) { /* onEdit 不可拋錯，避免影響使用者編輯 */ }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
