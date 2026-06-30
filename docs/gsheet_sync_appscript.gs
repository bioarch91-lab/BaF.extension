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
      return json({ ok: true, apiVersion: 'v12', splitTabs: made });
    }

    // 範圍限定同步：只回寫「某一圖紙類別(階段)」→ 取代總表中該類別的列、
    // 只重建該類別分表；其他類別與分表完全不動（分表匯入後的自動回寫用）。
    if (body.action === 'syncCategory') {
      const synced = syncCategory(ss, tabName, body.category, records, labels,
                                  checkboxLabel, dropdownLabels,
                                  body.updateDate, body.lockHeader);
      return json({ ok: true, apiVersion: 'v12', category: body.category,
                    synced: synced });
    }

    // 一般匯出：寫主頁籤（並標記為「總表」）
    let sh = ss.getSheetByName(tabName);
    if (!sh) sh = ss.insertSheet(tabName);
    const res = writeIndexToSheet(sh, records, labels, checkboxLabel,
                                  dropdownLabels, body.updateDate, body.lockHeader);
    setRole(sh, 'main');
    return json({
      ok: true, apiVersion: 'v12', tab: tabName, headerRow: res.headerRow,
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
        if (l === '修正備註') v = String(v).replace(/；/g, '\n');  // Sheet 內換行顯示
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

    // 「已完成」開頭的修正備註條目 → 反灰
    applyDoneGreyByText_(sh, dataStart, records, colOf['修正備註']);
    // 修正備註欄自動換行，讓「；」轉成的換行能完整顯示
    if (colOf['修正備註']) {
      try { sh.getRange(dataStart, colOf['修正備註'], nData, 1).setWrap(true); } catch (eW) {}
    }
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


// 依「已完成」規則，回傳整格「修正備註」的 RichTextValue：
// 以「；」分隔，凡(去前導空白後)以「已完成」開頭的條目塗灰(#999)，其餘明確設黑字(#000)。
// 明確設黑 → 在 Sheet 編輯時可把舊的灰字「重設」成黑字(新內容與舊內容可區分)。
function _noteRich_(full) {
  // 顯示用：把「；」轉成儲存格內換行（Google Sheet 檢視較友善）；依換行切條判斷反灰。
  var disp = String(full == null ? '' : full).replace(/；/g, '\n');
  var black = SpreadsheetApp.newTextStyle().setForegroundColor('#000000').build();
  var grey = SpreadsheetApp.newTextStyle().setForegroundColor('#999999').build();
  var rich = SpreadsheetApp.newRichTextValue().setText(disp).setTextStyle(black);
  var parts = disp.split('\n');
  var pos = 0;
  for (var j = 0; j < parts.length; j++) {
    var seg = parts[j];
    var segStart = pos, segEnd = pos + seg.length;
    if (segEnd > segStart && seg.replace(/^\s+/, '').indexOf('已完成') === 0) {
      rich.setTextStyle(segStart, segEnd, grey);
    }
    pos = segEnd + 1;   // 換行符為 1 字元
  }
  return rich.build();
}

// 匯出時：含「已完成」條目的「修正備註」格依規則上色(其餘列已是純值黑字，略過以省效能)。
function applyDoneGreyByText_(sh, dataStart, records, noteCol) {
  try {
    if (!noteCol) return;
    for (var i = 0; i < records.length; i++) {
      var rec = records[i] || {};
      var full = (rec['修正備註'] == null) ? '' : String(rec['修正備註']);
      if (!full || full.indexOf('已完成') < 0) continue;
      try { sh.getRange(dataStart + i, noteCol).setRichTextValue(_noteRich_(full)); } catch (e1) {}
    }
  } catch (e) {}
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

// 範圍限定同步：只回寫某一圖紙類別(階段)。
//   1) 總表：只取代屬於該類別的列，其他類別的列原封不動。
//   2) 只重建該類別的分表；其他分表完全不碰。
function syncCategory(ss, mainTabName, category, records, labels, checkboxLabel,
                      dropdownLabels, updateDate, lockHeader) {
  // 1) 總表
  let mainSh = ss.getSheetByName(mainTabName);
  if (!mainSh || getRole(mainSh) !== 'main') {
    const mn = findMainTabName(ss);
    if (mn) mainSh = ss.getSheetByName(mn);
  }
  if (mainSh) {
    replaceCategoryRowsInMain(mainSh, category, records, labels, checkboxLabel,
                              dropdownLabels, updateDate, lockHeader);
    setRole(mainSh, 'main');
  }
  // 2) 只重建該類別分表
  if (category) {
    let subSh = ss.getSheetByName(category);
    if (!subSh) subSh = ss.insertSheet(category);
    const bp = {};
    BAF_BATCH_LABELS.forEach(function (lbl) {
      bp[lbl] = (records[0] && records[0][lbl] != null) ? records[0][lbl] : '';
    });
    writeSubSheet(subSh, records, labels, checkboxLabel, dropdownLabels,
                  updateDate, lockHeader, bp);
    setRole(subSh, 'sub');
  }
  return [category];
}

// 在「總表」中只取代某類別的列：保留其他類別的列(用表上現值)，
// 把本類別的列換成傳入的 records(來自 Revit)，順序儘量維持。
function replaceCategoryRowsInMain(sh, category, records, labels, checkboxLabel,
                                   dropdownLabels, updateDate, lockHeader) {
  const existing = readSheet(sh, labels);
  if (!existing || !existing.ok) return false;  // 讀不到表頭 → 不動總表(避免誤刪)
  const combined = [];
  let inserted = false;
  (existing.records || []).forEach(function (r) {
    if (String(r['圖紙類別'] || '').trim() === category) {
      if (!inserted) {                       // 在本類別第一筆的位置插入新列
        records.forEach(function (nr) { combined.push(nr); });
        inserted = true;
      }
      // 丟掉舊的本類別列
    } else {
      combined.push(r);                       // 其他類別維持原值原順序
    }
  });
  if (!inserted) records.forEach(function (nr) { combined.push(nr); });
  writeIndexToSheet(sh, combined, labels, checkboxLabel, dropdownLabels,
                    updateDate, lockHeader);
  return true;
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

  // 固定參數/更新時間放在「圖紙號碼」欄(名稱)與「圖紙名稱」欄(值)的上方，
  // 因為這兩欄永遠不會被隱藏；其他欄位使用者可自行隱藏不受影響。
  const lblCol = (labels.indexOf('圖紙號碼') >= 0) ? labels.indexOf('圖紙號碼') + 1 : 1;
  const valCol = (labels.indexOf('圖紙名稱') >= 0) ? labels.indexOf('圖紙名稱') + 1 : 2;

  // 第1列：更新時間（名稱在「圖紙號碼」欄、值在「圖紙名稱」欄）
  sh.getRange(1, lblCol).setValue('更新時間：');
  if (updateDate) sh.getRange(1, valCol).setValue(updateDate);

  // 特定階段固定參數區（名稱在「圖紙號碼」欄、值在「圖紙名稱」欄、各佔一格）
  let r = 2;
  BAF_BATCH_LABELS.forEach(function (lbl) {
    sh.getRange(r, lblCol).setValue(lbl + '：');
    sh.getRange(r, valCol).setValue((batchParams && batchParams[lbl] != null) ? batchParams[lbl] : '');
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
        let v = rec[l]; if (v === undefined || v === null) v = '';
        if (l === '修正備註') v = String(v).replace(/；/g, '\n');  // Sheet 內換行顯示
        return [v];
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

    // 「已完成」開頭的修正備註條目 → 反灰
    applyDoneGreyByText_(sh, dataStart, records, colOf['修正備註']);
    // 修正備註欄自動換行，讓「；」轉成的換行能完整顯示
    if (colOf['修正備註']) {
      try { sh.getRange(dataStart, colOf['修正備註'], nData, 1).setWrap(true); } catch (eW) {}
    }
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
      // 修正備註：Sheet 內以換行顯示 → 回 Revit 前轉回「；」(Revit 明細表單行友善)
      if (rec['修正備註'] != null) {
        rec['修正備註'] = String(rec['修正備註']).replace(/\r?\n/g, '；');
      }
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
    const valCol = colOf['圖紙名稱'];  // 值放在「圖紙名稱」欄
    BAF_BATCH_LABELS.forEach(function (lbl) {
      const loc = findCell(sh, lbl, 10, lastCol);
      if (loc) {
        const vcol = valCol || (loc.col + 1);
        batchParams[lbl] = sh.getRange(loc.row, vcol).getValue();
      } else {
        batchParams[lbl] = '';
      }
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
// 分表↔總表「值編輯」自動雙向即時同步（依 UID 比對）：
//   編輯分表某格 → 寫回總表對應列；編輯總表某格 → 寫到含該 UID 的分表對應列。
// 不處理整列新增/刪除（onEdit 不會在刪/增列時觸發）；結構變更請在 Revit 端做後匯出。
// 安全：簡單觸發器 onEdit 由「使用者手動編輯」觸發；本函式以程式 setValue() 寫入對方，
//       程式寫入不會再觸發簡單觸發器，故不會無限迴圈。
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const role = getRole(sh);
    if (role !== 'sub' && role !== 'main') return;   // 只處理總表/分表
    const srcInfo = headerInfo(sh);
    if (!srcInfo || !srcInfo.colOf['UID']) return;

    // 修正備註即時上色：使用者在 Sheet 改「修正備註」格 → 依「已完成」規則重上色
    //   (新內容無已完成→黑字、已完成→灰字)，讓新舊內容可區分。不依賴是否有同步目標。
    const _ncol = srcInfo.colOf['修正備註'];
    if (_ncol) {
      const er = e.range;
      const rr0 = er.getRow(), cc0 = er.getColumn();
      const rrn = er.getNumRows(), ccn = er.getNumColumns();
      if (_ncol >= cc0 && _ncol <= cc0 + ccn - 1) {
        for (let k = 0; k < rrn; k++) {
          const rw = rr0 + k;
          if (rw <= srcInfo.headerRow) continue;
          const v = String(sh.getRange(rw, _ncol).getValue() || '');
          if (v) { try { sh.getRange(rw, _ncol).setRichTextValue(_noteRich_(v)); } catch (e0) {} }
        }
      }
    }

    const ss = e.source;
    const sheets = ss.getSheets();

    // 決定要寫入的目標：分表編輯→單一總表；總表編輯→所有分表
    const targets = [];
    for (let i = 0; i < sheets.length; i++) {
      const tr = getRole(sheets[i]);
      if (role === 'sub' && tr === 'main') { targets.push(sheets[i]); break; }
      if (role === 'main' && tr === 'sub') { targets.push(sheets[i]); }
    }
    if (!targets.length) return;

    // 預先建每個目標的 UID -> 列 對照
    const tgt = targets.map(function (t) {
      const info = headerInfo(t);
      const map = {};
      if (info && info.colOf['UID']) {
        const last = t.getLastRow(), start = info.headerRow + 1;
        if (last >= start) {
          const us = t.getRange(start, info.colOf['UID'], last - start + 1, 1).getValues();
          for (let i = 0; i < us.length; i++) {
            const u = String(us[i][0] || '').trim();
            if (u) map[u] = start + i;
          }
        }
      }
      return { sheet: t, info: info, map: map };
    });

    const r0 = e.range.getRow(), c0 = e.range.getColumn();
    const nR = e.range.getNumRows(), nC = e.range.getNumColumns();
    for (let rr = 0; rr < nR; rr++) {
      const row = r0 + rr;
      if (row <= srcInfo.headerRow) continue;
      const uid = String(sh.getRange(row, srcInfo.colOf['UID']).getValue() || '').trim();
      if (!uid) continue;
      for (let cc = 0; cc < nC; cc++) {
        const col = c0 + cc;
        const label = srcInfo.labelByCol[col];
        if (!label || label === 'UID') continue;     // UID 不改
        const val = sh.getRange(row, col).getValue();
        const isNote = (label === '修正備註');   // 來源格已在上面即時上色
        const sval = String(val == null ? '' : val);
        for (let ti = 0; ti < tgt.length; ti++) {
          const td = tgt[ti];
          if (!td.info || !td.info.colOf[label]) continue;
          const trow = td.map[uid];
          if (!trow) continue;                         // 對方沒有這個 UID 就略過
          const tcell = td.sheet.getRange(trow, td.info.colOf[label]);
          if (isNote) {
            if (sval) { try { tcell.setRichTextValue(_noteRich_(sval)); } catch (eT) {} }
            else { tcell.setValue(''); }
          } else {
            tcell.setValue(val);
          }
        }
      }
    }
  } catch (err) { /* onEdit 不可拋錯，避免影響使用者編輯 */ }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====================================================================
// 試算表上的「BaF」自訂選單（外掛）
// 純選單功能，用 appsscript.json「預設權限」(自動偵測)即可，不需 Apps Script API。
// 兩個功能都「鎖 UID」比對：
//   ① 分表 → 總表：用一或多張分表的列，依 UID 覆蓋總表對應列(總表沒有的附加)。
//   ② 總表 → 分表：用總表的列，依 UID 覆蓋多張分表中現有的列(分表列數不變)。
// 下拉只列「同時有 UID／圖紙名稱／圖紙號碼」的工作表；總表由使用者自己選
//   (允許多個總表，例如備份副本)。
// ====================================================================
const BAF_REQUIRED_LABELS = ['UID', '圖紙名稱', '圖紙號碼'];

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    // 分表↔總表的「值編輯」已由 onEdit 自動雙向同步，不再需要手動同步選單。
    // 新增/刪除圖紙等結構變更請在 Revit 端做，匯出時會自動重建總表與所有分表。
    ui.createMenu('BaF')
      .addSubMenu(ui.createMenu('設定目前工作表角色')
        .addItem('設為總表(main)', 'bafMarkMain')
        .addItem('設為分表(sub)', 'bafMarkSub')
        .addItem('清除角色', 'bafClearRole'))
      .addToUi();
  } catch (e) {}
}

// ---- 角色標記：把「目前正在看的工作表」設成總表/分表/清除 ----
// 用途：複製總表當存檔備份後，複製出的副本可能沒帶到角色標記，
//       用這個把副本也設成 main，它就會出現在「總表」下拉(允許多個總表)。
function bafSetActiveRole_(role) {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const name = sh.getName();
  if (role === 'main' || role === 'sub') {
    if (!bafQualifies_(sh)) {
      ui.alert('「' + name + '」沒有 UID／圖紙名稱／圖紙號碼 欄位，不能設為總表/分表。\n'
        + '請先確認這是 BaF 的明細表(總表或分表)。');
      return;
    }
    setRole(sh, role);
    ui.alert('已將工作表「' + name + '」設為：' + (role === 'main' ? '總表(main)' : '分表(sub)')
      + '。\n它現在會出現在對應的清單中。');
  } else {
    try {
      const md = sh.getDeveloperMetadata();
      for (let i = 0; i < md.length; i++) if (md[i].getKey() === 'BAF_ROLE') md[i].remove();
    } catch (e) {}
    ui.alert('已清除工作表「' + name + '」的角色。');
  }
}
function bafMarkMain() { bafSetActiveRole_('main'); }
function bafMarkSub() { bafSetActiveRole_('sub'); }
function bafClearRole() { bafSetActiveRole_(''); }

function nowStamp() {
  const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  return Utilities.formatDate(new Date(), tz, 'yyyyMMdd，HH:mm');
}

// 工作表是否同時具備 UID／圖紙名稱／圖紙號碼 欄位（掃前 10 列）
function bafQualifies_(sh) {
  const scan = Math.min(10, sh.getMaxRows());
  const lastCol = Math.max(sh.getLastColumn(), 8);
  const found = {};
  for (let r = 1; r <= scan; r++) {
    const vals = sh.getRange(r, 1, 1, lastCol).getValues()[0];
    for (let c = 0; c < vals.length; c++) {
      const t = String(vals[c]).trim();
      if (BAF_REQUIRED_LABELS.indexOf(t) >= 0) found[t] = true;
    }
  }
  return BAF_REQUIRED_LABELS.every(function (n) { return found[n]; });
}

// 給下拉用：列出所有合格工作表（不靠角色判斷；總表讓使用者自己選）
function bafListSyncSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = [];
  ss.getSheets().forEach(function (sh) {
    const name = sh.getName();
    if (name.indexOf('_BaF') === 0) return;
    if (bafQualifies_(sh)) out.push({ name: name, role: getRole(sh) });
  });
  return out;
}

function bafRecMapByUid_(records) {
  const m = {};
  records.forEach(function (r) {
    const u = String(r['UID'] || '').trim();
    if (u) m[u] = r;
  });
  return m;
}

function bafEsc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// 自我檢查：在 Apps Script 編輯器選此函式按 Run，到「執行記錄」看合格工作表清單。
function bafSelfTest() {
  const list = bafListSyncSheets();
  Logger.log('合格工作表 ' + list.length + ' 張：' + JSON.stringify(list));
  return list;
}

// ---- 差異比對小工具 ----
function bafNorm_(v) {
  if (v === true) return 'TRUE';
  if (v === false) return 'FALSE';
  if (v === null || v === undefined) return '';
  return String(v).trim();
}
const BAF_CMP_FIELDS = ['是否為Revit出圖', '圖紙類別', '圖紙號碼', '圖紙名稱', '繪圖員', '修正備註'];
function bafRowDiff_(a, b) {   // 回傳有差異的欄位名
  const ch = [];
  BAF_CMP_FIELDS.forEach(function (f) { if (bafNorm_(a[f]) !== bafNorm_(b[f])) ch.push(f); });
  return ch;
}
function bafItem_(rec, type, detail) {
  const num = String(rec['圖紙號碼'] || '').trim();
  const name = String(rec['圖紙名稱'] || '').trim();
  return { type: type, label: (num || '(無圖號)') + '  ' + (name || '(無圖名)'), detail: detail || '' };
}
function bafFindDupUids_(records) {
  const c = {};
  records.forEach(function (r) { const u = String(r['UID'] || '').trim(); if (u) c[u] = (c[u] || 0) + 1; });
  return Object.keys(c).filter(function (u) { return c[u] > 1; });
}
function bafFindBadRows_(records) {  // 有 UID 但圖號圖名皆空 → 匯入 Revit 會有問題
  const bad = [];
  records.forEach(function (r) {
    const u = String(r['UID'] || '').trim();
    if (u && !String(r['圖紙號碼'] || '').trim() && !String(r['圖紙名稱'] || '').trim()) bad.push(u.slice(0, 8));
  });
  return bad;
}
// 一列的種類：delete=有ID沒資料(刪除標記)；new=沒ID有資料(新增)；
//             normal=有ID有資料；junk=沒ID沒資料(空列，忽略)
function bafRowKind_(r) {
  const u = String(r['UID'] || '').trim();
  const has = !!(String(r['圖紙號碼'] || '').trim() || String(r['圖紙名稱'] || '').trim());
  if (u && !has) return 'delete';
  if (!u && has) return 'new';
  if (u && has) return 'normal';
  return 'junk';
}
function bafDiffToJson_(b) {
  let items = [].concat(b.added, b.modified, b.deleted);
  let truncated = false;
  if (items.length > 300) { items = items.slice(0, 300); truncated = true; }
  const blockers = [];
  if (b.dup && b.dup.length) {
    blockers.push('結果會有重複 UID（' + b.dup.slice(0, 8).join('、') + (b.dup.length > 8 ? '…' : '')
      + '）→ 已擋下，避免匯入 Revit 出錯。');
  }
  const warns = [];
  if (b.warn && b.warn.length) {
    warns.push('有 ' + b.warn.length + ' 列「有 UID 但無圖號圖名」，匯入 Revit 可能異常（建議先補圖號/圖名）。');
  }
  return {
    summary: { added: b.added.length, modified: b.modified.length, deleted: b.deleted.length },
    items: items, truncated: truncated, blockers: blockers, warns: warns
  };
}

// ---- 對話框（伺服器端先把清單塞進 HTML；兩段式：預覽差異 → 確認執行）----
function bafBuildDialogHtml_(kind) {
  const list = bafListSyncSheets();
  const subs = list.filter(function (s) { return s.role === 'sub'; });
  const mains = list.filter(function (s) { return s.role === 'main'; });
  let chk = '', opt = '';
  if (!subs.length) {
    chk = '<span style="color:#999">（沒有標記為分表[sub]的工作表，請先用 Revit「依類別拆分」建立）</span>';
  } else {
    subs.forEach(function (s) {
      const v = bafEsc_(s.name);
      chk += '<label><input type="checkbox" value="' + v + '"> ' + v + '</label>';
    });
  }
  if (!mains.length) {
    opt = '<option value="">（沒有標記為總表[main]的工作表，請先用 Revit 匯出建立）</option>';
  } else {
    mains.forEach(function (s) { const v = bafEsc_(s.name); opt += '<option value="' + v + '">' + v + '</option>'; });
  }
  const isS2M = (kind === 'subs2main');
  const hint = isS2M
    ? '用勾選的分表(可多選)，依 UID 同步到總表。新增/修改/刪除會先列出，確認後才執行。'
    : '用總表，依 UID 同步勾選分表(可多選)。新增/修改/刪除會先列出，確認後才執行。';
  const subsBlock = '<label>分表（可多選' + (isS2M ? '，來源' : '，目標')
    + '）</label><div id="subs" class="box">' + chk + '</div>';
  const mainBlock = '<label>總表（' + (isS2M ? '目標' : '來源')
    + '）</label><select id="main">' + opt + '</select>';
  const body = isS2M ? (subsBlock + mainBlock) : (mainBlock + subsBlock);
  return '<!DOCTYPE html><html><head><base target="_top"><style>' + BAF_DLG_CSS
    + '</style></head><body>'
    + '<div class="hint">' + hint + '</div>' + body
    + '<div><button onclick="preview()">預覽差異</button> '
    + '<button id="exec" onclick="exec()" disabled>確認執行</button></div>'
    + '<div id="msg"></div><div id="diff"></div>'
    + '<script>'
    + 'var KIND="' + kind + '";'
    + 'function getSel(){var main=document.getElementById("main").value;'
    + 'var cs=document.querySelectorAll("#subs input[type=checkbox]"),sel=[];'
    + 'for(var i=0;i<cs.length;i++)if(cs[i].checked)sel.push(cs[i].value);return {main:main,sel:sel};}'
    + 'function valid(s,m){if(!s.main){m.style.color="#c00";m.innerText="請選擇總表。";return false;}'
    + 'if(!s.sel.length){m.style.color="#c00";m.innerText="請至少勾選一張分表。";return false;}return true;}'
    + 'function showErr(e){var m=document.getElementById("msg");m.style.color="#c00";m.innerText="錯誤："+e.message;}'
    + 'function preview(){var s=getSel(),m=document.getElementById("msg");if(!valid(s,m))return;'
    + 'm.style.color="#666";m.innerText="計算差異中…";document.getElementById("exec").disabled=true;'
    + 'document.getElementById("diff").innerHTML="";'
    + 'var r=google.script.run.withSuccessHandler(showDiff).withFailureHandler(showErr);'
    + 'if(KIND==="subs2main")r.bafDiffSubsToMain(s.sel,s.main);else r.bafDiffMainToSubs(s.main,s.sel);}'
    + 'function showDiff(d){var m=document.getElementById("msg");m.style.color="#222";'
    + 'm.innerText="新增 "+d.summary.added+"　修改 "+d.summary.modified+"　刪除 "+d.summary.deleted;'
    + 'var box=document.getElementById("diff"),h="";'
    + 'd.items.forEach(function(it){var col=it.type.indexOf("刪除")>=0?"#c00":(it.type.indexOf("新增")>=0?"#080":"#06c");'
    + 'h+="<div style=\\"color:"+col+"\\">["+it.type+"] "+it.label+(it.detail?(" — "+it.detail):"")+"</div>";});'
    + 'if(d.truncated)h+="<div style=\\"color:#999\\">…只顯示前 "+d.items.length+" 筆</div>";'
    + 'if(d.warns&&d.warns.length)h+="<div style=\\"color:#b60\\">⚠ "+d.warns.join("<br>⚠ ")+"</div>";'
    + 'if(d.blockers&&d.blockers.length){h+="<div style=\\"color:#c00;font-weight:bold\\">⛔ "+d.blockers.join("<br>⛔ ")+"</div>";document.getElementById("exec").disabled=true;}'
    + 'else{document.getElementById("exec").disabled=(d.summary.added+d.summary.modified+d.summary.deleted===0);}'
    + 'box.innerHTML=h||"<div style=\\"color:#080\\">沒有差異，無需更新。</div>";}'
    + 'function exec(){var s=getSel(),m=document.getElementById("msg");if(!valid(s,m))return;'
    + 'm.style.color="#666";m.innerText="執行中…";document.getElementById("exec").disabled=true;'
    + 'var r=google.script.run.withSuccessHandler(function(x){m.style.color="#080";m.innerText=x;}).withFailureHandler(showErr);'
    + 'if(KIND==="subs2main")r.bafApplySubsToMain(s.sel,s.main);else r.bafApplyMainToSubs(s.main,s.sel);}'
    + '</scr' + 'ipt></body></html>';
}

// ====== 功能①：分表 → 總表 ======
// 範圍＝被勾選分表(類別)。範圍內：UID 在來源→修改、不在→刪除；來源多出的→新增。
// 其他類別不動。鎖 UID 比對；擋重複 UID；刪除會列出(不會誤刪)。
function bafShowSubsToMainDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(bafBuildDialogHtml_('subs2main'))
      .setWidth(500).setHeight(560),
    '① 分表 → 更新總表');
}

function bafBuildSubsToMain_(subNames, mainName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!mainName) throw new Error('請選擇總表。');
  if (!subNames || !subNames.length) throw new Error('請至少勾選一張分表。');
  if (subNames.indexOf(mainName) >= 0) throw new Error('總表不能同時是被勾選的分表。');
  const main = ss.getSheetByName(mainName);
  if (!main) throw new Error('找不到總表：' + mainName);
  const labels = BAF_ALL_LABELS;
  const md = readSheet(main, labels);
  if (!md || !md.ok) throw new Error('總表讀不到表頭：' + mainName);

  // 分類分表來源列：delete(刪除標記)/normal(更新)/new(新增)；scopeCats=分表涵蓋的類別
  const delSet = {}, srcByUid = {}, newRows = [], scopeCats = {};
  subNames.forEach(function (nm) {
    const sub = ss.getSheetByName(nm);
    if (!sub) return;
    const d = readSheet(sub, labels);
    if (!d || !d.ok) return;
    // 分表頁籤名即其圖紙類別（拆分時以類別命名）。把頁籤名納入 scopeCats，
    // 這樣即使分表的某類別被刪到一列不剩(分表變空)，仍能連動刪掉總表中該類別的列。
    if (String(nm || '').trim()) scopeCats[String(nm).trim()] = true;
    d.records.forEach(function (r) {
      const cat = String(r['圖紙類別'] || '').trim();
      if (cat) scopeCats[cat] = true;
      const k = bafRowKind_(r);
      const u = String(r['UID'] || '').trim();
      if (k === 'delete') delSet[u] = true;
      else if (k === 'normal') srcByUid[u] = r;
      else if (k === 'new') newRows.push(r);
    });
  });

  const added = [], modified = [], deleted = [], result = [], seen = {};
  md.records.forEach(function (r) {
    const u = String(r['UID'] || '').trim();
    const cat = String(r['圖紙類別'] || '').trim();
    const k = bafRowKind_(r);
    if (u && delSet[u]) { deleted.push(bafItem_(r, '刪除', '標記刪除(有ID無資料)')); return; }
    if (k === 'delete') { deleted.push(bafItem_(r, '刪除', '標記刪除(有ID無資料)')); if (u) delSet[u] = true; return; }
    if (u && srcByUid[u]) {
      const ch = bafRowDiff_(r, srcByUid[u]);
      if (ch.length) modified.push(bafItem_(srcByUid[u], '修改', ch.join('、')));
      result.push(srcByUid[u]); seen[u] = true; return;
    }
    if (u && scopeCats[cat]) { deleted.push(bafItem_(r, '刪除', '分表此類別已無此列')); return; }
    result.push(r); if (u) seen[u] = true;
  });
  Object.keys(srcByUid).forEach(function (u) {
    if (!seen[u]) { added.push(bafItem_(srcByUid[u], '新增', '')); result.push(srcByUid[u]); seen[u] = true; }
  });
  newRows.forEach(function (r) { added.push(bafItem_(r, '新增', '無ID，匯入Revit後配發')); result.push(r); });

  return {
    sheet: main, labels: labels, result: result, delSet: delSet, subNames: subNames,
    added: added, modified: modified, deleted: deleted,
    dup: bafFindDupUids_(result), warn: bafFindBadRows_(result)
  };
}

function bafDiffSubsToMain(subNames, mainName) { return bafDiffToJson_(bafBuildSubsToMain_(subNames, mainName)); }

function bafApplySubsToMain(subNames, mainName) {
  const b = bafBuildSubsToMain_(subNames, mainName);
  if (b.dup && b.dup.length) throw new Error('結果會有重複 UID，已中止：' + b.dup.join('、'));
  const stamp = nowStamp();
  bafFastWrite_(b.sheet, b.result, b.labels, '是否為Revit出圖', ['繪圖員'], stamp, false, null);
  setRole(b.sheet, 'main');
  // 刪除：把被刪 UID 與刪除標記列，從各分表也移除（同時刪總表與分表）
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  subNames.forEach(function (nm) {
    const sub = ss.getSheetByName(nm);
    if (!sub) return;
    const d = readSheet(sub, b.labels);
    if (!d || !d.ok) return;
    const kept = d.records.filter(function (r) {
      const u = String(r['UID'] || '').trim();
      if (u && b.delSet[u]) return false;
      if (bafRowKind_(r) === 'delete') return false;
      return true;
    });
    if (kept.length !== d.records.length) {
      bafFastWrite_(sub, kept, b.labels, '是否為Revit出圖', ['繪圖員'], stamp, true, d.batchParams || {});
    }
  });
  return '✅ 完成（總表「' + mainName + '」）：新增 ' + b.added.length + '、修改 '
       + b.modified.length + '、刪除 ' + b.deleted.length + '。';
}

// ====== 功能②：總表 → 分表 ======
// 每張分表(類別)＝總表中該類別的列。對該分表：UID 在總表→修改、總表多出→新增、
// 分表多出(總表沒有)→刪除。鎖 UID；擋重複 UID；刪除會列出。
function bafShowMainToSubsDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(bafBuildDialogHtml_('main2subs'))
      .setWidth(500).setHeight(560),
    '② 總表 → 更新分表');
}

function bafBuildMainToSubsPlan_(mainName, subNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!mainName) throw new Error('請選擇總表。');
  if (!subNames || !subNames.length) throw new Error('請至少勾選一張分表。');
  const main = ss.getSheetByName(mainName);
  if (!main) throw new Error('找不到總表：' + mainName);
  const labels = BAF_ALL_LABELS;
  const md = readSheet(main, labels);
  if (!md || !md.ok) throw new Error('總表讀不到表頭：' + mainName);

  // 總表中的刪除標記(有ID無資料) → 全域刪除(總表與分表都刪)
  const delSet = {};
  md.records.forEach(function (r) {
    if (bafRowKind_(r) === 'delete') { const u = String(r['UID'] || '').trim(); if (u) delSet[u] = true; }
  });
  const mainClean = md.records.filter(function (r) { return bafRowKind_(r) !== 'delete'; });
  const mainChangedByDel = (mainClean.length !== md.records.length);

  const plans = [], added = [], modified = [], deleted = [];
  let dup = [], warn = [];
  subNames.forEach(function (nm) {
    if (nm === mainName) return;
    const sub = ss.getSheetByName(nm);
    if (!sub) return;
    const d = readSheet(sub, labels);
    if (!d || !d.ok) return;
    const cat = nm;
    // 分表 ＝ 總表中此類別、非刪除標記的列
    const mainCat = mainClean.filter(function (r) { return String(r['圖紙類別'] || '').trim() === cat; });
    const mainCatByUid = bafRecMapByUid_(mainCat);
    const subByUid = bafRecMapByUid_(d.records);
    mainCat.forEach(function (r) {
      const u = String(r['UID'] || '').trim();
      if (u && subByUid[u]) {
        const ch = bafRowDiff_(subByUid[u], r);
        if (ch.length) modified.push(bafItem_(r, '修改(' + nm + ')', ch.join('、')));
      } else { added.push(bafItem_(r, '新增(' + nm + ')', u ? '' : '無ID')); }
    });
    d.records.forEach(function (r) {
      const u = String(r['UID'] || '').trim();
      if (u && delSet[u]) deleted.push(bafItem_(r, '刪除(' + nm + ')', '標記刪除'));
      else if (u && !mainCatByUid[u]) deleted.push(bafItem_(r, '刪除(' + nm + ')', '總表此類別已無此列'));
    });
    const result = mainCat.slice();
    dup = dup.concat(bafFindDupUids_(result));
    warn = warn.concat(bafFindBadRows_(result));
    plans.push({ sheet: sub, result: result, batchParams: (d.batchParams || {}) });
  });
  return {
    plans: plans, labels: labels, main: main, mainClean: mainClean,
    mainChangedByDel: mainChangedByDel,
    added: added, modified: modified, deleted: deleted, dup: dup, warn: warn
  };
}

function bafDiffMainToSubs(mainName, subNames) { return bafDiffToJson_(bafBuildMainToSubsPlan_(mainName, subNames)); }

function bafApplyMainToSubs(mainName, subNames) {
  const p = bafBuildMainToSubsPlan_(mainName, subNames);
  if (p.dup && p.dup.length) throw new Error('結果會有重複 UID，已中止：' + p.dup.join('、'));
  const stamp = nowStamp();
  // 先移除總表的刪除標記列(同時刪總表與分表)
  if (p.mainChangedByDel) {
    bafFastWrite_(p.main, p.mainClean, p.labels, '是否為Revit出圖', ['繪圖員'], stamp, false, null);
    setRole(p.main, 'main');
  }
  p.plans.forEach(function (pl) {
    bafFastWrite_(pl.sheet, pl.result, p.labels, '是否為Revit出圖', ['繪圖員'], stamp, true, pl.batchParams);
  });
  return '✅ 完成（' + p.plans.length + ' 張分表）：新增 ' + p.added.length + '、修改 '
       + p.modified.length + '、刪除 ' + p.deleted.length + '。';
}

// ---- 批次寫入（加速）：總表批次一次寫；分表沿用 writeSubSheet ----
function bafFastWrite_(sh, records, labels, checkboxLabel, dropdownLabels, updateDate, isSub, batchParams) {
  if (isSub) {
    writeSubSheet(sh, records, labels, checkboxLabel, dropdownLabels, updateDate, true, batchParams || {});
    return;
  }
  try { if (sh.getFilter()) sh.getFilter().remove(); } catch (e0) {}
  try { trimLeadingBlankCols(sh, labels); } catch (eT) {}
  const scan = Math.min(10, sh.getMaxRows());
  const lastCol0 = Math.max(sh.getLastColumn(), 8);
  let headerRow = 0, headerVals = null, best = 0;
  for (let r = 1; r <= scan; r++) {
    const vals = sh.getRange(r, 1, 1, lastCol0).getValues()[0];
    let hit = 0;
    for (let c = 0; c < vals.length; c++) if (labels.indexOf(String(vals[c]).trim()) >= 0) hit++;
    if (hit > best) { best = hit; headerRow = r; headerVals = vals; }
  }
  if (!headerRow) {
    headerRow = 2;
    sh.getRange(headerRow, 1, 1, labels.length).setValues([labels]);
    sh.getRange(1, 1).setValue('更新時間：');
    headerVals = sh.getRange(headerRow, 1, 1, Math.max(lastCol0, labels.length)).getValues()[0];
  }
  const colOf = {};
  for (let c = 0; c < headerVals.length; c++) {
    const t = String(headerVals[c]).trim();
    if (t && colOf[t] === undefined && labels.indexOf(t) >= 0) colOf[t] = c + 1;
  }
  const cols = labels.map(function (l) { return colOf[l]; }).filter(function (c) { return c; });
  if (!cols.length || (Math.max.apply(null, cols) - Math.min.apply(null, cols) + 1 !== cols.length)) {
    // 欄位不連續(被調換) → 用穩健的逐欄版
    writeIndexToSheet(sh, records, labels, checkboxLabel, dropdownLabels, updateDate, true);
    return;
  }
  const minC = Math.min.apply(null, cols), maxC = Math.max.apply(null, cols);
  const labelByCol = {};
  Object.keys(colOf).forEach(function (l) { labelByCol[colOf[l]] = l; });

  const dataStart = headerRow + 1, nData = records.length;
  const needRows = dataStart - 1 + nData;
  if (sh.getMaxRows() < needRows) sh.insertRowsAfter(sh.getMaxRows(), needRows - sh.getMaxRows());
  const usedRows = sh.getLastRow() - dataStart + 1;
  if (usedRows > 0) {
    const rng = sh.getRange(dataStart, minC, usedRows, maxC - minC + 1);
    try { rng.breakApart(); } catch (e1) {}
    rng.clearContent(); rng.clearDataValidations();
  }
  if (nData > 0) {
    const block = records.map(function (rec) {
      const row = [];
      for (let c = minC; c <= maxC; c++) {
        const l = labelByCol[c];
        let v = l ? rec[l] : '';
        if (v === undefined || v === null) v = '';
        row.push(v);
      }
      return row;
    });
    sh.getRange(dataStart, minC, nData, maxC - minC + 1).setValues(block);

    const cStatus = colOf[checkboxLabel];
    if (cStatus) {
      try {
        const cr = sh.getRange(dataStart, cStatus, nData, 1);
        cr.insertCheckboxes();
        cr.setValues(records.map(function (rec) {
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
          const rule = SpreadsheetApp.newDataValidation().requireValueInList(opts, true).setAllowInvalid(true).build();
          sh.getRange(dataStart, c, nData, 1).setDataValidation(rule);
        }
      } catch (e3) {}
    });
  }
  if (updateDate) {
    try {
      const lc = Math.max(sh.getLastColumn(), 8);
      const loc = findCell(sh, '更新時間', scan, lc);
      if (loc) sh.getRange(loc.row, Math.min(loc.col + 3, lc)).setValue(updateDate);
    } catch (e4) {}
  }
  try { sh.setFrozenRows(headerRow); } catch (eFz) {}
  try { lockHeaderRows(sh, headerRow, Math.max(sh.getLastColumn(), 8)); } catch (e5) {}
}

// ---- 對話框 HTML ----
const BAF_DLG_CSS =
'body{font-family:Arial,"Microsoft JhengHei";font-size:13px;padding:10px;color:#222}' +
'label{display:block;margin:10px 0 2px;font-weight:bold}' +
'select{width:100%;padding:5px;box-sizing:border-box}' +
'.hint{color:#666;font-size:12px;margin-bottom:6px}' +
'.box{border:1px solid #ccc;max-height:150px;overflow:auto;padding:6px}' +
'.box label{font-weight:normal;display:block;margin:2px 0}' +
'button{margin-top:14px;padding:7px 16px;font-weight:bold}' +
'button:disabled{opacity:0.5}' +
'#msg{margin-top:10px;white-space:pre-wrap;font-weight:bold}' +
'#diff{margin-top:8px;max-height:190px;overflow:auto;border:1px solid #eee;' +
'padding:6px;font-size:12px;line-height:1.5}';
