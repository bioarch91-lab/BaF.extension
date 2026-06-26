# BaF.extension

給 Revit 使用的 pyRevit 工具套件，作者：BaF（九典聯合 / BIM 工具）。
安裝後會在 Revit 上方出現「**BaF**」頁籤，內含多個常用的批次與輔助工具。

> 執行環境：pyRevit + Revit 內的 IronPython 2.7。

---

## 功能總覽

### 視圖工具（Views.panel）

| 工具 | 說明 |
|------|------|
| **視圖對齊** | 以樣板圖紙為基準，對齊其他圖紙上的同類型 Viewport 位置 |
| **批次套用視圖屬性** | 一次把多個視圖套用 Scope Box、視圖樣板等屬性，每個屬性可獨立勾選是否套用 |
| **批次更改圖框** | 一次更換多張圖紙的圖框（Title Block），可切換類型、替換 Family、為空白圖紙加圖框 |
| **批次新增/編輯圖紙** | 一次建立多張圖紙，或批次重新編號/改名既有圖紙；支援逐筆輸入、Excel 貼上、規則自動生成 |
| **視圖配對上圖** | 用連連看的方式，一次把多個視圖配對放置到圖紙上 |
| **取消元素隱藏** | 一鍵取消選定視圖中「被直接隱藏的元素」(Hide Elements)，不影響視圖樣板與類別可見性 |

### 需求書工具（Specs.panel）

| 工具 | 說明 |
|------|------|
| **需求書查閱** | 在 Revit 內開啟側邊視窗查閱統包需求書 PDF/Word/Excel（使用 WebView2） |

### 元件庫（Families.panel）

| 工具 | 說明 |
|------|------|
| **族群瀏覽器** | 連到本案的族群／圖框資料夾，直接在 Revit 內瀏覽、載入並放置族群（.rfa） |

---

## 安裝

需先安裝 [pyRevit](https://github.com/eirannejad/pyRevit/releases)。

### 方式 A：用 pyRevit 從 GitHub 安裝（建議，可自動更新）

開啟「命令提示字元」執行（每台電腦做一次）：

```bash
pyrevit extend ui BaF https://github.com/b00204002-spec/BaF.extension.git --branch=main
```

安裝後開啟 Revit，切到 **pyRevit** 頁籤點 **Reload**，上方就會出現「**BaF**」頁籤。

也可改用 GUI：Revit → **pyRevit → Extensions**，貼上 git 網址安裝。

### 方式 B：手動複製資料夾

1. 下載 / 解壓縮整個 `BaF.extension` 資料夾，放到：
   ```
   %APPDATA%\pyRevit\Extensions\
   ```
   （把這串貼到檔案總管網址列即可開啟）
2. 開 Revit → **pyRevit** 頁籤 → **Reload**。

> 注意：方式 A 與方式 B 不要並存，否則會出現兩個重複的 BaF 頁籤。

---

## 更新

- **方式 A 安裝者：** Revit → **pyRevit → Extensions** 視窗按 **Update**；
  或在「命令提示字元」執行 `pyrevit extensions update --all`；
  也可在 pyRevit Settings 勾選「啟動時檢查更新」，開 Revit 時自動拉取。
- **方式 B 安裝者：** 重新下載覆蓋資料夾後，回 Revit 按 **Reload**。

---

## 開發者：如何發佈更新

1. 修改工具程式碼。
2. commit 並 push 到 GitHub：
   ```bash
   git add .
   git commit -m "說明這次的修改"
   git push
   ```
3. 同事端依「更新」章節拉取即可。

---

## 專案結構與慣例

```
BaF.extension/
├── hooks/                                 ← pyRevit 事件掛勾
│   └── view-activated.py                  ← 切換到圖紙時，有「修正備註」就跳視窗列出
└── Tools.tab/                             ← Revit 頁籤 (title: BaF)
    ├── Views.panel/                        ← 視圖工具
    │   ├── AlignViewports.pushbutton/
    │   ├── BatchApplyViewProps.pushbutton/
    │   ├── BatchChangeTitleBlock.pushbutton/
    │   ├── BatchManageSheets.pushbutton/
    │   ├── PlaceViewsToSheets.pushbutton/
    │   └── UnhideElements.pushbutton/
    ├── Specs.panel/                        ← 需求書工具
    │   └── SpecReader.pushbutton/
    └── Families.panel/                     ← 元件庫
        └── FamilyBrowser.pushbutton/
```

慣例：
- **資料夾名稱用英文**（避免 IronPython 編碼問題），**UI 顯示名稱用中文**，透過各層 `bundle.yaml` 的 `title:` 設定。
- 每個 pushbutton 含 `bundle.yaml` + `script.py` + `icon.png`；共用程式放在工具的 `lib/` 下。
- 程式以 IronPython 2.7 撰寫，使用 Py2 相容語法，檔案開頭加 `# -*- coding: utf-8 -*-`。
- 改動後無法在一般終端機測試，需在 Revit 內 **Reload** 驗證。
