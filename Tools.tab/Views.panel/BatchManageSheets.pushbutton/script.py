# -*- coding: utf-8 -*-
"""
圖紙編輯與管理 (Batch Manage Sheets)
===========================================================
五種操作模式（Tab）：
  1. 編輯既有：勾選想處理的既有圖紙，輸入新的圖號/圖名（空白 = 不變更）
  2. 逐筆輸入：UI 上一行一行填寫圖號/圖名 → 新增圖紙
  3. 貼上文字：從 Excel/記事本複製貼上 → 新增圖紙
  4. 規則生成：例如 A2-{n:02d} 1~20 → 新增 20 張圖紙
  5. Google Sheet 同步：圖紙索引與 Google Sheet 雙向同步（開發中，目前為唯讀讀取測試）

「編輯既有」也支援規則式批次重新編號。
作者: BaF / BIM 工具
"""

import re
import clr
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")
clr.AddReference("System")

from System import DateTime

from System.Windows import (
    Window, Thickness, HorizontalAlignment, VerticalAlignment,
    WindowStartupLocation, Visibility, TextTrimming, FontWeights,
    GridLength, GridUnitType, CornerRadius, TextWrapping, Clipboard
)
from System.Windows.Controls import (
    StackPanel, Button, ScrollViewer, Grid, RowDefinition, ColumnDefinition,
    TextBlock, Border, Orientation, ComboBox, TextBox, TabControl, TabItem,
    CheckBox, ScrollBarVisibility, PasswordBox
)
from System.Windows.Media import SolidColorBrush, Color, FontFamily

from pyrevit import revit, DB, forms, script

import os
import json

doc = revit.doc
output = script.get_output()


# 部分同步參數其實是 Revit「內建」圖紙參數（名稱依語言/範本用字可能略有差異，
# 例如「圖紙發佈日期(佈)」vs 我們用的「圖紙發布日期(布)」）。
# 用 BuiltInParameter 對應，偵測/讀/寫/明細表都優先用內建，避免比錯字、建出重複。
def _resolve_bip(bipname):
    try:
        return getattr(DB.BuiltInParameter, bipname)
    except Exception:
        return None


BUILTIN_PARAM_MAP = {}
for _logical, _bipname in (
        (u"繪圖員", "SHEET_DRAWN_BY"),
        (u"審圖員", "SHEET_CHECKED_BY"),
        (u"設計者", "SHEET_DESIGNED_BY"),
        (u"批准者", "SHEET_APPROVED_BY"),
        (u"圖紙發布日期", "SHEET_ISSUE_DATE")):
    _b = _resolve_bip(_bipname)
    if _b is not None:
        BUILTIN_PARAM_MAP[_logical] = _b


def _lookup_sheet_param(sheet, pname):
    """取得圖紙參數。

    對照表內的名稱(內建圖紙參數)一律優先用 BuiltInParameter，
    避免被同名的自訂/誤建參數攔截(例如先前誤建的空白「圖紙發布日期」)。
    其餘名稱用一般 LookupParameter。
    """
    bip = BUILTIN_PARAM_MAP.get(pname)
    if bip is not None:
        try:
            p = sheet.get_Parameter(bip)
            if p is not None:
                return p
        except Exception:
            pass
    try:
        return sheet.LookupParameter(pname)
    except Exception:
        return None


# Google Sheet 寫入設定（存在本機，不進 git，避免 URL 外流）
GSHEET_CFG = os.path.join(os.getenv("APPDATA") or u"", "pyRevit", "baf_gsheet_config.json")


def load_gsheet_cfg():
    try:
        with open(GSHEET_CFG, "r") as f:
            return json.load(f)
    except Exception:
        return {}


def save_gsheet_cfg(cfg):
    """把 cfg 合併進現有設定（不覆蓋掉其他鍵，例如 script_master）。"""
    try:
        merged = load_gsheet_cfg()
        merged.update(cfg)
        d = os.path.dirname(GSHEET_CFG)
        if d and not os.path.exists(d):
            os.makedirs(d)
        with open(GSHEET_CFG, "w") as f:
            json.dump(merged, f)
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# AI 視覺判讀（「依紅筆圖註解」用）：Claude / Gemini / ChatGPT 三家可選
#   金鑰與模型存在同一份 baf_gsheet_config.json（不進 git）的 ai_keys / ai_models。
#   作法：把 PDF（base64）直接送給有視覺能力的模型，回傳合併好的修改描述 JSON。
# ---------------------------------------------------------------------------

# (內部代號, UI 顯示名稱)；順序＝下拉選單順序
AI_PROVIDER_LABELS = [
    (u"claude", u"Claude (Anthropic)"),
    (u"gemini", u"Gemini (Google)"),
    (u"openai", u"ChatGPT (OpenAI)"),
]
AI_DEFAULT_MODELS = {
    u"claude": u"claude-opus-4-8",
    u"gemini": u"gemini-2.5-pro",
    u"openai": u"gpt-4o",
}


def _ai_provider_label(provider):
    for pk, pl in AI_PROVIDER_LABELS:
        if pk == provider:
            return pl
    return provider


def _ai_model_for(provider, models):
    return (models.get(provider) or AI_DEFAULT_MODELS.get(provider) or u"").strip()


def _file_to_base64(path):
    """讀檔轉 base64，回傳 (base64 字串, 位元組數)。用 .NET 讀避免二進位編碼問題。"""
    from System import Convert
    from System.IO import File
    data = File.ReadAllBytes(path)
    return Convert.ToBase64String(data), data.Length


def _describe_web_exception(ex):
    """把 API 失敗原因組成可讀字串：HTTP 狀態碼 + 回應 body（API 錯誤訊息）；
    連線層級錯誤（逾時／DNS／TLS）則回連線狀態 + 例外訊息。"""
    try:
        resp = getattr(ex, "Response", None)
        if resp is None:
            inner = getattr(ex, "clsException", None)
            resp = getattr(inner, "Response", None) if inner is not None else None
        code_txt, body = u"", u""
        if resp is not None:
            try:
                sc = getattr(resp, "StatusCode", None)
                if sc is not None:
                    code_txt = u"HTTP {} {}".format(
                        int(sc), getattr(resp, "StatusDescription", u"") or u"").strip()
            except Exception:
                pass
            try:
                from System.IO import StreamReader
                from System.Text import Encoding
                reader = StreamReader(resp.GetResponseStream(), Encoding.UTF8)
                body = (reader.ReadToEnd() or u"").strip()
                reader.Close()
            except Exception:
                pass
        parts = []
        if code_txt:
            parts.append(code_txt)
        if body:
            parts.append(body)
        if not parts:
            # 沒有 HTTP 回應 → 連線層級錯誤
            status = getattr(ex, "Status", None)
            if status is not None:
                parts.append(u"連線狀態：{}".format(status))
            parts.append(unicode(getattr(ex, "Message", None) or ex))
        return u"　".join(p for p in parts if p)
    except Exception:
        return unicode(getattr(ex, "Message", None) or ex)


def _http_post_raw(url, body_text, headers, timeout_ms=300000):
    """POST 文字 body（application/json），回傳 (回應字串, 錯誤字串)。
    用 HttpWebRequest 以便設定較長的逾時（AI 判讀大張 PDF 可能要數十秒）。"""
    try:
        from System.Net import (WebRequest, ServicePointManager,
                                SecurityProtocolType)
        from System.Text import Encoding
        from System.IO import StreamReader
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        req = WebRequest.Create(url)
        req.Method = "POST"
        req.ContentType = "application/json"
        req.Timeout = timeout_ms
        req.ReadWriteTimeout = timeout_ms
        for k, v in headers.items():
            req.Headers.Add(k, v)
        data = Encoding.UTF8.GetBytes(body_text)
        req.ContentLength = data.Length
        stream = req.GetRequestStream()
        stream.Write(data, 0, data.Length)
        stream.Close()
        resp = req.GetResponse()
        reader = StreamReader(resp.GetResponseStream(), Encoding.UTF8)
        text = reader.ReadToEnd()
        reader.Close()
        return text, None
    except Exception as ex:
        return None, _describe_web_exception(ex)


def _redpen_build_prompt():
    """送給 AI 的指令：逐頁讀圖框、把紅筆標示轉成具體中文描述、同張圖合併、輸出 JSON。"""
    return (
        u"你是建築圖協作助理。這個 PDF 是建築工程圖（A1 或 A3），每一頁是一張圖，"
        u"頁面上有圖框（標題欄），圖框內含該圖的圖號、階段、日期、繪圖員等資訊；"
        u"圖面上可能有以紅筆（或其他顏色）標註的修改意見，常見形式為箭頭、方框、"
        u"圈選、底線、雲線（revision cloud）並搭配手寫或文字說明。\n\n"
        u"請逐頁處理，完成以下工作：\n"
        u"1. 從圖框讀出該頁的『圖號』(sheet_number)，盡量也讀出階段(stage)、日期(date)、"
        u"繪圖員(drawer)。圖號要與圖框上印的完全一致（含英數字與符號，不要自行加減空白）。\n"
        u"2. 找出該頁所有紅筆／手寫的修改標示，把每一個標示轉成一句**具體、可執行的中文描述**"
        u"（說明位置與要改什麼，例如：『將北側外牆窗戶尺寸由 W1200 改為 W1500』、"
        u"『樓梯間欄杆需補上高度標註』）。不要只寫『有紅筆』這類空泛內容。\n"
        u"3. 同一頁（同一張圖）若有多個修改標示，請合併成一段，用全形分號『；』分隔。\n"
        u"4. 沒有任何修改標示的頁面，請略過、不要列入結果。\n\n"
        u"只輸出 JSON，不要任何其他文字或 markdown 圍欄，格式如下：\n"
        u'{"sheets":[{"sheet_number":"圖號","stage":"階段","date":"日期",'
        u'"drawer":"繪圖員","notes":"合併後的修改描述"}]}\n')


def _redpen_extract_json(text):
    """從模型回覆中抽出最外層 JSON 物件（容忍 ```json 圍欄與前後雜訊）。"""
    if not text:
        return None
    t = text.strip()
    if t.startswith(u"```"):
        nl = t.find(u"\n")
        if nl != -1:
            t = t[nl + 1:]
        t = t.rstrip()
        if t.endswith(u"```"):
            t = t[:-3]
    t = t.strip()
    i = t.find(u"{")
    j = t.rfind(u"}")
    if i == -1 or j == -1 or j < i:
        return None
    try:
        return json.loads(t[i:j + 1])
    except Exception:
        return None


def _redpen_call_claude(model, key, b64, prompt):
    body = json.dumps({
        "model": model, "max_tokens": 8000,
        "messages": [{"role": "user", "content": [
            {"type": "document",
             "source": {"type": "base64",
                        "media_type": "application/pdf", "data": b64}},
            {"type": "text", "text": prompt}]}]})
    headers = {"x-api-key": key, "anthropic-version": "2023-06-01"}
    resp, err = _http_post_raw(
        "https://api.anthropic.com/v1/messages", body, headers)
    if err:
        return None, err
    try:
        d = json.loads(resp)
    except Exception:
        return None, u"回應非 JSON：{}".format(resp[:300])
    if d.get("type") == "error":
        e = d.get("error") or {}
        return None, u"{}: {}".format(
            e.get("type") or u"error", e.get("message") or unicode(d.get("error")))
    parts = d.get("content") or []
    txt = u"".join(p.get("text") or u""
                   for p in parts if p.get("type") == "text")
    if not txt:
        return None, u"Anthropic 回應沒有文字內容（stop_reason={}）：{}".format(
            d.get("stop_reason"), resp[:300])
    return txt, None


def _redpen_call_gemini(model, key, b64, prompt):
    url = ("https://generativelanguage.googleapis.com/v1beta/models/"
           + model + ":generateContent?key=" + key)
    body = json.dumps({
        "contents": [{"parts": [
            {"inline_data": {"mime_type": "application/pdf", "data": b64}},
            {"text": prompt}]}]})
    resp, err = _http_post_raw(url, body, {})
    if err:
        return None, err
    try:
        d = json.loads(resp)
    except Exception:
        return None, u"回應非 JSON：{}".format(resp[:300])
    if d.get("error"):
        return None, unicode(d.get("error"))
    cands = d.get("candidates") or []
    if not cands:
        return None, u"無 candidates：{}".format(resp[:300])
    parts = (cands[0].get("content") or {}).get("parts") or []
    txt = u"".join(p.get("text") or u"" for p in parts if p.get("text"))
    return txt, None


def _redpen_call_openai(model, key, b64, prompt):
    body = json.dumps({
        "model": model,
        "input": [{"role": "user", "content": [
            {"type": "input_file", "filename": "redpen.pdf",
             "file_data": "data:application/pdf;base64," + b64},
            {"type": "input_text", "text": prompt}]}]})
    headers = {"Authorization": "Bearer " + key}
    resp, err = _http_post_raw(
        "https://api.openai.com/v1/responses", body, headers)
    if err:
        return None, err
    try:
        d = json.loads(resp)
    except Exception:
        return None, u"回應非 JSON：{}".format(resp[:300])
    if d.get("error"):
        return None, unicode(d.get("error"))
    chunks = []
    for item in (d.get("output") or []):
        for c in (item.get("content") or []):
            if c.get("type") == "output_text" and c.get("text"):
                chunks.append(c.get("text"))
    txt = u"".join(chunks)
    if not txt and d.get("output_text"):
        txt = d.get("output_text")
    return txt, None


def _redpen_call_ai(provider, model, key, b64, prompt):
    """依後端代號分派。回傳 (模型文字回覆, 錯誤字串)。"""
    if provider == u"claude":
        return _redpen_call_claude(model, key, b64, prompt)
    if provider == u"gemini":
        return _redpen_call_gemini(model, key, b64, prompt)
    if provider == u"openai":
        return _redpen_call_openai(model, key, b64, prompt)
    return None, u"未知的 AI 後端：{}".format(provider)


def _doevents():
    """讓 WPF 把待處理的版面/繪製事件跑一輪（顯示『處理中』訊息用）。"""
    try:
        from System.Windows.Threading import (DispatcherFrame, Dispatcher,
                                              DispatcherPriority)
        from System import Action
        frame = DispatcherFrame()
        Dispatcher.CurrentDispatcher.BeginInvoke(
            DispatcherPriority.Background, Action(lambda: setattr(frame, "Continue", False)))
        Dispatcher.PushFrame(frame)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Revit 資料準備
# ---------------------------------------------------------------------------

def get_title_block_types(document):
    types = DB.FilteredElementCollector(document) \
        .OfCategory(DB.BuiltInCategory.OST_TitleBlocks) \
        .WhereElementIsElementType() \
        .ToElements()
    
    result = []
    for t in types:
        family_name = ""
        type_name = ""
        try:
            family_name = t.Family.Name
        except:
            pass
        try:
            tn_param = t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if tn_param:
                type_name = tn_param.AsString() or ""
        except:
            pass
        if not type_name:
            try:
                type_name = t.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            except:
                pass
        
        label = "{} - {}".format(family_name, type_name) if family_name else type_name
        result.append((label, t.Id))
    
    return sorted(result, key=lambda x: x[0])


def get_existing_sheets(document):
    """取得所有非 placeholder 的既有圖紙。"""
    sheets = DB.FilteredElementCollector(document).OfClass(DB.ViewSheet).ToElements()
    sheets = [s for s in sheets if not s.IsPlaceholder]
    return sorted(sheets, key=lambda s: s.SheetNumber)


def get_existing_sheet_numbers(document):
    sheets = DB.FilteredElementCollector(document).OfClass(DB.ViewSheet).ToElements()
    return set(s.SheetNumber for s in sheets)


# ---------------------------------------------------------------------------
# 規則生成
# ---------------------------------------------------------------------------

def expand_pattern(pattern_text, start, end):
    result = []
    for i in range(start, end + 1):
        try:
            result.append(pattern_text.format(n=i))
        except:
            result.append(pattern_text.replace("{n}", str(i)))
    return result


def expand_pattern_for_count(pattern_text, count, start=1):
    return expand_pattern(pattern_text, start, start + count - 1)


# ---------------------------------------------------------------------------
# WPF 視窗
# ---------------------------------------------------------------------------

class BatchManageSheetsWindow(Window):
    
    COLOR_BG = (245, 245, 250)
    COLOR_BTN = (255, 255, 255)
    COLOR_PRIMARY = (99, 102, 241)
    COLOR_TEXT = (30, 30, 40)
    COLOR_TEXT_LIGHT = (255, 255, 255)
    COLOR_GROUP_HEADER = (235, 238, 245)
    COLOR_ERROR = (220, 53, 69)
    COLOR_SUCCESS = (34, 139, 34)
    COLOR_HIGHLIGHT = (252, 248, 220)  # 淡黃，標記「有變更」
    
    MODE_EDIT_EXISTING = 0
    MODE_TABLE = 1
    MODE_PASTE = 2
    MODE_PATTERN = 3
    MODE_SYNC = 4

    # Google Sheet 同步要對照的『自訂文字參數』欄位（圖號/圖名為內建，不放這裡）
    SYNC_TEXT_PARAMS = [u"圖紙類別", u"繪圖員", u"修正備註"]

    # 寫入 Google Sheet 用的表頭名稱（程式會去試算表找同名的欄，可自由調換欄位順序）
    HEADER_LABELS = [u"UID", u"是否為Revit出圖", u"圖紙類別", u"圖紙號碼",
                     u"圖紙名稱", u"繪圖員", u"修正備註"]
    CHECKBOX_LABEL = u"是否為Revit出圖"   # 這欄在 Google Sheet 套核取方塊
    DROPDOWN_LABELS = [u"繪圖員"]          # 這些欄套下拉選單（圖紙類別改為一般文字）
    # 分表標頭的「整批共用」參數：同一圖紙類別批次出圖時一致的資訊
    BATCH_PARAMS = [u"審圖員", u"設計者", u"批准者", u"圖紙發布日期", u"校核2"]
    
    def __init__(self, title_blocks, existing_sheets):
        self.title_blocks = title_blocks
        self.existing_sheets = existing_sheets
        self.existing_numbers = set(s.SheetNumber for s in existing_sheets)
        
        self.result = None
        self.confirmed = False
        self._pending_plan = None  # 匯入預覽後暫存的套用計畫
        self._pending_export = None  # 匯出預覽後暫存的匯出資訊

        # 表格輸入模式的列容器
        self.table_rows = []
        
        # 編輯既有模式：每張圖紙的 (checkbox, 新圖號 textbox, 新圖名 textbox, sheet, row_border)
        self.edit_rows = []
        self.edit_search_text = ""
        # Shift 範圍勾選用：上一次點選的列
        self.last_clicked_index = None
        # 表頭總勾選框（避免事件回呼造成迴圈）
        self._suppress_header_check = False
        
        self._build_ui()
    
    @staticmethod
    def _brush(rgb):
        return SolidColorBrush(Color.FromRgb(rgb[0], rgb[1], rgb[2]))
    
    @staticmethod
    def _mono_font():
        return FontFamily("Consolas, Courier New, monospace")
    
    def _build_ui(self):
        self.Title = "圖紙編輯與管理"
        self.Width = 950
        self.Height = 760
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = self._brush(self.COLOR_BG)
        
        root = Grid()
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        
        # 上方：圖框選擇
        tb_panel = self._build_titleblock_panel()
        Grid.SetRow(tb_panel, 0)
        root.Children.Add(tb_panel)
        
        # 中間：5 個 Tab
        self.tabs = TabControl()
        self.tabs.Margin = Thickness(20, 5, 20, 5)
        self.tabs.Items.Add(self._build_edit_existing_tab())
        self.tabs.Items.Add(self._build_table_tab())
        self.tabs.Items.Add(self._build_paste_tab())
        self.tabs.Items.Add(self._build_pattern_tab())
        self.tabs.Items.Add(self._build_sync_tab())
        self.tabs.SelectedIndex = self.MODE_EDIT_EXISTING
        self.tabs.SelectionChanged += self._on_tab_changed
        Grid.SetRow(self.tabs, 1)
        root.Children.Add(self.tabs)
        
        # 底部按鈕
        bottom = Grid()
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(140)))
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(20)))
        
        self.status_text = TextBlock()
        self.status_text.Text = ""
        self.status_text.VerticalAlignment = VerticalAlignment.Center
        self.status_text.Margin = Thickness(20, 0, 0, 0)
        self.status_text.Foreground = self._brush(self.COLOR_TEXT)
        Grid.SetColumn(self.status_text, 0)
        bottom.Children.Add(self.status_text)
        
        self.run_btn = Button()
        self.run_btn.Content = "執行"
        self.run_btn.Margin = Thickness(0, 10, 5, 10)
        self.run_btn.Background = self._brush(self.COLOR_PRIMARY)
        self.run_btn.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
        self.run_btn.FontWeight = FontWeights.Bold
        self.run_btn.Click += self._on_run
        Grid.SetColumn(self.run_btn, 1)
        bottom.Children.Add(self.run_btn)
        
        Grid.SetRow(bottom, 2)
        root.Children.Add(bottom)
        
        self.Content = root
        self._update_run_button_label()
    
    def _on_tab_changed(self, sender, args):
        self._update_run_button_label()
        self._update_titleblock_panel_visibility()
    
    def _update_run_button_label(self):
        idx = self.tabs.SelectedIndex
        # 同步分頁用自己的唯讀按鈕，隱藏底部「執行」鈕
        if idx == self.MODE_SYNC:
            self.run_btn.Visibility = Visibility.Collapsed
            return
        self.run_btn.Visibility = Visibility.Visible
        if idx == self.MODE_EDIT_EXISTING:
            self.run_btn.Content = "套用變更"
        else:
            self.run_btn.Content = "建立圖紙"
    
    def _update_titleblock_panel_visibility(self):
        """編輯既有模式不需要圖框設定，隱藏。"""
        idx = self.tabs.SelectedIndex
        if idx == self.MODE_EDIT_EXISTING:
            self.tb_label.Foreground = self._brush((180, 185, 200))
            self.tb_combo.IsEnabled = False
            self.tb_combo.Opacity = 0.4
        else:
            self.tb_label.Foreground = self._brush(self.COLOR_TEXT)
            self.tb_combo.IsEnabled = bool(self.title_blocks)
            self.tb_combo.Opacity = 1.0
    
    # ---- 圖框選擇 ----
    
    def _build_titleblock_panel(self):
        outer = Grid()
        outer.Margin = Thickness(20, 12, 20, 0)
        outer.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(80)))
        outer.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        self.tb_label = TextBlock()
        self.tb_label.Text = "圖框類型"
        self.tb_label.FontWeight = FontWeights.Bold
        self.tb_label.FontSize = 13
        self.tb_label.VerticalAlignment = VerticalAlignment.Center
        self.tb_label.Foreground = self._brush(self.COLOR_TEXT)
        Grid.SetColumn(self.tb_label, 0)
        outer.Children.Add(self.tb_label)
        
        self.tb_combo = ComboBox()
        self.tb_combo.Padding = Thickness(4, 4, 4, 4)
        self.tb_combo.VerticalAlignment = VerticalAlignment.Center
        if not self.title_blocks:
            self.tb_combo.Items.Add("（專案中沒有圖框，將使用 Revit 預設）")
            self.tb_combo.IsEnabled = False
        else:
            for label_text, _ in self.title_blocks:
                self.tb_combo.Items.Add(label_text)
        self.tb_combo.SelectedIndex = 0
        Grid.SetColumn(self.tb_combo, 1)
        outer.Children.Add(self.tb_combo)
        
        return outer
    
    def _get_selected_title_block_id(self):
        if not self.title_blocks:
            return None
        idx = self.tb_combo.SelectedIndex
        if idx < 0 or idx >= len(self.title_blocks):
            return None
        return self.title_blocks[idx][1]
    
    # ---- Tab 0: 編輯既有圖紙 ----
    
    def _build_edit_existing_tab(self):
        tab = TabItem()
        tab.Header = "  編輯既有  "
        
        outer = Grid()
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(75)))   # 提示+搜尋+全選
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(70)))   # 規則式批次填入
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(28)))   # 表頭
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))  # 列表
        
        # === 上方：提示 + 工具列 ===
        top = StackPanel()
        top.Orientation = Orientation.Vertical
        top.Margin = Thickness(10, 8, 10, 4)
        
        hint = TextBlock()
        hint.Text = "點勾選框後 Shift+點另一個 = 範圍勾選。新欄位已預填原值，直接修改即可。"
        hint.FontSize = 11
        hint.Foreground = self._brush((110, 115, 130))
        hint.Margin = Thickness(0, 0, 0, 4)
        top.Children.Add(hint)
        
        toolbar = Grid()
        toolbar.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        toolbar.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(120)))
        toolbar.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(100)))
        toolbar.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(120)))
        
        self.edit_search_box = TextBox()
        self.edit_search_box.Padding = Thickness(4, 4, 4, 4)
        self.edit_search_box.Margin = Thickness(0, 0, 5, 0)
        self.edit_search_box.TextChanged += self._on_edit_search_changed
        Grid.SetColumn(self.edit_search_box, 0)
        toolbar.Children.Add(self.edit_search_box)
        
        sel_all_btn = Button()
        sel_all_btn.Content = "全選 (顯示中)"
        sel_all_btn.Padding = Thickness(4, 4, 4, 4)
        sel_all_btn.Margin = Thickness(0, 0, 5, 0)
        sel_all_btn.Click += self._on_edit_select_all
        Grid.SetColumn(sel_all_btn, 1)
        toolbar.Children.Add(sel_all_btn)
        
        clr_btn = Button()
        clr_btn.Content = "全部取消"
        clr_btn.Padding = Thickness(4, 4, 4, 4)
        clr_btn.Margin = Thickness(0, 0, 5, 0)
        clr_btn.Click += self._on_edit_clear_all
        Grid.SetColumn(clr_btn, 2)
        toolbar.Children.Add(clr_btn)
        
        reset_btn = Button()
        reset_btn.Content = "還原勾選列原值"
        reset_btn.Padding = Thickness(4, 4, 4, 4)
        reset_btn.Click += self._on_reset_selected
        Grid.SetColumn(reset_btn, 3)
        toolbar.Children.Add(reset_btn)
        
        top.Children.Add(toolbar)
        Grid.SetRow(top, 0)
        outer.Children.Add(top)
        
        # === 規則式批次填入區 ===
        rule_box = Border()
        rule_box.Background = self._brush((232, 240, 255))
        rule_box.BorderBrush = self._brush((99, 102, 241))
        rule_box.BorderThickness = Thickness(1)
        rule_box.CornerRadius = CornerRadius(4)
        rule_box.Margin = Thickness(10, 0, 10, 4)
        rule_box.Padding = Thickness(8, 6, 8, 6)
        
        rule_grid = Grid()
        rule_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(80)))
        rule_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        rule_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(60)))
        rule_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(60)))
        rule_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(120)))
        rule_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(28)))
        rule_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(28)))
        
        # Row 0
        l1 = TextBlock()
        l1.Text = "圖號規則"
        l1.FontSize = 11
        l1.VerticalAlignment = VerticalAlignment.Center
        Grid.SetRow(l1, 0)
        Grid.SetColumn(l1, 0)
        rule_grid.Children.Add(l1)
        
        self.rule_num_box = TextBox()
        self.rule_num_box.Padding = Thickness(4, 2, 4, 2)
        self.rule_num_box.Margin = Thickness(0, 2, 5, 2)
        self.rule_num_box.Text = "A-{n:02d}"
        Grid.SetRow(self.rule_num_box, 0)
        Grid.SetColumn(self.rule_num_box, 1)
        rule_grid.Children.Add(self.rule_num_box)
        
        l2 = TextBlock()
        l2.Text = "起始"
        l2.FontSize = 11
        l2.VerticalAlignment = VerticalAlignment.Center
        l2.HorizontalAlignment = HorizontalAlignment.Right
        l2.Margin = Thickness(0, 0, 4, 0)
        Grid.SetRow(l2, 0)
        Grid.SetColumn(l2, 2)
        rule_grid.Children.Add(l2)
        
        self.rule_start_box = TextBox()
        self.rule_start_box.Padding = Thickness(4, 2, 4, 2)
        self.rule_start_box.Margin = Thickness(0, 2, 5, 2)
        self.rule_start_box.Text = "1"
        Grid.SetRow(self.rule_start_box, 0)
        Grid.SetColumn(self.rule_start_box, 3)
        rule_grid.Children.Add(self.rule_start_box)
        
        apply_num_btn = Button()
        apply_num_btn.Content = "套用到勾選"
        apply_num_btn.Padding = Thickness(4, 2, 4, 2)
        apply_num_btn.Margin = Thickness(0, 2, 0, 2)
        apply_num_btn.Click += self._on_apply_rule_num
        Grid.SetRow(apply_num_btn, 0)
        Grid.SetColumn(apply_num_btn, 4)
        rule_grid.Children.Add(apply_num_btn)
        
        # Row 1
        l3 = TextBlock()
        l3.Text = "圖名規則"
        l3.FontSize = 11
        l3.VerticalAlignment = VerticalAlignment.Center
        Grid.SetRow(l3, 1)
        Grid.SetColumn(l3, 0)
        rule_grid.Children.Add(l3)
        
        self.rule_name_box = TextBox()
        self.rule_name_box.Padding = Thickness(4, 2, 4, 2)
        self.rule_name_box.Margin = Thickness(0, 2, 5, 2)
        self.rule_name_box.Text = "{n}層平面圖"
        Grid.SetRow(self.rule_name_box, 1)
        Grid.SetColumn(self.rule_name_box, 1)
        rule_grid.Children.Add(self.rule_name_box)
        
        apply_name_btn = Button()
        apply_name_btn.Content = "套用到勾選"
        apply_name_btn.Padding = Thickness(4, 2, 4, 2)
        apply_name_btn.Margin = Thickness(0, 2, 0, 2)
        apply_name_btn.Click += self._on_apply_rule_name
        Grid.SetRow(apply_name_btn, 1)
        Grid.SetColumn(apply_name_btn, 4)
        rule_grid.Children.Add(apply_name_btn)
        
        rule_box.Child = rule_grid
        Grid.SetRow(rule_box, 1)
        outer.Children.Add(rule_box)
        
        # === 表頭 ===
        head = Grid()
        head.Margin = Thickness(10, 4, 10, 4)
        head.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(28)))
        head.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(150)))
        head.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        head.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(150)))
        head.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        # 表頭總勾選框（三態：全選 / 部分選 / 全不選）
        self.header_check = CheckBox()
        self.header_check.IsThreeState = True
        self.header_check.VerticalAlignment = VerticalAlignment.Center
        self.header_check.HorizontalAlignment = HorizontalAlignment.Center
        self.header_check.Click += self._on_header_check_clicked
        Grid.SetColumn(self.header_check, 0)
        head.Children.Add(self.header_check)
        
        for col, text in [(1, "原圖號"), (2, "原圖名"), (3, "新圖號"), (4, "新圖名")]:
            h = TextBlock()
            h.Text = text
            h.FontSize = 11
            h.FontWeight = FontWeights.Bold
            h.Foreground = self._brush(self.COLOR_TEXT)
            h.Margin = Thickness(4, 0, 4, 0)
            Grid.SetColumn(h, col)
            head.Children.Add(h)
        
        Grid.SetRow(head, 2)
        outer.Children.Add(head)
        
        # === 列表 ===
        scroll = ScrollViewer()
        scroll.Margin = Thickness(10, 0, 10, 8)
        self.edit_panel = StackPanel()
        self.edit_panel.Orientation = Orientation.Vertical
        scroll.Content = self.edit_panel
        Grid.SetRow(scroll, 3)
        outer.Children.Add(scroll)
        
        # 建立每張圖紙的列
        for i, sheet in enumerate(self.existing_sheets):
            row = self._make_edit_row(sheet, i)
            self.edit_panel.Children.Add(row)
        
        if not self.existing_sheets:
            empty = TextBlock()
            empty.Text = "（專案中沒有圖紙）"
            empty.HorizontalAlignment = HorizontalAlignment.Center
            empty.Margin = Thickness(0, 30, 0, 0)
            empty.Foreground = self._brush((150, 155, 165))
            self.edit_panel.Children.Add(empty)
        
        tab.Content = outer
        return tab
    
    def _make_edit_row(self, sheet, index):
        border = Border()
        border.Background = self._brush(self.COLOR_BTN)
        border.BorderBrush = self._brush((220, 222, 230))
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(2)
        border.Margin = Thickness(0, 1, 0, 1)
        border.Padding = Thickness(4, 4, 4, 4)
        
        grid = Grid()
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(28)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(150)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(150)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        cb = CheckBox()
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Margin = Thickness(4, 0, 0, 0)
        cb.Tag = index  # 用來判斷 Shift 範圍
        # 用 PreviewMouseLeftButtonDown 攔住點擊，自己處理 Shift 邏輯
        cb.PreviewMouseLeftButtonDown += self._on_row_check_mousedown
        cb.Checked += self._on_edit_check_changed
        cb.Unchecked += self._on_edit_check_changed
        Grid.SetColumn(cb, 0)
        grid.Children.Add(cb)
        
        old_num = TextBlock()
        old_num.Text = sheet.SheetNumber
        old_num.FontSize = 12
        old_num.FontWeight = FontWeights.SemiBold
        old_num.VerticalAlignment = VerticalAlignment.Center
        old_num.Margin = Thickness(4, 0, 4, 0)
        old_num.TextTrimming = TextTrimming.CharacterEllipsis
        Grid.SetColumn(old_num, 1)
        grid.Children.Add(old_num)
        
        old_name = TextBlock()
        old_name.Text = sheet.Name
        old_name.FontSize = 12
        old_name.VerticalAlignment = VerticalAlignment.Center
        old_name.Margin = Thickness(4, 0, 4, 0)
        old_name.Foreground = self._brush((100, 105, 115))
        old_name.TextTrimming = TextTrimming.CharacterEllipsis
        Grid.SetColumn(old_name, 2)
        grid.Children.Add(old_name)
        
        # 預填原值，使用者只要修改差異
        new_num_tb = TextBox()
        new_num_tb.Padding = Thickness(4, 3, 4, 3)
        new_num_tb.Margin = Thickness(4, 0, 4, 0)
        new_num_tb.Text = sheet.SheetNumber
        new_num_tb.TextChanged += lambda s, e, b=border: self._on_edit_value_changed(b)
        Grid.SetColumn(new_num_tb, 3)
        grid.Children.Add(new_num_tb)
        
        new_name_tb = TextBox()
        new_name_tb.Padding = Thickness(4, 3, 4, 3)
        new_name_tb.Margin = Thickness(4, 0, 4, 0)
        new_name_tb.Text = sheet.Name
        new_name_tb.TextChanged += lambda s, e, b=border: self._on_edit_value_changed(b)
        Grid.SetColumn(new_name_tb, 4)
        grid.Children.Add(new_name_tb)
        
        border.Child = grid
        # Tag: (checkbox, num_textbox, name_textbox, sheet, index)
        border.Tag = (cb, new_num_tb, new_name_tb, sheet, index)
        self.edit_rows.append(border)
        return border
    
    def _on_edit_value_changed(self, border):
        cb, num_tb, name_tb, sheet, idx = border.Tag
        # 跟原值比對才算「有變更」
        cur_num = (num_tb.Text or "").strip()
        cur_name = (name_tb.Text or "").strip()
        has_change = (cur_num != sheet.SheetNumber) or (cur_name != sheet.Name)
        if has_change:
            border.Background = self._brush(self.COLOR_HIGHLIGHT)
        else:
            border.Background = self._brush(self.COLOR_BTN)
    
    def _on_row_check_mousedown(self, sender, args):
        """攔住 CheckBox 的點擊，若按住 Shift 就做範圍勾選。"""
        try:
            from System.Windows.Input import Keyboard, ModifierKeys
            shift_held = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift
        except:
            shift_held = False
        
        clicked_idx = sender.Tag
        if shift_held and self.last_clicked_index is not None:
            # 範圍勾選
            lo = min(self.last_clicked_index, clicked_idx)
            hi = max(self.last_clicked_index, clicked_idx)
            # 用「點擊那個的相反狀態」當目標值（仿 Excel 行為）
            target = not bool(sender.IsChecked)
            for border in self.edit_rows:
                cb, _, _, _, idx = border.Tag
                if border.Visibility != Visibility.Visible:
                    continue
                if lo <= idx <= hi:
                    cb.IsChecked = target
            # 阻止預設處理（不然會再 toggle 一次當前 cb）
            args.Handled = True
            self.last_clicked_index = clicked_idx
            self._update_header_check()
        else:
            # 一般點擊 → 記錄為錨點，讓預設處理切換狀態
            self.last_clicked_index = clicked_idx
    
    def _on_edit_check_changed(self, sender, args):
        self._update_edit_status()
        self._update_header_check()
    
    def _update_header_check(self):
        """根據實際狀態更新表頭勾選框（三態）。"""
        if not hasattr(self, "header_check"):
            return
        visible_rows = [b for b in self.edit_rows if b.Visibility == Visibility.Visible]
        if not visible_rows:
            return
        checked_count = sum(1 for b in visible_rows if b.Tag[0].IsChecked)
        self._suppress_header_check = True
        try:
            if checked_count == 0:
                self.header_check.IsChecked = False
            elif checked_count == len(visible_rows):
                self.header_check.IsChecked = True
            else:
                self.header_check.IsChecked = None  # 部分選 = indeterminate
        finally:
            self._suppress_header_check = False
    
    def _on_header_check_clicked(self, sender, args):
        """表頭勾選框被使用者點擊：全選顯示中 / 全不選顯示中。"""
        if self._suppress_header_check:
            return
        # 三態的 Click 後，IsChecked 會在 True/False/None 之間循環，
        # 我們強制只用 True/False（全選 or 全清）
        target = bool(self.header_check.IsChecked)
        if self.header_check.IsChecked is None:
            target = True
            self._suppress_header_check = True
            try:
                self.header_check.IsChecked = True
            finally:
                self._suppress_header_check = False
        for border in self.edit_rows:
            if border.Visibility == Visibility.Visible:
                border.Tag[0].IsChecked = target
    
    def _on_edit_select_all(self, sender, args):
        for border in self.edit_rows:
            if border.Visibility == Visibility.Visible:
                border.Tag[0].IsChecked = True
    
    def _on_edit_clear_all(self, sender, args):
        for border in self.edit_rows:
            border.Tag[0].IsChecked = False
    
    def _on_reset_selected(self, sender, args):
        """把勾選列的新欄位還原為原值。"""
        for border in self.edit_rows:
            cb, num_tb, name_tb, sheet, idx = border.Tag
            if cb.IsChecked:
                num_tb.Text = sheet.SheetNumber
                name_tb.Text = sheet.Name
    
    def _on_edit_search_changed(self, sender, args):
        self.edit_search_text = (sender.Text or "").lower()
        for border in self.edit_rows:
            cb, num_tb, name_tb, sheet, idx = border.Tag
            if not self.edit_search_text:
                border.Visibility = Visibility.Visible
                continue
            combined = "{} {}".format(sheet.SheetNumber, sheet.Name).lower()
            if self.edit_search_text in combined:
                border.Visibility = Visibility.Visible
            else:
                border.Visibility = Visibility.Collapsed
        self._update_header_check()
    
    def _on_apply_rule_num(self, sender, args):
        """把規則生成的圖號填入「目前勾選且顯示中」的列。"""
        try:
            start = int((self.rule_start_box.Text or "1").strip())
        except:
            forms.alert("起始編號必須是數字")
            return
        pattern = (self.rule_num_box.Text or "").strip()
        if not pattern:
            forms.alert("請填寫圖號規則")
            return
        
        targets = [b for b in self.edit_rows
                   if b.Visibility == Visibility.Visible and b.Tag[0].IsChecked]
        if not targets:
            forms.alert("請先勾選要套用的圖紙列")
            return
        
        nums = expand_pattern_for_count(pattern, len(targets), start)
        for border, num in zip(targets, nums):
            cb, num_tb, name_tb, sheet, idx = border.Tag
            num_tb.Text = num
    
    def _on_apply_rule_name(self, sender, args):
        try:
            start = int((self.rule_start_box.Text or "1").strip())
        except:
            forms.alert("起始編號必須是數字")
            return
        pattern = (self.rule_name_box.Text or "").strip()
        if not pattern:
            forms.alert("請填寫圖名規則")
            return
        
        targets = [b for b in self.edit_rows
                   if b.Visibility == Visibility.Visible and b.Tag[0].IsChecked]
        if not targets:
            forms.alert("請先勾選要套用的圖紙列")
            return
        
        names = expand_pattern_for_count(pattern, len(targets), start)
        for border, name in zip(targets, names):
            cb, num_tb, name_tb, sheet, idx = border.Tag
            name_tb.Text = name
    
    def _update_edit_status(self):
        n = sum(1 for b in self.edit_rows if b.Tag[0].IsChecked)
        if n == 0:
            self.status_text.Text = "尚未勾選任何圖紙"
        else:
            self.status_text.Text = "已勾選 {} 張圖紙".format(n)
    
    def _collect_edits(self):
        """回傳 [(sheet, new_num_or_None, new_name_or_None), ...]，僅包含勾選且至少有一項變更的列。"""
        result = []
        for border in self.edit_rows:
            cb, num_tb, name_tb, sheet, idx = border.Tag
            if not cb.IsChecked:
                continue
            new_num = (num_tb.Text or "").strip() or None
            new_name = (name_tb.Text or "").strip() or None
            if new_num is None and new_name is None:
                continue
            # 跟原值一樣就視為沒變
            if new_num == sheet.SheetNumber:
                new_num = None
            if new_name == sheet.Name:
                new_name = None
            if new_num is None and new_name is None:
                continue
            result.append((sheet, new_num, new_name))
        return result
    
    # ---- Tab 1: 表格輸入（新增） ----
    
    def _build_table_tab(self):
        tab = TabItem()
        tab.Header = "  逐筆輸入  "
        
        outer = Grid()
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
        
        head = Grid()
        head.Margin = Thickness(10, 8, 10, 4)
        head.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(180)))
        head.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        head.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(60)))
        
        for col, text in [(0, "圖號"), (1, "圖名")]:
            h = TextBlock()
            h.Text = text
            h.FontWeight = FontWeights.Bold
            h.Foreground = self._brush(self.COLOR_TEXT)
            Grid.SetColumn(h, col)
            head.Children.Add(h)
        
        Grid.SetRow(head, 0)
        outer.Children.Add(head)
        
        scroll = ScrollViewer()
        scroll.Margin = Thickness(10, 0, 10, 0)
        self.table_panel = StackPanel()
        self.table_panel.Orientation = Orientation.Vertical
        scroll.Content = self.table_panel
        Grid.SetRow(scroll, 1)
        outer.Children.Add(scroll)
        
        for _ in range(5):
            self._add_table_row()
        
        bottom = StackPanel()
        bottom.Orientation = Orientation.Horizontal
        bottom.Margin = Thickness(10, 8, 10, 8)
        
        add_btn = Button()
        add_btn.Content = "+ 新增一列"
        add_btn.Padding = Thickness(10, 4, 10, 4)
        add_btn.Margin = Thickness(0, 0, 8, 0)
        add_btn.Click += lambda s, e: self._add_table_row()
        bottom.Children.Add(add_btn)
        
        add5_btn = Button()
        add5_btn.Content = "+ 新增 5 列"
        add5_btn.Padding = Thickness(10, 4, 10, 4)
        add5_btn.Click += lambda s, e: [self._add_table_row() for _ in range(5)]
        bottom.Children.Add(add5_btn)
        
        Grid.SetRow(bottom, 2)
        outer.Children.Add(bottom)
        
        tab.Content = outer
        return tab
    
    def _add_table_row(self):
        row = Grid()
        row.Margin = Thickness(0, 2, 0, 2)
        row.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(180)))
        row.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        row.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(60)))
        
        num_tb = TextBox()
        num_tb.Padding = Thickness(4, 3, 4, 3)
        num_tb.Margin = Thickness(0, 0, 4, 0)
        Grid.SetColumn(num_tb, 0)
        row.Children.Add(num_tb)
        
        name_tb = TextBox()
        name_tb.Padding = Thickness(4, 3, 4, 3)
        Grid.SetColumn(name_tb, 1)
        row.Children.Add(name_tb)
        
        del_btn = Button()
        del_btn.Content = "✕"
        del_btn.Margin = Thickness(4, 0, 0, 0)
        del_btn.Click += lambda s, e, r=row: self._remove_table_row(r)
        Grid.SetColumn(del_btn, 2)
        row.Children.Add(del_btn)
        
        self.table_rows.append((num_tb, name_tb, row))
        self.table_panel.Children.Add(row)
    
    def _remove_table_row(self, row):
        for i, (a, b, r) in enumerate(self.table_rows):
            if r is row:
                self.table_rows.pop(i)
                break
        self.table_panel.Children.Remove(row)
    
    def _collect_from_table(self):
        result = []
        for num_tb, name_tb, _ in self.table_rows:
            num = (num_tb.Text or "").strip()
            name = (name_tb.Text or "").strip()
            if num or name:
                result.append((num, name))
        return result
    
    # ---- Tab 2: 貼上 ----
    
    def _build_paste_tab(self):
        tab = TabItem()
        tab.Header = "  貼上文字  "
        
        outer = Grid()
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(80)))
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        
        hint = Border()
        hint.Background = self._brush((255, 247, 230))
        hint.BorderBrush = self._brush((250, 204, 21))
        hint.BorderThickness = Thickness(1)
        hint.CornerRadius = CornerRadius(4)
        hint.Padding = Thickness(12, 8, 12, 8)
        hint.Margin = Thickness(10, 8, 10, 4)
        
        hint_text = TextBlock()
        hint_text.Text = ("從 Excel 或記事本貼上文字。每行一張圖紙，圖號和圖名用 Tab 鍵或多個空格分開。\n"
                          "範例：    A1-01    封面圖")
        hint_text.FontSize = 11
        hint_text.TextWrapping = TextWrapping.Wrap
        hint_text.Foreground = self._brush((120, 80, 0))
        hint.Child = hint_text
        Grid.SetRow(hint, 0)
        outer.Children.Add(hint)
        
        self.paste_textbox = TextBox()
        self.paste_textbox.Margin = Thickness(10, 0, 10, 10)
        self.paste_textbox.Padding = Thickness(8, 6, 8, 6)
        self.paste_textbox.AcceptsReturn = True
        self.paste_textbox.AcceptsTab = True
        self.paste_textbox.TextWrapping = TextWrapping.NoWrap
        self.paste_textbox.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.paste_textbox.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        self.paste_textbox.FontFamily = self._mono_font()
        Grid.SetRow(self.paste_textbox, 1)
        outer.Children.Add(self.paste_textbox)
        
        tab.Content = outer
        return tab
    
    def _collect_from_paste(self):
        text = self.paste_textbox.Text or ""
        result = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            if "\t" in line:
                parts = line.split("\t", 1)
            else:
                parts = re.split(r"\s{2,}", line, maxsplit=1)
                if len(parts) == 1:
                    parts = line.split(None, 1)
            num = parts[0].strip() if len(parts) > 0 else ""
            name = parts[1].strip() if len(parts) > 1 else ""
            result.append((num, name))
        return result
    
    # ---- Tab 3: 規則 ----
    
    def _build_pattern_tab(self):
        tab = TabItem()
        tab.Header = "  規則生成  "
        
        outer = StackPanel()
        outer.Orientation = Orientation.Vertical
        outer.Margin = Thickness(20, 16, 20, 16)
        
        hint = Border()
        hint.Background = self._brush((232, 240, 255))
        hint.BorderBrush = self._brush((99, 102, 241))
        hint.BorderThickness = Thickness(1)
        hint.CornerRadius = CornerRadius(4)
        hint.Padding = Thickness(12, 8, 12, 8)
        hint.Margin = Thickness(0, 0, 0, 16)
        
        hint_text = TextBlock()
        hint_text.Text = ("用 {n} 代表編號位置。{n:02d} 表示補零成 2 位數，{n:03d} 為 3 位數。\n"
                          "範例：圖號 A2-{n:02d}、圖名 {n}層平面圖，編號 1 ~ 20")
        hint_text.FontSize = 11
        hint_text.TextWrapping = TextWrapping.Wrap
        hint_text.Foreground = self._brush((40, 50, 80))
        hint.Child = hint_text
        outer.Children.Add(hint)
        
        outer.Children.Add(self._make_label("圖號樣式"))
        self.pattern_num_box = TextBox()
        self.pattern_num_box.Padding = Thickness(4, 4, 4, 4)
        self.pattern_num_box.Margin = Thickness(0, 0, 0, 12)
        self.pattern_num_box.Text = "A2-{n:02d}"
        outer.Children.Add(self.pattern_num_box)
        
        outer.Children.Add(self._make_label("圖名樣式"))
        self.pattern_name_box = TextBox()
        self.pattern_name_box.Padding = Thickness(4, 4, 4, 4)
        self.pattern_name_box.Margin = Thickness(0, 0, 0, 12)
        self.pattern_name_box.Text = "{n}層平面圖"
        outer.Children.Add(self.pattern_name_box)
        
        range_panel = Grid()
        range_panel.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        range_panel.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(20)))
        range_panel.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        range_panel.Margin = Thickness(0, 0, 0, 12)
        
        col1 = StackPanel()
        col1.Children.Add(self._make_label("起始"))
        self.start_box = TextBox()
        self.start_box.Padding = Thickness(4, 4, 4, 4)
        self.start_box.Text = "1"
        col1.Children.Add(self.start_box)
        Grid.SetColumn(col1, 0)
        range_panel.Children.Add(col1)
        
        col2 = StackPanel()
        col2.Children.Add(self._make_label("結束"))
        self.end_box = TextBox()
        self.end_box.Padding = Thickness(4, 4, 4, 4)
        self.end_box.Text = "20"
        col2.Children.Add(self.end_box)
        Grid.SetColumn(col2, 2)
        range_panel.Children.Add(col2)
        
        outer.Children.Add(range_panel)
        
        preview_btn = Button()
        preview_btn.Content = "預覽生成結果"
        preview_btn.Padding = Thickness(10, 5, 10, 5)
        preview_btn.HorizontalAlignment = HorizontalAlignment.Left
        preview_btn.Click += self._on_preview_pattern
        outer.Children.Add(preview_btn)
        
        self.pattern_preview = TextBlock()
        self.pattern_preview.Margin = Thickness(0, 8, 0, 0)
        self.pattern_preview.FontSize = 11
        self.pattern_preview.Foreground = self._brush((110, 115, 130))
        self.pattern_preview.TextWrapping = TextWrapping.Wrap
        outer.Children.Add(self.pattern_preview)
        
        tab.Content = outer
        return tab
    
    def _make_label(self, text):
        tb = TextBlock()
        tb.Text = text
        tb.FontWeight = FontWeights.SemiBold
        tb.FontSize = 12
        tb.Margin = Thickness(0, 0, 0, 4)
        tb.Foreground = self._brush(self.COLOR_TEXT)
        return tb
    
    def _on_preview_pattern(self, sender, args):
        try:
            data = self._collect_from_pattern()
        except Exception as ex:
            self.pattern_preview.Text = "❌ 錯誤: {}".format(ex)
            self.pattern_preview.Foreground = self._brush(self.COLOR_ERROR)
            return
        
        if not data:
            self.pattern_preview.Text = "（無結果）"
            return
        
        n = len(data)
        preview_lines = ["將生成 {} 張圖紙：".format(n)]
        for i, (num, name) in enumerate(data[:5]):
            preview_lines.append("  {}  →  {}".format(num, name))
        if n > 5:
            preview_lines.append("  ... 共 {} 張".format(n))
        self.pattern_preview.Text = "\n".join(preview_lines)
        self.pattern_preview.Foreground = self._brush(self.COLOR_SUCCESS)
    
    def _collect_from_pattern(self):
        num_pat = (self.pattern_num_box.Text or "").strip()
        name_pat = (self.pattern_name_box.Text or "").strip()
        try:
            start = int((self.start_box.Text or "").strip())
        except:
            raise ValueError("起始編號必須是數字")
        try:
            end = int((self.end_box.Text or "").strip())
        except:
            raise ValueError("結束編號必須是數字")
        if end < start:
            raise ValueError("結束編號必須 >= 起始編號")
        if not num_pat:
            raise ValueError("請填寫圖號樣式")
        
        nums = expand_pattern(num_pat, start, end)
        names = expand_pattern(name_pat, start, end) if name_pat else [""] * len(nums)
        return list(zip(nums, names))

    # ---- Tab 4: Google Sheet 同步（第①步：唯讀讀取測試）----

    def _build_sync_tab(self):
        tab = TabItem()
        tab.Header = "  Google Sheet 同步  "

        outer = Grid()
        outer.Margin = Thickness(20, 16, 20, 16)
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))  # 說明
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))  # 讀取按鈕列
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))  # 設定/匯出匯入
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))  # 結果(填滿剩餘)

        hint = Border()
        hint.Background = self._brush((232, 240, 255))
        hint.BorderBrush = self._brush((99, 102, 241))
        hint.BorderThickness = Thickness(1)
        hint.CornerRadius = CornerRadius(4)
        hint.Padding = Thickness(12, 8, 12, 8)
        hint.Margin = Thickness(0, 0, 0, 12)
        hint_text = TextBlock()
        hint_text.Text = (u"1、第一次使用請先讀「操作說明」，並在 Google Sheet 載入「最新腳本」。\n"
                          u"2、功能介紹：\n"
                          u"    匯出：把 Revit 圖紙索引寫進 Google Sheet。\n"
                          u"    匯入：從 Google Sheet 讀回 → 先預覽差異 → 按「確認執行」才套用；"
                          u"完成後自動回寫新 UID 並關閉視窗，異常會跳警示。\n"
                          u"    派工系統：藉由派工系統產出協作設計師工作清單。")
        hint_text.FontSize = 11
        hint_text.TextWrapping = TextWrapping.Wrap
        hint_text.Foreground = self._brush((40, 50, 80))
        hint.Child = hint_text
        Grid.SetRow(hint, 0)
        outer.Children.Add(hint)

        # 文件按鈕列：操作說明 / 最新腳本
        doc_row = StackPanel()
        doc_row.Orientation = Orientation.Horizontal
        doc_row.Margin = Thickness(0, 0, 0, 10)

        manual_btn = Button()
        manual_btn.Content = u"📖 操作說明"
        manual_btn.Padding = Thickness(10, 5, 10, 5)
        manual_btn.Margin = Thickness(0, 0, 8, 0)
        manual_btn.Click += self._on_open_manual
        doc_row.Children.Add(manual_btn)

        script_btn = Button()
        script_btn.Content = u"📄 最新腳本"
        script_btn.Padding = Thickness(10, 5, 10, 5)
        script_btn.Margin = Thickness(0, 0, 8, 0)
        script_btn.Click += self._on_open_script
        doc_row.Children.Add(script_btn)

        env_btn = Button()
        env_btn.Content = u"🛠️ 建立環境"
        env_btn.Padding = Thickness(10, 5, 10, 5)
        env_btn.Click += self._on_build_env
        doc_row.Children.Add(env_btn)

        Grid.SetRow(doc_row, 1)
        outer.Children.Add(doc_row)

        # 設定區 + 匯出/匯入/指派
        cfg = load_gsheet_cfg()
        cfg_box = Border()
        cfg_box.Background = self._brush((245, 247, 250))
        cfg_box.BorderBrush = self._brush((210, 215, 225))
        cfg_box.BorderThickness = Thickness(1)
        cfg_box.CornerRadius = CornerRadius(4)
        cfg_box.Padding = Thickness(12, 8, 12, 10)
        cfg_box.Margin = Thickness(0, 0, 0, 10)
        cfg_panel = StackPanel()
        cfg_panel.Orientation = Orientation.Vertical

        cfg_panel.Children.Add(self._make_label(u"Web App URL（/exec 結尾）"))
        self.gs_url_box = TextBox()
        self.gs_url_box.Padding = Thickness(4, 4, 4, 4)
        self.gs_url_box.Margin = Thickness(0, 0, 0, 8)
        self.gs_url_box.Text = cfg.get("url", u"")
        cfg_panel.Children.Add(self.gs_url_box)

        cfg_panel.Children.Add(self._make_label(u"目標頁籤名稱"))
        self.gs_tab_box = TextBox()
        self.gs_tab_box.Padding = Thickness(4, 4, 4, 4)
        self.gs_tab_box.Margin = Thickness(0, 0, 0, 10)
        self.gs_tab_box.Text = cfg.get("tab", u"")
        cfg_panel.Children.Add(self.gs_tab_box)

        wbtn_row = StackPanel()
        wbtn_row.Orientation = Orientation.Horizontal

        write_btn = Button()
        write_btn.Content = u"匯出（寫入 Google Sheet）"
        write_btn.Padding = Thickness(12, 6, 12, 6)
        write_btn.Margin = Thickness(0, 0, 8, 0)
        write_btn.Background = self._brush(self.COLOR_PRIMARY)
        write_btn.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
        write_btn.FontWeight = FontWeights.Bold
        write_btn.Click += self._on_write_gsheet
        wbtn_row.Children.Add(write_btn)

        import_btn = Button()
        import_btn.Content = u"匯入（預覽差異）"
        import_btn.Padding = Thickness(12, 6, 12, 6)
        import_btn.Margin = Thickness(0, 0, 8, 0)
        import_btn.Background = self._brush((217, 119, 6))
        import_btn.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
        import_btn.FontWeight = FontWeights.Bold
        import_btn.Click += self._on_import_preview
        wbtn_row.Children.Add(import_btn)

        assign_btn = Button()
        assign_btn.Content = u"指派工作任務"
        assign_btn.Padding = Thickness(12, 6, 12, 6)
        assign_btn.Margin = Thickness(0, 0, 8, 0)
        assign_btn.FontWeight = FontWeights.Bold
        assign_btn.Click += self._on_assign_tasks
        wbtn_row.Children.Add(assign_btn)

        redpen_btn = Button()
        redpen_btn.Content = u"🖍 依紅筆圖註解"
        redpen_btn.Padding = Thickness(12, 6, 12, 6)
        redpen_btn.FontWeight = FontWeights.Bold
        redpen_btn.Click += self._on_edit_by_redpen
        wbtn_row.Children.Add(redpen_btn)

        cfg_panel.Children.Add(wbtn_row)

        cfg_box.Child = cfg_panel
        Grid.SetRow(cfg_box, 2)
        outer.Children.Add(cfg_box)

        sv = ScrollViewer()
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.Background = self._brush((255, 255, 255))
        sv.BorderBrush = self._brush((200, 205, 215))
        sv.BorderThickness = Thickness(1)
        self.sync_sv = sv
        placeholder = TextBlock()
        placeholder.Margin = Thickness(12, 12, 12, 12)
        placeholder.Foreground = self._brush((110, 115, 130))
        placeholder.Text = u"（按「匯入（預覽差異）」開始）"
        sv.Content = placeholder
        Grid.SetRow(sv, 3)
        outer.Children.Add(sv)

        tab.Content = outer
        return tab

    # ---- 同步分頁：唯讀讀取邏輯 ----

    def _read_text_param(self, sheet, pname):
        """讀參數文字值。None = 此圖紙沒有這個參數；'' = 有參數但空值。"""
        p = _lookup_sheet_param(sheet, pname)
        if p is None:
            return None
        try:
            if p.StorageType == DB.StorageType.String:
                val = p.AsString()
            else:
                val = p.AsValueString()
            return val if val is not None else u""
        except Exception:
            return u""

    def _all_sheets_sorted(self):
        sheets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements()

        # 依 Revit 明細表的編排原則：先分組(圖紙類別)再依圖號排序
        def _key(s):
            cat = self._read_text_param(s, u"圖紙類別") or u""
            return (cat, s.SheetNumber or u"")
        return sorted(sheets, key=_key)

    def _make_cell(self, text, bold=False, color=None):
        tb = TextBlock()
        tb.Text = text if text is not None else u""
        tb.FontSize = 12
        tb.Margin = Thickness(6, 2, 14, 2)
        tb.TextTrimming = TextTrimming.CharacterEllipsis
        if bold:
            tb.FontWeight = FontWeights.Bold
        tb.Foreground = self._brush(color or self.COLOR_TEXT)
        return tb

    def _build_grid_table(self, headers, rows):
        """用 WPF Grid 排版，欄位自動對齊（不受中英文字寬影響）。"""
        grid = Grid()
        for _ in headers:
            grid.ColumnDefinitions.Add(
                ColumnDefinition(Width=GridLength(1, GridUnitType.Auto)))
        for _ in range(len(rows) + 1):
            grid.RowDefinitions.Add(
                RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        for c, h in enumerate(headers):
            cell = self._make_cell(h, bold=True)
            Grid.SetRow(cell, 0)
            Grid.SetColumn(cell, c)
            grid.Children.Add(cell)
        for r, rowdata in enumerate(rows):
            for c, val in enumerate(rowdata):
                cell = self._make_cell(val)
                Grid.SetRow(cell, r + 1)
                Grid.SetColumn(cell, c)
                grid.Children.Add(cell)
        return grid

    def _on_read_index(self, sender, args):
        sheets = self._all_sheets_sorted()
        container = StackPanel()
        container.Margin = Thickness(10, 10, 10, 10)
        if not sheets:
            container.Children.Add(self._make_cell(u"（模型內沒有圖紙）"))
            self.sync_sv.Content = container
            return
        params = self.SYNC_TEXT_PARAMS
        headers = [u"UID(前8)", u"狀態", u"圖號", u"圖名"] + params
        rows = []
        n_place = 0
        for s in sheets:
            if s.IsPlaceholder:
                n_place += 1
                state = u"僅文字資訊"
            else:
                state = u"真實圖紙"
            row = [s.UniqueId[:8], state, s.SheetNumber or u"", s.Name or u""]
            for pn in params:
                v = self._read_text_param(s, pn)
                if v is None:
                    row.append(u"⚠ 無此參數")
                else:
                    row.append(v)  # 空值就留白
            rows.append(row)
        summary = self._make_cell(
            u"共 {} 張圖紙　（真實圖紙 {} ／ 僅文字資訊 {}）".format(
                len(sheets), len(sheets) - n_place, n_place),
            bold=True)
        summary.Margin = Thickness(6, 2, 6, 10)
        container.Children.Add(summary)
        container.Children.Add(self._build_grid_table(headers, rows))
        self.sync_sv.Content = container

    def _on_list_params(self, sender, args):
        sheets = self._all_sheets_sorted()
        container = StackPanel()
        container.Margin = Thickness(10, 10, 10, 10)
        if not sheets:
            container.Children.Add(self._make_cell(u"（模型內沒有圖紙）"))
            self.sync_sv.Content = container
            return
        first = sheets[0]
        title = self._make_cell(
            u"第一張圖紙：{}  {}　（核對「圖紙類別/繪圖員/修正備註」的正確參數名稱）".format(
                first.SheetNumber, first.Name),
            bold=True)
        title.Margin = Thickness(6, 2, 6, 10)
        container.Children.Add(title)
        items = []
        for p in first.Parameters:
            try:
                name = p.Definition.Name
                if p.StorageType == DB.StorageType.String:
                    val = p.AsString()
                else:
                    val = p.AsValueString()
                items.append([name, val if val else u""])
            except Exception:
                pass
        items = sorted(items, key=lambda x: x[0])
        container.Children.Add(self._build_grid_table([u"參數名稱", u"目前值"], items))
        self.sync_sv.Content = container

    # ---- 同步分頁：匯出（Revit → 剪貼簿，照 Google Sheet 排版）----

    @staticmethod
    def _clean_cell(x):
        """去掉會破壞 TSV 欄位的 tab/換行。"""
        s = u"" if x is None else unicode(x)
        return s.replace(u"\t", u" ").replace(u"\r", u" ").replace(u"\n", u" ")

    def _build_export_rows(self):
        """回傳 (二維字串陣列, 圖紙數)。欄位順序: A空 B:UID C:狀態 D:類別 E:圖號 F:圖名 G:繪圖員 H:備註"""
        sheets = self._all_sheets_sorted()
        date_str = DateTime.Now.ToString("yyyyMMdd，HH:mm")

        rows = []
        # 第1列：更新時間（日期放 E 欄）
        rows.append([u"", u"更新時間：", u"", u"", date_str, u"", u"", u""])
        # 第2列：表頭
        rows.append([u"", u"UID", u"是否為Revit出圖", u"圖紙類別", u"圖紙號碼",
                     u"圖紙名稱", u"繪圖員", u"修正備註"])
        # 第3列起：資料
        for s in sheets:
            cat = self._read_text_param(s, u"圖紙類別")
            drawer = self._read_text_param(s, u"繪圖員")
            note = self._read_text_param(s, u"修正備註")
            row = [
                u"",                                       # A 空白（保留擴充）
                s.UniqueId,                                # B 完整 UniqueId（主鍵）
                u"FALSE" if s.IsPlaceholder else u"TRUE",   # C 狀態（核取方塊：勾=真實）
                cat if cat else u"",                       # D 圖紙類別（下拉）
                s.SheetNumber or u"",                      # E 圖紙號碼
                s.Name or u"",                             # F 圖紙名稱
                drawer if drawer else u"",                 # G 繪圖員（下拉）
                note if note else u"",                     # H 修正備註
            ]
            rows.append([self._clean_cell(x) for x in row])
        return rows, len(sheets)

    def _on_export_clipboard(self, sender, args):
        rows, n = self._build_export_rows()
        tsv = u"\r\n".join(u"\t".join(r) for r in rows)

        ok = True
        err = u""
        try:
            Clipboard.SetText(tsv)
        except Exception as ex:
            ok = False
            err = str(ex)

        container = StackPanel()
        container.Margin = Thickness(10, 10, 10, 10)
        if ok:
            msg = (u"✅ 已複製 {} 張圖紙到剪貼簿（含更新時間列與表頭）。\n"
                   u"請到 Google Sheet 的 A1 儲存格貼上 → 檢查下拉選單與核取方塊有沒有吃進去。").format(n)
            head = self._make_cell(msg, bold=True, color=self.COLOR_SUCCESS)
        else:
            head = self._make_cell(u"❌ 複製失敗：{}".format(err), bold=True, color=self.COLOR_ERROR)
        head.Margin = Thickness(6, 2, 6, 10)
        container.Children.Add(head)

        # 預覽：用第2列(表頭)＋資料列做表格
        if len(rows) >= 2:
            container.Children.Add(self._build_grid_table(rows[1], rows[2:]))
        self.sync_sv.Content = container

    # ---- 同步分頁：③ 一鍵寫入 Google Sheet ----

    def _show_sync_msg(self, text, color):
        container = StackPanel()
        container.Margin = Thickness(10, 10, 10, 10)
        cell = self._make_cell(text, bold=True, color=color)
        cell.TextWrapping = TextWrapping.Wrap
        container.Children.Add(cell)
        self.sync_sv.Content = container

    def _build_export_records(self):
        """回傳 (records, 圖紙數)。每筆是 dict，key = 表頭名稱（含批次參數，供拆分填標頭）。"""
        sheets = self._all_sheets_sorted()
        recs = []
        for s in sheets:
            rec = {
                u"UID": s.UniqueId,
                u"是否為Revit出圖": (not s.IsPlaceholder),   # True=Revit出圖(真實), False=CAD(預留)
                u"圖紙類別": self._read_text_param(s, u"圖紙類別") or u"",
                u"圖紙號碼": s.SheetNumber or u"",
                u"圖紙名稱": s.Name or u"",
                u"繪圖員": self._read_text_param(s, u"繪圖員") or u"",
                u"修正備註": self._read_text_param(s, u"修正備註") or u"",
            }
            for k in self.BATCH_PARAMS:
                rec[k] = self._read_text_param(s, k) or u""
            recs.append(rec)
        return recs, len(sheets)

    def _do_export(self, url, tab):
        """不經 UI 的匯出（給套用後自動回寫）。回傳 (result, err, 張數)。"""
        recs, n = self._build_export_records()
        payload = {
            "secret": u"",
            "tab": tab,
            "headerLabels": self.HEADER_LABELS,
            "checkboxLabel": self.CHECKBOX_LABEL,
            "dropdownLabels": self.DROPDOWN_LABELS,
            "records": recs,
            "updateDate": DateTime.Now.ToString("yyyyMMdd，HH:mm"),
            "lockHeader": True,
        }
        result, err = self._post_json(url, payload)
        return result, err, n

    def _do_split(self, url, tab):
        """依圖紙類別拆成多個工作表。回傳 (result, err, 張數)。"""
        recs, n = self._build_export_records()
        payload = {
            "secret": u"",
            "tab": tab,
            "action": "split",
            "headerLabels": self.HEADER_LABELS,
            "checkboxLabel": self.CHECKBOX_LABEL,
            "dropdownLabels": self.DROPDOWN_LABELS,
            "records": recs,
            "updateDate": DateTime.Now.ToString("yyyyMMdd，HH:mm"),
            "lockHeader": True,
        }
        result, err = self._post_json(url, payload)
        return result, err, n

    def _build_export_records_for_category(self, category):
        """只取某一圖紙類別(階段)的匯出 records。"""
        recs, _ = self._build_export_records()
        cat = unicode(category or u"")
        scoped = [r for r in recs
                  if unicode(r.get(u"圖紙類別") or u"") == cat]
        return scoped

    def _do_sync_category(self, url, main_tab, category):
        """範圍限定回寫：只更新總表中該類別的列、只重建該類別分表。

        其他階段與分表完全不動。回傳 (result, err, 該類別張數)。
        """
        recs = self._build_export_records_for_category(category)
        payload = {
            "secret": u"",
            "tab": main_tab,
            "action": "syncCategory",
            "category": category,
            "headerLabels": self.HEADER_LABELS,
            "checkboxLabel": self.CHECKBOX_LABEL,
            "dropdownLabels": self.DROPDOWN_LABELS,
            "records": recs,
            "updateDate": DateTime.Now.ToString("yyyyMMdd，HH:mm"),
            "lockHeader": True,
        }
        result, err = self._post_json(url, payload)
        return result, err, len(recs)

    @staticmethod
    def _norm(v):
        if v is True:
            return u"TRUE"
        if v is False:
            return u"FALSE"
        if v is None:
            return u""
        return unicode(v).strip()

    def _compute_export_diff(self, revit_recs, sheet_recs):
        """比對『要寫出的 Revit 資料』與『Google Sheet 現況』，回傳 (新增, 修改, 移除)。"""
        sheet_by_uid = {}
        for r in sheet_recs:
            u = unicode(r.get(u"UID") or u"").strip()
            if u:
                sheet_by_uid[u] = r
        fields = [u"是否為Revit出圖", u"圖紙類別", u"圖紙號碼",
                  u"圖紙名稱", u"繪圖員", u"修正備註"]
        revit_uids = set()
        added, modified, removed = [], [], []
        for rec in revit_recs:
            u = unicode(rec.get(u"UID") or u"").strip()
            revit_uids.add(u)
            num = unicode(rec.get(u"圖紙號碼") or u"")
            nm = unicode(rec.get(u"圖紙名稱") or u"")
            old = sheet_by_uid.get(u)
            if old is None:
                added.append([u"🆕 新增到表", num, nm, u""])
            else:
                diffs = [f for f in fields
                         if self._norm(rec.get(f)) != self._norm(old.get(f))]
                if diffs:
                    modified.append([u"✏️ 修改", num, nm, u"、".join(diffs)])
        for u, r in sheet_by_uid.items():
            if u not in revit_uids:
                removed.append([u"🗑️ 從表移除",
                                unicode(r.get(u"圖紙號碼") or u""),
                                unicode(r.get(u"圖紙名稱") or u""),
                                u"Revit 已無此圖"])
        return added, modified, removed

    def _on_write_gsheet(self, sender, args):
        url = (self.gs_url_box.Text or u"").strip()
        tab = (self.gs_tab_box.Text or u"").strip()
        if not url:
            forms.alert(u"請先貼上 Web App URL。")
            return
        if not tab:
            forms.alert(u"請填寫目標頁籤名稱。")
            return
        save_gsheet_cfg({"url": url, "tab": tab, "secret": u""})

        # 先讀目前 Google Sheet，算出『匯出後會做的變更』給你預覽
        revit_recs, n = self._build_export_records()
        result, err = self._post_json(
            url, {"secret": u"", "tab": tab, "action": "read",
                  "headerLabels": self.HEADER_LABELS})
        if err:
            self._show_sync_msg(u"❌ 連線失敗：{}".format(err), self.COLOR_ERROR)
            return

        # 防呆：避免把「總表」整批匯出蓋掉「某圖紙類別的分表」
        target_role = result.get("role") if (result and result.get("ok")) else u""
        if target_role == u"sub":
            target_cat = result.get("category") or tab
            main_tab = result.get("mainTab") or u""
            forms.alert(
                u"目標頁籤「{}」是圖紙類別「{}」的【分表】，\n"
                u"不能把總表整批匯出到這裡（會蓋掉分表內容）。\n\n"
                u"請把「目標頁籤名稱」改成總表頁籤{}，再匯出。\n"
                u"分表會在匯出後由「依圖紙類別拆分」自動更新。".format(
                    tab, target_cat,
                    u"（目前偵測到：{}）".format(main_tab) if main_tab else u""))
            return

        sheet_recs = (result.get("records") if (result and result.get("ok")) else []) or []

        added, modified, removed = self._compute_export_diff(revit_recs, sheet_recs)
        self._pending_export = {"url": url, "tab": tab, "n": n}
        self._show_export_diff(tab, n, added, modified, removed)

    def _show_export_diff(self, tab, n, added, modified, removed):
        container = StackPanel()
        container.Margin = Thickness(10, 10, 10, 10)
        summary = self._make_cell(
            u"匯出預覽（尚未寫入，按下方「確認匯出」才會寫）\n"
            u"目標頁籤：{}　共 {} 張圖紙\n"
            u"新增到表 {}／修改 {}／從表移除 {}".format(
                tab, n, len(added), len(modified), len(removed)),
            bold=True)
        summary.TextWrapping = TextWrapping.Wrap
        summary.Margin = Thickness(6, 2, 6, 10)
        container.Children.Add(summary)

        rows = added + modified + removed
        if not rows:
            container.Children.Add(self._make_cell(
                u"（內容與 Google Sheet 一致，匯出只會刷新更新時間）"))
        else:
            container.Children.Add(
                self._build_grid_table([u"動作", u"圖號", u"圖名", u"細節"], rows))

        exec_btn = Button()
        exec_btn.Content = u"✅ 確認匯出（寫入 Google Sheet）"
        exec_btn.Padding = Thickness(14, 8, 14, 8)
        exec_btn.Margin = Thickness(0, 14, 0, 0)
        exec_btn.HorizontalAlignment = HorizontalAlignment.Left
        exec_btn.Background = self._brush(self.COLOR_PRIMARY)
        exec_btn.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
        exec_btn.FontWeight = FontWeights.Bold
        exec_btn.Click += self._on_export_execute
        container.Children.Add(exec_btn)
        self.sync_sv.Content = container

    def _on_export_execute(self, sender, args):
        pe = getattr(self, "_pending_export", None)
        if not pe:
            forms.alert(u"請先按「匯出」預覽。")
            return
        url, tab, n = pe["url"], pe["tab"], pe["n"]
        result, err, _ = self._do_export(url, tab)
        if err:
            self._show_sync_msg(u"❌ 連線失敗：{}".format(err), self.COLOR_ERROR)
            return
        if not (result and result.get("ok")):
            self._show_sync_msg(
                u"❌ 寫入失敗：{}".format(result.get("error") if result else u"無回應"),
                self.COLOR_ERROR)
            return
        ver = result.get("apiVersion") or u"舊版(請重新部署新版本!)"
        msg = (u"✅ 匯出完成！　Google端程式版本：{}\n"
               u"寫入 {} 張圖紙到「{}」。{}").format(
                   ver, result.get("wrote", n), result.get("tab", tab),
                   u"　🔒 表頭已鎖定。" if result.get("locked") else u"")
        note = result.get("note") or u""
        if note:
            msg += u"\n⚠ 備註：{}".format(note)
        # 第二步：一律依「圖紙類別」一併更新所有分表（不再詢問），讓總表與分表
        # 的所有欄位(含修正備註)每次匯出都同步。
        sres, serr, _ = self._do_split(url, tab)
        if serr:
            msg += u"\n\n⚠ 分表更新失敗：{}".format(serr)
        elif sres and sres.get("ok"):
            made = sres.get("splitTabs") or []
            msg += (u"\n\n🗂️ 已一併更新 {} 個分表：{}".format(len(made), u"、".join(made))
                    if made else u"\n\n（沒有可依圖紙類別拆分的資料）")
        else:
            msg += u"\n\n⚠ 分表更新回應異常：{}".format(sres)
        self._show_sync_msg(msg, self.COLOR_SUCCESS)

    # ---- 同步分頁：④ 匯入（從 Google Sheet 讀回 + 預覽差異，唯讀不套用）----

    def _post_json(self, url, payload):
        """POST JSON，回傳 (dict 或 None, 錯誤字串 或 None)。"""
        body = json.dumps(payload)
        try:
            from System.Net import (WebClient, ServicePointManager,
                                    SecurityProtocolType)
            from System.Text import Encoding
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            wc = WebClient()
            wc.Encoding = Encoding.UTF8
            wc.Headers.Add("Content-Type", "application/json")
            resp = wc.UploadString(url, "POST", body)
        except Exception as ex:
            return None, u"連線失敗：{}".format(ex)
        try:
            return json.loads(resp), None
        except Exception:
            return None, u"回應解析失敗：{}".format(resp)

    @staticmethod
    def _truthy(v):
        if v is True:
            return True
        if v is False or v is None:
            return False
        return unicode(v).strip().upper() in (u"TRUE", u"1", u"V", u"X", u"是", u"YES")

    def _build_uid_sheet_map(self):
        """UID → (圖號, 圖名)，供匯入資料檢查顯示對應的現有圖紙。"""
        m = {}
        for s in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements():
            try:
                m[s.UniqueId] = (s.SheetNumber or u"", s.Name or u"")
            except Exception:
                pass
        return m

    def _on_import_preview(self, sender, args):
        self._pending_plan = None
        url = (self.gs_url_box.Text or u"").strip()
        tab = (self.gs_tab_box.Text or u"").strip()
        if not url or not tab:
            forms.alert(u"請先填 Web App URL 與目標頁籤名稱。")
            return
        save_gsheet_cfg({"url": url, "tab": tab, "secret": u""})
        result, err = self._post_json(
            url, {"secret": u"", "tab": tab, "action": "read",
                  "headerLabels": self.HEADER_LABELS})
        if err:
            self._show_sync_msg(u"❌ {}".format(err), self.COLOR_ERROR)
            return
        if not (result and result.get("ok")):
            self._show_sync_msg(
                u"❌ 讀取失敗：{}".format(result.get("error") if result else u"無回應"),
                self.COLOR_ERROR)
            return
        records = result.get("records") or []
        if not records:
            forms.alert(u"Google Sheet 讀不到任何資料列，為避免誤刪已中止。\n"
                        u"請確認頁籤名稱正確、表格有資料。")
            return

        # 先檢查 UID 異常（有 UID 但無值 / 重複 UID）→ 跳對話框讓使用者處理
        if _ImportCleanupWindow.has_issues(records):
            dlg = _ImportCleanupWindow(records, self._build_uid_sheet_map())
            dlg.ShowDialog()
            if not dlg.confirmed:
                return  # 取消匯入
            records = dlg.cleaned_records
            if not records:
                forms.alert(u"處理後沒有可匯入的列，已中止。")
                return

        resolved_dup = False
        if self._check_dup_numbers(records):
            # 跳出對話框，讓使用者就地把重複圖號改掉（直接改 records）
            dlg = _DupResolveWindow(records)
            dlg.ShowDialog()
            if not dlg.confirmed:
                return  # 取消匯入
            if self._check_dup_numbers(records):
                forms.alert(u"仍有重複圖號，已中止匯入。請再試一次。")
                return
            resolved_dup = True
        role = result.get("role") or u"main"
        category = result.get("category") or u""
        batch = result.get("batchParams") or {}
        main_tab = result.get("mainTab") or tab

        if role == u"sub" and category:
            plan = self._build_plan(records, scope_category=category, batch=batch)
            scope_note = u"【分表匯入】只影響圖紙類別＝「{}」的圖紙。".format(category)
        else:
            plan = self._build_plan(records)
            scope_note = u"【總表匯入】影響全部圖紙。"

        plan["tb_id"] = self._get_selected_title_block_id() or DB.ElementId.InvalidElementId
        plan["url"] = url
        plan["tab"] = tab
        plan["role"] = role
        plan["mainTab"] = main_tab
        plan["category"] = category
        plan["resolvedDup"] = resolved_dup

        # 新增圖紙的圖號是否與總表(現有 Revit 圖紙)撞號 → 彈窗顯示衝突圖紙資訊並中止。
        # 特別針對分表匯入：同類別比對看不到別類別已占用的圖號。
        conflicts = self._find_new_number_conflicts(plan)
        if conflicts:
            self._pending_plan = None
            self._show_new_number_conflicts(conflicts)
            return

        n_changes = (len(plan["new"]) + len(plan["edit"]) + len(plan["toReal"])
                     + len(plan["toPlace"]) + len(plan["delete"]))
        self._pending_plan = plan if n_changes else None
        self._show_diff(plan, scope_note)

    def _edit_detail(self, d):
        """產生『修改了哪些欄位』的說明文字（給預覽用）。"""
        s = d.get("sheet")
        if not s:
            return u""
        out = []
        if d["num"] and d["num"] != (s.SheetNumber or u""):
            out.append(u"圖號→{}".format(d["num"]))
        if d["name"] and d["name"] != (s.Name or u""):
            out.append(u"圖名→{}".format(d["name"]))
        # 修正備註：以 Sheet 為準（含清空）→ 顯示變更，清空標示「(清空)」
        cur_note = self._read_text_param(s, u"修正備註") or u""
        if d["note"] != cur_note:
            out.append(u"修正備註→{}".format(d["note"] if d["note"] else u"(清空)"))
        for label, val in ((u"圖紙類別", d["cat"]), (u"繪圖員", d["drawer"])):
            if not val:
                continue
            cur = self._read_text_param(s, label) or u""
            if val != cur:
                out.append(u"{}→{}".format(label, val))
        for k, v in d.get("extra", {}).items():
            if not v:
                continue
            cur = self._read_text_param(s, k) or u""
            if v != cur:
                out.append(u"{}→{}".format(k, v))
        return u"; ".join(out)

    def _plan_to_rows(self, plan):
        """把套用計畫轉成預覽表格列（與實際套用同一份計畫，不會有落差）。"""
        rows = []
        for d in plan["new"]:
            tag = u"🆕 新增" + (u"(真實圖紙)" if d["real"] else u"(僅文字資訊)")
            extra = u"; ".join([u"{}={}".format(k, v)
                                for k, v in d.get("extra", {}).items() if v])
            rows.append([tag, d["num"], d["name"], extra])
        for d in plan["toReal"]:
            s = d["sheet"]
            rows.append([u"⬆️ 轉真實圖紙", d["num"] or s.SheetNumber,
                         d["name"] or s.Name, self._edit_detail(d)])
        for d in plan["edit"]:
            s = d["sheet"]
            rows.append([u"✏️ 修改", d["num"] or s.SheetNumber,
                         d["name"] or s.Name, self._edit_detail(d)])
        for d in plan["toPlace"]:
            s = d["sheet"]
            rows.append([u"⚠️ 真實圖紙→僅文字資訊", s.SheetNumber, s.Name,
                         u"刪原圖改建為僅文字資訊"])
        for d in plan["delete"]:
            s = d["sheet"]
            rows.append([u"🗑️ 刪除", s.SheetNumber, s.Name, u"Sheet 已移除"])
        return rows

    def _show_diff(self, plan, scope_note=u""):
        container = StackPanel()
        container.Margin = Thickness(10, 10, 10, 10)
        rows = self._plan_to_rows(plan)

        summary = self._make_cell(
            (scope_note + u"\n" if scope_note else u"") +
            u"差異預覽（唯讀，按下方「確認執行」才會套用）\n"
            u"新增 {}／修改 {}／轉真實圖紙 {}／降為僅文字資訊 {}／刪除 {}".format(
                len(plan["new"]), len(plan["edit"]), len(plan["toReal"]),
                len(plan["toPlace"]), len(plan["delete"])),
            bold=True)
        summary.TextWrapping = TextWrapping.Wrap
        summary.Margin = Thickness(6, 2, 6, 10)
        container.Children.Add(summary)

        if not rows:
            container.Children.Add(
                self._make_cell(u"✅ 沒有差異，Revit 與 Google Sheet 一致。",
                                bold=True, color=self.COLOR_SUCCESS))
        else:
            container.Children.Add(
                self._build_grid_table([u"動作", u"圖號", u"圖名", u"細節"], rows))
            exec_btn = Button()
            exec_btn.Content = u"✅ 確認執行（套用到 Revit）"
            exec_btn.Padding = Thickness(14, 8, 14, 8)
            exec_btn.Margin = Thickness(0, 14, 0, 0)
            exec_btn.HorizontalAlignment = HorizontalAlignment.Left
            exec_btn.Background = self._brush(self.COLOR_SUCCESS)
            exec_btn.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
            exec_btn.FontWeight = FontWeights.Bold
            exec_btn.Click += self._on_import_execute
            container.Children.Add(exec_btn)

        self.sync_sv.Content = container

    # ---- 同步分頁：⑤ 套用匯入（會修改 Revit）----

    def _set_param(self, sheet, pname, value):
        p = _lookup_sheet_param(sheet, pname)
        if p is None or p.IsReadOnly:
            return
        try:
            if p.StorageType == DB.StorageType.String:
                p.Set(value if value is not None else u"")
        except Exception:
            pass

    def _needs_edit(self, s, d):
        # 其他欄位：空白 = 不變更（與「編輯既有」一致，不會清掉既有資料）
        if d["num"] and d["num"] != (s.SheetNumber or u""):
            return True
        if d["name"] and d["name"] != (s.Name or u""):
            return True
        # 修正備註：以 Google Sheet 為準（含清空）→ 與現值不同即視為變更，
        # 讓「Sheet 清空 → 移除 Revit 修正備註」也能被偵測到。
        cur_note = self._read_text_param(s, u"修正備註")
        cur_note = cur_note if cur_note is not None else u""
        if d["note"] != cur_note:
            return True
        for label, val in ((u"圖紙類別", d["cat"]), (u"繪圖員", d["drawer"])):
            if not val:
                continue
            cur = self._read_text_param(s, label)
            cur = cur if cur is not None else u""
            if val != cur:
                return True
        for k, v in d.get("extra", {}).items():
            if not v:
                continue
            cur = self._read_text_param(s, k)
            cur = cur if cur is not None else u""
            if v != cur:
                return True
        return False

    def _build_plan(self, records, scope_category=None, batch=None):
        all_sheets = list(
            DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements())
        if scope_category is not None:
            # 分表匯入：只在「該圖紙類別」範圍內比對(刪除也只限這個類別)
            scoped = [s for s in all_sheets
                      if (self._read_text_param(s, u"圖紙類別") or u"") == scope_category]
            by_uid = dict((s.UniqueId, s) for s in scoped)
        else:
            by_uid = dict((s.UniqueId, s) for s in all_sheets)
        seen = set()
        plan = {"new": [], "edit": [], "toReal": [], "toPlace": [],
                "delete": [], "orphan": 0}
        # 標頭批次參數(分表用)：套用到該類別所有圖紙
        batch_extra = {}
        if batch:
            for k in self.BATCH_PARAMS:
                v = batch.get(k)
                if v not in (None, u""):
                    batch_extra[k] = unicode(v)
        for rec in records:
            uid = unicode(rec.get(u"UID") or u"").strip()
            d = {
                "num": unicode(rec.get(u"圖紙號碼") or u""),
                "name": unicode(rec.get(u"圖紙名稱") or u""),
                "cat": unicode(rec.get(u"圖紙類別") or (scope_category or u"")),
                "drawer": unicode(rec.get(u"繪圖員") or u""),
                # 修正備註：把 Sheet 顯示用的換行正規化回「；」，避免換行寫進 Revit
                # 參數後明細表只顯示第一行（與匯出時「；→換行」對應）。
                "note": unicode(rec.get(u"修正備註") or u"").replace(
                    u"\r\n", u"；").replace(u"\n", u"；").replace(u"\r", u"；"),
                "real": self._truthy(rec.get(u"是否為Revit出圖")),
                "extra": dict(batch_extra),
            }
            if not uid:
                plan["new"].append(d)
                continue
            seen.add(uid)
            s = by_uid.get(uid)
            if s is None:
                # 沒有相對應的 id → 當新圖紙建立（套用後回寫正確的新 UID）
                plan["new"].append(d)
                plan["orphan"] += 1
                continue
            d["sheet"] = s
            if d["real"] and s.IsPlaceholder:
                plan["toReal"].append(d)
            elif (not d["real"]) and (not s.IsPlaceholder):
                plan["toPlace"].append(d)
            elif self._needs_edit(s, d):
                plan["edit"].append(d)
        # 範圍內：Sheet 已把整列(含 id)刪除 → Revit 仍有 → 視為刪除
        for uid, s in by_uid.items():
            if uid not in seen:
                plan["delete"].append({"sheet": s})
        return plan

    def _check_dup_numbers(self, records):
        """檢查圖號是否重複；有重複回傳警示文字，否則 None。"""
        num_count = {}
        for rec in records:
            rnum = unicode(rec.get(u"圖紙號碼") or u"").strip()
            if not rnum:
                continue  # 空白圖號保留 Revit 預設，不列入重複檢查
            num_count[rnum] = num_count.get(rnum, 0) + 1
        dups = sorted([n for n, c in num_count.items() if c > 1])
        if dups:
            return (u"發現重複圖號，無法執行匯入。\n"
                    u"請先到 Google Sheet 把下列圖號改成唯一，再按一次「匯入」：\n\n"
                    + u"\n".join(u"  • {} （重複 {} 次）".format(n, num_count[n])
                                 for n in dups))
        return None

    def _find_new_number_conflicts(self, plan):
        """找出『新增圖紙』的圖號與總表（現有 Revit 圖紙）撞號者。

        圖號在 Revit 必須全域唯一。分表匯入只比對同類別，無法察覺新增圖紙的圖號
        其實已被『其他類別』的圖紙占用 → 套用時才會失敗。這裡先抓出來。
        會排除『這次匯入會釋放掉舊圖號』的圖紙（被刪除/降為僅文字、或被改成別的圖號）。
        回傳 [(新增圖號, 新增圖名, 既有衝突圖紙)]。
        """
        freeing = set()  # 這次匯入後會空出圖號的既有圖紙 UID
        for key in (u"delete", u"toPlace"):
            for d in plan.get(key, []):
                s = d.get("sheet")
                if s is not None:
                    freeing.add(s.UniqueId)
        for key in (u"edit", u"toReal"):
            for d in plan.get(key, []):
                s = d.get("sheet")
                newnum = unicode(d.get("num") or u"").strip()
                if s is not None and newnum and \
                        newnum != (s.SheetNumber or u"").strip():
                    freeing.add(s.UniqueId)

        exist_by_num = {}
        for s in DB.FilteredElementCollector(doc).OfClass(
                DB.ViewSheet).ToElements():
            if s.UniqueId in freeing:
                continue
            num = (s.SheetNumber or u"").strip()
            if num and num not in exist_by_num:
                exist_by_num[num] = s

        conflicts = []
        for d in plan.get("new", []):
            num = unicode(d.get("num") or u"").strip()
            if not num:
                continue
            s = exist_by_num.get(num)
            if s is not None:
                conflicts.append((num, unicode(d.get("name") or u""), s))
        return conflicts

    def _show_new_number_conflicts(self, conflicts):
        """列出『新增圖紙圖號與總表重複』的衝突資訊（唯讀），並中止匯入。"""
        rows = []
        for num, new_name, s in conflicts:
            cat = self._read_text_param(s, u"圖紙類別") or u""
            rows.append([num, new_name or u"（未命名）",
                         s.SheetNumber or u"", s.Name or u"", cat])
        dlg = _ListConfirmWindow(
            u"新增圖紙的圖號與總表重複",
            u"以下 {} 筆『新增圖紙』的圖號，在總表（現有 Revit 圖紙）中已被占用。\n"
            u"圖號在 Revit 必須唯一，已中止匯入。請到 Google Sheet 分表把這些新增列的"
            u"圖號改成唯一後再匯入；若該列其實是既有圖紙，請改回它原本的 UID。".format(
                len(rows)),
            [u"新增圖號", u"新增圖名", u"總表既有圖號", u"總表既有圖名", u"既有圖紙類別"],
            rows, confirm_label=u"我知道了", cancel_label=u"關閉")
        try:
            dlg.Owner = self
        except Exception:
            pass
        dlg.ShowDialog()
        self._show_sync_msg(
            u"⚠ 新增圖紙圖號與總表重複，已中止匯入（請改圖號後重試）。",
            self.COLOR_ERROR)

    def _on_import_execute(self, sender, args):
        plan = getattr(self, "_pending_plan", None)
        if not plan:
            forms.alert(u"請先按「匯入（預覽差異）」載入要套用的內容。")
            return
        # 刪除/重建前：用可捲動視窗列出「全部」將被刪除或重建的圖紙
        del_rows = []
        for d in plan["toPlace"]:
            s = d["sheet"]
            del_rows.append([u"轉為僅文字資訊（原圖刪除重建）",
                             s.SheetNumber or u"", s.Name or u""])
        for d in plan["delete"]:
            s = d["sheet"]
            del_rows.append([u"刪除", s.SheetNumber or u"", s.Name or u""])
        if del_rows:
            dlg = _ListConfirmWindow(
                u"確認刪除 / 重建圖紙",
                u"以下 {} 張圖紙將被刪除或重建，確定執行嗎？".format(len(del_rows)),
                [u"動作", u"圖號", u"圖名"], del_rows,
                confirm_label=u"確定執行")
            dlg.ShowDialog()
            if not dlg.confirmed:
                return
        # 不在 modal 視窗內改模型(會被還原)；關窗後由 main() 執行交易
        self.result = {"mode": "import_apply", "plan": plan}
        self.confirmed = True
        self.Close()

    def _on_assign_tasks(self, sender, args):
        forms.alert(u"「指派工作任務」功能開發中。\n\n"
                    u"未來按下後，會把『修正備註』欄裡的工作，\n"
                    u"分配給對應的『繪圖員』同事。\n（分配方式待討論）")

    def _existing_sheet_param_names(self, wanted):
        """回傳 wanted 之中『圖紙已具備』的參數名稱集合。

        用實際圖紙的 LookupParameter 判斷 → 內建/專案/共用參數都算數。
        （內建參數如 繪圖員/審圖員/設計者/批准者/圖紙發布日期 不會出現在
          doc.ParameterBindings，所以不能只看綁定。）
        空專案無圖紙時，退而用 ParameterBindings（只看得到專案/共用參數）。
        """
        found = set()
        sheets = list(
            DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements())
        if sheets:
            # 優先用真實圖紙；沒有就用第一張（預留圖紙也有識別資料參數）
            sample = next((s for s in sheets if not s.IsPlaceholder), sheets[0])
            for name in wanted:
                try:
                    if _lookup_sheet_param(sample, name) is not None:
                        found.add(name)
                except Exception:
                    pass
            return found

        # 後備（空專案）：用 ParameterBindings 找已綁到圖紙類別的專案/共用參數
        try:
            sheets_cat = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Sheets)
        except Exception:
            sheets_cat = None
        try:
            it = doc.ParameterBindings.ForwardIterator()
            it.Reset()
            while it.MoveNext():
                definition = it.Key
                binding = it.Current
                try:
                    cats = binding.Categories
                except Exception:
                    cats = None
                if cats is None:
                    continue
                for c in cats:
                    if sheets_cat is not None and \
                            c.Id.IntegerValue == sheets_cat.Id.IntegerValue:
                        if definition.Name in wanted:
                            found.add(definition.Name)
                        break
        except Exception:
            pass
        return found

    def _on_build_env(self, sender, args):
        needed = list(self.SYNC_TEXT_PARAMS) + list(self.BATCH_PARAMS)
        existing = self._existing_sheet_param_names(needed)
        missing = [n for n in needed if n not in existing]
        have = [n for n in needed if n in existing]

        if not missing:
            # 參數都在了，仍可能要補建「圖紙明細表」→ 走 build_env(只建明細表)
            if not forms.alert(
                    u"同步所需的 {} 個圖紙文字參數都已存在。\n\n{}\n\n"
                    u"要（重新）確認並建立「BaF_Sheet專用圖紙明細表」嗎？".format(
                        len(needed), u"、".join(needed)),
                    yes=True, no=True):
                return
            self.result = {"mode": "build_env", "missing": [],
                           "have": have, "all_params": needed}
            self.confirmed = True
            self.Close()
            return

        msg = u"將在「圖紙(Sheets)」類別建立下列文字參數：\n\n"
        msg += u"\n".join(u"  ＋ " + n for n in missing)
        if have:
            msg += u"\n\n已存在，將略過：\n" + u"\n".join(u"  ✓ " + n for n in have)
        msg += (u"\n\n說明：公司共用參數檔已有的會直接沿用共用參數；其餘建為專案參數，"
                u"並建立「BaF_Sheet專用圖紙明細表」。不會更動 Google Sheet。\n\n"
                u"確定建立嗎？")
        if not forms.alert(msg, yes=True, no=True):
            return

        # 不在 modal 視窗內改文件（與匯入一致）；關窗後由 main() 在交易內建立
        self.result = {"mode": "build_env", "missing": missing,
                       "have": have, "all_params": needed}
        self.confirmed = True
        self.Close()

    # ---- 依紅筆圖註解：選後端 → 選 PDF → AI 判讀 → 預覽 → 寫入修正備註 ----

    def _on_edit_by_redpen(self, sender, args):
        # 1. 跳出「AI 視覺判讀後端」選擇視窗（沒填金鑰會在視窗內提醒並停留）
        start = _RedpenStartWindow()
        start.ShowDialog()
        if not start.confirmed:
            return
        provider = start.provider

        cfg = load_gsheet_cfg()
        keys = cfg.get("ai_keys") or {}
        models = cfg.get("ai_models") or {}
        key = (keys.get(provider) or u"").strip()
        if not key:
            # 理論上選擇視窗已擋掉；保險再檢一次
            forms.alert(u"尚未設定「{}」的 API 金鑰。".format(
                _ai_provider_label(provider)))
            return
        model = _ai_model_for(provider, models)

        # 2. 選 PDF（單一檔）
        path = forms.pick_file(file_ext='pdf')
        if not path:
            return
        if isinstance(path, list):
            if not path:
                return
            path = path[0]

        # 3. 讀檔 → base64
        try:
            b64, nbytes = _file_to_base64(path)
        except Exception as ex:
            forms.alert(u"讀取 PDF 失敗：{}".format(ex))
            return
        mb = nbytes / (1024.0 * 1024.0)
        if mb > 20:
            if not forms.alert(
                    u"這個 PDF 約 {:.1f} MB，可能超過 AI 服務的單次上限而失敗。\n"
                    u"建議分頁或壓縮後再試。仍要繼續嗎？".format(mb),
                    yes=True, no=True):
                return

        # 4. 開「處理進度」小視窗，逐步說明呼叫 AI 判讀過程中發生了什麼
        prog = _RedpenProgressWindow(_ai_provider_label(provider))
        try:
            prog.Owner = self
        except Exception:
            pass
        prog.Show()
        prog.log(u"① 已讀取 PDF：{}（約 {:.1f} MB）".format(
            os.path.basename(unicode(path)), mb))
        prog.log(u"② 正在把整份 PDF 上傳給「{}」並請它判讀紅筆標示…".format(
            _ai_provider_label(provider)))
        prog.log(u"　（這一步需要數十秒，期間視窗會像沒有反應，屬正常，請勿關閉）")

        text, err = _redpen_call_ai(
            provider, model, key, b64, _redpen_build_prompt())
        if err:
            err = unicode(err)
            prog.log(u"✗ 呼叫失敗。")
            prog.close()
            # 把「完整」錯誤寫到暫存檔，方便整段複製(forms.alert 會截斷/難複製)
            logp = u""
            try:
                import tempfile
                import io
                logp = os.path.join(tempfile.gettempdir(), u"baf_redpen_error.txt")
                with io.open(logp, "w", encoding="utf-8") as f:
                    f.write(u"provider={}\nmodel={}\n\n{}".format(
                        provider, model, err))
            except Exception:
                logp = u""
            self._show_sync_msg(
                u"❌ AI 呼叫失敗：{}".format(err[:200]), self.COLOR_ERROR)
            forms.alert(u"AI（{}，模型 {}）呼叫失敗。\n原因／錯誤代碼：\n\n{}{}".format(
                _ai_provider_label(provider), model, err[:1500],
                (u"\n\n（完整錯誤已寫到：{}）".format(logp) if logp else u"")))
            return

        prog.log(u"③ AI 已回覆，正在解析判讀結果（JSON）…")
        data = _redpen_extract_json(text)
        if not data or not isinstance(data, dict):
            prog.close()
            self._show_sync_msg(u"❌ AI 回應無法解析為 JSON。", self.COLOR_ERROR)
            forms.alert(u"AI 回應無法解析為 JSON。\n原始回應前 800 字：\n\n{}".format(
                (text or u"")[:800]))
            return
        ai_sheets = data.get("sheets") or []
        if not ai_sheets:
            prog.log(u"③ 解析完成：AI 沒有判讀到任何紅筆修改標示。")
            prog.close()
            self._show_sync_msg(
                u"AI 在此 PDF 沒有判讀到任何紅筆修改標示。", self.COLOR_TEXT)
            forms.alert(u"AI 沒有在這份 PDF 找到紅筆修改標示。")
            return
        prog.log(u"③ 解析完成：AI 判讀到 {} 筆修改。".format(len(ai_sheets)))

        # 5. 用圖框上的「圖號」對應 Revit 圖紙（掃描檔沒有 UID）
        #    Revit 圖號唯一 → 一個圖號對一張圖。PDF 若有「同圖號、不同圖名」的多頁，
        #    會全部對到同一張 Revit 圖紙 → 先把這些頁的備註「合併」成一筆(用「；」)，
        #    避免互相覆蓋而遺失。圖名不參與比對，僅供顯示。
        prog.log(u"④ 正在用圖號比對目前 Revit 的圖紙…")
        by_num = {}
        for s in DB.FilteredElementCollector(doc).OfClass(
                DB.ViewSheet).ToElements():
            by_num[(s.SheetNumber or u"").strip()] = s
        acc = {}        # uid -> {"sheet", "num", "notes": [..]}
        acc_order = []  # 保持出現順序
        unmatched = []  # (num, note)
        for item in ai_sheets:
            try:
                num = unicode(item.get("sheet_number") or u"").strip()
                note = unicode(item.get("notes") or u"").strip()
            except Exception:
                continue
            if not note:
                continue
            s = by_num.get(num)
            if s is None:
                unmatched.append((num, note))
                continue
            uid = s.UniqueId
            if uid not in acc:
                acc[uid] = {"sheet": s, "num": num, "notes": []}
                acc_order.append(uid)
            acc[uid]["notes"].append(note)
        matched = []   # (sheet, num, name, old_note, new_note)
        for uid in acc_order:
            a = acc[uid]
            s = a["sheet"]
            old = self._read_text_param(s, u"修正備註") or u""
            combined = u"；".join(a["notes"])   # 同圖號多頁的備註合併
            matched.append((s, a["num"], s.Name or u"", old, combined))

        if not matched:
            prog.log(u"④ 比對完成：{} 個圖號都對不到 Revit 圖紙。".format(len(unmatched)))
            prog.close()
            msg = u"AI 判讀到 {} 筆，但圖號都對不到目前 Revit 的圖紙。".format(
                len(ai_sheets))
            if unmatched:
                msg += u"\n\n對不到的圖號：\n" + u"\n".join(
                    u"・{}".format(n) for n, _ in unmatched[:20])
            self._show_sync_msg(u"⚠ 圖號對不到 Revit 圖紙。", self.COLOR_ERROR)
            forms.alert(msg)
            return

        prog.log(u"⑤ 比對完成：對到 {} 張、對不到 {} 個圖號。即將開啟預覽…".format(
            len(matched), len(unmatched)))
        prog.close()

        # 6. 預覽（逐張選 附加/覆蓋）
        url = (self.gs_url_box.Text or u"").strip()
        tab = (self.gs_tab_box.Text or u"").strip()
        dlg = _RedpenPreviewWindow(matched, unmatched, can_sync=bool(url and tab))
        try:
            dlg.Owner = self
        except Exception:
            pass
        dlg.ShowDialog()
        if not dlg.confirmed:
            return
        updates = dlg.get_updates()   # [(uid, final_note), ...]
        if not updates:
            forms.alert(u"沒有勾選任何要寫入的圖紙。")
            return

        # 關窗後由 main() 在交易內套用（與匯入一致，避免被還原）
        self.result = {"mode": "redpen", "updates": updates,
                       "url": url, "tab": tab,
                       "push_to_sheet": bool(dlg.push_to_sheet and url and tab)}
        self.confirmed = True
        self.Close()

    def _docs_path(self, filename):
        """回傳擴充功能 docs 資料夾下的檔案完整路徑。"""
        here = os.path.dirname(__file__)
        return os.path.normpath(
            os.path.join(here, u"..", u"..", u"..", u"docs", filename))

    def _open_doc(self, filename, what):
        p = self._docs_path(filename)
        if not os.path.exists(p):
            forms.alert(u"找不到{}：\n{}".format(what, p))
            return
        try:
            os.startfile(p)
        except Exception as ex:
            forms.alert(u"開啟{}失敗：{}\n\n你可以手動到這個位置打開：\n{}".format(what, ex, p))

    def _on_open_manual(self, sender, args):
        self._open_doc(u"使用說明_從零設定.txt", u"操作說明")

    def _on_open_script(self, sender, args):
        self._open_doc(u"gsheet_sync_appscript.gs", u"最新腳本")

    def _apply_plan(self, plan):
        done = {"new": 0, "edit": 0, "toReal": 0, "toPlace": 0, "delete": 0}
        fails = []
        tb_id = plan.get("tb_id") or DB.ElementId.InvalidElementId
        try:
            with revit.Transaction("BaF 匯入套用"):
                targets = []  # (sheet, data, is_new)

                # 刪除：Sheet 整列(含 id)被移除的圖紙
                for d in plan.get("delete", []):
                    try:
                        doc.Delete(d["sheet"].Id)
                        done["delete"] += 1
                    except Exception as ex:
                        fails.append((d["sheet"].SheetNumber, u"刪除:" + unicode(ex)))

                for d in plan["toReal"]:
                    try:
                        d["sheet"].ConvertToRealSheet(tb_id)
                        targets.append((d["sheet"], d, False))
                        done["toReal"] += 1
                    except Exception as ex:
                        fails.append((d.get("num"), u"轉真實圖紙:" + unicode(ex)))

                for d in plan["edit"]:
                    targets.append((d["sheet"], d, False))
                    done["edit"] += 1

                for d in plan["new"]:
                    try:
                        if d["real"]:
                            ns = DB.ViewSheet.Create(doc, tb_id)
                        else:
                            ns = DB.ViewSheet.CreatePlaceholder(doc)
                        targets.append((ns, d, True))
                        done["new"] += 1
                    except Exception as ex:
                        fails.append((d.get("num"), u"新增:" + unicode(ex)))

                for d in plan["toPlace"]:
                    try:
                        doc.Delete(d["sheet"].Id)
                        ns = DB.ViewSheet.CreatePlaceholder(doc)
                        targets.append((ns, d, True))
                        done["toPlace"] += 1
                    except Exception as ex:
                        fails.append((d.get("num"), u"降為僅文字資訊:" + unicode(ex)))

                # 只處理「有指定圖號」的；沒給圖號就保留 Revit 預設(不再指派未命名)
                renum = [t for t in targets if t[1]["num"]]

                # 兩階段：先全設暫時號避免互撞
                tmp_i = 0
                for sheet, d, isnew in renum:
                    try:
                        sheet.SheetNumber = u"__BAFTMP_{}__".format(tmp_i)
                        tmp_i += 1
                    except Exception:
                        pass
                occupied = set(
                    s.SheetNumber for s in
                    DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements()
                    if not (s.SheetNumber or u"").startswith(u"__BAFTMP_"))

                # 設最終圖號（重複已在套用前擋掉，不自動加尾碼）
                for sheet, d, isnew in renum:
                    try:
                        sheet.SheetNumber = d["num"]
                        occupied.add(d["num"])
                    except Exception as ex:
                        fails.append((d.get("num"), u"設定圖號:" + unicode(ex)))

                # 設圖名與參數（空白=不覆蓋，不再自動命名）
                for sheet, d, isnew in targets:
                    try:
                        if d["name"]:
                            sheet.Name = d["name"]
                        if d["cat"]:
                            self._set_param(sheet, u"圖紙類別", d["cat"])
                        if d["drawer"]:
                            self._set_param(sheet, u"繪圖員", d["drawer"])
                        # 修正備註以 Google Sheet 為準（空白＝清空，移除 Revit 的修正備註）
                        self._set_param(sheet, u"修正備註", d["note"])
                        for k, v in d.get("extra", {}).items():
                            if v:
                                self._set_param(sheet, k, v)
                    except Exception as ex:
                        fails.append((d.get("num"), u"套用欄位:" + unicode(ex)))
        except Exception as ex:
            return u"❌ 套用失敗，已自動復原：{}".format(ex), True, done

        report = (u"🆕 新增 {} ／ ✏️ 修改 {} ／ ⬆️ 轉真實圖紙 {} ／ "
                  u"⚠️ 降為僅文字資訊 {} ／ 🗑️ 刪除 {}".format(
                      done["new"], done["edit"], done["toReal"],
                      done["toPlace"], done["delete"]))
        if fails:
            report += u"\n\n⚠ {} 筆有狀況：\n".format(len(fails)) + u"\n".join(
                u"{}: {}".format(a, b) for a, b in fails[:15])
        return report, False, done

    # ---- 執行 ----
    
    def _on_run(self, sender, args):
        idx = self.tabs.SelectedIndex
        if idx == self.MODE_SYNC:
            return  # 同步分頁不走這個按鈕
        if idx == self.MODE_EDIT_EXISTING:
            self._run_edit_existing()
        elif idx == self.MODE_TABLE:
            self._run_create(self._collect_from_table())
        elif idx == self.MODE_PASTE:
            self._run_create(self._collect_from_paste())
        elif idx == self.MODE_PATTERN:
            try:
                data = self._collect_from_pattern()
            except Exception as ex:
                forms.alert("輸入有誤：{}".format(ex))
                return
            self._run_create(data)
    
    def _run_edit_existing(self):
        edits = self._collect_edits()
        if not edits:
            forms.alert("沒有可套用的變更。請勾選圖紙並填寫新值。")
            return
        
        # 驗證新圖號的合法性
        new_nums_in_batch = set()
        errors = []
        # 將被改掉的舊圖號（這些圖號將不再佔用）
        nums_being_changed = set(s.SheetNumber for s, n, _ in edits if n)
        
        for sheet, new_num, new_name in edits:
            if new_num is None:
                continue
            if not new_num:
                errors.append("'{}' 的新圖號不能是空白".format(sheet.SheetNumber))
                continue
            if new_num in new_nums_in_batch:
                errors.append("新圖號 '{}' 在這次變更中重複".format(new_num))
                continue
            # 跟其他既有圖號衝突（但被改掉的舊圖號不算）
            if new_num in self.existing_numbers and new_num not in nums_being_changed:
                errors.append("新圖號 '{}' 已被其他圖紙使用".format(new_num))
                continue
            new_nums_in_batch.add(new_num)
        
        if errors:
            msg = "發現 {} 個問題：\n\n".format(len(errors))
            msg += "\n".join(errors[:10])
            if len(errors) > 10:
                msg += "\n... 還有 {} 個".format(len(errors) - 10)
            forms.alert(msg)
            return
        
        self.result = {"mode": "edit", "edits": edits}
        self.confirmed = True
        self.Close()
    
    def _run_create(self, data):
        if not data:
            forms.alert("沒有可建立的圖紙資料。")
            return
        
        errors = []
        seen = set()
        clean = []
        for i, (num, name) in enumerate(data):
            line_no = i + 1
            if not num:
                errors.append("第 {} 列：圖號為空".format(line_no))
                continue
            if num in self.existing_numbers:
                errors.append("第 {} 列：圖號 '{}' 已存在".format(line_no, num))
                continue
            if num in seen:
                errors.append("第 {} 列：圖號 '{}' 重複".format(line_no, num))
                continue
            seen.add(num)
            clean.append((num, name or "Unnamed"))
        
        if errors:
            msg = "發現 {} 個問題：\n\n".format(len(errors))
            msg += "\n".join(errors[:10])
            if len(errors) > 10:
                msg += "\n... 還有 {} 個".format(len(errors) - 10)
            if not clean:
                forms.alert(msg)
                return
            msg += "\n\n要繼續建立沒問題的 {} 張圖紙嗎？".format(len(clean))
            ok = forms.alert(msg, yes=True, no=True)
            if not ok:
                return
        
        self.result = {
            "mode": "create",
            "sheets": clean,
            "title_block_id": self._get_selected_title_block_id(),
        }
        self.confirmed = True
        self.Close()


# ---------------------------------------------------------------------------
# 輔助對話框：可捲動清單確認 / 重複圖號就地修改
# ---------------------------------------------------------------------------

def _brush(rgb):
    return SolidColorBrush(Color.FromRgb(rgb[0], rgb[1], rgb[2]))


class _ListConfirmWindow(Window):
    """可捲動的清單確認視窗：列出『全部』項目（右側有捲軸），按確定/取消。"""

    def __init__(self, title, intro, headers, rows,
                 confirm_label=u"確定執行", cancel_label=u"取消"):
        self.confirmed = False
        self.Title = title
        self.Width = 640
        self.Height = 560
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = _brush((245, 245, 250))

        root = Grid()
        root.Margin = Thickness(16, 14, 16, 14)
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))

        head = TextBlock()
        head.Text = intro
        head.TextWrapping = TextWrapping.Wrap
        head.FontSize = 13
        head.FontWeight = FontWeights.Bold
        head.Margin = Thickness(0, 0, 0, 10)
        head.Foreground = _brush((180, 60, 60))
        Grid.SetRow(head, 0)
        root.Children.Add(head)

        sv = ScrollViewer()
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Visible
        sv.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.Background = _brush((255, 255, 255))
        sv.BorderBrush = _brush((200, 205, 215))
        sv.BorderThickness = Thickness(1)
        sv.Padding = Thickness(8, 8, 8, 8)
        sv.Content = self._build_table(headers, rows)
        Grid.SetRow(sv, 1)
        root.Children.Add(sv)

        btns = StackPanel()
        btns.Orientation = Orientation.Horizontal
        btns.HorizontalAlignment = HorizontalAlignment.Right
        btns.Margin = Thickness(0, 12, 0, 0)

        ok = Button()
        ok.Content = u"{}（共 {} 項）".format(confirm_label, len(rows))
        ok.Padding = Thickness(14, 6, 14, 6)
        ok.Margin = Thickness(0, 0, 8, 0)
        ok.Background = _brush((220, 53, 69))
        ok.Foreground = _brush((255, 255, 255))
        ok.FontWeight = FontWeights.Bold
        ok.Click += self._on_ok
        btns.Children.Add(ok)

        cancel = Button()
        cancel.Content = cancel_label
        cancel.Padding = Thickness(14, 6, 14, 6)
        cancel.Click += self._on_cancel
        btns.Children.Add(cancel)

        Grid.SetRow(btns, 2)
        root.Children.Add(btns)
        self.Content = root

    def _build_table(self, headers, rows):
        grid = Grid()
        for _ in headers:
            grid.ColumnDefinitions.Add(
                ColumnDefinition(Width=GridLength(1, GridUnitType.Auto)))
        for _ in range(len(rows) + 1):
            grid.RowDefinitions.Add(
                RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        for c, h in enumerate(headers):
            tb = TextBlock()
            tb.Text = h
            tb.FontWeight = FontWeights.Bold
            tb.Margin = Thickness(6, 2, 14, 6)
            Grid.SetRow(tb, 0)
            Grid.SetColumn(tb, c)
            grid.Children.Add(tb)
        for r, rowdata in enumerate(rows):
            for c, val in enumerate(rowdata):
                tb = TextBlock()
                tb.Text = val if val is not None else u""
                tb.Margin = Thickness(6, 2, 14, 2)
                Grid.SetRow(tb, r + 1)
                Grid.SetColumn(tb, c)
                grid.Children.Add(tb)
        return grid

    def _on_ok(self, sender, args):
        self.confirmed = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.Close()


class _DupResolveWindow(Window):
    """重複圖號就地修改：列出每組重複圖號，使用者改到不重複後按「確定」。

    確定後直接更新傳入的 records（rec[u'圖紙號碼']）。
    """

    def __init__(self, records):
        self.confirmed = False
        self.records = records
        self._edit_boxes = []  # (textbox, rec)

        num_to_recs = {}
        order = []
        for rec in records:
            num = unicode(rec.get(u"圖紙號碼") or u"").strip()
            if not num:
                continue
            if num not in num_to_recs:
                num_to_recs[num] = []
                order.append(num)
            num_to_recs[num].append(rec)
        groups = [(num, num_to_recs[num]) for num in order
                  if len(num_to_recs[num]) > 1]
        self._build_ui(groups)

    def _build_ui(self, groups):
        self.Title = u"解決重複圖號"
        self.Width = 680
        self.Height = 600
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = _brush((245, 245, 250))

        root = Grid()
        root.Margin = Thickness(16, 14, 16, 14)
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))

        intro = TextBlock()
        intro.Text = (u"Google Sheet 有重複的圖號，無法匯入。\n"
                      u"請把每組重複的圖號改成不重複（直接修改下方圖號欄），再按「確定」。")
        intro.TextWrapping = TextWrapping.Wrap
        intro.FontSize = 13
        intro.FontWeight = FontWeights.Bold
        intro.Foreground = _brush((180, 90, 0))
        intro.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(intro, 0)
        root.Children.Add(intro)

        sv = ScrollViewer()
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Visible
        sv.Background = _brush((255, 255, 255))
        sv.BorderBrush = _brush((200, 205, 215))
        sv.BorderThickness = Thickness(1)
        sv.Padding = Thickness(8, 8, 8, 8)
        panel = StackPanel()
        panel.Orientation = Orientation.Vertical
        for num, recs in groups:
            gh = TextBlock()
            gh.Text = u"重複圖號：{}（{} 張）".format(num, len(recs))
            gh.FontWeight = FontWeights.Bold
            gh.Margin = Thickness(2, 8, 2, 4)
            gh.Foreground = _brush((30, 30, 40))
            panel.Children.Add(gh)
            for rec in recs:
                panel.Children.Add(self._make_row(rec))
        sv.Content = panel
        Grid.SetRow(sv, 1)
        root.Children.Add(sv)

        self.err_text = TextBlock()
        self.err_text.Foreground = _brush((220, 53, 69))
        self.err_text.TextWrapping = TextWrapping.Wrap
        self.err_text.Margin = Thickness(2, 8, 2, 0)
        Grid.SetRow(self.err_text, 2)
        root.Children.Add(self.err_text)

        btns = StackPanel()
        btns.Orientation = Orientation.Horizontal
        btns.HorizontalAlignment = HorizontalAlignment.Right
        btns.Margin = Thickness(0, 10, 0, 0)
        ok = Button()
        ok.Content = u"確定"
        ok.Padding = Thickness(16, 6, 16, 6)
        ok.Margin = Thickness(0, 0, 8, 0)
        ok.Background = _brush((99, 102, 241))
        ok.Foreground = _brush((255, 255, 255))
        ok.FontWeight = FontWeights.Bold
        ok.Click += self._on_ok
        btns.Children.Add(ok)
        cancel = Button()
        cancel.Content = u"取消匯入"
        cancel.Padding = Thickness(16, 6, 16, 6)
        cancel.Click += self._on_cancel
        btns.Children.Add(cancel)
        Grid.SetRow(btns, 3)
        root.Children.Add(btns)
        self.Content = root

    def _make_row(self, rec):
        row = Grid()
        row.Margin = Thickness(0, 1, 0, 1)
        row.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(150)))
        row.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        row.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(90)))

        tb = TextBox()
        tb.Text = unicode(rec.get(u"圖紙號碼") or u"")
        tb.Padding = Thickness(4, 3, 4, 3)
        tb.Margin = Thickness(0, 0, 6, 0)
        Grid.SetColumn(tb, 0)
        row.Children.Add(tb)
        self._edit_boxes.append((tb, rec))

        nm = TextBlock()
        nm.Text = unicode(rec.get(u"圖紙名稱") or u"")
        nm.VerticalAlignment = VerticalAlignment.Center
        nm.TextTrimming = TextTrimming.CharacterEllipsis
        nm.Foreground = _brush((90, 95, 110))
        Grid.SetColumn(nm, 1)
        row.Children.Add(nm)

        ut = TextBlock()
        ut.Text = unicode(rec.get(u"UID") or u"")[:8]
        ut.VerticalAlignment = VerticalAlignment.Center
        ut.Foreground = _brush((150, 155, 165))
        ut.FontSize = 11
        Grid.SetColumn(ut, 2)
        row.Children.Add(ut)
        return row

    def _on_ok(self, sender, args):
        proposed = {}
        for tb, rec in self._edit_boxes:
            new_num = (tb.Text or u"").strip()
            if not new_num:
                self.err_text.Text = u"圖號不可空白，請填寫。"
                return
            proposed[id(rec)] = new_num
        # 全部圖號（編輯過用新值、其餘用原值）做唯一性檢查
        counts = {}
        for rec in self.records:
            if id(rec) in proposed:
                num = proposed[id(rec)]
            else:
                num = unicode(rec.get(u"圖紙號碼") or u"").strip()
            if not num:
                continue
            counts[num] = counts.get(num, 0) + 1
        still_dup = sorted([n for n, c in counts.items() if c > 1])
        if still_dup:
            self.err_text.Text = u"仍有重複：{}　請再修改。".format(u"、".join(still_dup))
            return
        for tb, rec in self._edit_boxes:
            rec[u"圖紙號碼"] = proposed[id(rec)]
        self.confirmed = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.Close()


class _ImportCleanupWindow(Window):
    """匯入前資料檢查：擋下並條列「有 UID 但無值」與「重複 UID」的列。

    勾選 = 這列不要匯入（移除）。每個 UID 最多只能保留 1 列。
    某 UID 若一列都不留 → 對應的 Revit 圖紙會被刪除（之後刪除確認視窗會再列一次）。
    確定後把保留的列放在 self.cleaned_records。
    """

    @staticmethod
    def _uid(rec):
        return unicode(rec.get(u"UID") or u"").strip()

    @staticmethod
    def _is_blank(rec):
        num = unicode(rec.get(u"圖紙號碼") or u"").strip()
        name = unicode(rec.get(u"圖紙名稱") or u"").strip()
        return (not num) and (not name)

    @classmethod
    def has_issues(cls, records):
        seen = {}
        for r in records:
            u = cls._uid(r)
            if u:
                seen[u] = seen.get(u, 0) + 1
        for r in records:
            u = cls._uid(r)
            if u and cls._is_blank(r):
                return True            # 有 UID 但無值
            if u and seen[u] > 1:
                return True            # 重複 UID
        return False

    def __init__(self, records, uid_to_sheet):
        self.confirmed = False
        self.records = records
        self.cleaned_records = list(records)
        self._uid_to_sheet = uid_to_sheet or {}
        self._row_checks = []  # (checkbox, rec)
        self._build_ui()

    def _ctx(self, rec):
        """該列 UID 對應現有 Revit 圖紙的圖號/圖名（沒有則顯示提示）。"""
        info = self._uid_to_sheet.get(self._uid(rec))
        if info:
            return u"{}  {}".format(info[0], info[1])
        return u"（無對應的現有圖紙）"

    def _build_ui(self):
        self.Title = u"匯入資料檢查（UID 異常）"
        self.Width = 720
        self.Height = 620
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = _brush((245, 245, 250))

        root = Grid()
        root.Margin = Thickness(16, 14, 16, 14)
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))

        intro = TextBlock()
        intro.Text = (u"Google Sheet 有異常的列，匯入前請先處理。\n"
                      u"勾選 = 這列不要匯入（移除）。每個 UID 只能保留 1 列；\n"
                      u"某 UID 若一列都不保留，對應的 Revit 圖紙會被刪除（之後會再確認一次）。")
        intro.TextWrapping = TextWrapping.Wrap
        intro.FontSize = 13
        intro.FontWeight = FontWeights.Bold
        intro.Foreground = _brush((180, 90, 0))
        intro.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(intro, 0)
        root.Children.Add(intro)

        sv = ScrollViewer()
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Visible
        sv.Background = _brush((255, 255, 255))
        sv.BorderBrush = _brush((200, 205, 215))
        sv.BorderThickness = Thickness(1)
        sv.Padding = Thickness(8, 8, 8, 8)
        panel = StackPanel()
        panel.Orientation = Orientation.Vertical

        # 分類：重複 UID 群組 / 空白列(UID 唯一)
        uid_count = {}
        order = []
        for r in self.records:
            u = self._uid(r)
            if u:
                if u not in uid_count:
                    uid_count[u] = 0
                    order.append(u)
                uid_count[u] += 1

        # 1) 重複 UID
        dup_uids = [u for u in order if uid_count[u] > 1]
        if dup_uids:
            panel.Children.Add(self._section_header(
                u"重複 UID（同一 UID 多列；每個 UID 只能保留 1 列）"))
            for u in dup_uids:
                group = [r for r in self.records if self._uid(r) == u]
                gh = TextBlock()
                gh.Text = u"UID {}…（{} 列）　對應：{}".format(
                    u[:8], len(group), self._ctx(group[0]))
                gh.FontWeight = FontWeights.Bold
                gh.Margin = Thickness(2, 8, 2, 4)
                gh.Foreground = _brush((30, 30, 40))
                panel.Children.Add(gh)
                # 預設保留第一筆非空白列，其餘勾選移除
                keep_idx = next((i for i, r in enumerate(group)
                                 if not self._is_blank(r)), 0)
                for i, r in enumerate(group):
                    panel.Children.Add(self._make_row(r, default_remove=(i != keep_idx)))

        # 2) 空白列（UID 唯一、圖號圖名皆空）
        blank_unique = [r for r in self.records
                        if self._uid(r) and self._is_blank(r) and uid_count[self._uid(r)] == 1]
        if blank_unique:
            panel.Children.Add(self._section_header(
                u"空白列（有 UID、無圖號圖名）；勾選＝刪除對應圖紙"))
            for r in blank_unique:
                panel.Children.Add(self._make_row(r, default_remove=False))

        sv.Content = panel
        Grid.SetRow(sv, 1)
        root.Children.Add(sv)

        self.err_text = TextBlock()
        self.err_text.Foreground = _brush((220, 53, 69))
        self.err_text.TextWrapping = TextWrapping.Wrap
        self.err_text.Margin = Thickness(2, 8, 2, 0)
        Grid.SetRow(self.err_text, 2)
        root.Children.Add(self.err_text)

        btns = StackPanel()
        btns.Orientation = Orientation.Horizontal
        btns.HorizontalAlignment = HorizontalAlignment.Right
        btns.Margin = Thickness(0, 10, 0, 0)
        ok = Button()
        ok.Content = u"確定"
        ok.Padding = Thickness(16, 6, 16, 6)
        ok.Margin = Thickness(0, 0, 8, 0)
        ok.Background = _brush((99, 102, 241))
        ok.Foreground = _brush((255, 255, 255))
        ok.FontWeight = FontWeights.Bold
        ok.Click += self._on_ok
        btns.Children.Add(ok)
        cancel = Button()
        cancel.Content = u"取消匯入"
        cancel.Padding = Thickness(16, 6, 16, 6)
        cancel.Click += self._on_cancel
        btns.Children.Add(cancel)
        Grid.SetRow(btns, 3)
        root.Children.Add(btns)
        self.Content = root

    def _section_header(self, text):
        tb = TextBlock()
        tb.Text = text
        tb.FontWeight = FontWeights.Bold
        tb.FontSize = 13
        tb.Margin = Thickness(0, 6, 0, 2)
        tb.Foreground = _brush((99, 102, 241))
        return tb

    def _make_row(self, rec, default_remove):
        row = Grid()
        row.Margin = Thickness(0, 1, 0, 1)
        row.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(50)))
        row.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(140)))
        row.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        row.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(90)))

        cb = CheckBox()
        cb.IsChecked = bool(default_remove)
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.HorizontalAlignment = HorizontalAlignment.Center
        Grid.SetColumn(cb, 0)
        row.Children.Add(cb)
        self._row_checks.append((cb, rec))

        num = TextBlock()
        num.Text = unicode(rec.get(u"圖紙號碼") or u"") or u"（空）"
        num.VerticalAlignment = VerticalAlignment.Center
        num.Foreground = _brush((30, 30, 40))
        Grid.SetColumn(num, 1)
        row.Children.Add(num)

        nm = TextBlock()
        nm.Text = unicode(rec.get(u"圖紙名稱") or u"") or u"（空）"
        nm.VerticalAlignment = VerticalAlignment.Center
        nm.TextTrimming = TextTrimming.CharacterEllipsis
        nm.Foreground = _brush((90, 95, 110))
        Grid.SetColumn(nm, 2)
        row.Children.Add(nm)

        ut = TextBlock()
        ut.Text = self._uid(rec)[:8]
        ut.VerticalAlignment = VerticalAlignment.Center
        ut.Foreground = _brush((150, 155, 165))
        ut.FontSize = 11
        Grid.SetColumn(ut, 3)
        row.Children.Add(ut)
        return row

    def _on_ok(self, sender, args):
        drop = set(id(rec) for cb, rec in self._row_checks if cb.IsChecked)
        kept = [r for r in self.records if id(r) not in drop]
        # 每個 UID 最多保留 1 列
        counts = {}
        for r in kept:
            u = self._uid(r)
            if u:
                counts[u] = counts.get(u, 0) + 1
        still = sorted([u[:8] for u, c in counts.items() if c > 1])
        if still:
            self.err_text.Text = (u"這些 UID 仍保留超過 1 列，請再勾選移除："
                                  + u"、".join(still))
            return
        self.cleaned_records = kept
        self.confirmed = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.Close()


# ---------------------------------------------------------------------------
# 執行（Revit）
# ---------------------------------------------------------------------------

def create_sheets(result, document):
    sheets_data = result["sheets"]
    tb_id = result.get("title_block_id") or DB.ElementId.InvalidElementId
    
    created = []
    failed = []
    with revit.Transaction("Batch Create Sheets"):
        for num, name in sheets_data:
            try:
                sheet = DB.ViewSheet.Create(document, tb_id)
                sheet.SheetNumber = num
                sheet.Name = name
                created.append((num, name))
            except Exception as ex:
                failed.append((num, name, str(ex)))
    return created, failed


def edit_sheets(result, document):
    edits = result["edits"]
    edited = []
    failed = []
    
    with revit.Transaction("Batch Edit Sheets"):
        # 因為新圖號可能跟其他圖紙的舊圖號衝突，採用兩階段：
        # 1. 先把所有需要改圖號的圖紙暫時改成獨特的暫時編號（避免互撞）
        # 2. 再改成最終值
        sheets_changing_num = [(s, n, name) for s, n, name in edits if n]
        sheets_only_name = [(s, n, name) for s, n, name in edits if not n]
        
        # 處理只改圖名的（不會撞）
        for sheet, _, new_name in sheets_only_name:
            try:
                old_num = sheet.SheetNumber
                old_name = sheet.Name
                sheet.Name = new_name
                edited.append((old_num, old_name, old_num, new_name))
            except Exception as ex:
                failed.append((sheet.SheetNumber, str(ex)))
        
        # 第一階段：暫時編號
        temp_map = {}
        for i, (sheet, new_num, _) in enumerate(sheets_changing_num):
            try:
                temp_num = "__TMP_{}_{}__".format(i, sheet.Id.IntegerValue)
                temp_map[sheet.Id.IntegerValue] = (sheet, temp_num, new_num)
                old_num = sheet.SheetNumber
                old_name = sheet.Name
                sheet.SheetNumber = temp_num
                # 把原始資料先存起來，後面要用
                temp_map[sheet.Id.IntegerValue] = (sheet, temp_num, new_num, old_num, old_name)
            except Exception as ex:
                failed.append((sheet.SheetNumber, "暫時編號失敗: {}".format(ex)))
        
        # 第二階段：套用最終值
        for sheet_id_int, value in temp_map.items():
            sheet, temp_num, new_num, old_num, old_name = value
            try:
                sheet.SheetNumber = new_num
                # 對應的新圖名（如果有）
                _, _, new_name = next(((s, n, nm) for s, n, nm in sheets_changing_num
                                       if s.Id.IntegerValue == sheet_id_int), (None, None, None))
                if new_name:
                    sheet.Name = new_name
                final_name = sheet.Name
                edited.append((old_num, old_name, new_num, final_name))
            except Exception as ex:
                failed.append((old_num, "套用新圖號 '{}' 失敗: {}".format(new_num, ex)))
    
    return edited, failed


# ---------------------------------------------------------------------------
# 建立同步環境（在 Sheets 類別建立文字參數 + 圖紙明細表）
# ---------------------------------------------------------------------------

# 公司共用參數檔所在資料夾（請勿改動該資料夾內容）
COMPANY_SP_DIR = (u"\\\\data.bioarch.com.tw\\Public\\工作區"
                  u"\\03 技術中心專區\\06-2 REVIT\\03 Revit 資料庫\\03 共用參數")


def _text_spec():
    """文字資料型別：Revit 2022+ 用 SpecTypeId.String.Text，舊版用 ParameterType.Text。"""
    try:
        return DB.SpecTypeId.String.Text
    except Exception:
        return DB.ParameterType.Text


# Revit 共用參數檔表頭（空檔，Revit 之後會自行改寫）
_SP_HEADER = (
    u"# This is a Revit shared parameter file.\n"
    u"# Do not edit manually.\n"
    u"*META\tVERSION\tMINVERSION\n"
    u"META\t2\t1\n"
    u"*GROUP\tID\tNAME\n"
    u"*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE"
    u"\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE\n")


def _find_company_sp_file():
    """回傳公司共用參數檔(.txt)的完整路徑；資料夾不可達/無 txt 時回 None。"""
    try:
        names = os.listdir(COMPANY_SP_DIR)
    except Exception:
        return None
    txts = []
    for n in names:
        if n.lower().endswith(u".txt"):
            full = os.path.join(COMPANY_SP_DIR, n)
            try:
                if os.path.isfile(full):
                    txts.append(full)
            except Exception:
                pass
    if not txts:
        return None
    try:
        txts.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    except Exception:
        pass
    return txts[0]


def build_env(result, document):
    """在「圖紙(Sheets)」類別建立缺少的文字參數，並建立「圖紙明細表」。

    優先沿用公司共用參數（若該名稱已在公司共用參數檔中，保留其 GUID）；
    公司檔沒有的，才用本機暫存共用參數檔建為專案參數。
    不會對 Google Sheet 做任何動作。

    回傳 (created_shared, created_project, failed, sched_msg)。
    """
    import tempfile

    missing = list(result.get("missing") or [])
    all_params = result.get("all_params") or missing
    created_shared = []
    created_project = []
    failed = []
    sched_msg = u""
    if not missing:
        # 仍嘗試建立明細表（參數都已存在的情況）
        try:
            sched_msg = _ensure_sheet_schedule(document, all_params)
        except Exception as ex:
            sched_msg = u"⚠ 建立圖紙明細表失敗：{}".format(ex)
        return created_shared, created_project, failed, sched_msg

    app = document.Application
    spec = _text_spec()
    sheets_cat = DB.Category.GetCategory(document, DB.BuiltInCategory.OST_Sheets)

    def _bind(ext_def):
        cs = app.Create.NewCategorySet()
        cs.Insert(sheets_cat)
        binding = app.Create.NewInstanceBinding(cs)
        try:
            return document.ParameterBindings.Insert(
                ext_def, binding, DB.BuiltInParameterGroup.PG_TEXT)
        except Exception:
            # 新版本可能移除 BuiltInParameterGroup 多載 → 用兩參數版
            return document.ParameterBindings.Insert(ext_def, binding)

    old_sp = None
    try:
        old_sp = app.SharedParametersFilename
    except Exception:
        old_sp = None

    company_file = _find_company_sp_file()

    try:
        with revit.Transaction(u"BaF 建立同步參數"):
            remaining = list(missing)

            # (1) 公司共用參數檔已有的 → 直接沿用共用參數定義（保留 GUID）
            if company_file:
                cdef = None
                try:
                    app.SharedParametersFilename = company_file
                    cdef = app.OpenSharedParameterFile()
                except Exception:
                    cdef = None
                if cdef is not None:
                    found = {}
                    for g in cdef.Groups:
                        for d in g.Definitions:
                            if d.Name in remaining and d.Name not in found:
                                found[d.Name] = d
                    for name in list(remaining):
                        ext_def = found.get(name)
                        if ext_def is None:
                            continue
                        try:
                            if _bind(ext_def):
                                created_shared.append(name)
                            else:
                                failed.append((name, u"共用參數綁定未成功"))
                        except Exception as ex:
                            failed.append((name, u"共用:" + unicode(ex)))
                        remaining.remove(name)

            # (2) 公司檔沒有的 → 用本機暫存共用參數檔建為專案參數
            if remaining:
                sp_path = os.path.join(tempfile.gettempdir(),
                                       u"BaF_shared_params.txt")
                tdef = None
                try:
                    with open(sp_path, "w") as f:
                        f.write(_SP_HEADER.encode("utf-8"))
                    app.SharedParametersFilename = sp_path
                    tdef = app.OpenSharedParameterFile()
                except Exception as ex:
                    for name in remaining:
                        failed.append((name, u"建暫存共用參數檔失敗:" + unicode(ex)))
                    remaining = []
                if remaining and tdef is None:
                    for name in remaining:
                        failed.append((name, u"無法開啟暫存共用參數檔"))
                elif remaining:
                    group = None
                    for g in tdef.Groups:
                        if g.Name == u"BaF":
                            group = g
                            break
                    if group is None:
                        group = tdef.Groups.Create(u"BaF")
                    for name in remaining:
                        try:
                            ext_def = None
                            for d in group.Definitions:
                                if d.Name == name:
                                    ext_def = d
                                    break
                            if ext_def is None:
                                opts = DB.ExternalDefinitionCreationOptions(name, spec)
                                ext_def = group.Definitions.Create(opts)
                            if _bind(ext_def):
                                created_project.append(name)
                            else:
                                failed.append((name, u"專案參數綁定未成功"))
                        except Exception as ex:
                            failed.append((name, unicode(ex)))
    finally:
        # 還原使用者原本的共用參數檔設定
        try:
            if old_sp:
                app.SharedParametersFilename = old_sp
        except Exception:
            pass

    # (3) 建立「圖紙明細表」（參數已綁定並提交後才建，欄位才抓得到）
    try:
        sched_msg = _ensure_sheet_schedule(document, all_params)
    except Exception as ex:
        sched_msg = u"⚠ 建立圖紙明細表失敗：{}".format(ex)

    return created_shared, created_project, failed, sched_msg


def _ensure_sheet_schedule(document, want_params):
    """建立「圖紙明細表」(若同名的已存在則不重複建立)。回傳說明文字。"""
    sched_name = u"BaF_Sheet專用圖紙明細表"
    for v in DB.FilteredElementCollector(document).OfClass(DB.ViewSchedule).ToElements():
        try:
            if v.Name == sched_name:
                return u"「{}」已存在，未重複建立。".format(sched_name)
        except Exception:
            pass

    added = []
    with revit.Transaction(u"BaF 建立圖紙明細表"):
        # 圖紙清單必須用 CreateSheetList()；用 CreateSchedule(OST_Sheets) 會丟例外
        try:
            sched = DB.ViewSchedule.CreateSheetList(document)
        except Exception:
            sheets_cat_id = DB.Category.GetCategory(
                document, DB.BuiltInCategory.OST_Sheets).Id
            sched = DB.ViewSchedule.CreateSchedule(document, sheets_cat_id)
        try:
            sched.Name = sched_name
        except Exception:
            pass
        definition = sched.Definition

        # CreateSheetList 可能已預設帶圖號/圖名欄位 → 記下避免重複加
        existing_pids = set()
        try:
            for i in range(definition.GetFieldCount()):
                try:
                    existing_pids.add(definition.GetField(i).ParameterId.IntegerValue)
                except Exception:
                    pass
        except Exception:
            pass

        sf_by_name = {}
        sf_by_pid = {}
        for sf in definition.GetSchedulableFields():
            try:
                nm = sf.GetName(document)
            except Exception:
                nm = None
            if nm and nm not in sf_by_name:
                sf_by_name[nm] = sf
            try:
                pid = sf.ParameterId.IntegerValue
                if pid not in sf_by_pid:
                    sf_by_pid[pid] = sf
            except Exception:
                pass

        def _field_for(name):
            sf = sf_by_name.get(name)
            if sf is not None:
                return sf
            # 內建參數(如 圖紙發布日期)用 ParameterId 對應，避免比錯中文用字
            bip = BUILTIN_PARAM_MAP.get(name)
            if bip is not None:
                try:
                    pid = DB.ElementId(bip).IntegerValue
                    return sf_by_pid.get(pid)
                except Exception:
                    return None
            return None

        # 先放圖號/圖名(內建欄位名稱依語言而異，逐一嘗試)，再放 8 個同步參數
        ordered = []  # (顯示名稱, SchedulableField)
        for cand in [u"圖紙編號", u"圖紙號碼", u"Sheet Number"]:
            if cand in sf_by_name:
                ordered.append((cand, sf_by_name[cand]))
                break
        for cand in [u"圖紙名稱", u"Sheet Name"]:
            if cand in sf_by_name:
                ordered.append((cand, sf_by_name[cand]))
                break
        seen_names = set(nm for nm, _ in ordered)
        for nm in want_params:
            if nm in seen_names:
                continue
            sf = _field_for(nm)
            if sf is not None:
                ordered.append((nm, sf))
                seen_names.add(nm)

        for nm, sf in ordered:
            try:
                pid = sf.ParameterId.IntegerValue
            except Exception:
                pid = None
            if pid is not None and pid in existing_pids:
                added.append(nm)  # 已是預設欄位，不重複加
                continue
            try:
                definition.AddField(sf)
                added.append(nm)
                if pid is not None:
                    existing_pids.add(pid)
            except Exception:
                pass

    if added:
        return u"已建立「{}」，欄位：{}".format(sched_name, u"、".join(added))
    return u"已建立「{}」（但未加入欄位，請手動設定）".format(sched_name)


class _RedpenStartWindow(Window):
    """按下「依紅筆圖註解」後跳出：選 AI 視覺判讀後端、可開金鑰設定，再開始。
    沒填金鑰時提醒並停留在本視窗（不中止），讓使用者當場補金鑰。"""

    def __init__(self):
        self.confirmed = False
        self.provider = u"claude"

        self.Title = u"依紅筆圖註解 — 選擇 AI 視覺判讀後端"
        self.Width = 460
        self.Height = 240
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = _brush((245, 245, 250))

        cfg = load_gsheet_cfg()

        root = StackPanel()
        root.Margin = Thickness(18, 16, 18, 16)

        intro = TextBlock()
        intro.Text = (u"選擇要用哪個 AI 視覺模型判讀紅筆 PDF，按「開始」後再選 PDF 檔。\n"
                      u"第一次使用請先按「🔑 金鑰／模型設定」填入金鑰。")
        intro.TextWrapping = TextWrapping.Wrap
        intro.FontSize = 12
        intro.Margin = Thickness(0, 0, 0, 14)
        root.Children.Add(intro)

        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.Margin = Thickness(0, 0, 0, 8)

        lab = TextBlock()
        lab.Text = u"後端："
        lab.VerticalAlignment = VerticalAlignment.Center
        lab.FontWeight = FontWeights.SemiBold
        lab.Margin = Thickness(0, 0, 8, 0)
        row.Children.Add(lab)

        self._combo = ComboBox()
        self._combo.Width = 200
        self._combo.Padding = Thickness(4, 3, 4, 3)
        for _pk, _pl in AI_PROVIDER_LABELS:
            self._combo.Items.Add(_pl)
        try:
            _idx = [pk for pk, _ in AI_PROVIDER_LABELS].index(
                cfg.get("ai_provider") or u"claude")
        except Exception:
            _idx = 0
        self._combo.SelectedIndex = _idx
        row.Children.Add(self._combo)

        set_btn = Button()
        set_btn.Content = u"🔑 金鑰／模型設定"
        set_btn.Padding = Thickness(10, 4, 10, 4)
        set_btn.Margin = Thickness(10, 0, 0, 0)
        set_btn.Click += self._on_settings
        row.Children.Add(set_btn)

        root.Children.Add(row)

        btns = StackPanel()
        btns.Orientation = Orientation.Horizontal
        btns.HorizontalAlignment = HorizontalAlignment.Right
        btns.Margin = Thickness(0, 18, 0, 0)

        ok = Button()
        ok.Content = u"開始（選 PDF）"
        ok.Padding = Thickness(16, 6, 16, 6)
        ok.Margin = Thickness(0, 0, 8, 0)
        ok.Background = _brush((99, 102, 241))
        ok.Foreground = _brush((255, 255, 255))
        ok.FontWeight = FontWeights.Bold
        ok.Click += self._on_ok
        btns.Children.Add(ok)

        cancel = Button()
        cancel.Content = u"取消"
        cancel.Padding = Thickness(16, 6, 16, 6)
        cancel.Click += self._on_cancel
        btns.Children.Add(cancel)

        root.Children.Add(btns)
        self.Content = root

    def _selected_provider(self):
        i = self._combo.SelectedIndex
        if i is None or i < 0 or i >= len(AI_PROVIDER_LABELS):
            i = 0
        return AI_PROVIDER_LABELS[i][0]

    def _has_key(self, provider):
        cfg = load_gsheet_cfg()
        return bool(((cfg.get("ai_keys") or {}).get(provider) or u"").strip())

    def _on_settings(self, sender, args):
        dlg = _AiSettingsWindow()
        dlg.ShowDialog()

    def _on_ok(self, sender, args):
        provider = self._selected_provider()
        if not self._has_key(provider):
            # 沒設金鑰：直接開「金鑰／模型設定」視窗讓使用者填，不跳警告
            dlg = _AiSettingsWindow()
            dlg.ShowDialog()
            if not self._has_key(provider):
                # 取消或仍未填該後端金鑰 → 留在本視窗（不中止）
                return
        # 記住這次選的後端
        try:
            save_gsheet_cfg({"ai_provider": provider})
        except Exception:
            pass
        self.provider = provider
        self.confirmed = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.Close()


class _RedpenProgressWindow(Window):
    """依紅筆圖註解的處理進度小視窗：逐步說明呼叫 AI 判讀過程中發生了什麼。
    非模態（Show）；AI 呼叫在主執行緒同步阻塞，期間視窗會停在最後一句訊息上。"""

    def __init__(self, provider_label):
        self._closed = False
        self.Title = u"依紅筆圖註解 — 處理進度"
        self.Width = 470
        self.Height = 300
        self.WindowStartupLocation = WindowStartupLocation.CenterOwner
        self.Background = _brush((245, 245, 250))
        self.Closed += self._on_closed

        root = Grid()
        root.Margin = Thickness(16, 14, 16, 14)
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))

        head = TextBlock()
        head.Text = u"正在以「{}」判讀紅筆 PDF".format(provider_label)
        head.FontWeight = FontWeights.Bold
        head.FontSize = 13
        head.Margin = Thickness(0, 0, 0, 8)
        Grid.SetRow(head, 0)
        root.Children.Add(head)

        sv = ScrollViewer()
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.Background = _brush((255, 255, 255))
        sv.BorderBrush = _brush((200, 205, 215))
        sv.BorderThickness = Thickness(1)
        sv.Padding = Thickness(8, 8, 8, 8)
        self._sv = sv
        self._log = TextBlock()
        self._log.Text = u""
        self._log.TextWrapping = TextWrapping.Wrap
        self._log.FontSize = 12
        sv.Content = self._log
        Grid.SetRow(sv, 1)
        root.Children.Add(sv)

        foot = TextBlock()
        foot.Text = u"完成後會自動關閉並開啟預覽，請勿手動關閉本視窗。"
        foot.FontSize = 11
        foot.Foreground = _brush((130, 135, 145))
        foot.Margin = Thickness(0, 8, 0, 0)
        Grid.SetRow(foot, 2)
        root.Children.Add(foot)

        self.Content = root

    def _on_closed(self, sender, args):
        self._closed = True

    def log(self, line):
        if self._closed:
            return
        try:
            self._log.Text = (self._log.Text or u"") + line + u"\n"
            self._sv.ScrollToEnd()
            _doevents()
        except Exception:
            pass

    def close(self):
        if self._closed:
            return
        try:
            self.Close()
        except Exception:
            pass


class _AiSettingsWindow(Window):
    """AI 金鑰／模型設定：三家各一組（金鑰遮蔽、模型可改），存進 baf_gsheet_config.json。"""

    def __init__(self):
        self.Title = u"AI 金鑰／模型設定"
        self.Width = 560
        self.Height = 420
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = _brush((245, 245, 250))

        cfg = load_gsheet_cfg()
        keys = cfg.get("ai_keys") or {}
        models = cfg.get("ai_models") or {}

        root = StackPanel()
        root.Margin = Thickness(18, 16, 18, 16)

        intro = TextBlock()
        intro.Text = (u"填入你要使用的 AI 服務金鑰（只存在本機設定檔，不會進 git）。\n"
                      u"模型留空＝使用預設。")
        intro.TextWrapping = TextWrapping.Wrap
        intro.FontSize = 12
        intro.Margin = Thickness(0, 0, 0, 12)
        root.Children.Add(intro)

        self._key_boxes = {}    # provider -> PasswordBox
        self._model_boxes = {}  # provider -> TextBox
        for pk, pl in AI_PROVIDER_LABELS:
            lab = TextBlock()
            lab.Text = pl
            lab.FontWeight = FontWeights.Bold
            lab.FontSize = 12
            lab.Margin = Thickness(0, 6, 0, 4)
            root.Children.Add(lab)

            row = Grid()
            row.ColumnDefinitions.Add(
                ColumnDefinition(Width=GridLength(2, GridUnitType.Star)))
            row.ColumnDefinitions.Add(
                ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
            row.Margin = Thickness(0, 0, 0, 6)

            kb = PasswordBox()
            kb.Padding = Thickness(4, 4, 4, 4)
            kb.Margin = Thickness(0, 0, 6, 0)
            try:
                kb.Password = keys.get(pk) or u""
            except Exception:
                pass
            Grid.SetColumn(kb, 0)
            row.Children.Add(kb)
            self._key_boxes[pk] = kb

            mb = TextBox()
            mb.Padding = Thickness(4, 4, 4, 4)
            mb.Text = models.get(pk) or u""
            try:
                mb.ToolTip = u"預設：{}".format(AI_DEFAULT_MODELS.get(pk) or u"")
            except Exception:
                pass
            Grid.SetColumn(mb, 1)
            row.Children.Add(mb)
            self._model_boxes[pk] = mb

            hint = TextBlock()
            hint.Text = u"（金鑰　／　模型，預設 {}）".format(
                AI_DEFAULT_MODELS.get(pk) or u"")
            hint.FontSize = 10
            hint.Foreground = _brush((130, 135, 145))
            hint.Margin = Thickness(0, 0, 0, 4)
            root.Children.Add(row)
            root.Children.Add(hint)

        btns = StackPanel()
        btns.Orientation = Orientation.Horizontal
        btns.HorizontalAlignment = HorizontalAlignment.Right
        btns.Margin = Thickness(0, 14, 0, 0)

        save = Button()
        save.Content = u"儲存"
        save.Padding = Thickness(16, 6, 16, 6)
        save.Margin = Thickness(0, 0, 8, 0)
        save.Background = _brush((99, 102, 241))
        save.Foreground = _brush((255, 255, 255))
        save.FontWeight = FontWeights.Bold
        save.Click += self._on_save
        btns.Children.Add(save)

        cancel = Button()
        cancel.Content = u"取消"
        cancel.Padding = Thickness(16, 6, 16, 6)
        cancel.Click += self._on_cancel
        btns.Children.Add(cancel)

        root.Children.Add(btns)
        self.Content = root

    def _on_save(self, sender, args):
        keys, models = {}, {}
        for pk, _ in AI_PROVIDER_LABELS:
            try:
                keys[pk] = (self._key_boxes[pk].Password or u"").strip()
            except Exception:
                keys[pk] = u""
            models[pk] = (self._model_boxes[pk].Text or u"").strip()
        save_gsheet_cfg({"ai_keys": keys, "ai_models": models})
        self.Close()

    def _on_cancel(self, sender, args):
        self.Close()


class _RedpenPreviewWindow(Window):
    """依紅筆圖註解預覽：逐張顯示現有備註 vs AI 判讀，並各自選『附加／覆蓋』。"""

    SEP = u"；"  # 附加時，舊備註與新內容之間的分隔（用全形分號而非換行，
    #             因 Revit 明細表儲存格遇換行只顯示第一行 → 改用「；」同一行顯示完整內容）

    def __init__(self, matched, unmatched, can_sync):
        self.confirmed = False
        self.push_to_sheet = False
        self._matched = matched          # (sheet, num, name, old, new)
        self._can_sync = can_sync
        self._rows = []                  # {sheet, old, new, chk, mode}

        self.Title = u"依紅筆圖註解 — 預覽（逐張確認）"
        self.Width = 980
        self.Height = 640
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = _brush((245, 245, 250))

        root = Grid()
        root.Margin = Thickness(16, 14, 16, 14)
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))

        head = TextBlock()
        head.Text = (u"AI 判讀到 {} 張可對應的圖。請確認每張的『處理方式』"
                     u"（附加＝保留原備註並接在後面；覆蓋＝以 AI 內容取代）。".format(
                         len(matched)))
        head.TextWrapping = TextWrapping.Wrap
        head.FontSize = 13
        head.FontWeight = FontWeights.Bold
        head.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(head, 0)
        root.Children.Add(head)

        sv = ScrollViewer()
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Visible
        sv.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.Background = _brush((255, 255, 255))
        sv.BorderBrush = _brush((200, 205, 215))
        sv.BorderThickness = Thickness(1)
        sv.Padding = Thickness(8, 8, 8, 8)
        sv.Content = self._build_table(matched, unmatched)
        Grid.SetRow(sv, 1)
        root.Children.Add(sv)

        bottom = StackPanel()
        bottom.Orientation = Orientation.Horizontal
        bottom.Margin = Thickness(0, 12, 0, 0)

        self._sync_chk = CheckBox()
        self._sync_chk.Content = u"同時同步到 Google Sheet（總表＋分表）"
        self._sync_chk.VerticalAlignment = VerticalAlignment.Center
        self._sync_chk.IsChecked = bool(can_sync)
        self._sync_chk.IsEnabled = bool(can_sync)
        if not can_sync:
            self._sync_chk.Content = u"同時同步到 Google Sheet（需先填 URL／頁籤）"
        bottom.Children.Add(self._sync_chk)

        spacer = TextBlock()
        spacer.Width = 30
        bottom.Children.Add(spacer)

        ok = Button()
        ok.Content = u"確認寫入"
        ok.Padding = Thickness(16, 6, 16, 6)
        ok.Margin = Thickness(0, 0, 8, 0)
        ok.Background = _brush((34, 139, 34))
        ok.Foreground = _brush((255, 255, 255))
        ok.FontWeight = FontWeights.Bold
        ok.Click += self._on_ok
        bottom.Children.Add(ok)

        cancel = Button()
        cancel.Content = u"取消"
        cancel.Padding = Thickness(16, 6, 16, 6)
        cancel.Click += self._on_cancel
        bottom.Children.Add(cancel)

        Grid.SetRow(bottom, 2)
        root.Children.Add(bottom)
        self.Content = root

    def _build_table(self, matched, unmatched):
        outer = StackPanel()

        grid = Grid()
        headers = [u"套用", u"圖號", u"圖名", u"現有修正備註", u"AI 判讀內容", u"處理方式"]
        widths = [GridLength(1, GridUnitType.Auto),
                  GridLength(1, GridUnitType.Auto),
                  GridLength(150),
                  GridLength(230),
                  GridLength(300),
                  GridLength(1, GridUnitType.Auto)]
        for w in widths:
            grid.ColumnDefinitions.Add(ColumnDefinition(Width=w))
        for _ in range(len(matched) + 1):
            grid.RowDefinitions.Add(
                RowDefinition(Height=GridLength(1, GridUnitType.Auto)))

        for c, h in enumerate(headers):
            tb = TextBlock()
            tb.Text = h
            tb.FontWeight = FontWeights.Bold
            tb.Margin = Thickness(6, 2, 12, 6)
            Grid.SetRow(tb, 0)
            Grid.SetColumn(tb, c)
            grid.Children.Add(tb)

        for r, (sheet, num, name, old, new) in enumerate(matched):
            chk = CheckBox()
            chk.IsChecked = True
            chk.VerticalAlignment = VerticalAlignment.Center
            chk.Margin = Thickness(6, 4, 12, 4)
            Grid.SetRow(chk, r + 1)
            Grid.SetColumn(chk, 0)
            grid.Children.Add(chk)

            grid.Children.Add(self._cell(num, r + 1, 1, bold=True))
            grid.Children.Add(self._cell(name, r + 1, 2))
            grid.Children.Add(self._cell(old if old else u"（空）", r + 1, 3,
                                         color=(120, 120, 130)))
            grid.Children.Add(self._cell(new, r + 1, 4))

            mode = ComboBox()
            mode.Items.Add(u"附加")
            mode.Items.Add(u"覆蓋")
            mode.SelectedIndex = 0 if old else 1
            mode.Margin = Thickness(6, 4, 6, 4)
            mode.VerticalAlignment = VerticalAlignment.Center
            Grid.SetRow(mode, r + 1)
            Grid.SetColumn(mode, 5)
            grid.Children.Add(mode)

            self._rows.append({"sheet": sheet, "old": old, "new": new,
                               "chk": chk, "mode": mode})

        outer.Children.Add(grid)

        if unmatched:
            warn = TextBlock()
            warn.Text = (u"\n⚠ 以下 {} 個圖號對不到目前 Revit 的圖紙，已略過：\n".format(
                len(unmatched)) + u"\n".join(
                u"・{}　{}".format(n, (note[:40] + u"…") if len(note) > 40 else note)
                for n, note in unmatched))
            warn.TextWrapping = TextWrapping.Wrap
            warn.FontSize = 11
            warn.Foreground = _brush((180, 60, 60))
            warn.Margin = Thickness(2, 10, 2, 2)
            outer.Children.Add(warn)

        return outer

    def _cell(self, text, row, col, bold=False, color=None):
        tb = TextBlock()
        tb.Text = text if text is not None else u""
        tb.TextWrapping = TextWrapping.Wrap
        tb.Margin = Thickness(6, 4, 12, 4)
        if bold:
            tb.FontWeight = FontWeights.Bold
        if color is not None:
            tb.Foreground = _brush(color)
        Grid.SetRow(tb, row)
        Grid.SetColumn(tb, col)
        return tb

    def get_updates(self):
        """回傳勾選的 [(UniqueId, 最終備註)]。附加＝舊＋分隔＋新；覆蓋＝新。"""
        out = []
        for row in self._rows:
            if not row["chk"].IsChecked:
                continue
            old, new = row["old"], row["new"]
            overwrite = (row["mode"].SelectedIndex == 1)
            if overwrite or not old:
                final = new
            else:
                final = old + self.SEP + new
            out.append((row["sheet"].UniqueId, final))
        return out

    def _on_ok(self, sender, args):
        self.push_to_sheet = bool(self._can_sync and self._sync_chk.IsChecked)
        self.confirmed = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.Close()


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def _open_manager_window(initial_tab=None):
    """開啟管理器視窗（可指定初始分頁），回傳關閉後的視窗物件。"""
    from System import AppDomain
    _WKEY = "BAF_SheetMgr_Window"
    prev = AppDomain.CurrentDomain.GetData(_WKEY)
    if prev is not None:
        try:
            prev.Close()
        except Exception:
            pass
        AppDomain.CurrentDomain.SetData(_WKEY, None)

    title_blocks = get_title_block_types(doc)
    existing_sheets = get_existing_sheets(doc)

    win = BatchManageSheetsWindow(title_blocks, existing_sheets)
    if initial_tab is not None:
        try:
            win.tabs.SelectedIndex = initial_tab
        except Exception:
            pass
    # 讓視窗固定停在 Revit 主視窗之上（不會被 Revit 蓋住），直到使用者自行關閉。
    # 作法：把 Revit 主視窗設為本視窗的 Owner（本腳本就跑在 Revit 行程內）。
    try:
        from System.Windows.Interop import WindowInteropHelper
        from System.Diagnostics import Process
        from System import IntPtr
        rvt_handle = Process.GetCurrentProcess().MainWindowHandle
        if rvt_handle != IntPtr.Zero:
            WindowInteropHelper(win).Owner = rvt_handle
        else:
            win.Topmost = True
    except Exception:
        try:
            win.Topmost = True
        except Exception:
            pass

    AppDomain.CurrentDomain.SetData(_WKEY, win)
    win.ShowDialog()
    # 只有當記錄的還是自己時才清掉（避免清掉後開的新視窗）
    if AppDomain.CurrentDomain.GetData(_WKEY) is win:
        AppDomain.CurrentDomain.SetData(_WKEY, None)
    return win


def _prompt_sync_after_change():
    """新增/編輯/刪除圖紙後，詢問是否要與 Google Sheet 同步。"""
    return forms.alert(
        u"Revit 圖紙清單已更新。\n\n要現在與 Google Sheet 同步嗎？\n"
        u"（選「是」會切換到「Google Sheet 同步」分頁；選「否」則只更新 Revit）",
        yes=True, no=True)


def main():
    next_tab = None
    while True:
        win = _open_manager_window(next_tab)
        next_tab = None

        if not win.confirmed:
            script.exit()
            return

        mode = win.result["mode"]

        if mode == "create":
            created, failed = create_sheets(win.result, doc)
            output.print_md("# ✅ 批次新增圖紙完成")
            output.print_md("- 成功: **{}** 張".format(len(created)))
            output.print_md("- 失敗: **{}** 張".format(len(failed)))
            if created:
                output.print_md("\n## 已建立")
                for num, name in created:
                    output.print_md("- `{}` - {}".format(num, name))
            if failed:
                output.print_md("\n## 失敗")
                for num, name, reason in failed:
                    output.print_md("- `{}` - {}: {}".format(num, name, reason))
            if created and _prompt_sync_after_change():
                next_tab = BatchManageSheetsWindow.MODE_SYNC
                continue
            return

        elif mode == "edit":
            edited, failed = edit_sheets(win.result, doc)
            output.print_md("# ✅ 批次編輯圖紙完成")
            output.print_md("- 成功: **{}** 張".format(len(edited)))
            output.print_md("- 失敗: **{}** 張".format(len(failed)))
            if edited:
                output.print_md("\n## 已變更")
                for old_num, old_name, new_num, new_name in edited:
                    output.print_md("- `{} | {}` → `{} | {}`".format(old_num, old_name, new_num, new_name))
            if failed:
                output.print_md("\n## 失敗")
                for old_num, reason in failed:
                    output.print_md("- `{}`: {}".format(old_num, reason))
            if edited and _prompt_sync_after_change():
                next_tab = BatchManageSheetsWindow.MODE_SYNC
                continue
            return

        elif mode == "build_env":
            created_shared, created_project, failed, sched_msg = build_env(win.result, doc)
            have = win.result.get("have") or []
            output.print_md("# 🛠️ 建立同步環境")
            output.print_md("- 沿用公司共用參數: **{}** 個".format(len(created_shared)))
            output.print_md("- 新建專案參數: **{}** 個".format(len(created_project)))
            output.print_md("- 已存在略過: **{}** 個".format(len(have)))
            output.print_md("- 失敗: **{}** 個".format(len(failed)))
            if created_shared:
                output.print_md("\n## 沿用公司共用參數（保留 GUID）")
                for name in created_shared:
                    output.print_md("- `{}`".format(name))
            if created_project:
                output.print_md("\n## 新建專案參數")
                for name in created_project:
                    output.print_md("- `{}`".format(name))
            if have:
                output.print_md("\n## 已存在（略過）")
                for name in have:
                    output.print_md("- `{}`".format(name))
            if failed:
                output.print_md("\n## 失敗")
                for name, reason in failed:
                    output.print_md("- `{}`: {}".format(name, reason))
            if sched_msg:
                output.print_md("\n## 圖紙明細表")
                output.print_md("- {}".format(sched_msg))
            if not failed:
                output.print_md("\n✅ 環境就緒，可在圖紙屬性填寫並與 Google Sheet 同步。"
                                "（本步驟未更動 Google Sheet）")
            return

        elif mode == "import_apply":
            plan = win.result["plan"]
            # 在 modal 視窗關閉「之後」才改模型，交易才不會被還原
            report, is_err, done = win._apply_plan(plan)
            output.print_md("# {} 匯入套用".format(u"❌" if is_err else u"✅"))
            output.print_md(report)
            role = plan.get("role") or "main"
            main_tab = plan.get("mainTab") or plan.get("tab")
            if not is_err and role == "sub":
                # 分表匯入後：範圍限定回寫 —— 只更新總表中「該階段」的列、
                # 只重建「該階段」分表；其他階段與分表完全不動（避免誤刪/覆蓋）。
                cat = plan.get("category") or u""
                res, err, _ = win._do_sync_category(plan["url"], main_tab, cat)
                if err:
                    output.print_md("\n⚠ 自動回寫（範圍：{}）失敗：{}".format(cat, err))
                elif res and res.get("ok"):
                    output.print_md(
                        "\n🔄 已只回寫「{}」階段：更新總表中該階段的列、重建該階段分表"
                        "（其他階段與分表未變動）。".format(cat))
                else:
                    output.print_md("\n⚠ 自動回寫回應異常：{}".format(res))
            elif not is_err:
                # 總表匯入後的自動回寫
                changed = any(done.get(k, 0) for k in
                              ("new", "edit", "toReal", "toPlace", "delete"))
                # 只有 UID 會變動(新建/降為僅文字)或解決過重複圖號時，才需要回寫總表本身
                need_main = (done.get("new", 0) > 0 or done.get("toPlace", 0) > 0
                             or plan.get("resolvedDup"))
                if need_main:
                    res, err, n = win._do_export(plan["url"], main_tab)
                    if err:
                        output.print_md("\n⚠ 自動回寫總表失敗：{}".format(err))
                    elif res and res.get("ok"):
                        output.print_md("\n🔄 已自動回寫總表（更新變動/新建的 UID、已修正的重複圖號）。")
                    else:
                        output.print_md("\n⚠ 自動回寫總表回應異常：{}".format(res))
                # 修正：先前這裡只更新總表、漏了分表。只要有變動就一併重建所有分表，
                # 讓分表內容與總表同步（與手動「匯出後拆分」相同的拆分邏輯）。
                if changed or plan.get("resolvedDup"):
                    sres, serr, _ = win._do_split(plan["url"], main_tab)
                    if serr:
                        output.print_md("\n⚠ 分表自動更新失敗：{}".format(serr))
                    elif sres and sres.get("ok"):
                        made = sres.get("splitTabs") or []
                        output.print_md(
                            "\n🗂️ 已一併更新所有分表（{} 個：{}）。".format(
                                len(made), u"、".join(made))
                            if made else
                            "\n🗂️ 已重建分表（目前沒有可依圖紙類別拆分的資料）。")
                    else:
                        output.print_md("\n⚠ 分表自動更新回應異常：{}".format(sres))
            return

        elif mode == "redpen":
            updates = win.result.get("updates") or []
            url = win.result.get("url") or u""
            tab = win.result.get("tab") or u""
            push = win.result.get("push_to_sheet")
            by_uid = dict(
                (s.UniqueId, s) for s in
                DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements())
            ok_n, fails = 0, []
            try:
                with revit.Transaction("BaF 依紅筆圖註解寫入修正備註"):
                    for uid, note in updates:
                        s = by_uid.get(uid)
                        if s is None:
                            fails.append((uid, u"找不到圖紙"))
                            continue
                        # 明確檢查「修正備註」參數是否存在/可寫，避免靜默跳過卻誤報成功
                        p = _lookup_sheet_param(s, u"修正備註")
                        if p is None:
                            fails.append((s.SheetNumber or uid[:8],
                                          u"此圖沒有『修正備註』參數，請先按「🛠️ 建立環境」"))
                            continue
                        if p.IsReadOnly:
                            fails.append((s.SheetNumber or uid[:8], u"『修正備註』為唯讀"))
                            continue
                        try:
                            win._set_param(s, u"修正備註", note)
                            ok_n += 1
                        except Exception as ex:
                            fails.append((s.SheetNumber or uid[:8], unicode(ex)))
            except Exception as ex:
                output.print_md("# ❌ 依紅筆圖註解：寫入失敗，已自動復原")
                output.print_md("- {}".format(ex))
                return

            output.print_md("# 🖍 依紅筆圖註解")
            output.print_md("- ✍️ 已**寫入 Revit「修正備註」參數**：**{}** 張"
                            "（在圖紙屬性面板或圖紙明細表可看到；開圖也會彈出顯示）".format(ok_n))
            if fails:
                output.print_md("- ⚠ 失敗 {} 張".format(len(fails)))
                for uid, reason in fails[:15]:
                    output.print_md("  - `{}`: {}".format(uid[:8], reason))

            if push and ok_n and url and tab:
                # 先同步總表
                res, err, _ = win._do_export(url, tab)
                if err:
                    output.print_md("\n⚠ 同步總表失敗：{}".format(err))
                elif res and res.get("ok"):
                    output.print_md(
                        "\n🔄 已把含新修正備註的總表同步到 Google Sheet。")
                    # 再一併重建所有分表（讓分表也帶到新的修正備註）
                    sres, serr, _ = win._do_split(url, tab)
                    if serr:
                        output.print_md("\n⚠ 分表同步失敗：{}".format(serr))
                    elif sres and sres.get("ok"):
                        made = sres.get("splitTabs") or []
                        output.print_md(
                            "\n🗂️ 已一併更新所有分表（{} 個：{}）。".format(
                                len(made), u"、".join(made))
                            if made else
                            "\n🗂️ 已重建分表（目前沒有可依圖紙類別拆分的資料）。")
                    else:
                        output.print_md("\n⚠ 分表同步回應異常：{}".format(sres))
                else:
                    output.print_md("\n⚠ 同步回應異常：{}".format(res))
            return

        # 其他模式：處理完就結束
        return


if __name__ == '__main__':
    main()
