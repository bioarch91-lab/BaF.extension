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
    CheckBox, ScrollBarVisibility
)
from System.Windows.Media import SolidColorBrush, Color, FontFamily

from pyrevit import revit, DB, forms, script

import os
import json

doc = revit.doc
output = script.get_output()

# Google Sheet 寫入設定（存在本機，不進 git，避免 URL 外流）
GSHEET_CFG = os.path.join(os.getenv("APPDATA") or u"", "pyRevit", "baf_gsheet_config.json")


def load_gsheet_cfg():
    try:
        with open(GSHEET_CFG, "r") as f:
            return json.load(f)
    except Exception:
        return {}


def save_gsheet_cfg(cfg):
    try:
        d = os.path.dirname(GSHEET_CFG)
        if d and not os.path.exists(d):
            os.makedirs(d)
        with open(GSHEET_CFG, "w") as f:
            json.dump(cfg, f)
        return True
    except Exception:
        return False


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

        minutes_btn = Button()
        minutes_btn.Content = u"👁 依會議紀錄編輯圖紙清單"
        minutes_btn.Padding = Thickness(12, 6, 12, 6)
        minutes_btn.FontWeight = FontWeights.Bold
        minutes_btn.Click += self._on_edit_by_minutes
        wbtn_row.Children.Add(minutes_btn)

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
        p = sheet.LookupParameter(pname)
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
                state = u"預留"
            else:
                state = u"真實"
            row = [s.UniqueId[:8], state, s.SheetNumber or u"", s.Name or u""]
            for pn in params:
                v = self._read_text_param(s, pn)
                if v is None:
                    row.append(u"⚠ 無此參數")
                else:
                    row.append(v)  # 空值就留白
            rows.append(row)
        summary = self._make_cell(
            u"共 {} 張圖紙　（真實 {} ／ 預留 {}）".format(
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
        # 第二步：選擇是否依「圖紙類別」拆成多個工作表
        if forms.alert(u"匯出完成。\n\n要再依『圖紙類別』拆成多個工作表嗎？\n"
                       u"（每個類別一個分頁、以類別命名）",
                       yes=True, no=True):
            sres, serr, _ = self._do_split(url, tab)
            if serr:
                msg += u"\n\n⚠ 拆分失敗：{}".format(serr)
            elif sres and sres.get("ok"):
                made = sres.get("splitTabs") or []
                msg += (u"\n\n🗂️ 已拆出 {} 個分表：{}".format(len(made), u"、".join(made))
                        if made else u"\n\n（沒有可拆分的圖紙類別）")
            else:
                msg += u"\n\n⚠ 拆分回應異常：{}".format(sres)
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
        dup_msg = self._check_dup_numbers(records)
        if dup_msg:
            forms.alert(dup_msg)
            return
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
        for label, val in ((u"圖紙類別", d["cat"]), (u"繪圖員", d["drawer"]),
                           (u"修正備註", d["note"])):
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
            tag = u"🆕 新增" + (u"(真實)" if d["real"] else u"(預留)")
            extra = u"; ".join([u"{}={}".format(k, v)
                                for k, v in d.get("extra", {}).items() if v])
            rows.append([tag, d["num"], d["name"], extra])
        for d in plan["toReal"]:
            s = d["sheet"]
            rows.append([u"⬆️ 轉真實", d["num"] or s.SheetNumber,
                         d["name"] or s.Name, self._edit_detail(d)])
        for d in plan["edit"]:
            s = d["sheet"]
            rows.append([u"✏️ 修改", d["num"] or s.SheetNumber,
                         d["name"] or s.Name, self._edit_detail(d)])
        for d in plan["toPlace"]:
            s = d["sheet"]
            rows.append([u"⚠️ 真實→預留", s.SheetNumber, s.Name, u"刪原圖建預留"])
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
            u"新增 {}／修改 {}／轉真實 {}／降預留刪除 {}／刪除 {}".format(
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
        p = sheet.LookupParameter(pname)
        if p is None or p.IsReadOnly:
            return
        try:
            if p.StorageType == DB.StorageType.String:
                p.Set(value if value is not None else u"")
        except Exception:
            pass

    def _needs_edit(self, s, d):
        # 空白 = 不變更（與「編輯既有」一致，不會清掉既有資料）
        if d["num"] and d["num"] != (s.SheetNumber or u""):
            return True
        if d["name"] and d["name"] != (s.Name or u""):
            return True
        for label, val in ((u"圖紙類別", d["cat"]), (u"繪圖員", d["drawer"]),
                           (u"修正備註", d["note"])):
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
                "note": unicode(rec.get(u"修正備註") or u""),
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

    def _on_import_execute(self, sender, args):
        plan = getattr(self, "_pending_plan", None)
        if not plan:
            forms.alert(u"請先按「匯入（預覽差異）」載入要套用的內容。")
            return
        # 刪除前清單確認（你的要求）
        del_lines = [u"  {}  {}".format(d["sheet"].SheetNumber, d["sheet"].Name)
                     for d in plan["toPlace"]]
        del_lines += [u"  {}  {}".format(d["sheet"].SheetNumber, d["sheet"].Name)
                      for d in plan["delete"]]
        if del_lines:
            msg = u"以下圖紙將被刪除，確定執行嗎？\n\n" + u"\n".join(del_lines[:30])
            if len(del_lines) > 30:
                msg += u"\n  …還有 {} 張".format(len(del_lines) - 30)
            if not forms.alert(msg, yes=True, no=True):
                return
        # 不在 modal 視窗內改模型(會被還原)；關窗後由 main() 執行交易
        self.result = {"mode": "import_apply", "plan": plan}
        self.confirmed = True
        self.Close()

    def _on_assign_tasks(self, sender, args):
        forms.alert(u"「指派工作任務」功能開發中。\n\n"
                    u"未來按下後，會把『修正備註』欄裡的工作，\n"
                    u"分配給對應的『繪圖員』同事。\n（分配方式待討論）")

    def _on_build_env(self, sender, args):
        forms.alert(u"「建立環境」功能開發中（佔位）。\n\n"
                    u"規劃：一鍵自動在 Revit 建立同步所需的圖紙文字參數\n"
                    u"（圖紙類別/繪圖員/修正備註/審圖員/設計者/批准者/圖紙發布日期/校核2），\n"
                    u"已存在的略過，並提示 Google 端設定。")

    def _on_edit_by_minutes(self, sender, args):
        forms.alert(u"「依會議紀錄編輯圖紙清單」功能開發中（佔位）。\n\n"
                    u"規劃：貼上/讀取會議紀錄 → 用 AI 抓出每張圖的待辦 →\n"
                    u"預覽 → 寫進對應圖紙的『修正備註』。\n（需接 Claude API）")

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
                        fails.append((d.get("num"), u"轉真實:" + unicode(ex)))

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
                        fails.append((d.get("num"), u"降預留:" + unicode(ex)))

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
                        if d["note"]:
                            self._set_param(sheet, u"修正備註", d["note"])
                        for k, v in d.get("extra", {}).items():
                            if v:
                                self._set_param(sheet, k, v)
                    except Exception as ex:
                        fails.append((d.get("num"), u"套用欄位:" + unicode(ex)))
        except Exception as ex:
            return u"❌ 套用失敗，已自動復原：{}".format(ex), True, done

        report = (u"🆕 新增 {} ／ ✏️ 修改 {} ／ ⬆️ 轉真實 {} ／ "
                  u"⚠️ 降預留 {} ／ 🗑️ 刪除 {}".format(
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
# 主流程
# ---------------------------------------------------------------------------

def main():
    # 開新視窗前，先關掉上一個還開著的（用 AppDomain 跨次記住）
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
    AppDomain.CurrentDomain.SetData(_WKEY, win)
    win.ShowDialog()
    # 只有當記錄的還是自己時才清掉（避免清掉後開的新視窗）
    if AppDomain.CurrentDomain.GetData(_WKEY) is win:
        AppDomain.CurrentDomain.SetData(_WKEY, None)

    if not win.confirmed:
        script.exit()
    
    if win.result["mode"] == "create":
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
    
    elif win.result["mode"] == "edit":
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

    elif win.result["mode"] == "import_apply":
        plan = win.result["plan"]
        # 在 modal 視窗關閉「之後」才改模型，交易才不會被還原
        report, is_err, done = win._apply_plan(plan)
        output.print_md("# {} 匯入套用".format(u"❌" if is_err else u"✅"))
        output.print_md(report)
        role = plan.get("role") or "main"
        main_tab = plan.get("mainTab") or plan.get("tab")
        if not is_err and role == "sub":
            # 分表匯入後：自動回寫總表 + 重新拆分，讓全部對齊（你的需求3）
            r1, e1, _ = win._do_export(plan["url"], main_tab)
            r2, e2, _ = win._do_split(plan["url"], main_tab)
            if e1 or e2:
                output.print_md("\n⚠ 自動同步總表/分表時有狀況：{} {}".format(e1 or "", e2 or ""))
            else:
                output.print_md("\n🔄 已自動回寫總表並重新拆分各分表（全部對齊）。")
        elif not is_err and (done.get("new", 0) > 0 or done.get("toPlace", 0) > 0):
            # 總表匯入：有新增/刪除重建 → UID 會變動，自動回寫
            res, err, n = win._do_export(plan["url"], main_tab)
            if err:
                output.print_md("\n⚠ 自動回寫 Google Sheet 失敗：{}".format(err))
            elif res and res.get("ok"):
                output.print_md("\n🔄 已自動回寫 Google Sheet（更新變動/新建的 UID）。")
            else:
                output.print_md("\n⚠ 自動回寫回應異常：{}".format(res))


if __name__ == '__main__':
    main()
