# -*- coding: utf-8 -*-
"""
批次新增 / 編輯圖紙 (Batch Manage Sheets)
===========================================================
四種操作模式（Tab）：
  1. 編輯既有：勾選想處理的既有圖紙，輸入新的圖號/圖名（空白 = 不變更）
  2. 逐筆輸入：UI 上一行一行填寫圖號/圖名 → 新增圖紙
  3. 貼上文字：從 Excel/記事本複製貼上 → 新增圖紙
  4. 規則生成：例如 A2-{n:02d} 1~20 → 新增 20 張圖紙

「編輯既有」也支援規則式批次重新編號。
作者: BaF / BIM 工具
"""

import re
import clr
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")

from System.Windows import (
    Window, Thickness, HorizontalAlignment, VerticalAlignment,
    WindowStartupLocation, Visibility, TextTrimming, FontWeights,
    GridLength, GridUnitType, CornerRadius, TextWrapping
)
from System.Windows.Controls import (
    StackPanel, Button, ScrollViewer, Grid, RowDefinition, ColumnDefinition,
    TextBlock, Border, Orientation, ComboBox, TextBox, TabControl, TabItem,
    CheckBox, ScrollBarVisibility
)
from System.Windows.Media import SolidColorBrush, Color, FontFamily

from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()


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
    
    def __init__(self, title_blocks, existing_sheets):
        self.title_blocks = title_blocks
        self.existing_sheets = existing_sheets
        self.existing_numbers = set(s.SheetNumber for s in existing_sheets)
        
        self.result = None
        self.confirmed = False
        
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
        self.Title = "批次新增 / 編輯圖紙"
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
        
        # 中間：4 個 Tab
        self.tabs = TabControl()
        self.tabs.Margin = Thickness(20, 5, 20, 5)
        self.tabs.Items.Add(self._build_edit_existing_tab())
        self.tabs.Items.Add(self._build_table_tab())
        self.tabs.Items.Add(self._build_paste_tab())
        self.tabs.Items.Add(self._build_pattern_tab())
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
    
    # ---- 執行 ----
    
    def _on_run(self, sender, args):
        idx = self.tabs.SelectedIndex
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
    title_blocks = get_title_block_types(doc)
    existing_sheets = get_existing_sheets(doc)
    
    win = BatchManageSheetsWindow(title_blocks, existing_sheets)
    win.ShowDialog()
    
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


if __name__ == '__main__':
    main()
