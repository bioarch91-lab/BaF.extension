# -*- coding: utf-8 -*-
"""
批次更改圖框 (Batch Change Title Block)
===========================================================
功能:
  一次更換多張圖紙最外圍的圖框 (Title Block)。
  
  涵蓋三種使用情境（同一個 UI 處理）：
  - 換成同一族群的另一個類型（例如 A1 換 A0）
  - 換成完全不同的圖框族群（例如事務所版換業主版）
  - 為原本沒有圖框的圖紙批次加上圖框

設計重點:
  - 左欄列出所有圖紙，副標顯示目前的圖框（包含「目前: 無」）
  - 上方有「目前圖框」過濾器，方便鎖定特定批次
  - 支援 Shift+點擊範圍勾選、表頭總勾選框
  - 右欄選擇要套用的目標圖框（兩階段：Family → Type）

作者: BaF / BIM 工具
"""

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
    TextBlock, Border, Orientation, ComboBox, TextBox, CheckBox
)
from System.Windows.Media import SolidColorBrush, Color

from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()


# ---------------------------------------------------------------------------
# 常數
# ---------------------------------------------------------------------------

NO_TITLEBLOCK = "（無圖框）"
ALL_FAMILIES = "（全部 Family）"


# ---------------------------------------------------------------------------
# Revit 資料準備
# ---------------------------------------------------------------------------

def get_title_block_types(document):
    """取得所有圖框類型，回傳 [{family, type_name, id, label}, ...]"""
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
            tn = t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if tn:
                type_name = tn.AsString() or ""
        except:
            pass
        if not type_name:
            try:
                type_name = t.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() or ""
            except:
                pass
        
        label = "{} - {}".format(family_name, type_name) if family_name else type_name
        result.append({
            "family": family_name,
            "type_name": type_name,
            "id": t.Id,
            "label": label,
        })
    
    return sorted(result, key=lambda x: (x["family"], x["type_name"]))


def get_title_block_instance(sheet, document):
    """取得圖紙上的圖框實例。回傳第一個找到的或 None。"""
    instances = DB.FilteredElementCollector(document, sheet.Id) \
        .OfCategory(DB.BuiltInCategory.OST_TitleBlocks) \
        .WhereElementIsNotElementType() \
        .ToElements()
    return list(instances)[0] if instances else None


def get_current_title_block_label(sheet, document):
    """圖紙目前圖框的描述字串，沒有則 NO_TITLEBLOCK。"""
    inst = get_title_block_instance(sheet, document)
    if inst is None:
        return NO_TITLEBLOCK
    try:
        type_elem = document.GetElement(inst.GetTypeId())
        family_name = type_elem.Family.Name
        type_name = ""
        try:
            tn = type_elem.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if tn:
                type_name = tn.AsString() or ""
        except:
            pass
        if not type_name:
            try:
                type_name = type_elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() or ""
            except:
                pass
        return "{} - {}".format(family_name, type_name)
    except:
        return "(unknown)"


def get_all_sheets(document):
    sheets = DB.FilteredElementCollector(document).OfClass(DB.ViewSheet).ToElements()
    sheets = [s for s in sheets if not s.IsPlaceholder]
    return sorted(sheets, key=lambda s: s.SheetNumber)


# ---------------------------------------------------------------------------
# WPF 視窗
# ---------------------------------------------------------------------------

class ChangeTitleBlockWindow(Window):
    
    COLOR_BG = (245, 245, 250)
    COLOR_BTN = (255, 255, 255)
    COLOR_PRIMARY = (99, 102, 241)
    COLOR_TEXT = (30, 30, 40)
    COLOR_TEXT_LIGHT = (255, 255, 255)
    COLOR_NO_TB = (200, 100, 100)  # 「無圖框」用紅色強調
    COLOR_GROUP_HEADER = (235, 238, 245)
    COLOR_HIGHLIGHT = (252, 248, 220)
    
    FILTER_ALL = "全部圖紙"
    
    def __init__(self, sheets, title_block_types):
        self.sheets = sheets
        self.title_block_types = title_block_types
        # 預先計算每張圖紙目前的圖框 label（避免反覆查詢）
        self.sheet_current_tb = {}
        for s in sheets:
            self.sheet_current_tb[s.Id.IntegerValue] = get_current_title_block_label(s, doc)
        
        # 收集「目前圖框」的所有可能值，用於過濾下拉
        self.current_tb_values = sorted(set(self.sheet_current_tb.values()))
        
        # Families 列表
        self.families = sorted(set(t["family"] for t in title_block_types if t["family"]))
        
        # 狀態
        self.sheet_rows = []  # [(border, checkbox, sheet, index)]
        self.last_clicked_index = None
        self._suppress_header_check = False
        self.search_text = ""
        self.filter_current_tb = self.FILTER_ALL
        
        # 結果
        self.result = None
        self.confirmed = False
        
        self._build_ui()
    
    @staticmethod
    def _brush(rgb):
        return SolidColorBrush(Color.FromRgb(rgb[0], rgb[1], rgb[2]))
    
    def _build_ui(self):
        self.Title = "批次更改圖框"
        self.Width = 1050
        self.Height = 720
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = self._brush(self.COLOR_BG)
        
        root = Grid()
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        
        # 頂部說明
        header = TextBlock()
        header.Text = "1. 在左邊勾選要更改圖框的圖紙   2. 在右邊選擇目標圖框   3. 按「執行」"
        header.FontSize = 13
        header.Foreground = self._brush(self.COLOR_TEXT)
        header.VerticalAlignment = VerticalAlignment.Center
        header.Margin = Thickness(20, 0, 0, 0)
        Grid.SetRow(header, 0)
        root.Children.Add(header)
        
        # 中間: 左 (圖紙) / 右 (目標圖框)
        mid = Grid()
        mid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        mid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(380)))
        
        left = self._build_sheet_column()
        Grid.SetColumn(left, 0)
        mid.Children.Add(left)
        
        right = self._build_target_column()
        Grid.SetColumn(right, 1)
        mid.Children.Add(right)
        
        Grid.SetRow(mid, 1)
        root.Children.Add(mid)
        
        # 底部
        bottom = Grid()
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(120)))
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(120)))
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(20)))
        
        self.status_text = TextBlock()
        self.status_text.Text = "尚未勾選任何圖紙"
        self.status_text.VerticalAlignment = VerticalAlignment.Center
        self.status_text.Margin = Thickness(20, 0, 0, 0)
        self.status_text.Foreground = self._brush(self.COLOR_TEXT)
        Grid.SetColumn(self.status_text, 0)
        bottom.Children.Add(self.status_text)
        
        cancel_btn = Button()
        cancel_btn.Content = "取消"
        cancel_btn.Margin = Thickness(0, 10, 5, 10)
        cancel_btn.Click += self._on_cancel
        Grid.SetColumn(cancel_btn, 1)
        bottom.Children.Add(cancel_btn)
        
        run_btn = Button()
        run_btn.Content = "執行"
        run_btn.Margin = Thickness(0, 10, 5, 10)
        run_btn.Background = self._brush(self.COLOR_PRIMARY)
        run_btn.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
        run_btn.FontWeight = FontWeights.Bold
        run_btn.Click += self._on_run
        Grid.SetColumn(run_btn, 2)
        bottom.Children.Add(run_btn)
        
        Grid.SetRow(bottom, 2)
        root.Children.Add(bottom)
        
        self.Content = root
    
    # ---- 左欄: 圖紙清單 ----
    
    def _build_sheet_column(self):
        outer = Grid()
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(115)))
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(28)))   # 表頭
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        outer.Margin = Thickness(20, 0, 10, 0)
        
        # 過濾區
        filter_panel = StackPanel()
        filter_panel.Orientation = Orientation.Vertical
        filter_panel.Margin = Thickness(0, 0, 0, 5)
        
        self.count_label = TextBlock()
        self.count_label.Text = "圖紙 ({})".format(len(self.sheets))
        self.count_label.FontWeight = FontWeights.Bold
        self.count_label.FontSize = 13
        self.count_label.Margin = Thickness(2, 0, 0, 4)
        self.count_label.Foreground = self._brush(self.COLOR_TEXT)
        filter_panel.Children.Add(self.count_label)
        
        # 「目前圖框」過濾
        tb_filter_label = TextBlock()
        tb_filter_label.Text = "目前圖框過濾"
        tb_filter_label.FontSize = 11
        tb_filter_label.Foreground = self._brush((110, 115, 130))
        tb_filter_label.Margin = Thickness(2, 0, 0, 2)
        filter_panel.Children.Add(tb_filter_label)
        
        self.tb_filter_combo = ComboBox()
        self.tb_filter_combo.Margin = Thickness(0, 0, 0, 4)
        self.tb_filter_combo.Items.Add(self.FILTER_ALL)
        for v in self.current_tb_values:
            self.tb_filter_combo.Items.Add(v)
        self.tb_filter_combo.SelectedIndex = 0
        self.tb_filter_combo.SelectionChanged += self._on_filter_changed
        filter_panel.Children.Add(self.tb_filter_combo)
        
        # 搜尋 + 全選
        toolbar = Grid()
        toolbar.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        toolbar.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(120)))
        toolbar.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(100)))
        
        self.search_box = TextBox()
        self.search_box.Padding = Thickness(4, 4, 4, 4)
        self.search_box.Margin = Thickness(0, 0, 5, 0)
        self.search_box.TextChanged += self._on_search_changed
        Grid.SetColumn(self.search_box, 0)
        toolbar.Children.Add(self.search_box)
        
        sel_btn = Button()
        sel_btn.Content = "全選 (顯示中)"
        sel_btn.Padding = Thickness(4, 4, 4, 4)
        sel_btn.Margin = Thickness(0, 0, 5, 0)
        sel_btn.Click += self._on_select_all_visible
        Grid.SetColumn(sel_btn, 1)
        toolbar.Children.Add(sel_btn)
        
        clr_btn = Button()
        clr_btn.Content = "全部取消"
        clr_btn.Padding = Thickness(4, 4, 4, 4)
        clr_btn.Click += self._on_clear_all
        Grid.SetColumn(clr_btn, 2)
        toolbar.Children.Add(clr_btn)
        
        filter_panel.Children.Add(toolbar)
        
        Grid.SetRow(filter_panel, 0)
        outer.Children.Add(filter_panel)
        
        # 表頭
        head = Grid()
        head.Margin = Thickness(10, 4, 10, 4)
        head.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(28)))
        head.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(180)))
        head.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        head.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(2, GridUnitType.Star)))
        
        # 表頭總勾選框
        self.header_check = CheckBox()
        self.header_check.IsThreeState = True
        self.header_check.VerticalAlignment = VerticalAlignment.Center
        self.header_check.HorizontalAlignment = HorizontalAlignment.Center
        self.header_check.Click += self._on_header_check_clicked
        Grid.SetColumn(self.header_check, 0)
        head.Children.Add(self.header_check)
        
        for col, text in [(1, "圖號"), (2, "圖名"), (3, "目前圖框")]:
            h = TextBlock()
            h.Text = text
            h.FontSize = 11
            h.FontWeight = FontWeights.Bold
            h.Foreground = self._brush(self.COLOR_TEXT)
            h.Margin = Thickness(4, 0, 4, 0)
            Grid.SetColumn(h, col)
            head.Children.Add(h)
        
        Grid.SetRow(head, 1)
        outer.Children.Add(head)
        
        # 列表
        scroll = ScrollViewer()
        scroll.Margin = Thickness(10, 0, 10, 8)
        self.list_panel = StackPanel()
        self.list_panel.Orientation = Orientation.Vertical
        scroll.Content = self.list_panel
        Grid.SetRow(scroll, 2)
        outer.Children.Add(scroll)
        
        for i, sheet in enumerate(self.sheets):
            row = self._make_sheet_row(sheet, i)
            self.list_panel.Children.Add(row)
        
        return outer
    
    def _make_sheet_row(self, sheet, index):
        border = Border()
        border.Background = self._brush(self.COLOR_BTN)
        border.BorderBrush = self._brush((220, 220, 230))
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(2)
        border.Margin = Thickness(0, 1, 0, 1)
        border.Padding = Thickness(4, 4, 4, 4)
        
        grid = Grid()
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(28)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(180)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(2, GridUnitType.Star)))
        
        cb = CheckBox()
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Margin = Thickness(4, 0, 0, 0)
        cb.Tag = index
        cb.PreviewMouseLeftButtonDown += self._on_row_check_mousedown
        cb.Checked += self._on_check_changed
        cb.Unchecked += self._on_check_changed
        Grid.SetColumn(cb, 0)
        grid.Children.Add(cb)
        
        num_tb = TextBlock()
        num_tb.Text = sheet.SheetNumber
        num_tb.FontSize = 12
        num_tb.FontWeight = FontWeights.SemiBold
        num_tb.VerticalAlignment = VerticalAlignment.Center
        num_tb.Margin = Thickness(4, 0, 4, 0)
        num_tb.TextTrimming = TextTrimming.CharacterEllipsis
        Grid.SetColumn(num_tb, 1)
        grid.Children.Add(num_tb)
        
        name_tb = TextBlock()
        name_tb.Text = sheet.Name
        name_tb.FontSize = 12
        name_tb.VerticalAlignment = VerticalAlignment.Center
        name_tb.Margin = Thickness(4, 0, 4, 0)
        name_tb.Foreground = self._brush((100, 105, 115))
        name_tb.TextTrimming = TextTrimming.CharacterEllipsis
        Grid.SetColumn(name_tb, 2)
        grid.Children.Add(name_tb)
        
        cur_tb_label = self.sheet_current_tb[sheet.Id.IntegerValue]
        cur_tb = TextBlock()
        cur_tb.Text = cur_tb_label
        cur_tb.FontSize = 11
        cur_tb.VerticalAlignment = VerticalAlignment.Center
        cur_tb.Margin = Thickness(4, 0, 4, 0)
        cur_tb.TextTrimming = TextTrimming.CharacterEllipsis
        if cur_tb_label == NO_TITLEBLOCK:
            cur_tb.Foreground = self._brush(self.COLOR_NO_TB)
        else:
            cur_tb.Foreground = self._brush((80, 85, 95))
        Grid.SetColumn(cur_tb, 3)
        grid.Children.Add(cur_tb)
        
        border.Child = grid
        # Tag: (checkbox, sheet, current_tb_label, index)
        border.Tag = (cb, sheet, cur_tb_label, index)
        self.sheet_rows.append(border)
        return border
    
    # ---- 右欄: 目標圖框選擇 ----
    
    def _build_target_column(self):
        outer = StackPanel()
        outer.Orientation = Orientation.Vertical
        outer.Margin = Thickness(10, 0, 20, 0)
        
        title = TextBlock()
        title.Text = "目標圖框"
        title.FontWeight = FontWeights.Bold
        title.FontSize = 14
        title.Margin = Thickness(0, 5, 0, 8)
        title.Foreground = self._brush(self.COLOR_TEXT)
        outer.Children.Add(title)
        
        # 提示
        hint = Border()
        hint.Background = self._brush((232, 240, 255))
        hint.BorderBrush = self._brush((99, 102, 241))
        hint.BorderThickness = Thickness(1)
        hint.CornerRadius = CornerRadius(4)
        hint.Padding = Thickness(10, 8, 10, 8)
        hint.Margin = Thickness(0, 0, 0, 14)
        
        hint_text = TextBlock()
        hint_text.Text = ("選擇圖框 Family 和 Type，點「執行」會把勾選圖紙的圖框替換成這個。\n"
                          "原本沒圖框的圖紙會直接套上指定的圖框。")
        hint_text.FontSize = 11
        hint_text.TextWrapping = TextWrapping.Wrap
        hint_text.Foreground = self._brush((40, 50, 80))
        hint.Child = hint_text
        outer.Children.Add(hint)
        
        # Family 下拉
        outer.Children.Add(self._make_label("Family（族群）"))
        self.family_combo = ComboBox()
        self.family_combo.Padding = Thickness(4, 4, 4, 4)
        self.family_combo.Margin = Thickness(0, 0, 0, 12)
        self.family_combo.Items.Add(ALL_FAMILIES)
        for f in self.families:
            self.family_combo.Items.Add(f)
        self.family_combo.SelectedIndex = 0
        self.family_combo.SelectionChanged += self._on_family_changed
        outer.Children.Add(self.family_combo)
        
        # Type 下拉
        outer.Children.Add(self._make_label("Type（類型）"))
        self.type_combo = ComboBox()
        self.type_combo.Padding = Thickness(4, 4, 4, 4)
        self.type_combo.Margin = Thickness(0, 0, 0, 12)
        self.type_combo.SelectionChanged += self._on_type_changed
        outer.Children.Add(self.type_combo)
        
        self._refresh_type_combo()
        
        # 預覽
        preview_box = Border()
        preview_box.Background = self._brush((255, 255, 255))
        preview_box.BorderBrush = self._brush((220, 222, 230))
        preview_box.BorderThickness = Thickness(1)
        preview_box.CornerRadius = CornerRadius(4)
        preview_box.Padding = Thickness(12, 10, 12, 10)
        preview_box.Margin = Thickness(0, 8, 0, 0)
        
        self.preview_text = TextBlock()
        self.preview_text.Text = "（請先選擇目標 Type）"
        self.preview_text.FontSize = 12
        self.preview_text.TextWrapping = TextWrapping.Wrap
        self.preview_text.Foreground = self._brush(self.COLOR_TEXT)
        preview_box.Child = self.preview_text
        outer.Children.Add(preview_box)
        
        return outer
    
    def _make_label(self, text):
        tb = TextBlock()
        tb.Text = text
        tb.FontWeight = FontWeights.SemiBold
        tb.FontSize = 12
        tb.Margin = Thickness(0, 0, 0, 4)
        tb.Foreground = self._brush(self.COLOR_TEXT)
        return tb
    
    def _refresh_type_combo(self):
        """根據目前選的 Family 重新填 Type 下拉。"""
        self.type_combo.Items.Clear()
        selected_family = self.family_combo.SelectedItem
        
        if selected_family == ALL_FAMILIES:
            candidates = self.title_block_types
        else:
            candidates = [t for t in self.title_block_types if t["family"] == selected_family]
        
        if not candidates:
            self.type_combo.Items.Add("（沒有可用類型）")
            self.type_combo.IsEnabled = False
            self.type_combo.SelectedIndex = 0
            return
        
        self.type_combo.IsEnabled = True
        for t in candidates:
            # 顯示時若 family 全部模式，加上 family 名稱避免重名混淆
            if selected_family == ALL_FAMILIES:
                self.type_combo.Items.Add(t["label"])
            else:
                self.type_combo.Items.Add(t["type_name"])
        self.type_combo.SelectedIndex = 0
    
    def _get_selected_target(self):
        """回傳目前選中的圖框類型 dict 或 None。"""
        selected_family = self.family_combo.SelectedItem
        idx = self.type_combo.SelectedIndex
        if idx < 0:
            return None
        
        if selected_family == ALL_FAMILIES:
            candidates = self.title_block_types
        else:
            candidates = [t for t in self.title_block_types if t["family"] == selected_family]
        
        if not candidates or idx >= len(candidates):
            return None
        return candidates[idx]
    
    def _update_preview(self):
        # 防呆：UI 還沒建立完就被 SelectionChanged 觸發到（建構階段）
        if not hasattr(self, "preview_text"):
            return
        target = self._get_selected_target()
        if target is None:
            self.preview_text.Text = "（請先選擇目標 Type）"
            return
        n_checked = sum(1 for b in self.sheet_rows if b.Tag[0].IsChecked)
        self.preview_text.Text = (
            "將套用：{}\n"
            "影響：{} 張勾選的圖紙"
        ).format(target["label"], n_checked)
    
    # ---- 事件處理 ----
    
    def _on_family_changed(self, sender, args):
        self._refresh_type_combo()
        self._update_preview()
    
    def _on_type_changed(self, sender, args):
        self._update_preview()
    
    def _on_filter_changed(self, sender, args):
        self.filter_current_tb = sender.SelectedItem or self.FILTER_ALL
        self._apply_filter()
    
    def _on_search_changed(self, sender, args):
        self.search_text = (sender.Text or "").lower()
        self._apply_filter()
    
    def _apply_filter(self):
        visible_count = 0
        for border in self.sheet_rows:
            cb, sheet, cur_tb, idx = border.Tag
            
            # 「目前圖框」過濾
            if self.filter_current_tb != self.FILTER_ALL:
                if cur_tb != self.filter_current_tb:
                    border.Visibility = Visibility.Collapsed
                    continue
            
            # 文字搜尋
            if self.search_text:
                combined = "{} {} {}".format(sheet.SheetNumber, sheet.Name, cur_tb).lower()
                if self.search_text not in combined:
                    border.Visibility = Visibility.Collapsed
                    continue
            
            border.Visibility = Visibility.Visible
            visible_count += 1
        
        self.count_label.Text = "圖紙 ({} / {})".format(visible_count, len(self.sheets))
        self._update_header_check()
    
    def _on_check_changed(self, sender, args):
        self._update_status()
        self._update_header_check()
        self._update_preview()
    
    def _on_row_check_mousedown(self, sender, args):
        try:
            from System.Windows.Input import Keyboard, ModifierKeys
            shift_held = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift
        except:
            shift_held = False
        
        clicked_idx = sender.Tag
        if shift_held and self.last_clicked_index is not None:
            lo = min(self.last_clicked_index, clicked_idx)
            hi = max(self.last_clicked_index, clicked_idx)
            target = not bool(sender.IsChecked)
            for border in self.sheet_rows:
                cb, _, _, idx = border.Tag
                if border.Visibility != Visibility.Visible:
                    continue
                if lo <= idx <= hi:
                    cb.IsChecked = target
            args.Handled = True
            self.last_clicked_index = clicked_idx
            self._update_header_check()
        else:
            self.last_clicked_index = clicked_idx
    
    def _update_header_check(self):
        if not hasattr(self, "header_check"):
            return
        visible_rows = [b for b in self.sheet_rows if b.Visibility == Visibility.Visible]
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
                self.header_check.IsChecked = None
        finally:
            self._suppress_header_check = False
    
    def _on_header_check_clicked(self, sender, args):
        if self._suppress_header_check:
            return
        target = bool(self.header_check.IsChecked)
        if self.header_check.IsChecked is None:
            target = True
            self._suppress_header_check = True
            try:
                self.header_check.IsChecked = True
            finally:
                self._suppress_header_check = False
        for border in self.sheet_rows:
            if border.Visibility == Visibility.Visible:
                border.Tag[0].IsChecked = target
    
    def _on_select_all_visible(self, sender, args):
        for border in self.sheet_rows:
            if border.Visibility == Visibility.Visible:
                border.Tag[0].IsChecked = True
    
    def _on_clear_all(self, sender, args):
        for border in self.sheet_rows:
            border.Tag[0].IsChecked = False
    
    def _update_status(self):
        if not hasattr(self, "status_text"):
            return
        n = sum(1 for b in self.sheet_rows if b.Tag[0].IsChecked)
        if n == 0:
            self.status_text.Text = "尚未勾選任何圖紙"
        else:
            self.status_text.Text = "已勾選 {} 張圖紙".format(n)
    
    def _on_run(self, sender, args):
        target = self._get_selected_target()
        if target is None:
            forms.alert("請先選擇目標圖框類型。")
            return
        
        checked_sheets = [b.Tag[1] for b in self.sheet_rows if b.Tag[0].IsChecked]
        if not checked_sheets:
            forms.alert("請至少勾選一張圖紙。")
            return
        
        self.result = {
            "sheets": checked_sheets,
            "target_type_id": target["id"],
            "target_label": target["label"],
        }
        self.confirmed = True
        self.Close()
    
    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.Close()


# ---------------------------------------------------------------------------
# 執行替換
# ---------------------------------------------------------------------------

def change_title_blocks(result, document):
    """把勾選圖紙的圖框實例替換成目標類型。沒圖框的圖紙會新加一個。"""
    target_type_id = result["target_type_id"]
    target_type = document.GetElement(target_type_id)
    sheets = result["sheets"]
    
    success_replaced = 0   # 既有圖框換成新的
    success_added = 0      # 新加圖框
    failed = []
    
    # 確保目標 family symbol 已啟用（不啟用無法用 NewFamilyInstance）
    if not target_type.IsActive:
        try:
            with revit.Transaction("Activate Title Block Symbol"):
                target_type.Activate()
                document.Regenerate()
        except Exception as ex:
            failed.append(("(全部)", "啟用目標圖框失敗: {}".format(ex)))
            return success_replaced, success_added, failed
    
    with revit.Transaction("Batch Change Title Blocks"):
        for sheet in sheets:
            try:
                inst = get_title_block_instance(sheet, document)
                
                if inst is None:
                    # 沒圖框 → 新加一個
                    # 預設放在原點 (0,0)，使用者之後可移動
                    new_inst = document.Create.NewFamilyInstance(
                        DB.XYZ(0, 0, 0),
                        target_type,
                        sheet
                    )
                    success_added += 1
                else:
                    current_type_id = inst.GetTypeId()
                    if current_type_id == target_type_id:
                        # 已經是目標類型，跳過
                        continue
                    # 直接切換 type，位置會保留
                    inst.ChangeTypeId(target_type_id)
                    success_replaced += 1
            except Exception as ex:
                failed.append((sheet.SheetNumber, str(ex)))
    
    return success_replaced, success_added, failed


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main():
    sheets = get_all_sheets(doc)
    if not sheets:
        forms.alert("專案中沒有圖紙。", exitscript=True)
    
    title_block_types = get_title_block_types(doc)
    if not title_block_types:
        forms.alert("專案中沒有任何圖框類型。請先載入圖框 Family。", exitscript=True)
    
    win = ChangeTitleBlockWindow(sheets, title_block_types)
    win.ShowDialog()
    
    if not win.confirmed:
        script.exit()
    
    replaced, added, failed = change_title_blocks(win.result, doc)
    
    output.print_md("# ✅ 批次更改圖框完成")
    output.print_md("- 套用目標：**{}**".format(win.result["target_label"]))
    output.print_md("- 替換既有圖框：**{}** 張".format(replaced))
    output.print_md("- 為空白圖紙加圖框：**{}** 張".format(added))
    output.print_md("- 失敗：**{}** 張".format(len(failed)))
    
    if failed:
        output.print_md("\n## 失敗清單")
        for sheet_num, reason in failed:
            output.print_md("- `{}`: {}".format(sheet_num, reason))


if __name__ == '__main__':
    main()
