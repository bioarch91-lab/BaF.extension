# -*- coding: utf-8 -*-
"""
視圖配對上圖 (Place Views to Sheets - Drag-to-Match)
===========================================================
功能:
  以「連連看」UI 讓使用者把視圖配對到圖紙上，一次大量處理。

操作流程:
  1. 用上方搜尋框/下拉過濾左右兩欄
  2. 點左欄一個視圖
  3. 點右欄一個圖紙 → 自動連線完成配對
  4. 想取消某條連線：再點一次該視圖或該圖紙
  5. 全部配對好後按「執行」→ 視圖會被放置到圖紙中央

作者: BaF / BIM 工具
"""

import clr
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")

from System.Windows import (
    Window, Thickness, HorizontalAlignment, VerticalAlignment,
    WindowStartupLocation, Visibility, TextTrimming, FontWeights, Point,
    GridLength, GridUnitType, CornerRadius
)
from System.Windows.Controls import (
    StackPanel, Button, ScrollViewer, Grid, RowDefinition, ColumnDefinition,
    TextBlock, Border, Canvas, Orientation, ComboBox, ComboBoxItem, TextBox
)
from System.Windows.Media import (
    SolidColorBrush, Color, Brushes, PenLineCap
)
from System.Windows.Shapes import Line

from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()


# ---------------------------------------------------------------------------
# Revit 資料準備
# ---------------------------------------------------------------------------

def get_view_type_name(view):
    """取得 Revit 視圖類型名稱（例如「出圖 1/200」）。"""
    try:
        vt = doc.GetElement(view.GetTypeId())
        if vt:
            tn = vt.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            if tn:
                return tn
    except:
        pass
    return str(view.ViewType)


def get_view_family_name(view):
    """取得 Revit 視圖大分類（中文化），對應專案瀏覽器頂層分類。"""
    family_map = {
        DB.ViewType.FloorPlan: "樓板平面圖",
        DB.ViewType.CeilingPlan: "天花板平面圖",
        DB.ViewType.Elevation: "立面圖",
        DB.ViewType.Section: "剖面",
        DB.ViewType.ThreeD: "3D 視圖",
        DB.ViewType.Detail: "詳圖",
        DB.ViewType.DraftingView: "圖說",
        DB.ViewType.Legend: "圖例",
        DB.ViewType.Schedule: "明細表",
        DB.ViewType.Rendering: "彩現",
        DB.ViewType.AreaPlan: "面積圖",
        DB.ViewType.EngineeringPlan: "結構平面",
        DB.ViewType.Walkthrough: "漫遊",
    }
    return family_map.get(view.ViewType, str(view.ViewType))


def get_view_group_label(view):
    """組合「大分類（類型名）」當分組標籤，跟 Revit 專案瀏覽器一致。"""
    return "{} ({})".format(get_view_family_name(view), get_view_type_name(view))


def get_placeable_views(document):
    """取得所有「尚未放置在任何圖紙上」的視圖。"""
    views = DB.FilteredElementCollector(document).OfClass(DB.View).ToElements()
    
    placed_view_ids = set()
    sheets = DB.FilteredElementCollector(document).OfClass(DB.ViewSheet).ToElements()
    for sheet in sheets:
        if sheet.IsPlaceholder:
            continue
        for vp_id in sheet.GetAllViewports():
            vp = document.GetElement(vp_id)
            placed_view_ids.add(vp.ViewId.IntegerValue)
    
    excluded_types = (
        DB.ViewType.SystemBrowser, DB.ViewType.ProjectBrowser,
        DB.ViewType.Internal, DB.ViewType.Undefined
    )
    
    result = []
    for v in views:
        if v.IsTemplate:
            continue
        if isinstance(v, DB.ViewSheet):
            continue
        if v.ViewType in excluded_types:
            continue
        if v.Id.IntegerValue in placed_view_ids:
            continue
        result.append(v)
    
    return sorted(result, key=lambda x: (get_view_group_label(x), x.Name))


def get_all_sheets(document):
    sheets = DB.FilteredElementCollector(document).OfClass(DB.ViewSheet).ToElements()
    sheets = [s for s in sheets if not s.IsPlaceholder]
    return sorted(sheets, key=lambda s: s.SheetNumber)


# ---------------------------------------------------------------------------
# 連連看 WPF 視窗
# ---------------------------------------------------------------------------

class MatchWindow(Window):
    
    COLOR_BG = (245, 245, 250)
    COLOR_BTN = (255, 255, 255)
    COLOR_SELECTED = (99, 102, 241)
    COLOR_MATCHED = (167, 139, 250)
    COLOR_LINE = (139, 92, 246)
    COLOR_TEXT = (30, 30, 40)
    COLOR_TEXT_LIGHT = (255, 255, 255)
    COLOR_GROUP_HEADER = (235, 238, 245)
    COLOR_GROUP_TEXT = (90, 95, 110)
    
    SHEET_FILTER_ALL = "全部圖紙"
    SHEET_FILTER_EMPTY = "只看空圖紙"
    SHEET_FILTER_HAS_VIEWS = "只看已有視圖"
    
    VIEW_GROUP_ALL = "全部分類"
    
    def __init__(self, views, sheets):
        self.views = views
        self.sheets = sheets
        
        # 蒐集分組標籤（保持順序）
        self.view_groups = []
        seen = set()
        for v in views:
            label = get_view_group_label(v)
            if label not in seen:
                seen.add(label)
                self.view_groups.append(label)
        
        self.view_buttons = {}
        self.sheet_buttons = {}
        self.matches = {}
        self.lines = {}
        self.selected_view_id = None
        self.selected_sheet_id = None
        self.confirmed = False
        self.view_group_headers = {}
        
        self.view_search_text = ""
        self.view_group_filter = self.VIEW_GROUP_ALL
        self.sheet_search_text = ""
        self.sheet_filter = self.SHEET_FILTER_ALL
        
        self._build_ui()
    
    @staticmethod
    def _brush(rgb):
        return SolidColorBrush(Color.FromRgb(rgb[0], rgb[1], rgb[2]))
    
    def _build_ui(self):
        self.Title = "視圖配對上圖 - 連連看"
        self.Width = 1100
        self.Height = 720
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = self._brush(self.COLOR_BG)
        
        root = Grid()
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        
        header = TextBlock()
        header.Text = "點選左邊的視圖、再點右邊的圖紙即可配對。再點一次可取消。"
        header.FontSize = 13
        header.Foreground = self._brush(self.COLOR_TEXT)
        header.VerticalAlignment = VerticalAlignment.Center
        header.Margin = Thickness(20, 0, 0, 0)
        Grid.SetRow(header, 0)
        root.Children.Add(header)
        
        mid = Grid()
        mid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        mid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(120)))
        mid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        left_col = self._build_view_column()
        Grid.SetColumn(left_col, 0)
        mid.Children.Add(left_col)
        
        self.canvas = Canvas()
        self.canvas.Background = self._brush(self.COLOR_BG)
        Grid.SetColumn(self.canvas, 1)
        mid.Children.Add(self.canvas)
        
        right_col = self._build_sheet_column()
        Grid.SetColumn(right_col, 2)
        mid.Children.Add(right_col)
        
        Grid.SetRow(mid, 1)
        root.Children.Add(mid)
        
        bottom = Grid()
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(120)))
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(120)))
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(20)))
        
        self.status_text = TextBlock()
        self.status_text.Text = "尚未配對"
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
        run_btn.Background = self._brush(self.COLOR_SELECTED)
        run_btn.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
        run_btn.FontWeight = FontWeights.Bold
        run_btn.Click += self._on_run
        Grid.SetColumn(run_btn, 2)
        bottom.Children.Add(run_btn)
        
        Grid.SetRow(bottom, 2)
        root.Children.Add(bottom)
        
        self.Content = root
        
        self.SizeChanged += lambda s, e: self._redraw_all_lines()
        self.Loaded += lambda s, e: self._redraw_all_lines()
    
    def _build_view_column(self):
        outer = Grid()
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(80)))
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        outer.Margin = Thickness(20, 0, 5, 0)
        
        filter_panel = StackPanel()
        filter_panel.Orientation = Orientation.Vertical
        filter_panel.Margin = Thickness(0, 0, 0, 5)
        
        self.view_count_label = TextBlock()
        self.view_count_label.Text = "視圖 ({})".format(len(self.views))
        self.view_count_label.FontWeight = FontWeights.Bold
        self.view_count_label.FontSize = 13
        self.view_count_label.Margin = Thickness(2, 0, 0, 4)
        self.view_count_label.Foreground = self._brush(self.COLOR_TEXT)
        filter_panel.Children.Add(self.view_count_label)
        
        self.view_group_combo = ComboBox()
        self.view_group_combo.Margin = Thickness(0, 0, 0, 4)
        self.view_group_combo.Items.Add(self.VIEW_GROUP_ALL)
        for g in self.view_groups:
            self.view_group_combo.Items.Add(g)
        self.view_group_combo.SelectedIndex = 0
        self.view_group_combo.SelectionChanged += self._on_view_group_changed
        filter_panel.Children.Add(self.view_group_combo)
        
        self.view_search_box = TextBox()
        self.view_search_box.Padding = Thickness(4, 3, 4, 3)
        self.view_search_box.TextChanged += self._on_view_search_changed
        filter_panel.Children.Add(self.view_search_box)
        
        Grid.SetRow(filter_panel, 0)
        outer.Children.Add(filter_panel)
        
        self.view_list_panel = StackPanel()
        self.view_list_panel.Orientation = Orientation.Vertical
        
        current_group = None
        for v in self.views:
            label = get_view_group_label(v)
            if label != current_group:
                hdr = self._make_group_header(label)
                self.view_group_headers[label] = hdr
                self.view_list_panel.Children.Add(hdr)
                current_group = label
            
            btn = self._make_item_button(
                v.Name,
                self._view_subtitle(v),
                lambda s, e, vid=v.Id.IntegerValue: self._on_view_clicked(vid)
            )
            t1, t2 = btn.Tag
            btn.Tag = (t1, t2, label, v.Name)
            self.view_buttons[v.Id.IntegerValue] = btn
            self.view_list_panel.Children.Add(btn)
        
        scroll = ScrollViewer()
        scroll.Content = self.view_list_panel
        scroll.ScrollChanged += lambda s, e: self._redraw_all_lines()
        self.view_scroll = scroll
        Grid.SetRow(scroll, 1)
        outer.Children.Add(scroll)
        
        return outer
    
    def _build_sheet_column(self):
        outer = Grid()
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(80)))
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        outer.Margin = Thickness(5, 0, 20, 0)
        
        filter_panel = StackPanel()
        filter_panel.Orientation = Orientation.Vertical
        filter_panel.Margin = Thickness(0, 0, 0, 5)
        
        self.sheet_count_label = TextBlock()
        self.sheet_count_label.Text = "圖紙 ({})".format(len(self.sheets))
        self.sheet_count_label.FontWeight = FontWeights.Bold
        self.sheet_count_label.FontSize = 13
        self.sheet_count_label.Margin = Thickness(2, 0, 0, 4)
        self.sheet_count_label.Foreground = self._brush(self.COLOR_TEXT)
        filter_panel.Children.Add(self.sheet_count_label)
        
        self.sheet_filter_combo = ComboBox()
        self.sheet_filter_combo.Margin = Thickness(0, 0, 0, 4)
        for opt in [self.SHEET_FILTER_ALL, self.SHEET_FILTER_EMPTY, self.SHEET_FILTER_HAS_VIEWS]:
            self.sheet_filter_combo.Items.Add(opt)
        self.sheet_filter_combo.SelectedIndex = 0
        self.sheet_filter_combo.SelectionChanged += self._on_sheet_filter_changed
        filter_panel.Children.Add(self.sheet_filter_combo)
        
        self.sheet_search_box = TextBox()
        self.sheet_search_box.Padding = Thickness(4, 3, 4, 3)
        self.sheet_search_box.TextChanged += self._on_sheet_search_changed
        filter_panel.Children.Add(self.sheet_search_box)
        
        Grid.SetRow(filter_panel, 0)
        outer.Children.Add(filter_panel)
        
        self.sheet_list_panel = StackPanel()
        self.sheet_list_panel.Orientation = Orientation.Vertical
        
        for s in self.sheets:
            vp_count = len(s.GetAllViewports())
            subtitle = "空白圖紙" if vp_count == 0 else "已有 {} 個視圖".format(vp_count)
            btn = self._make_item_button(
                "{} - {}".format(s.SheetNumber, s.Name),
                subtitle,
                lambda sender, e, sid=s.Id.IntegerValue: self._on_sheet_clicked(sid)
            )
            t1, t2 = btn.Tag
            btn.Tag = (t1, t2, s.SheetNumber, s.Name, vp_count)
            self.sheet_buttons[s.Id.IntegerValue] = btn
            self.sheet_list_panel.Children.Add(btn)
        
        scroll = ScrollViewer()
        scroll.Content = self.sheet_list_panel
        scroll.ScrollChanged += lambda s, e: self._redraw_all_lines()
        self.sheet_scroll = scroll
        Grid.SetRow(scroll, 1)
        outer.Children.Add(scroll)
        
        return outer
    
    def _make_group_header(self, label):
        border = Border()
        border.Background = self._brush(self.COLOR_GROUP_HEADER)
        border.Padding = Thickness(8, 4, 8, 4)
        border.Margin = Thickness(0, 6, 0, 2)
        border.CornerRadius = CornerRadius(3)
        
        tb = TextBlock()
        tb.Text = label
        tb.FontSize = 11
        tb.FontWeight = FontWeights.SemiBold
        tb.Foreground = self._brush(self.COLOR_GROUP_TEXT)
        border.Child = tb
        return border
    
    def _view_subtitle(self, view):
        type_name = get_view_type_name(view)
        return "{} · ID {}".format(type_name, view.Id.IntegerValue)
    
    def _make_item_button(self, title, subtitle, on_click):
        btn = Button()
        btn.HorizontalContentAlignment = HorizontalAlignment.Stretch
        btn.Margin = Thickness(0, 1, 0, 1)
        btn.Padding = Thickness(10, 6, 10, 6)
        btn.Background = self._brush(self.COLOR_BTN)
        btn.BorderBrush = self._brush((220, 220, 230))
        btn.BorderThickness = Thickness(1)
        
        sp = StackPanel()
        sp.Orientation = Orientation.Vertical
        
        t1 = TextBlock()
        t1.Text = title
        t1.FontSize = 12
        t1.FontWeight = FontWeights.SemiBold
        t1.Foreground = self._brush(self.COLOR_TEXT)
        t1.TextTrimming = TextTrimming.CharacterEllipsis
        sp.Children.Add(t1)
        
        t2 = TextBlock()
        t2.Text = subtitle
        t2.FontSize = 10
        t2.Foreground = self._brush((120, 120, 130))
        t2.Margin = Thickness(0, 2, 0, 0)
        sp.Children.Add(t2)
        
        btn.Content = sp
        btn.Click += on_click
        btn.Tag = (t1, t2)
        return btn
    
    # -- 過濾邏輯 --
    
    def _on_view_group_changed(self, sender, args):
        self.view_group_filter = sender.SelectedItem or self.VIEW_GROUP_ALL
        self._apply_view_filter()
    
    def _on_view_search_changed(self, sender, args):
        self.view_search_text = (sender.Text or "").lower()
        self._apply_view_filter()
    
    def _on_sheet_filter_changed(self, sender, args):
        self.sheet_filter = sender.SelectedItem or self.SHEET_FILTER_ALL
        self._apply_sheet_filter()
    
    def _on_sheet_search_changed(self, sender, args):
        self.sheet_search_text = (sender.Text or "").lower()
        self._apply_sheet_filter()
    
    def _apply_view_filter(self):
        visible_count = 0
        group_has_visible = {g: False for g in self.view_groups}
        
        for vid, btn in self.view_buttons.items():
            _, _, group_label, view_name = btn.Tag
            
            if self.view_group_filter != self.VIEW_GROUP_ALL:
                if group_label != self.view_group_filter:
                    btn.Visibility = Visibility.Collapsed
                    continue
            
            if self.view_search_text:
                if self.view_search_text not in view_name.lower() \
                   and self.view_search_text not in group_label.lower():
                    btn.Visibility = Visibility.Collapsed
                    continue
            
            btn.Visibility = Visibility.Visible
            group_has_visible[group_label] = True
            visible_count += 1
        
        for label, hdr in self.view_group_headers.items():
            hdr.Visibility = Visibility.Visible if group_has_visible.get(label) else Visibility.Collapsed
        
        self.view_count_label.Text = "視圖 ({} / {})".format(visible_count, len(self.views))
        self._redraw_all_lines()
    
    def _apply_sheet_filter(self):
        visible_count = 0
        for sid, btn in self.sheet_buttons.items():
            _, _, sheet_num, sheet_name, vp_count = btn.Tag
            
            if self.sheet_filter == self.SHEET_FILTER_EMPTY and vp_count > 0:
                btn.Visibility = Visibility.Collapsed
                continue
            if self.sheet_filter == self.SHEET_FILTER_HAS_VIEWS and vp_count == 0:
                btn.Visibility = Visibility.Collapsed
                continue
            
            if self.sheet_search_text:
                combined = "{} {}".format(sheet_num, sheet_name).lower()
                if self.sheet_search_text not in combined:
                    btn.Visibility = Visibility.Collapsed
                    continue
            
            btn.Visibility = Visibility.Visible
            visible_count += 1
        
        self.sheet_count_label.Text = "圖紙 ({} / {})".format(visible_count, len(self.sheets))
        self._redraw_all_lines()
    
    # -- 點擊邏輯 --
    
    def _on_view_clicked(self, view_id):
        if view_id in self.matches:
            self._unmatch(view_id)
            return
        if self.selected_view_id == view_id:
            self.selected_view_id = None
        else:
            self.selected_view_id = view_id
        self._try_make_match()
        self._refresh_button_styles()
    
    def _on_sheet_clicked(self, sheet_id):
        for vid, sid in list(self.matches.items()):
            if sid == sheet_id:
                self._unmatch(vid)
                return
        if self.selected_sheet_id == sheet_id:
            self.selected_sheet_id = None
        else:
            self.selected_sheet_id = sheet_id
        self._try_make_match()
        self._refresh_button_styles()
    
    def _try_make_match(self):
        if self.selected_view_id is not None and self.selected_sheet_id is not None:
            self.matches[self.selected_view_id] = self.selected_sheet_id
            self.selected_view_id = None
            self.selected_sheet_id = None
            self._redraw_all_lines()
            self._update_status()
    
    def _unmatch(self, view_id):
        if view_id in self.matches:
            del self.matches[view_id]
        if view_id in self.lines:
            self.canvas.Children.Remove(self.lines[view_id])
            del self.lines[view_id]
        self._refresh_button_styles()
        self._update_status()
    
    def _refresh_button_styles(self):
        matched_view_ids = set(self.matches.keys())
        matched_sheet_ids = set(self.matches.values())
        
        for vid, btn in self.view_buttons.items():
            t1, t2 = btn.Tag[0], btn.Tag[1]
            if vid == self.selected_view_id:
                btn.Background = self._brush(self.COLOR_SELECTED)
                t1.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
                t2.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
            elif vid in matched_view_ids:
                btn.Background = self._brush(self.COLOR_MATCHED)
                t1.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
                t2.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
            else:
                btn.Background = self._brush(self.COLOR_BTN)
                t1.Foreground = self._brush(self.COLOR_TEXT)
                t2.Foreground = self._brush((120, 120, 130))
        
        for sid, btn in self.sheet_buttons.items():
            t1, t2 = btn.Tag[0], btn.Tag[1]
            if sid == self.selected_sheet_id:
                btn.Background = self._brush(self.COLOR_SELECTED)
                t1.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
                t2.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
            elif sid in matched_sheet_ids:
                btn.Background = self._brush(self.COLOR_MATCHED)
                t1.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
                t2.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
            else:
                btn.Background = self._brush(self.COLOR_BTN)
                t1.Foreground = self._brush(self.COLOR_TEXT)
                t2.Foreground = self._brush((120, 120, 130))
    
    def _redraw_all_lines(self):
        self.canvas.Children.Clear()
        self.lines.clear()
        for vid, sid in self.matches.items():
            self._draw_line(vid, sid)
    
    def _draw_line(self, view_id, sheet_id):
        vbtn = self.view_buttons.get(view_id)
        sbtn = self.sheet_buttons.get(sheet_id)
        if vbtn is None or sbtn is None:
            return
        if vbtn.Visibility != Visibility.Visible or sbtn.Visibility != Visibility.Visible:
            return
        try:
            v_point = vbtn.TranslatePoint(
                Point(vbtn.ActualWidth, vbtn.ActualHeight / 2), self.canvas
            )
            s_point = sbtn.TranslatePoint(
                Point(0, sbtn.ActualHeight / 2), self.canvas
            )
        except:
            return
        
        # Canvas 高度範圍 — 把跑到範圍外的 Y 座標夾住，避免線飛到天邊
        canvas_h = self.canvas.ActualHeight
        if canvas_h <= 0:
            return
        
        def clamp_y(y):
            if y < 0:
                return 0.0
            if y > canvas_h:
                return canvas_h
            return y
        
        # 任一端完全超出可見範圍 → 不畫（避免線從邊緣突兀地凸出來）
        margin = 30  # 容許一點點超出，讓滾動時的視覺有連續感
        if v_point.Y < -margin or v_point.Y > canvas_h + margin:
            return
        if s_point.Y < -margin or s_point.Y > canvas_h + margin:
            return
        
        line = Line()
        line.X1 = v_point.X
        line.Y1 = clamp_y(v_point.Y)
        line.X2 = s_point.X
        line.Y2 = clamp_y(s_point.Y)
        line.Stroke = self._brush(self.COLOR_LINE)
        line.StrokeThickness = 2.5
        line.StrokeStartLineCap = PenLineCap.Round
        line.StrokeEndLineCap = PenLineCap.Round
        self.canvas.Children.Add(line)
        self.lines[view_id] = line
    
    def _update_status(self):
        if not self.matches:
            self.status_text.Text = "尚未配對"
        else:
            self.status_text.Text = "已配對 {} 組".format(len(self.matches))
    
    def _on_run(self, sender, args):
        if not self.matches:
            forms.alert("還沒有任何配對，至少配一組才能執行。")
            return
        self.confirmed = True
        self.Close()
    
    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.Close()


# ---------------------------------------------------------------------------
# 執行配對 → 放置視圖
# ---------------------------------------------------------------------------

def place_views_to_sheets(matches, document):
    success = 0
    failed = []
    
    with revit.Transaction("Place Views to Sheets"):
        for view_id_int, sheet_id_int in matches.items():
            view_id = DB.ElementId(view_id_int)
            sheet_id = DB.ElementId(sheet_id_int)
            view = document.GetElement(view_id)
            sheet = document.GetElement(sheet_id)
            
            try:
                if not DB.Viewport.CanAddViewToSheet(document, sheet_id, view_id):
                    failed.append((view.Name, sheet.SheetNumber, "此視圖無法放上此圖紙"))
                    continue
                
                outline = sheet.Outline
                center = DB.XYZ(
                    (outline.Min.U + outline.Max.U) / 2.0,
                    (outline.Min.V + outline.Max.V) / 2.0,
                    0
                )
                
                DB.Viewport.Create(document, sheet_id, view_id, center)
                success += 1
            except Exception as ex:
                failed.append((view.Name, sheet.SheetNumber, str(ex)))
    
    return success, failed


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main():
    views = get_placeable_views(doc)
    sheets = get_all_sheets(doc)
    
    if not views:
        forms.alert("沒有可配對的視圖（所有視圖可能都已放上圖紙）。", exitscript=True)
    if not sheets:
        forms.alert("專案中沒有可用的圖紙。", exitscript=True)
    
    win = MatchWindow(views, sheets)
    win.ShowDialog()
    
    if not win.confirmed:
        script.exit()
    
    success, failed = place_views_to_sheets(win.matches, doc)
    
    output.print_md("# ✅ 視圖配對上圖完成")
    output.print_md("- 成功放置: **{}** 個視圖".format(success))
    output.print_md("- 失敗: **{}** 個".format(len(failed)))
    
    if failed:
        output.print_md("\n## 失敗清單")
        for vname, snum, reason in failed:
            output.print_md("- `{}` → `{}`: {}".format(vname, snum, reason))


if __name__ == '__main__':
    main()
