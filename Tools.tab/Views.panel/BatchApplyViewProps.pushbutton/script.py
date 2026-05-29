# -*- coding: utf-8 -*-
"""
批次套用視圖屬性 (Batch Apply View Properties)
===========================================================
功能:
  在左欄勾選想處理的視圖，在右欄獨立勾選想套用的屬性
  （Scope Box、View Template...），按執行批次套用。

特色:
  - 每個屬性前面有「啟用」勾選框，沒勾就不動該屬性
  - 視圖列表可依分類過濾、搜尋
  - 副標題顯示視圖目前的 Scope Box / Template，讓你知道現況

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
# Revit 資料準備
# ---------------------------------------------------------------------------

VIEW_FAMILY_MAP = {
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


def get_view_type_name(view):
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
    return VIEW_FAMILY_MAP.get(view.ViewType, str(view.ViewType))


def get_view_group_label(view):
    return "{} ({})".format(get_view_family_name(view), get_view_type_name(view))


def get_all_views(document):
    """取得所有可批次處理的視圖。"""
    views = DB.FilteredElementCollector(document).OfClass(DB.View).ToElements()
    excluded_types = {
        DB.ViewType.SystemBrowser, DB.ViewType.ProjectBrowser,
        DB.ViewType.Internal, DB.ViewType.Undefined
    }
    result = []
    for v in views:
        if v.IsTemplate or isinstance(v, DB.ViewSheet) or v.ViewType in excluded_types:
            continue
        result.append(v)
    return sorted(result, key=lambda x: (get_view_group_label(x), x.Name))


def get_all_scope_boxes(document):
    boxes = DB.FilteredElementCollector(document) \
        .OfCategory(DB.BuiltInCategory.OST_VolumeOfInterest) \
        .WhereElementIsNotElementType() \
        .ToElements()
    return sorted(boxes, key=lambda b: b.Name)


def get_all_view_templates(document):
    views = DB.FilteredElementCollector(document).OfClass(DB.View).ToElements()
    templates = [v for v in views if v.IsTemplate]
    return sorted(templates, key=lambda t: (str(t.ViewType), t.Name))


def get_current_scope_box_name(view):
    try:
        p = view.get_Parameter(DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
        if p and p.HasValue:
            box_id = p.AsElementId()
            if box_id and box_id.IntegerValue > 0:
                box = doc.GetElement(box_id)
                if box:
                    return box.Name
    except:
        pass
    return None


def get_current_template_name(view):
    try:
        tpl_id = view.ViewTemplateId
        if tpl_id and tpl_id.IntegerValue > 0:
            tpl = doc.GetElement(tpl_id)
            if tpl:
                return tpl.Name
    except:
        pass
    return None


# ---------------------------------------------------------------------------
# WPF 視窗
# ---------------------------------------------------------------------------

class BatchApplyWindow(Window):
    
    COLOR_BG = (245, 245, 250)
    COLOR_BTN = (255, 255, 255)
    COLOR_PRIMARY = (99, 102, 241)
    COLOR_TEXT = (30, 30, 40)
    COLOR_TEXT_LIGHT = (255, 255, 255)
    COLOR_GROUP_HEADER = (235, 238, 245)
    COLOR_GROUP_TEXT = (90, 95, 110)
    COLOR_DISABLED = (160, 165, 175)
    
    OPTION_CLEAR = "（清除 / 不指定）"
    VIEW_GROUP_ALL = "全部分類"
    
    def __init__(self, views, scope_boxes, view_templates):
        self.views = views
        self.scope_boxes = scope_boxes
        self.view_templates = view_templates
        
        self.view_groups = []
        seen = set()
        for v in views:
            label = get_view_group_label(v)
            if label not in seen:
                seen.add(label)
                self.view_groups.append(label)
        
        # 視圖狀態
        self.view_checks = {}
        self.view_rows = {}
        self.view_group_headers = {}
        
        self.view_search_text = ""
        self.view_group_filter = self.VIEW_GROUP_ALL
        
        # 屬性勾選結果（給主流程使用）
        # apply_scope_box / scope_box_choice (None 代表清除)
        # apply_template / template_choice
        self.result = None
        self.confirmed = False
        
        self._build_ui()
    
    @staticmethod
    def _brush(rgb):
        return SolidColorBrush(Color.FromRgb(rgb[0], rgb[1], rgb[2]))
    
    def _build_ui(self):
        self.Title = "批次套用視圖屬性"
        self.Width = 1000
        self.Height = 720
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = self._brush(self.COLOR_BG)
        
        root = Grid()
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        
        header = TextBlock()
        header.Text = "1. 在左邊勾選要處理的視圖   2. 在右邊勾選並設定要套用的屬性   3. 按「執行」"
        header.FontSize = 13
        header.Foreground = self._brush(self.COLOR_TEXT)
        header.VerticalAlignment = VerticalAlignment.Center
        header.Margin = Thickness(20, 0, 0, 0)
        Grid.SetRow(header, 0)
        root.Children.Add(header)
        
        mid = Grid()
        mid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        mid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(380)))
        
        left = self._build_view_column()
        Grid.SetColumn(left, 0)
        mid.Children.Add(left)
        
        right = self._build_properties_column()
        Grid.SetColumn(right, 1)
        mid.Children.Add(right)
        
        Grid.SetRow(mid, 1)
        root.Children.Add(mid)
        
        bottom = Grid()
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(120)))
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(120)))
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(20)))
        
        self.status_text = TextBlock()
        self.status_text.Text = "尚未選擇視圖"
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
    
    def _build_view_column(self):
        outer = Grid()
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(115)))
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        outer.Margin = Thickness(20, 0, 10, 0)
        
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
        
        action_panel = StackPanel()
        action_panel.Orientation = Orientation.Horizontal
        action_panel.Margin = Thickness(0, 4, 0, 0)
        
        sel_all = Button()
        sel_all.Content = "全選 (顯示中)"
        sel_all.Margin = Thickness(0, 0, 5, 0)
        sel_all.Padding = Thickness(8, 2, 8, 2)
        sel_all.Click += self._on_select_all_visible
        action_panel.Children.Add(sel_all)
        
        clr_all = Button()
        clr_all.Content = "全部取消"
        clr_all.Padding = Thickness(8, 2, 8, 2)
        clr_all.Click += self._on_clear_all
        action_panel.Children.Add(clr_all)
        
        filter_panel.Children.Add(action_panel)
        
        Grid.SetRow(filter_panel, 0)
        outer.Children.Add(filter_panel)
        
        # 視圖列表
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
            
            row = self._make_view_row(v, label)
            self.view_rows[v.Id.IntegerValue] = row
            self.view_list_panel.Children.Add(row)
        
        scroll = ScrollViewer()
        scroll.Content = self.view_list_panel
        Grid.SetRow(scroll, 1)
        outer.Children.Add(scroll)
        
        return outer
    
    def _build_properties_column(self):
        """右欄: 各種屬性區塊，每個都有自己的「啟用」勾選框。"""
        outer = ScrollViewer()
        outer.Margin = Thickness(10, 0, 20, 0)
        
        panel = StackPanel()
        panel.Orientation = Orientation.Vertical
        
        # === 區塊 1: Scope Box ===
        self.sb_block, self.sb_enable_check, self.sb_combo = self._make_property_block(
            title="範圍框 (Scope Box)",
            hint="套用後視圖會被裁切到範圍框邊界。",
            options=[self.OPTION_CLEAR] + [b.Name for b in self.scope_boxes],
            on_enable_changed=self._on_sb_enable_changed
        )
        panel.Children.Add(self.sb_block)
        
        # === 區塊 2: View Template ===
        self.tpl_block, self.tpl_enable_check, self.tpl_combo = self._make_property_block(
            title="視圖樣板 (View Template)",
            hint="樣板會控制可見性、線型、顯示樣式等。被樣板控制的屬性無法手動修改。",
            options=[self.OPTION_CLEAR] + [t.Name for t in self.view_templates],
            on_enable_changed=self._on_tpl_enable_changed
        )
        panel.Children.Add(self.tpl_block)
        
        # 預設都未啟用
        self._set_block_enabled(self.sb_combo, False)
        self._set_block_enabled(self.tpl_combo, False)
        
        outer.Content = panel
        return outer
    
    def _make_property_block(self, title, hint, options, on_enable_changed):
        """產生一個屬性區塊：上方有啟用勾選框、下方下拉選單。"""
        border = Border()
        border.Background = self._brush(self.COLOR_BTN)
        border.BorderBrush = self._brush((220, 222, 230))
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(5)
        border.Padding = Thickness(14, 12, 14, 14)
        border.Margin = Thickness(0, 6, 0, 6)
        
        sp = StackPanel()
        sp.Orientation = Orientation.Vertical
        
        # 標題列：[啟用勾選] [標題]
        head = StackPanel()
        head.Orientation = Orientation.Horizontal
        
        enable_check = CheckBox()
        enable_check.VerticalAlignment = VerticalAlignment.Center
        enable_check.Margin = Thickness(0, 0, 8, 0)
        enable_check.Checked += on_enable_changed
        enable_check.Unchecked += on_enable_changed
        head.Children.Add(enable_check)
        
        title_tb = TextBlock()
        title_tb.Text = title
        title_tb.FontWeight = FontWeights.Bold
        title_tb.FontSize = 14
        title_tb.Foreground = self._brush(self.COLOR_TEXT)
        title_tb.VerticalAlignment = VerticalAlignment.Center
        head.Children.Add(title_tb)
        
        sp.Children.Add(head)
        
        # 提示
        hint_tb = TextBlock()
        hint_tb.Text = hint
        hint_tb.FontSize = 11
        hint_tb.Foreground = self._brush((110, 115, 130))
        hint_tb.Margin = Thickness(24, 4, 0, 8)
        hint_tb.TextWrapping = TextWrapping.Wrap
        sp.Children.Add(hint_tb)
        
        # 下拉
        combo = ComboBox()
        combo.Padding = Thickness(4, 4, 4, 4)
        combo.Margin = Thickness(24, 0, 0, 0)
        for opt in options:
            combo.Items.Add(opt)
        combo.SelectedIndex = 0
        sp.Children.Add(combo)
        
        border.Child = sp
        return border, enable_check, combo
    
    def _set_block_enabled(self, combo, enabled):
        combo.IsEnabled = enabled
        combo.Opacity = 1.0 if enabled else 0.5
    
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
    
    def _make_view_row(self, view, group_label):
        border = Border()
        border.Background = self._brush(self.COLOR_BTN)
        border.BorderBrush = self._brush((220, 220, 230))
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(2)
        border.Margin = Thickness(0, 1, 0, 1)
        border.Padding = Thickness(8, 4, 8, 4)
        
        grid = Grid()
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(28)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        cb = CheckBox()
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Tag = view.Id.IntegerValue
        cb.Checked += self._on_view_check_changed
        cb.Unchecked += self._on_view_check_changed
        Grid.SetColumn(cb, 0)
        grid.Children.Add(cb)
        
        info = StackPanel()
        info.Orientation = Orientation.Vertical
        Grid.SetColumn(info, 1)
        
        t1 = TextBlock()
        t1.Text = view.Name
        t1.FontSize = 12
        t1.FontWeight = FontWeights.SemiBold
        t1.Foreground = self._brush(self.COLOR_TEXT)
        t1.TextTrimming = TextTrimming.CharacterEllipsis
        info.Children.Add(t1)
        
        # 副標：類型 + 目前 Scope Box / Template
        sub_parts = [get_view_type_name(view)]
        sb_name = get_current_scope_box_name(view)
        tpl_name = get_current_template_name(view)
        if sb_name:
            sub_parts.append("框: " + sb_name)
        if tpl_name:
            sub_parts.append("樣板: " + tpl_name)
        sub_parts.append("ID " + str(view.Id.IntegerValue))
        
        t2 = TextBlock()
        t2.Text = " · ".join(sub_parts)
        t2.FontSize = 10
        t2.Foreground = self._brush((120, 120, 130))
        t2.Margin = Thickness(0, 2, 0, 0)
        t2.TextTrimming = TextTrimming.CharacterEllipsis
        info.Children.Add(t2)
        
        grid.Children.Add(info)
        border.Child = grid
        
        self.view_checks[view.Id.IntegerValue] = cb
        border.Tag = (group_label, view.Name)
        return border
    
    # ---- 事件 ----
    
    def _on_view_check_changed(self, sender, args):
        self._update_status()
    
    def _on_select_all_visible(self, sender, args):
        for vid, row in self.view_rows.items():
            if row.Visibility == Visibility.Visible:
                self.view_checks[vid].IsChecked = True
    
    def _on_clear_all(self, sender, args):
        for cb in self.view_checks.values():
            cb.IsChecked = False
    
    def _on_view_group_changed(self, sender, args):
        self.view_group_filter = sender.SelectedItem or self.VIEW_GROUP_ALL
        self._apply_view_filter()
    
    def _on_view_search_changed(self, sender, args):
        self.view_search_text = (sender.Text or "").lower()
        self._apply_view_filter()
    
    def _on_sb_enable_changed(self, sender, args):
        self._set_block_enabled(self.sb_combo, sender.IsChecked)
    
    def _on_tpl_enable_changed(self, sender, args):
        self._set_block_enabled(self.tpl_combo, sender.IsChecked)
    
    def _apply_view_filter(self):
        visible_count = 0
        group_has_visible = {g: False for g in self.view_groups}
        
        for vid, row in self.view_rows.items():
            group_label, view_name = row.Tag
            
            if self.view_group_filter != self.VIEW_GROUP_ALL:
                if group_label != self.view_group_filter:
                    row.Visibility = Visibility.Collapsed
                    continue
            
            if self.view_search_text:
                if self.view_search_text not in view_name.lower() \
                   and self.view_search_text not in group_label.lower():
                    row.Visibility = Visibility.Collapsed
                    continue
            
            row.Visibility = Visibility.Visible
            group_has_visible[group_label] = True
            visible_count += 1
        
        for label, hdr in self.view_group_headers.items():
            hdr.Visibility = Visibility.Visible if group_has_visible.get(label) else Visibility.Collapsed
        
        self.view_count_label.Text = "視圖 ({} / {})".format(visible_count, len(self.views))
    
    def _update_status(self):
        n = sum(1 for cb in self.view_checks.values() if cb.IsChecked)
        if n == 0:
            self.status_text.Text = "尚未選擇視圖"
        else:
            self.status_text.Text = "已選擇 {} 個視圖".format(n)
    
    def _on_run(self, sender, args):
        # 收集勾選的視圖
        checked_view_ids = [vid for vid, cb in self.view_checks.items() if cb.IsChecked]
        if not checked_view_ids:
            forms.alert("請至少勾選一個視圖。")
            return
        
        if not (self.sb_enable_check.IsChecked or self.tpl_enable_check.IsChecked):
            forms.alert("請至少勾選一個要套用的屬性（Scope Box 或視圖樣板）。")
            return
        
        # 整理結果
        result = {
            "view_ids": checked_view_ids,
            "apply_scope_box": bool(self.sb_enable_check.IsChecked),
            "scope_box": None,  # ElementId or None=clear
            "apply_template": bool(self.tpl_enable_check.IsChecked),
            "template": None,
        }
        
        if result["apply_scope_box"]:
            sel = self.sb_combo.SelectedItem
            if sel and sel != self.OPTION_CLEAR:
                for b in self.scope_boxes:
                    if b.Name == sel:
                        result["scope_box"] = b.Id
                        break
        
        if result["apply_template"]:
            sel = self.tpl_combo.SelectedItem
            if sel and sel != self.OPTION_CLEAR:
                for t in self.view_templates:
                    if t.Name == sel:
                        result["template"] = t.Id
                        break
        
        self.result = result
        self.confirmed = True
        self.Close()
    
    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.Close()


# ---------------------------------------------------------------------------
# 執行套用
# ---------------------------------------------------------------------------

def apply_properties(result, document):
    """根據 UI 回傳的 result 執行批次套用。"""
    success = 0
    failed = []  # (view_name, reason)
    
    with revit.Transaction("Batch Apply View Properties"):
        for vid_int in result["view_ids"]:
            view_id = DB.ElementId(vid_int)
            view = document.GetElement(view_id)
            if view is None:
                continue
            
            errors_for_view = []
            
            # 套用 Scope Box
            if result["apply_scope_box"]:
                try:
                    p = view.get_Parameter(DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
                    if p is None:
                        errors_for_view.append("此視圖不支援 Scope Box")
                    else:
                        target = result["scope_box"] or DB.ElementId.InvalidElementId
                        p.Set(target)
                except Exception as ex:
                    errors_for_view.append("Scope Box 失敗: {}".format(ex))
            
            # 套用視圖樣板
            if result["apply_template"]:
                try:
                    target = result["template"] or DB.ElementId.InvalidElementId
                    view.ViewTemplateId = target
                except Exception as ex:
                    errors_for_view.append("視圖樣板失敗: {}".format(ex))
            
            if errors_for_view:
                failed.append((view.Name, "; ".join(errors_for_view)))
            else:
                success += 1
    
    return success, failed


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main():
    views = get_all_views(doc)
    if not views:
        forms.alert("專案中沒有可處理的視圖。", exitscript=True)
    
    scope_boxes = get_all_scope_boxes(doc)
    view_templates = get_all_view_templates(doc)
    
    win = BatchApplyWindow(views, scope_boxes, view_templates)
    win.ShowDialog()
    
    if not win.confirmed:
        script.exit()
    
    success, failed = apply_properties(win.result, doc)
    
    output.print_md("# ✅ 批次套用屬性完成")
    output.print_md("- 成功: **{}** 個視圖".format(success))
    output.print_md("- 失敗: **{}** 個".format(len(failed)))
    
    if failed:
        output.print_md("\n## 失敗清單")
        for vname, reason in failed:
            output.print_md("- `{}`: {}".format(vname, reason))


if __name__ == '__main__':
    main()
