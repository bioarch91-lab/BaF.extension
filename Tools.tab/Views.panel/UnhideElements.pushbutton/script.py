# -*- coding: utf-8 -*-
"""
取消元素隱藏 (Unhide Directly-Hidden Elements)
===================================================
功能:
  一鍵取消選定視圖中「被直接隱藏 (Hide in View > Hide Elements)」的元素，
  也就是用快捷鍵 H 或右鍵隱藏的個別物件。

特別注意 (重要):
  本工具只處理「元素層級隱藏」(Element.IsHidden / view.UnhideElements)，
  → 不會 動到視圖樣板 (View Template) 控制的可見性
  → 不會 改動類別可見性 (Hide Category / V/G)
  → 不會 影響可見性篩選器 (Filter)
  因此用視圖樣板關掉的東西，執行後仍維持關閉。

使用方式:
  1. 執行此腳本
  2. 在清單勾選要處理的視圖 (預設只勾「目前視圖」，可用「全選/全不選/反選」)
  3. 按「執行」
  4. 完成後輸出每個視圖取消隱藏的元素數量報告

作者: BaF / BIM 工具
"""

import clr
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")

from System.Collections.Generic import List
from System.Windows import (
    Window, Thickness, VerticalAlignment, HorizontalAlignment,
    WindowStartupLocation, GridLength, GridUnitType, FontWeights, TextWrapping
)
from System.Windows.Controls import (
    StackPanel, Button, ScrollViewer, CheckBox, TextBlock, Border,
    Grid, RowDefinition, ColumnDefinition, Orientation, ScrollBarVisibility
)
from System.Windows.Media import SolidColorBrush, Color

from pyrevit import revit, DB, forms, script

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


# ---------------------------------------------------------------------------
# Revit 資料準備
# ---------------------------------------------------------------------------

# 不適合 / 不可能有「元素層級隱藏」的視圖類型，從清單排除
_EXCLUDED_VIEW_TYPES = (
    DB.ViewType.SystemBrowser, DB.ViewType.ProjectBrowser,
    DB.ViewType.Internal, DB.ViewType.Undefined,
    DB.ViewType.Schedule, DB.ViewType.ColumnSchedule,
    DB.ViewType.PanelSchedule,
)

_FAMILY_MAP = {
    DB.ViewType.FloorPlan: "樓板平面圖",
    DB.ViewType.CeilingPlan: "天花板平面圖",
    DB.ViewType.Elevation: "立面圖",
    DB.ViewType.Section: "剖面",
    DB.ViewType.ThreeD: "3D 視圖",
    DB.ViewType.Detail: "詳圖",
    DB.ViewType.DraftingView: "圖說",
    DB.ViewType.Legend: "圖例",
    DB.ViewType.Rendering: "彩現",
    DB.ViewType.AreaPlan: "面積圖",
    DB.ViewType.EngineeringPlan: "結構平面",
    DB.ViewType.Walkthrough: "漫遊",
    DB.ViewType.DrawingSheet: "圖紙",
}


def get_view_family_name(view):
    """視圖大分類 (中文化)。"""
    return _FAMILY_MAP.get(view.ViewType, str(view.ViewType))


def get_view_type_name(view):
    """視圖類型名稱 (例如「出圖 1/200」)。"""
    try:
        vt = doc.GetElement(view.GetTypeId())
        if vt:
            tn = vt.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            if tn:
                return tn
    except:
        pass
    return get_view_family_name(view)


def get_view_group_label(view):
    """分組標籤，跟專案瀏覽器一致：大分類 (類型名)。圖紙統一歸到「圖紙」。"""
    if isinstance(view, DB.ViewSheet):
        return "圖紙"
    return "{} ({})".format(get_view_family_name(view), get_view_type_name(view))


def get_view_display_name(view):
    """清單與報告顯示名稱。圖紙顯示「圖號 - 圖名」，其餘用視圖名。"""
    if isinstance(view, DB.ViewSheet):
        return "{} - {}".format(view.SheetNumber, view.Name)
    return view.Name


def get_selectable_views(document):
    """取得可處理的視圖 (含圖紙；排除樣板、佔位圖紙與非圖面視圖)。"""
    views = DB.FilteredElementCollector(document).OfClass(DB.View).ToElements()
    result = []
    for v in views:
        if v.IsTemplate:
            continue
        if isinstance(v, DB.ViewSheet) and v.IsPlaceholder:
            continue
        if v.ViewType in _EXCLUDED_VIEW_TYPES:
            continue
        result.append(v)
    return sorted(result, key=lambda x: (get_view_group_label(x), get_view_display_name(x)))


def find_hidden_element_ids(view, all_elements):
    """
    回傳該視圖中「被直接隱藏 (元素層級)」的 ElementId 清單。

    註: 以 view 為範圍的 FilteredElementCollector 會排除被隱藏的元素，
    因此改為對全文件非類型元素逐一以 IsHidden(view) 判斷。
    IsHidden 對部分元素會丟例外，故包 try/except。
    """
    hidden = []
    for el in all_elements:
        try:
            if el.IsHidden(view):
                hidden.append(el.Id)
        except:
            pass
    return hidden


# ---------------------------------------------------------------------------
# 視圖勾選 WPF 視窗
# ---------------------------------------------------------------------------

class UnhideWindow(Window):

    COLOR_BG = (245, 245, 250)
    COLOR_TEXT = (30, 30, 40)
    COLOR_GROUP = (90, 95, 110)
    COLOR_ACCENT = (99, 102, 241)
    COLOR_TEXT_LIGHT = (255, 255, 255)

    def __init__(self, views, active_view_id):
        self.views = views
        self.active_view_id = active_view_id
        self.checkboxes = []        # list of (CheckBox, view)
        self.confirmed = False
        self.selected_views = []
        self._build_ui()

    @staticmethod
    def _brush(rgb):
        return SolidColorBrush(Color.FromRgb(rgb[0], rgb[1], rgb[2]))

    def _build_ui(self):
        self.Title = "取消元素隱藏 - 選擇視圖"
        self.Width = 520
        self.Height = 640
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = self._brush(self.COLOR_BG)

        root = Grid()
        root.Margin = Thickness(16)
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(46)))           # 說明
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(36)))           # 工具列
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))  # 清單
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(52)))           # 底部按鈕

        # --- 說明 ---
        hint = TextBlock()
        hint.Text = ("勾選要處理的視圖。只會取消「被直接隱藏的元素」，"
                     "不影響視圖樣板與類別可見性。")
        hint.TextWrapping = TextWrapping.Wrap
        hint.FontSize = 12
        hint.Foreground = self._brush(self.COLOR_TEXT)
        hint.VerticalAlignment = VerticalAlignment.Center
        Grid.SetRow(hint, 0)
        root.Children.Add(hint)

        # --- 工具列 (全選 / 全不選 / 反選) ---
        toolbar = StackPanel()
        toolbar.Orientation = Orientation.Horizontal
        Grid.SetRow(toolbar, 1)
        for text, handler in (("全選", self._on_select_all),
                              ("全不選", self._on_select_none),
                              ("反選", self._on_invert)):
            b = Button()
            b.Content = text
            b.MinWidth = 70
            b.Margin = Thickness(0, 0, 8, 0)
            b.Padding = Thickness(6, 2, 6, 2)
            b.Click += handler
            toolbar.Children.Add(b)
        root.Children.Add(toolbar)

        # --- 視圖清單 (可捲動 + 分組) ---
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Margin = Thickness(0, 6, 0, 6)
        Grid.SetRow(scroll, 2)

        list_panel = StackPanel()
        list_panel.Orientation = Orientation.Vertical

        current_group = None
        for v in self.views:
            group = get_view_group_label(v)
            if group != current_group:
                current_group = group
                header = TextBlock()
                header.Text = group
                header.FontWeight = FontWeights.Bold
                header.FontSize = 12
                header.Foreground = self._brush(self.COLOR_GROUP)
                header.Margin = Thickness(2, 8, 0, 2)
                list_panel.Children.Add(header)

            cb = CheckBox()
            cb.Content = get_view_display_name(v)
            cb.FontSize = 13
            cb.Margin = Thickness(14, 2, 0, 2)
            cb.Foreground = self._brush(self.COLOR_TEXT)
            if (self.active_view_id is not None
                    and v.Id.IntegerValue == self.active_view_id.IntegerValue):
                cb.IsChecked = True
                cb.Content = get_view_display_name(v) + "  (目前視圖)"
            self.checkboxes.append((cb, v))
            list_panel.Children.Add(cb)

        scroll.Content = list_panel
        root.Children.Add(scroll)

        # --- 底部按鈕 ---
        bottom = Grid()
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(110)))
        bottom.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(110)))
        Grid.SetRow(bottom, 3)

        cancel_btn = Button()
        cancel_btn.Content = "取消"
        cancel_btn.Margin = Thickness(0, 10, 6, 6)
        cancel_btn.Click += self._on_cancel
        Grid.SetColumn(cancel_btn, 1)
        bottom.Children.Add(cancel_btn)

        run_btn = Button()
        run_btn.Content = "執行"
        run_btn.Margin = Thickness(0, 10, 0, 6)
        run_btn.Background = self._brush(self.COLOR_ACCENT)
        run_btn.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
        run_btn.FontWeight = FontWeights.Bold
        run_btn.Click += self._on_run
        Grid.SetColumn(run_btn, 2)
        bottom.Children.Add(run_btn)

        root.Children.Add(bottom)
        self.Content = root

    # --- 工具列事件 ---
    def _on_select_all(self, sender, args):
        for cb, _ in self.checkboxes:
            cb.IsChecked = True

    def _on_select_none(self, sender, args):
        for cb, _ in self.checkboxes:
            cb.IsChecked = False

    def _on_invert(self, sender, args):
        for cb, _ in self.checkboxes:
            cb.IsChecked = not bool(cb.IsChecked)

    # --- 底部事件 ---
    def _on_run(self, sender, args):
        self.selected_views = [v for cb, v in self.checkboxes if cb.IsChecked]
        if not self.selected_views:
            forms.alert("請至少勾選一個視圖。")
            return
        self.confirmed = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.Close()


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main():
    views = get_selectable_views(doc)
    if not views:
        forms.alert("專案中沒有可處理的視圖。", exitscript=True)

    active_view = uidoc.ActiveView
    active_view_id = active_view.Id if active_view is not None else None

    win = UnhideWindow(views, active_view_id)
    win.ShowDialog()
    if not win.confirmed:
        script.exit()

    # 全文件非類型元素只蒐集一次，供各視圖判斷
    all_elements = list(
        DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    )

    report = []
    total = 0
    with revit.Transaction("取消元素隱藏"):
        for v in win.selected_views:
            hidden_ids = find_hidden_element_ids(v, all_elements)
            if hidden_ids:
                v.UnhideElements(List[DB.ElementId](hidden_ids))
            total += len(hidden_ids)
            report.append((v, len(hidden_ids)))

    # --- 輸出報告 ---
    output.print_md("# ✅ 取消元素隱藏完成")
    output.print_md("共處理 **{}** 個視圖，總計取消隱藏 **{}** 個元素。\n".format(
        len(win.selected_views), total))
    output.print_md("## 各視圖結果")
    for v, n in report:
        mark = "—" if n == 0 else "取消隱藏 **{}** 個".format(n)
        output.print_md("- {} `{}`: {}".format(
            get_view_group_label(v), get_view_display_name(v), mark))

    if total == 0:
        output.print_md("\n> 選定視圖中沒有「被直接隱藏的元素」。"
                        "（用視圖樣板/類別關掉的不在處理範圍內）")


if __name__ == '__main__':
    main()
