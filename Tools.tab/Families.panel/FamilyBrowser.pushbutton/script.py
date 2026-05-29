# -*- coding: utf-8 -*-
"""
族群瀏覽器 (Family Browser) - Phase 1
===========================================================
連到本案的族群／圖框資料夾，直接在 Revit 內瀏覽 .rfa 清單，
一鍵「載入」到專案，或「載入並放置」（載入後直接進入 Revit 放置模式）。

設計重點:
  - 側邊欄列出使用者設定的「資料夾來源」，點一下掃描該資料夾的 .rfa。
  - 主區列出族群，可用上方搜尋框過濾，每個族群有「載入」與「載入並放置」。
  - 視窗為「非 Modal」(Show)，所以開著也能繼續操作 Revit。
  - 因為非 Modal 視窗在事件回呼時已脫離 Revit API context，所有 Revit API
    動作都透過 ExternalEvent + IExternalEventHandler 在合法 context 內執行。

已知限制 (Phase 1):
  - 只支援資料夾(.rfa)來源；從 .rvt 專案檔抓族群為 Phase 2。
  - 尚未顯示縮圖（IronPython 無法乾淨讀取未載入 .rfa 的內嵌預覽），為 Phase 2。

作者: BaF / BIM 工具
"""

import io
import os
import sys
import clr

# --- 把 lib 目錄加入 path ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LIB_DIR = os.path.join(SCRIPT_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

import settings as settings_mod

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")

from System.Windows import (
    Window, Thickness, HorizontalAlignment, VerticalAlignment,
    WindowStartupLocation, TextTrimming, FontWeights,
    GridLength, GridUnitType, CornerRadius, TextWrapping
)
from System.Windows.Controls import (
    StackPanel, Button, ScrollViewer, Grid, RowDefinition, ColumnDefinition,
    TextBlock, Border, Orientation, TextBox
)
from System.Windows.Media import SolidColorBrush, Color

from pyrevit import revit, forms, script
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


# ---------------------------------------------------------------------------
# 設定檔位置（存在使用者 AppData）
# ---------------------------------------------------------------------------

APPDATA_DIR = os.path.join(os.environ.get("APPDATA", ""), "BaF_FamilyBrowser")
SETTINGS_FILE = os.path.join(APPDATA_DIR, "library.json")


# ---------------------------------------------------------------------------
# 掃描資料夾取得 .rfa 清單
# ---------------------------------------------------------------------------

def scan_folder_families(folder, recursive):
    """掃描資料夾的 .rfa，回傳 [{'name','path','group'}, ...]（已排序）。

    group = 相對於來源根目錄的子資料夾，用來在 UI 分群（例如「圖框」子夾）。
    """
    results = []
    if not folder or not os.path.isdir(folder):
        return results
    folder = os.path.abspath(folder)
    if recursive:
        for root, dirs, files in os.walk(folder):
            for fn in files:
                if fn.lower().endswith(".rfa"):
                    full = os.path.join(root, fn)
                    rel = os.path.relpath(root, folder)
                    group = "" if rel == "." else rel.replace("\\", "/")
                    results.append({
                        "name": os.path.splitext(fn)[0],
                        "path": full,
                        "group": group,
                    })
    else:
        for fn in os.listdir(folder):
            if fn.lower().endswith(".rfa"):
                full = os.path.join(folder, fn)
                if os.path.isfile(full):
                    results.append({
                        "name": os.path.splitext(fn)[0],
                        "path": full,
                        "group": "",
                    })
    results.sort(key=lambda d: (d["group"].lower(), d["name"].lower()))
    return results


# ---------------------------------------------------------------------------
# ExternalEvent 處理常式：在合法 Revit API context 內做載入 / 放置
# ---------------------------------------------------------------------------

class FamilyActionHandler(IExternalEventHandler):
    """非 Modal 視窗無法直接呼叫 Revit API，透過此 handler 在 idle 時執行。

    視窗端設定 self.action / self.payload 後呼叫 ExternalEvent.Raise()，
    Revit 會在合適時機呼叫本 Execute()。
    """

    def __init__(self):
        self.action = None      # "load" 或 "place"
        self.payload = None     # 目前是 .rfa 路徑字串

    def GetName(self):
        return "BaF 族群瀏覽器"

    def Execute(self, uiapp):
        # 重要：非 Modal 情境下，模組層級全域在 main() 結束後會被清掉，
        # 事件回呼時讀不到，因此所有東西都在這裡區域 import。
        from Autodesk.Revit.UI import TaskDialog
        try:
            uidoc_ = uiapp.ActiveUIDocument
            if uidoc_ is None:
                TaskDialog.Show("族群瀏覽器", "沒有開啟中的專案文件，無法載入族群。")
                return
            doc_ = uidoc_.Document
            action = self.action
            path = self.payload
            if action == "load":
                self._load_family(doc_, path, place=False, uidoc_=None)
            elif action == "place":
                self._load_family(doc_, path, place=True, uidoc_=uidoc_)
        except Exception as ex:
            import traceback
            try:
                TaskDialog.Show(
                    "族群瀏覽器 - 錯誤",
                    "{}\n\n--- 詳細 ---\n{}".format(ex, traceback.format_exc())
                )
            except:
                pass

    def _load_family(self, doc_, path, place, uidoc_):
        import os as _os
        import clr as _clr
        from pyrevit import DB as _DB
        from Autodesk.Revit.UI import TaskDialog

        if not path or not _os.path.isfile(path):
            TaskDialog.Show("族群瀏覽器", "找不到檔案：\n{}".format(path))
            return

        fam_name = _os.path.splitext(_os.path.basename(path))[0]

        # IFamilyLoadOptions：族群若已存在，預設「保留現有」不覆寫參數值，
        # 避免改動使用者專案內既有族群。
        class _LoadOpts(_DB.IFamilyLoadOptions):
            def OnFamilyFound(self, familyInUse, overwriteParameterValues):
                overwriteParameterValues.Value = False
                return True

            def OnSharedFamilyFound(self, sharedFamily, familyInUse,
                                    source, overwriteParameterValues):
                source.Value = _DB.FamilySource.Family
                overwriteParameterValues.Value = False
                return True

        # LoadFamily 自行管理交易，不要外包 Transaction。
        fam_ref = _clr.Reference[_DB.Family]()
        loaded = False
        try:
            loaded = doc_.LoadFamily(path, _LoadOpts(), fam_ref)
        except Exception:
            # 某些版本沒有 out 參數的多載差異，退而求其次用單參數版
            loaded = doc_.LoadFamily(path)

        family = None
        try:
            family = fam_ref.Value
        except Exception:
            family = None

        # loaded=False 通常代表「族群已經在專案裡」，此時 out family 可能為 None，
        # 改用名稱在既有族群中尋找，才能繼續放置。
        if family is None:
            family = self._find_family_by_name(_DB, doc_, fam_name)

        if family is None:
            TaskDialog.Show(
                "族群瀏覽器",
                "無法載入族群「{}」。\n（檔案可能損毀，或不是有效的 .rfa）".format(fam_name)
            )
            return

        if not place:
            verb = "已載入" if loaded else "已存在於專案中"
            TaskDialog.Show("族群瀏覽器", "族群「{}」{}。".format(fam_name, verb))
            return

        # ---- 載入並放置 ----
        sym_ids = list(family.GetFamilySymbolIds())
        if not sym_ids:
            TaskDialog.Show("族群瀏覽器", "族群「{}」沒有可放置的類型。".format(fam_name))
            return

        symbol = doc_.GetElement(sym_ids[0])

        # 啟用類型需要交易
        if symbol is not None and not symbol.IsActive:
            from Autodesk.Revit.DB import Transaction
            t = Transaction(doc_, "啟用族群類型")
            t.Start()
            try:
                symbol.Activate()
                doc_.Regenerate()
                t.Commit()
            except Exception:
                t.RollBack()
                raise

        # PostRequest 必須在交易外、API context 內呼叫；Execute 結束後 Revit
        # 才會進入放置模式（此時非 Modal 視窗仍開著，可重複點選）。
        uidoc_.PostRequestForElementTypePlacement(symbol)

    @staticmethod
    def _find_family_by_name(_DB, doc_, fam_name):
        try:
            collector = _DB.FilteredElementCollector(doc_).OfClass(_DB.Family)
            for fam in collector:
                if fam.Name == fam_name:
                    return fam
        except Exception:
            pass
        return None


# ---------------------------------------------------------------------------
# WPF 主視窗
# ---------------------------------------------------------------------------

class FamilyBrowserWindow(Window):

    COLOR_BG = (245, 245, 250)
    COLOR_SIDEBAR = (250, 250, 254)
    COLOR_BTN = (255, 255, 255)
    COLOR_PRIMARY = (99, 102, 241)
    COLOR_PRIMARY_SOFT = (224, 231, 255)
    COLOR_TEXT = (30, 30, 40)
    COLOR_TEXT_LIGHT = (255, 255, 255)
    COLOR_ACTIVE = (224, 231, 255)

    def __init__(self, settings_path, ext_event, handler):
        self.settings_path = settings_path
        self._ext_event = ext_event
        self._handler = handler

        import settings as _settings_mod
        self._settings_mod = _settings_mod
        self.settings = _settings_mod.load_settings(settings_path)

        # 存 dialog class 參考：非 Modal 視窗在事件回呼時模組全域已被清掉，
        # 屆時直接用名稱 AddSourceDialog 可能取不到，先在建構時（command context）存起來。
        self._add_source_dialog_cls = AddSourceDialog

        self.source_buttons = {}     # source_id -> Border
        self.active_source_id = None
        self.current_families = []    # 目前來源掃到的族群
        self.search_text = ""

        self._build_ui()
        self._refresh_source_list()

    @staticmethod
    def _brush(rgb):
        from System.Windows.Media import SolidColorBrush, Color
        return SolidColorBrush(Color.FromRgb(rgb[0], rgb[1], rgb[2]))

    def _build_ui(self):
        self.Title = "族群瀏覽器"
        self.Width = 980
        self.Height = 720
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = self._brush(self.COLOR_BG)

        root = Grid()
        root.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(260)))
        root.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))

        sidebar = self._build_sidebar()
        Grid.SetColumn(sidebar, 0)
        root.Children.Add(sidebar)

        main = self._build_main()
        Grid.SetColumn(main, 1)
        root.Children.Add(main)

        self.Content = root

    def _build_sidebar(self):
        import os as _os
        border = Border()
        border.Background = self._brush(self.COLOR_SIDEBAR)
        border.BorderBrush = self._brush((220, 222, 230))
        border.BorderThickness = Thickness(0, 0, 1, 0)

        outer = Grid()
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))   # 標題
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))  # 來源清單
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))   # 新增按鈕
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(35)))   # 狀態列

        title = TextBlock()
        title.Text = "📦 族群來源"
        title.FontSize = 16
        title.FontWeight = FontWeights.Bold
        title.Foreground = self._brush(self.COLOR_TEXT)
        title.Margin = Thickness(16, 14, 0, 0)
        Grid.SetRow(title, 0)
        outer.Children.Add(title)

        scroll = ScrollViewer()
        scroll.Margin = Thickness(8, 0, 8, 0)
        self.source_list_panel = StackPanel()
        self.source_list_panel.Orientation = Orientation.Vertical
        scroll.Content = self.source_list_panel
        Grid.SetRow(scroll, 1)
        outer.Children.Add(scroll)

        add_btn = Button()
        add_btn.Content = "+ 新增資料夾來源"
        add_btn.Margin = Thickness(12, 4, 12, 4)
        add_btn.Padding = Thickness(8, 8, 8, 8)
        add_btn.Background = self._brush(self.COLOR_PRIMARY)
        add_btn.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
        add_btn.FontWeight = FontWeights.SemiBold
        add_btn.Click += self._on_add_folder
        Grid.SetRow(add_btn, 2)
        outer.Children.Add(add_btn)

        status = TextBlock()
        status.Text = "設定檔: {}".format(_os.path.basename(self.settings_path or "未設定"))
        status.FontSize = 10
        status.Foreground = self._brush((130, 135, 145))
        status.Margin = Thickness(12, 4, 12, 4)
        status.TextTrimming = TextTrimming.CharacterEllipsis
        try:
            status.ToolTip = self.settings_path
        except:
            pass
        Grid.SetRow(status, 3)
        outer.Children.Add(status)

        border.Child = outer
        return border

    def _build_main(self):
        container = Grid()
        container.Background = self._brush((255, 255, 255))
        container.RowDefinitions.Add(RowDefinition(Height=GridLength(56)))   # 標題列
        container.RowDefinitions.Add(RowDefinition(Height=GridLength(44)))   # 搜尋列
        container.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))  # 清單

        # 標題列
        header = Grid()
        header.Background = self._brush((250, 250, 254))
        self.main_title = TextBlock()
        self.main_title.Text = "請從左側選擇一個來源"
        self.main_title.FontSize = 15
        self.main_title.FontWeight = FontWeights.SemiBold
        self.main_title.Foreground = self._brush(self.COLOR_TEXT)
        self.main_title.Margin = Thickness(16, 0, 0, 0)
        self.main_title.VerticalAlignment = VerticalAlignment.Center
        header.Children.Add(self.main_title)
        Grid.SetRow(header, 0)
        container.Children.Add(header)

        # 搜尋列
        search_box = TextBox()
        search_box.Margin = Thickness(16, 8, 16, 8)
        search_box.Padding = Thickness(6, 4, 6, 4)
        search_box.FontSize = 12
        search_box.VerticalContentAlignment = VerticalAlignment.Center
        search_box.TextChanged += self._on_search_changed
        self.search_box = search_box
        Grid.SetRow(search_box, 1)
        container.Children.Add(search_box)

        # 族群清單
        scroll = ScrollViewer()
        scroll.Margin = Thickness(8, 0, 8, 8)
        self.family_list_panel = StackPanel()
        self.family_list_panel.Orientation = Orientation.Vertical
        scroll.Content = self.family_list_panel
        Grid.SetRow(scroll, 2)
        container.Children.Add(scroll)

        self._show_family_placeholder("選擇來源後，這裡會列出資料夾中的族群 (.rfa)。")
        return container

    # ---- 來源清單 ----

    def _refresh_source_list(self):
        from System.Windows import Thickness, TextWrapping
        from System.Windows.Controls import TextBlock
        self.source_list_panel.Children.Clear()
        self.source_buttons.clear()

        sources = self.settings.get("sources", [])
        if not sources:
            empty = TextBlock()
            empty.Text = "（尚無來源，點下方「新增資料夾來源」設定本案族群庫位置）"
            empty.FontSize = 11
            empty.Foreground = self._brush((150, 155, 165))
            empty.Margin = Thickness(8, 12, 8, 0)
            empty.TextWrapping = TextWrapping.Wrap
            self.source_list_panel.Children.Add(empty)
            return

        for s in sources:
            btn = self._make_source_button(s)
            self.source_buttons[s["id"]] = btn
            self.source_list_panel.Children.Add(btn)

    def _make_source_button(self, src):
        import os as _os
        from System.Windows import (Thickness, HorizontalAlignment, FontWeights,
                                     TextTrimming, CornerRadius, GridLength,
                                     GridUnitType)
        from System.Windows.Controls import (Border, Grid, ColumnDefinition,
                                              Button, StackPanel, TextBlock,
                                              Orientation)
        from System.Windows.Media import SolidColorBrush, Color

        outer = Border()
        outer.Background = self._brush(self.COLOR_BTN)
        outer.BorderBrush = self._brush((220, 222, 230))
        outer.BorderThickness = Thickness(1)
        outer.CornerRadius = CornerRadius(4)
        outer.Margin = Thickness(0, 2, 0, 2)

        grid = Grid()
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(28)))

        main_btn = Button()
        main_btn.HorizontalContentAlignment = HorizontalAlignment.Stretch
        main_btn.Background = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
        main_btn.BorderThickness = Thickness(0)
        main_btn.Padding = Thickness(10, 6, 4, 6)
        main_btn.Click += lambda s, e, sd=src: self._on_select_source(sd)

        sp = StackPanel()
        sp.Orientation = Orientation.Vertical

        t1 = TextBlock()
        t1.Text = src.get("name", "(未命名)")
        t1.FontSize = 12
        t1.FontWeight = FontWeights.SemiBold
        t1.Foreground = self._brush(self.COLOR_TEXT)
        t1.TextTrimming = TextTrimming.CharacterEllipsis
        sp.Children.Add(t1)

        t2 = TextBlock()
        t2.Text = src.get("path", "")
        t2.FontSize = 10
        t2.Foreground = self._brush((130, 135, 145))
        t2.Margin = Thickness(0, 2, 0, 0)
        t2.TextTrimming = TextTrimming.CharacterEllipsis
        sp.Children.Add(t2)

        main_btn.Content = sp
        try:
            main_btn.ToolTip = src.get("path", "")
        except:
            pass
        Grid.SetColumn(main_btn, 0)
        grid.Children.Add(main_btn)

        del_btn = Button()
        del_btn.Content = "✕"
        del_btn.FontSize = 11
        del_btn.Background = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
        del_btn.BorderThickness = Thickness(0)
        del_btn.Foreground = self._brush((160, 165, 175))
        del_btn.Click += lambda s, e, sd=src: self._on_remove_source(sd)
        try:
            del_btn.ToolTip = "從清單移除此來源（不刪除任何檔案）"
        except:
            pass
        Grid.SetColumn(del_btn, 1)
        grid.Children.Add(del_btn)

        outer.Child = grid
        outer.Tag = src["id"]
        return outer

    def _update_source_highlight(self):
        for sid, btn in self.source_buttons.items():
            if sid == self.active_source_id:
                btn.Background = self._brush(self.COLOR_ACTIVE)
                btn.BorderBrush = self._brush(self.COLOR_PRIMARY)
            else:
                btn.Background = self._brush(self.COLOR_BTN)
                btn.BorderBrush = self._brush((220, 222, 230))

    # ---- 族群清單 ----

    def _show_family_placeholder(self, message):
        from System.Windows import (Thickness, HorizontalAlignment,
                                     VerticalAlignment, TextWrapping)
        from System.Windows.Controls import TextBlock
        self.family_list_panel.Children.Clear()
        tb = TextBlock()
        tb.Text = message
        tb.FontSize = 13
        tb.Foreground = self._brush((120, 125, 135))
        tb.TextWrapping = TextWrapping.Wrap
        tb.HorizontalAlignment = HorizontalAlignment.Center
        tb.Margin = Thickness(40, 40, 40, 0)
        tb.MaxWidth = 480
        self.family_list_panel.Children.Add(tb)

    def _render_families(self):
        from System.Windows import Thickness, FontWeights, TextTrimming
        from System.Windows.Controls import TextBlock
        self.family_list_panel.Children.Clear()

        q = (self.search_text or "").strip().lower()
        items = self.current_families
        if q:
            items = [d for d in items
                     if q in d["name"].lower() or q in d["group"].lower()]

        if not items:
            if not self.current_families:
                self._show_family_placeholder(
                    "此來源資料夾沒有找到任何 .rfa 檔案。"
                )
            else:
                self._show_family_placeholder(
                    "沒有符合「{}」的族群。".format(self.search_text)
                )
            return

        last_group = None
        for d in items:
            grp = d["group"]
            if grp != last_group:
                last_group = grp
                hdr = TextBlock()
                hdr.Text = grp if grp else "（根目錄）"
                hdr.FontSize = 11
                hdr.FontWeight = FontWeights.Bold
                hdr.Foreground = self._brush((120, 125, 140))
                hdr.Margin = Thickness(8, 10, 0, 2)
                self.family_list_panel.Children.Add(hdr)
            self.family_list_panel.Children.Add(self._make_family_row(d))

    def _make_family_row(self, fam):
        from System.Windows import (Thickness, HorizontalAlignment, VerticalAlignment,
                                     FontWeights, TextTrimming, CornerRadius,
                                     GridLength, GridUnitType)
        from System.Windows.Controls import (Border, Grid, ColumnDefinition,
                                              Button, StackPanel, TextBlock,
                                              Orientation)

        outer = Border()
        outer.Background = self._brush(self.COLOR_BTN)
        outer.BorderBrush = self._brush((224, 226, 234))
        outer.BorderThickness = Thickness(1)
        outer.CornerRadius = CornerRadius(4)
        outer.Margin = Thickness(2, 2, 2, 2)
        outer.Padding = Thickness(10, 6, 8, 6)

        grid = Grid()
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(70)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(96)))

        name_tb = TextBlock()
        name_tb.Text = fam["name"]
        name_tb.FontSize = 13
        name_tb.FontWeight = FontWeights.SemiBold
        name_tb.Foreground = self._brush(self.COLOR_TEXT)
        name_tb.VerticalAlignment = VerticalAlignment.Center
        name_tb.TextTrimming = TextTrimming.CharacterEllipsis
        try:
            name_tb.ToolTip = fam["path"]
        except:
            pass
        Grid.SetColumn(name_tb, 0)
        grid.Children.Add(name_tb)

        load_btn = Button()
        load_btn.Content = "載入"
        load_btn.Margin = Thickness(4, 0, 4, 0)
        load_btn.Padding = Thickness(4, 4, 4, 4)
        load_btn.FontSize = 12
        load_btn.Click += lambda s, e, p=fam["path"]: self._raise_action("load", p)
        try:
            load_btn.ToolTip = "只載入到專案，不放置"
        except:
            pass
        Grid.SetColumn(load_btn, 1)
        grid.Children.Add(load_btn)

        place_btn = Button()
        place_btn.Content = "載入並放置"
        place_btn.Margin = Thickness(0, 0, 0, 0)
        place_btn.Padding = Thickness(4, 4, 4, 4)
        place_btn.FontSize = 12
        place_btn.Background = self._brush(self.COLOR_PRIMARY_SOFT)
        place_btn.FontWeight = FontWeights.SemiBold
        place_btn.Click += lambda s, e, p=fam["path"]: self._raise_action("place", p)
        try:
            place_btn.ToolTip = "載入後立刻進入 Revit 放置模式，在畫布上點一下即可擺放"
        except:
            pass
        Grid.SetColumn(place_btn, 2)
        grid.Children.Add(place_btn)

        outer.Child = grid
        return outer

    # ---- 事件 ----

    def _on_select_source(self, src):
        try:
            self.active_source_id = src["id"]
            self._update_source_highlight()
            self.main_title.Text = src.get("name", "")
            path = src.get("path", "")
            recursive = src.get("recursive", True)
            if not os.path.isdir(path):
                self.current_families = []
                self._show_family_placeholder(
                    "找不到資料夾：\n{}\n\n可能已被移動或網路磁碟未連線。".format(path)
                )
                return
            self.current_families = scan_folder_families(path, recursive)
            self._render_families()
        except Exception as ex:
            import traceback
            self._show_error("載入來源時發生錯誤：\n\n{}\n\n{}".format(
                ex, traceback.format_exc()))

    def _on_search_changed(self, sender, args):
        try:
            self.search_text = sender.Text or ""
            if self.active_source_id is not None:
                self._render_families()
        except:
            pass

    def _raise_action(self, action, path):
        """設定 handler 的請求並觸發 ExternalEvent。"""
        try:
            self._handler.action = action
            self._handler.payload = path
            self._ext_event.Raise()
        except Exception as ex:
            import traceback
            self._show_error("觸發動作時發生錯誤：\n\n{}\n\n{}".format(
                ex, traceback.format_exc()))

    def _on_add_folder(self, sender, args):
        try:
            import os as _os
            dlg = self._add_source_dialog_cls()
            try:
                dlg.Owner = self
            except:
                pass
            ok = dlg.ShowDialog()
            if not ok:   # 取消或關閉（DialogResult 非 True）
                return

            # 容忍使用者從檔案總管「複製為路徑」貼進來時包的引號
            folder = (dlg.path_text or "").strip().strip('"').strip()
            name = (dlg.name_text or "").strip()

            if not folder:
                return
            if not _os.path.isdir(folder):
                self._show_error(
                    "找不到資料夾或不是有效的路徑：\n{}\n\n"
                    "請確認路徑正確、網路磁碟已連線。".format(folder))
                return
            if not name:
                name = _os.path.basename(_os.path.normpath(folder)) or folder

            self._settings_mod.add_folder_source(
                self.settings, name, folder, recursive=True)
            ok2, err = self._settings_mod.save_settings(self.settings_path, self.settings)
            if not ok2:
                self._show_error("儲存設定檔失敗：{}".format(err))
                return
            self._refresh_source_list()
        except Exception as ex:
            import traceback
            self._show_error("新增來源時發生錯誤：\n\n{}\n\n{}".format(
                ex, traceback.format_exc()))

    def _on_remove_source(self, src):
        try:
            from System.Windows import MessageBox, MessageBoxButton, MessageBoxResult
            result = MessageBox.Show(
                "確定要移除來源「{}」？\n\n（只是從清單移除，不會刪除任何檔案）"
                .format(src.get("name", "")),
                "確認移除", MessageBoxButton.YesNo)
            if result != MessageBoxResult.Yes:
                return
            self._settings_mod.remove_source(self.settings, src["id"])
            self._settings_mod.save_settings(self.settings_path, self.settings)
            if self.active_source_id == src["id"]:
                self.active_source_id = None
                self.current_families = []
                self.main_title.Text = "請從左側選擇一個來源"
                self._show_family_placeholder("選擇來源後，這裡會列出資料夾中的族群 (.rfa)。")
            self._refresh_source_list()
        except Exception as ex:
            import traceback
            self._show_error("移除來源時發生錯誤：\n\n{}\n\n{}".format(
                ex, traceback.format_exc()))

    # ---- 工具 ----

    def _show_error(self, message):
        try:
            from System.Windows import MessageBox
            MessageBox.Show(message, "錯誤")
        except:
            self._show_family_placeholder("⚠️ {}".format(message))


# ---------------------------------------------------------------------------
# 新增來源對話框（可貼路徑 / 也可瀏覽）
# ---------------------------------------------------------------------------

class AddSourceDialog(Window):
    """Modal 小對話框：讓使用者直接貼上資料夾路徑，或用「瀏覽…」選。

    成功按「確定」後，外部讀取 self.path_text / self.name_text。
    """

    def __init__(self):
        self.path_text = ""
        self.name_text = ""
        self._build_ui()

    @staticmethod
    def _brush(rgb):
        from System.Windows.Media import SolidColorBrush, Color
        return SolidColorBrush(Color.FromRgb(rgb[0], rgb[1], rgb[2]))

    def _build_ui(self):
        from System.Windows import (Thickness, WindowStartupLocation, ResizeMode,
                                     FontWeights, HorizontalAlignment, VerticalAlignment,
                                     GridLength, GridUnitType)
        from System.Windows.Controls import (Grid, RowDefinition, ColumnDefinition,
                                              TextBlock, TextBox, Button, StackPanel,
                                              Orientation)

        self.Title = "新增資料夾來源"
        self.Width = 580
        self.Height = 250
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize
        self.Background = self._brush((245, 245, 250))

        root = Grid()
        root.Margin = Thickness(18, 16, 18, 16)
        for h in (24, 38, 12, 24, 36, 1, 44):
            rd = RowDefinition()
            if h == 1:
                rd.Height = GridLength(1, GridUnitType.Star)
            else:
                rd.Height = GridLength(h)
            root.RowDefinitions.Add(rd)

        # row0: 路徑標籤
        lbl1 = TextBlock()
        lbl1.Text = "資料夾路徑（可直接貼上）："
        lbl1.FontSize = 12
        lbl1.FontWeight = FontWeights.SemiBold
        lbl1.Foreground = self._brush((40, 40, 55))
        Grid.SetRow(lbl1, 0)
        root.Children.Add(lbl1)

        # row1: 路徑輸入框 + 瀏覽
        path_grid = Grid()
        path_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        path_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(88)))

        self.path_box = TextBox()
        self.path_box.FontSize = 12
        self.path_box.VerticalContentAlignment = VerticalAlignment.Center
        self.path_box.Padding = Thickness(6, 4, 6, 4)
        self.path_box.Margin = Thickness(0, 0, 8, 0)
        Grid.SetColumn(self.path_box, 0)
        path_grid.Children.Add(self.path_box)

        browse_btn = Button()
        browse_btn.Content = "瀏覽…"
        browse_btn.FontSize = 12
        browse_btn.Click += self._on_browse
        Grid.SetColumn(browse_btn, 1)
        path_grid.Children.Add(browse_btn)

        Grid.SetRow(path_grid, 1)
        root.Children.Add(path_grid)

        # row3: 名稱標籤
        lbl2 = TextBlock()
        lbl2.Text = "顯示名稱（留空則用資料夾名稱）："
        lbl2.FontSize = 12
        lbl2.FontWeight = FontWeights.SemiBold
        lbl2.Foreground = self._brush((40, 40, 55))
        Grid.SetRow(lbl2, 3)
        root.Children.Add(lbl2)

        # row4: 名稱輸入框
        self.name_box = TextBox()
        self.name_box.FontSize = 12
        self.name_box.VerticalContentAlignment = VerticalAlignment.Center
        self.name_box.Padding = Thickness(6, 4, 6, 4)
        Grid.SetRow(self.name_box, 4)
        root.Children.Add(self.name_box)

        # row6: 按鈕列
        btns = StackPanel()
        btns.Orientation = Orientation.Horizontal
        btns.HorizontalAlignment = HorizontalAlignment.Right
        btns.VerticalAlignment = VerticalAlignment.Center

        cancel_btn = Button()
        cancel_btn.Content = "取消"
        cancel_btn.FontSize = 12
        cancel_btn.Width = 84
        cancel_btn.Margin = Thickness(0, 0, 10, 0)
        cancel_btn.Padding = Thickness(6, 4, 6, 4)
        cancel_btn.IsCancel = True   # Esc 觸發
        cancel_btn.Click += self._on_cancel
        btns.Children.Add(cancel_btn)

        ok_btn = Button()
        ok_btn.Content = "確定"
        ok_btn.FontSize = 12
        ok_btn.Width = 84
        ok_btn.Padding = Thickness(6, 4, 6, 4)
        ok_btn.FontWeight = FontWeights.SemiBold
        ok_btn.Background = self._brush((99, 102, 241))
        ok_btn.Foreground = self._brush((255, 255, 255))
        ok_btn.IsDefault = True   # Enter 觸發
        ok_btn.Click += self._on_ok
        btns.Children.Add(ok_btn)

        Grid.SetRow(btns, 6)
        root.Children.Add(btns)

        self.Content = root
        try:
            self.path_box.Focus()
        except:
            pass

    def _on_browse(self, sender, args):
        try:
            import clr as _clr
            _clr.AddReference("System.Windows.Forms")
            import System.Windows.Forms as _WinForms
            dlg = _WinForms.FolderBrowserDialog()
            dlg.Description = "選擇本案族群／圖框所在的資料夾"
            dlg.ShowNewFolderButton = False
            # 若輸入框已有有效路徑，當作起始位置
            try:
                import os as _os
                cur = (self.path_box.Text or "").strip().strip('"')
                if cur and _os.path.isdir(cur):
                    dlg.SelectedPath = cur
            except:
                pass
            result = dlg.ShowDialog()
            if str(result) == "OK":
                self.path_box.Text = dlg.SelectedPath
        except Exception as ex:
            from System.Windows import MessageBox
            MessageBox.Show("無法開啟資料夾選擇對話框：{}".format(ex), "錯誤")

    def _on_ok(self, sender, args):
        path = (self.path_box.Text or "").strip().strip('"').strip()
        if not path:
            from System.Windows import MessageBox
            MessageBox.Show("請貼上或選擇一個資料夾路徑。", "尚未填寫")
            return
        self.path_text = self.path_box.Text
        self.name_text = self.name_box.Text
        self.DialogResult = True   # 關閉並讓 ShowDialog 回傳 True

    def _on_cancel(self, sender, args):
        self.DialogResult = False


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main():
    handler = FamilyActionHandler()
    ext_event = ExternalEvent.Create(handler)
    win = FamilyBrowserWindow(SETTINGS_FILE, ext_event, handler)
    win.Show()   # 非 Modal：開著也能操作 Revit、進入放置模式


if __name__ == '__main__':
    main()
