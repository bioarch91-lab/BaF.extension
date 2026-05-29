# -*- coding: utf-8 -*-
"""
需求書查閱 (Spec Reader) - Phase 1
===========================================================
側邊欄查閱統包需求書，目前支援 PDF。
Word/Excel 將在 Phase 1.5 加入（自動轉 PDF）。

使用流程:
  1. 點按鈕 → 開啟側邊視窗
  2. 第一次使用會詢問設定檔要存哪
  3. 點「新增文件」載入 PDF
  4. 在左側清單切換文件，右側用 WebView2 顯示
  5. 用瀏覽器原生 Ctrl+F 搜尋

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

# 註：OpenFileDialog 的載入移到 _get_open_file_dialog() 內，
# 因為 IronPython 在 pyRevit 環境下的命名空間解析有時很挑剔。

# 嘗試載入 WebView2（pyRevit 通常自帶或可從系統載入）
WEBVIEW2_AVAILABLE = False
WEBVIEW2_LOAD_ERROR = None
try:
    clr.AddReference("Microsoft.Web.WebView2.Core")
    clr.AddReference("Microsoft.Web.WebView2.Wpf")
    from Microsoft.Web.WebView2.Wpf import WebView2
    WEBVIEW2_AVAILABLE = True
except Exception as ex:
    WEBVIEW2_LOAD_ERROR = str(ex)

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
from System import Uri

from pyrevit import revit, forms, script

doc = revit.doc
output = script.get_output()


# ---------------------------------------------------------------------------
# 設定檔位置管理
# ---------------------------------------------------------------------------

# 我們在使用者 AppData 存一個「指引檔」，記錄使用者選的設定檔位置
APPDATA_DIR = os.path.join(os.environ.get("APPDATA", ""), "BaF_SpecReader")
POINTER_FILE = os.path.join(APPDATA_DIR, "settings_location.txt")


def get_saved_settings_path():
    """讀取上次使用的設定檔路徑。"""
    if not os.path.isfile(POINTER_FILE):
        return None
    try:
        with io.open(POINTER_FILE, "r", encoding="utf-8-sig") as f:
            path = f.read().strip()
        return path if path else None
    except:
        return None


def save_settings_path(path):
    """記憶使用者選的設定檔路徑。"""
    try:
        if not os.path.isdir(APPDATA_DIR):
            os.makedirs(APPDATA_DIR)
        # 路徑可能含中文，需用 UTF-8 寫，否則內建 open() 會以 ASCII 編碼失敗
        if isinstance(path, bytes):
            path = path.decode("utf-8")
        with io.open(POINTER_FILE, "w", encoding="utf-8") as f:
            f.write(path)
    except:
        pass


def ask_user_for_settings_path():
    """彈出對話框讓使用者選設定檔位置（首次使用或找不到時）。"""
    # 用 forms.save_file，較直觀
    suggested_dir = os.path.dirname(doc.PathName) if doc.PathName else ""
    suggested_name = "spec-reader.json"
    
    result = forms.save_file(
        file_ext="json",
        default_name=suggested_name,
        unc_paths=False,
        title="選擇設定檔的儲存位置（書籤、文件清單會存在這裡）"
    )
    return result


# ---------------------------------------------------------------------------
# WPF 主視窗
# ---------------------------------------------------------------------------

class SpecReaderWindow(Window):
    
    COLOR_BG = (245, 245, 250)
    COLOR_SIDEBAR = (250, 250, 254)
    COLOR_BTN = (255, 255, 255)
    COLOR_PRIMARY = (99, 102, 241)
    COLOR_TEXT = (30, 30, 40)
    COLOR_TEXT_LIGHT = (255, 255, 255)
    COLOR_ACTIVE = (224, 231, 255)
    
    def __init__(self, settings_path):
        self.settings_path = settings_path
        # 先 import 給下方使用
        import settings as _settings_mod
        self._settings_mod = _settings_mod
        self.settings = _settings_mod.load_settings(settings_path)
        self.doc_buttons = {}   # doc_id -> Button
        self.active_doc_id = None
        self.web_view = None
        # 把 WebView2 狀態存到 self：modeless 視窗在 main() 結束後模組全域
        # 會被清掉，事件回呼時讀不到 WEBVIEW2_AVAILABLE 等全域常數。
        self._webview2_available = WEBVIEW2_AVAILABLE
        self._webview2_load_error = WEBVIEW2_LOAD_ERROR
        self._WebView2 = WebView2 if WEBVIEW2_AVAILABLE else None

        self._build_ui()
        self._refresh_doc_list()
    
    @staticmethod
    def _brush(rgb):
        # 區域 import：此方法也會在事件回呼時被呼叫，那時模組全域已被清掉
        from System.Windows.Media import SolidColorBrush, Color
        return SolidColorBrush(Color.FromRgb(rgb[0], rgb[1], rgb[2]))
    
    def _build_ui(self):
        self.Title = "需求書查閱"
        self.Width = 1200
        self.Height = 800
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = self._brush(self.COLOR_BG)
        
        root = Grid()
        root.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(240)))
        root.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        # === 左側：文件清單 ===
        sidebar = self._build_sidebar()
        Grid.SetColumn(sidebar, 0)
        root.Children.Add(sidebar)
        
        # === 右側：內容區（WebView2 或提示） ===
        self.content_area = Grid()
        self.content_area.Background = self._brush((255, 255, 255))
        Grid.SetColumn(self.content_area, 1)
        root.Children.Add(self.content_area)
        
        # 預設顯示提示
        self._show_placeholder("請從左側清單選擇文件，或點「新增文件」載入新的 PDF。")
        
        self.Content = root
        
        # 視窗載入完成後再初始化 WebView2（要等視窗 handle 有了）
        self.Loaded += self._on_window_loaded
    
    def _build_sidebar(self):
        import os as _os
        border = Border()
        border.Background = self._brush(self.COLOR_SIDEBAR)
        border.BorderBrush = self._brush((220, 222, 230))
        border.BorderThickness = Thickness(0, 0, 1, 0)
        
        outer = Grid()
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))   # 標題
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))  # 文件清單
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))   # 新增按鈕
        outer.RowDefinitions.Add(RowDefinition(Height=GridLength(35)))   # 狀態列
        
        # 標題
        title = TextBlock()
        title.Text = "📄 需求文件"
        title.FontSize = 16
        title.FontWeight = FontWeights.Bold
        title.Foreground = self._brush(self.COLOR_TEXT)
        title.Margin = Thickness(16, 14, 0, 0)
        Grid.SetRow(title, 0)
        outer.Children.Add(title)
        
        # 文件清單
        scroll = ScrollViewer()
        scroll.Margin = Thickness(8, 0, 8, 0)
        self.doc_list_panel = StackPanel()
        self.doc_list_panel.Orientation = Orientation.Vertical
        scroll.Content = self.doc_list_panel
        Grid.SetRow(scroll, 1)
        outer.Children.Add(scroll)
        
        # 新增按鈕
        add_btn = Button()
        add_btn.Content = "+ 新增文件"
        add_btn.Margin = Thickness(12, 4, 12, 4)
        add_btn.Padding = Thickness(8, 8, 8, 8)
        add_btn.Background = self._brush(self.COLOR_PRIMARY)
        add_btn.Foreground = self._brush(self.COLOR_TEXT_LIGHT)
        add_btn.FontWeight = FontWeights.SemiBold
        add_btn.Click += self._on_add_document
        Grid.SetRow(add_btn, 2)
        outer.Children.Add(add_btn)
        
        # 狀態列：顯示設定檔位置
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
    
    def _on_window_loaded(self, sender, args):
        """視窗載入完成後嘗試建立 WebView2 元件。"""
        if not self._webview2_available:
            self._show_placeholder(
                "⚠️ WebView2 元件無法載入。\n\n"
                "錯誤訊息：{}\n\n"
                "請確認系統已安裝 Microsoft Edge WebView2 Runtime。"
                .format(self._webview2_load_error or "未知錯誤")
            )
            return

        try:
            self.web_view = self._WebView2()
            self.content_area.Children.Clear()
            self.content_area.Children.Add(self.web_view)
        except Exception as ex:
            self._show_placeholder(
                "⚠️ 建立 WebView2 元件失敗：\n\n{}\n\n"
                "可能是 Runtime 版本不相容。"
                .format(ex)
            )
            self.web_view = None
    
    def _show_placeholder(self, message):
        """在內容區顯示提示文字（沒有文件或載入失敗時）。"""
        # 區域 import：此方法會在事件回呼時被呼叫，那時模組全域已被清掉
        from System.Windows import (Thickness, HorizontalAlignment,
                                    VerticalAlignment, TextWrapping)
        from System.Windows.Controls import TextBlock
        self.content_area.Children.Clear()
        tb = TextBlock()
        tb.Text = message
        tb.FontSize = 13
        tb.Foreground = self._brush((100, 105, 115))
        tb.TextWrapping = TextWrapping.Wrap
        tb.HorizontalAlignment = HorizontalAlignment.Center
        tb.VerticalAlignment = VerticalAlignment.Center
        tb.Margin = Thickness(40)
        tb.MaxWidth = 500
        self.content_area.Children.Add(tb)
    
    def _refresh_doc_list(self):
        """重建文件清單側邊欄按鈕。"""
        # 區域 import：此方法會在事件回呼時被呼叫，那時模組全域已被清掉
        from System.Windows import Thickness, TextWrapping
        from System.Windows.Controls import TextBlock
        self.doc_list_panel.Children.Clear()
        self.doc_buttons.clear()

        docs = self.settings.get("documents", [])
        if not docs:
            empty = TextBlock()
            empty.Text = "（尚無文件，點下方「新增」載入）"
            empty.FontSize = 11
            empty.Foreground = self._brush((150, 155, 165))
            empty.Margin = Thickness(8, 12, 8, 0)
            empty.TextWrapping = TextWrapping.Wrap
            self.doc_list_panel.Children.Add(empty)
            return
        
        for d in docs:
            btn = self._make_doc_button(d)
            self.doc_buttons[d["id"]] = btn
            self.doc_list_panel.Children.Add(btn)
    
    def _make_doc_button(self, doc_data):
        """產生一個文件按鈕（含名稱、路徑提示、移除按鈕）。"""
        import os as _os
        # 區域 import：此方法會在事件回呼時被呼叫，那時模組全域已被清掉
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
        outer.Padding = Thickness(0)
        
        grid = Grid()
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(28)))
        
        # 主要點擊區（打開文件）
        main_btn = Button()
        main_btn.HorizontalContentAlignment = HorizontalAlignment.Stretch
        main_btn.Background = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))  # 透明
        main_btn.BorderThickness = Thickness(0)
        main_btn.Padding = Thickness(10, 6, 4, 6)
        main_btn.Click += lambda s, e, d=doc_data: self._open_document(d)
        
        sp = StackPanel()
        sp.Orientation = Orientation.Vertical
        
        t1 = TextBlock()
        t1.Text = doc_data.get("name", "(未命名)")
        t1.FontSize = 12
        t1.FontWeight = FontWeights.SemiBold
        t1.Foreground = self._brush(self.COLOR_TEXT)
        t1.TextTrimming = TextTrimming.CharacterEllipsis
        sp.Children.Add(t1)
        
        t2 = TextBlock()
        t2.Text = _os.path.basename(doc_data.get("path", ""))
        t2.FontSize = 10
        t2.Foreground = self._brush((130, 135, 145))
        t2.Margin = Thickness(0, 2, 0, 0)
        t2.TextTrimming = TextTrimming.CharacterEllipsis
        sp.Children.Add(t2)
        
        main_btn.Content = sp
        try:
            main_btn.ToolTip = doc_data.get("path", "")
        except:
            pass
        Grid.SetColumn(main_btn, 0)
        grid.Children.Add(main_btn)
        
        # 移除按鈕
        del_btn = Button()
        del_btn.Content = "✕"
        del_btn.FontSize = 11
        del_btn.Background = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
        del_btn.BorderThickness = Thickness(0)
        del_btn.Foreground = self._brush((160, 165, 175))
        del_btn.Click += lambda s, e, d=doc_data: self._remove_document(d)
        try:
            del_btn.ToolTip = "從清單中移除（不刪除原檔案）"
        except:
            pass
        Grid.SetColumn(del_btn, 1)
        grid.Children.Add(del_btn)
        
        outer.Child = grid
        outer.Tag = doc_data["id"]
        return outer
    
    # ---- 事件 ----
    
    def _get_open_file_dialog(self):
        """每次需要時才嘗試載入 OpenFileDialog。回傳 (dialog_instance, use_winforms, errors)。"""
        errors = []
        # 嘗試 WPF 版
        try:
            import Microsoft.Win32 as _MsWin32
            return _MsWin32.OpenFileDialog(), False, []
        except Exception as e:
            errors.append("Microsoft.Win32: {}".format(e))
        # 嘗試 WinForms 版
        try:
            clr.AddReference("System.Windows.Forms")
            import System.Windows.Forms as _WinForms
            return _WinForms.OpenFileDialog(), True, errors
        except Exception as e:
            errors.append("System.Windows.Forms: {}".format(e))
        return None, False, errors
    
    def _on_add_document(self, sender, args):
        """打開檔案選擇對話框，加入新文件（支援多選）。"""
        # 重要：所有外部模組都在 method 內 import，避免 pyRevit/IronPython
        # 在某些情況下 method 看不到 module 頂層名稱的怪問題
        try:
            import os as _os
            
            dlg, use_winforms, errors = self._get_open_file_dialog()
            if dlg is None:
                errs = "\n".join(errors) if errors else "(無詳細)"
                self._show_error(
                    "無法載入檔案選擇對話框（OpenFileDialog）。\n\n"
                    "載入嘗試失敗原因：\n{}".format(errs)
                )
                return
            
            dlg.Filter = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
            dlg.Title = "選擇需求書 PDF（可一次選多個，按住 Ctrl 或 Shift）"
            dlg.Multiselect = True
            
            result = dlg.ShowDialog()
            if use_winforms:
                if str(result) != "OK":
                    return
            else:
                if not result:
                    return
            
            # 取得所有選到的檔案路徑
            try:
                file_names = list(dlg.FileNames)
            except:
                # 退而求其次
                file_names = [dlg.FileName] if dlg.FileName else []
            
            if not file_names:
                return
            
            # 多檔時，預設名稱用檔名（不再逐一問），單檔才問
            added_docs = []
            skipped = []
            
            if len(file_names) == 1:
                path = file_names[0]
                if not _os.path.isfile(path):
                    self._show_error("檔案不存在：\n{}".format(path))
                    return
                default_name = _os.path.splitext(_os.path.basename(path))[0]
                name = self._prompt_text("請輸入顯示名稱：",
                                         default_name, "新增文件")
                if name is None:
                    return
                if not name.strip():
                    name = default_name
                new_doc = self._settings_mod.add_document(self.settings,
                                                    name.strip(), path)
                added_docs.append(new_doc)
            else:
                # 多檔：自動用檔名，不一個個問
                for path in file_names:
                    if not _os.path.isfile(path):
                        skipped.append((path, "檔案不存在"))
                        continue
                    name = _os.path.splitext(_os.path.basename(path))[0]
                    new_doc = self._settings_mod.add_document(self.settings, name, path)
                    added_docs.append(new_doc)
            
            # 一次儲存（不要每個都存）
            ok, err = self._settings_mod.save_settings(self.settings_path, self.settings)
            if not ok:
                self._show_error("儲存設定檔失敗：{}".format(err))
                return
            
            self._refresh_doc_list()
            
            # 多檔的話顯示摘要、單檔直接打開
            if len(added_docs) == 1 and not skipped:
                self._open_document(added_docs[0])
            else:
                msg = "成功加入 {} 份文件。".format(len(added_docs))
                if skipped:
                    msg += "\n\n跳過 {} 個：\n".format(len(skipped))
                    for p, reason in skipped[:5]:
                        msg += "- {} ({})\n".format(_os.path.basename(p), reason)
                # 自動打開第一個
                if added_docs:
                    self._open_document(added_docs[0])
                try:
                    from System.Windows import MessageBox
                    MessageBox.Show(msg, "新增完成")
                except:
                    pass
        except Exception as ex:
            import traceback
            self._show_error("新增文件時發生錯誤：\n\n{}\n\n--- 詳細 ---\n{}".format(
                ex, traceback.format_exc()))
    
    def _show_error(self, message):
        """顯示錯誤對話框（直接用 WPF MessageBox 比 forms.alert 可靠）。"""
        try:
            from System.Windows import MessageBox
            MessageBox.Show(message, "錯誤")
        except:
            # 退而求其次寫到 placeholder
            self._show_placeholder("⚠️ {}".format(message))
    
    def _prompt_text(self, prompt, default_value, title):
        """簡單的文字輸入對話框，用 VB.NET 的 InputBox（最簡單可靠的方式）。"""
        try:
            import clr as _clr
            _clr.AddReference("Microsoft.VisualBasic")
            import Microsoft.VisualBasic as _VB
            result = _VB.Interaction.InputBox(prompt, title, default_value or "", -1, -1)
            # InputBox 按取消會回傳空字串。為了區分「取消」和「輸入空字串」，
            # 我們把空字串視為取消（使用者通常不會刻意輸入空字串當名稱）
            if not result:
                return None
            return result
        except Exception as ex:
            # 後備：如果 VB InputBox 也不行，至少給個預設值不要炸
            self._show_error("無法顯示輸入框：{}\n將使用預設名稱。".format(ex))
            return default_value
    
    def _remove_document(self, doc_data):
        """從清單移除文件（不刪除原檔）。"""
        try:
            from System.Windows import MessageBox, MessageBoxButton, MessageBoxResult
            result = MessageBox.Show(
                "確定要從清單移除「{}」？\n\n（原檔案不會被刪除，只是不再顯示在清單）"
                .format(doc_data.get("name", "")),
                "確認移除",
                MessageBoxButton.YesNo
            )
            if result != MessageBoxResult.Yes:
                return
            
            self._settings_mod.remove_document(self.settings, doc_data["id"])
            self._settings_mod.save_settings(self.settings_path, self.settings)
            
            # 如果剛好正在看這個被移除的文件，回到 placeholder
            if self.active_doc_id == doc_data["id"]:
                self.active_doc_id = None
                self._show_placeholder("文件已從清單移除。")
            
            self._refresh_doc_list()
        except Exception as ex:
            import traceback
            self._show_error("移除文件時發生錯誤：\n\n{}\n\n--- 詳細 ---\n{}".format(
                ex, traceback.format_exc()))
    
    def _open_document(self, doc_data):
        """在 WebView2 中開啟文件。"""
        try:
            import os as _os
            from System import Uri as _Uri
            
            path = doc_data.get("path", "")
            if not _os.path.isfile(path):
                self._show_error(
                    "找不到檔案：\n{}\n\n檔案可能已被移動或刪除。"
                    .format(path)
                )
                return
            
            if not self._webview2_available or self.web_view is None:
                self._show_placeholder("WebView2 元件不可用，無法顯示文件。")
                return
            
            # 確保 WebView2 已經在內容區顯示
            if self.web_view not in [c for c in self.content_area.Children]:
                self.content_area.Children.Clear()
                self.content_area.Children.Add(self.web_view)
            
            # 用 file:// URI 載入 PDF，瀏覽器會自動用內建 PDF viewer
            uri = _Uri("file:///" + path.replace("\\", "/"))
            self.web_view.Source = uri
            self.active_doc_id = doc_data["id"]
            self._update_active_highlight()
        except Exception as ex:
            import traceback
            self._show_error("開啟文件時發生錯誤：\n\n{}\n\n--- 詳細 ---\n{}".format(
                ex, traceback.format_exc()))
    
    def _update_active_highlight(self):
        """把目前打開的文件按鈕高亮。"""
        for doc_id, btn in self.doc_buttons.items():
            if doc_id == self.active_doc_id:
                btn.Background = self._brush(self.COLOR_ACTIVE)
                btn.BorderBrush = self._brush(self.COLOR_PRIMARY)
            else:
                btn.Background = self._brush(self.COLOR_BTN)
                btn.BorderBrush = self._brush((220, 222, 230))


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main():
    # 確認設定檔位置
    settings_path = get_saved_settings_path()
    
    if not settings_path or not os.path.isfile(settings_path):
        # 第一次使用，或設定檔已被移動
        if settings_path:
            msg = ("找不到上次的設定檔位置：\n{}\n\n"
                   "請重新選擇設定檔位置。").format(settings_path)
            forms.alert(msg)
        else:
            forms.alert(
                "首次使用「需求書查閱」。\n\n"
                "請選擇設定檔（書籤、文件清單）的儲存位置。\n"
                "建議放在專案資料夾中，方便分享。"
            )
        
        new_path = ask_user_for_settings_path()
        if not new_path:
            return
        
        # 確保檔案存在（若不存在就建立空的）
        if not os.path.isfile(new_path):
            import settings as _settings_mod
            empty = _settings_mod.default_settings()
            ok, err = _settings_mod.save_settings(new_path, empty)
            if not ok:
                forms.alert("建立設定檔失敗：{}".format(err))
                return
        
        save_settings_path(new_path)
        settings_path = new_path
    
    # 開啟主視窗
    win = SpecReaderWindow(settings_path)
    win.Show()  # 非 Modal，可以邊看 Revit 邊看需求書


if __name__ == '__main__':
    main()
