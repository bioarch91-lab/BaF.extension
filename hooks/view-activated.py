# -*- coding: utf-8 -*-
"""ViewActivated 事件掛勾：只在「新開啟」一張圖紙時，跳出 modeless 小視窗列出其
「修正備註」。

兩個重點：
1. modeless（.Show()，非 ShowDialog）→ 視窗開著也能繼續編輯圖紙，不必按關閉。
2. 只在「這次才被開啟的視圖」觸發（例如在專案瀏覽器雙擊開圖）；純粹在已開啟的分頁
   之間切換不會跳。作法：比對「目前開啟的視圖分頁集合」與上一次的快照，若作用圖紙
   先前不在開啟集合中，視為剛開啟。

跨多次掛勾執行要保存的狀態（上次開啟集合、上一個備註視窗）放在 AppDomain 資料中，
沿用本擴充功能既有的 AppDomain 保存模式。改檔後需 Revit → pyRevit → Reload 重新註冊。
"""

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import AppDomain, IntPtr
from System.Diagnostics import Process
from System.Windows import (
    Window, Thickness, WindowStartupLocation, TextWrapping, FontWeights,
    SystemParameters, HorizontalAlignment)
from System.Windows.Controls import (
    TextBlock, ScrollViewer, ScrollBarVisibility, Grid, RowDefinition,
    Button)
from System.Windows.Interop import WindowInteropHelper

from pyrevit import EXEC_PARAMS, DB, revit

NOTE_PARAM = u"修正備註"
OPEN_KEY = "BAF_RedpenOpenViewIds"   # 上次「目前開啟的視圖」集合（逗號字串）
WIN_KEY = "BAF_RedpenNoteWindow"     # 上一個備註視窗（避免疊一堆）


# ---- 取得作用圖紙 / 讀備註 / 排版 -------------------------------------------

def _get_active_sheet():
    try:
        args = EXEC_PARAMS.event_args
    except Exception:
        return None
    view = getattr(args, "CurrentActiveView", None)
    if isinstance(view, DB.ViewSheet):
        return view
    return None


def _read_note(sheet):
    try:
        p = sheet.LookupParameter(NOTE_PARAM)
        if p is not None and p.StorageType == DB.StorageType.String:
            val = p.AsString()
            if val and val.strip():
                return val
    except Exception:
        pass
    return None


def _format_note(note):
    """把合併的多筆標示拆行條列（支援全形分號『；』與換行分隔）。"""
    items = []
    for chunk in note.replace(u"\n", u"；").split(u"；"):
        chunk = chunk.strip()
        if chunk:
            items.append(u"・" + chunk)
    return u"\n".join(items) if items else note.strip()


# ---- 「目前開啟的視圖」快照（用來分辨『新開啟』vs『切換』） ------------------

def _open_view_ids(uidoc):
    ids = set()
    try:
        for uiv in uidoc.GetOpenUIViews():
            try:
                ids.add(uiv.ViewId.IntegerValue)
            except Exception:
                pass
    except Exception:
        pass
    return ids


def _get_prev_ids():
    s = AppDomain.CurrentDomain.GetData(OPEN_KEY)
    if not s:
        return set()
    try:
        return set(int(x) for x in str(s).split(",") if x)
    except Exception:
        return set()


def _set_prev_ids(ids):
    try:
        AppDomain.CurrentDomain.SetData(
            OPEN_KEY, ",".join(str(i) for i in ids))
    except Exception:
        pass


# ---- modeless 備註視窗 -------------------------------------------------------

class _NoteWindow(Window):
    def __init__(self, num, name, body):
        self.Title = (u"修正備註　{}　{}".format(num, name)).strip()
        self.Width = 380
        self.Height = 260
        self.ShowInTaskbar = False
        self.WindowStartupLocation = WindowStartupLocation.Manual
        try:
            wa = SystemParameters.WorkArea
            self.Left = wa.Right - self.Width - 24
            self.Top = wa.Top + 80
        except Exception:
            pass

        root = Grid()
        root.Margin = Thickness(14, 12, 14, 12)
        root.RowDefinitions.Add(RowDefinition())
        root.RowDefinitions.Add(RowDefinition())
        root.RowDefinitions.Add(RowDefinition())
        root.RowDefinitions[0].Height = self._auto()
        root.RowDefinitions[1].Height = self._star()
        root.RowDefinitions[2].Height = self._auto()

        head = TextBlock()
        head.Text = u"{}　{}".format(num, name).strip()
        head.FontWeight = FontWeights.Bold
        head.FontSize = 13
        head.TextWrapping = TextWrapping.Wrap
        head.Margin = Thickness(0, 0, 0, 8)
        Grid.SetRow(head, 0)
        root.Children.Add(head)

        sv = ScrollViewer()
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        txt = TextBlock()
        txt.Text = body
        txt.FontSize = 12
        txt.TextWrapping = TextWrapping.Wrap
        sv.Content = txt
        Grid.SetRow(sv, 1)
        root.Children.Add(sv)

        close = Button()
        close.Content = u"關閉"
        close.Padding = Thickness(14, 4, 14, 4)
        close.Margin = Thickness(0, 8, 0, 0)
        close.HorizontalAlignment = HorizontalAlignment.Right
        close.Click += self._on_close
        Grid.SetRow(close, 2)
        root.Children.Add(close)

        self.Content = root
        self.Closed += self._on_closed

    @staticmethod
    def _auto():
        from System.Windows import GridLength, GridUnitType
        return GridLength(1, GridUnitType.Auto)

    @staticmethod
    def _star():
        from System.Windows import GridLength, GridUnitType
        return GridLength(1, GridUnitType.Star)

    def _on_close(self, sender, args):
        self.Close()

    def _on_closed(self, sender, args):
        try:
            if AppDomain.CurrentDomain.GetData(WIN_KEY) is self:
                AppDomain.CurrentDomain.SetData(WIN_KEY, None)
        except Exception:
            pass


def _show_note_window(sheet, note):
    # 關掉上一個備註視窗（避免一直疊上去）
    try:
        prevwin = AppDomain.CurrentDomain.GetData(WIN_KEY)
        if prevwin is not None:
            prevwin.Close()
    except Exception:
        pass

    win = _NoteWindow(sheet.SheetNumber or u"", sheet.Name or u"",
                      _format_note(note))
    # 以 Revit 主視窗為 Owner → 浮在 Revit 上方但「不阻擋」操作（modeless）
    try:
        h = Process.GetCurrentProcess().MainWindowHandle
        if h != IntPtr.Zero:
            WindowInteropHelper(win).Owner = h
    except Exception:
        pass
    # 存進 AppDomain 保持參考，避免掛勾結束後被回收
    try:
        AppDomain.CurrentDomain.SetData(WIN_KEY, win)
    except Exception:
        pass
    win.Show()


# ---- 判斷此次啟用是否「從專案瀏覽器開圖」 ----------------------------------
# 訊號：事件當下游標落在 (a) 瀏覽器樹狀清單(SysTreeView32) → 雙擊開圖，或
#       (b) 快顯選單(#32768/menu) → 右鍵→「開啟」。
# 瀏覽器單擊只選取、不觸發 ViewActivated；雙擊/右鍵開啟才會；分頁列切換不經樹/選單。
# 因此「圖紙被啟用 + 游標在樹/選單上」≒『從瀏覽器開圖』(新開/已開/關後重開都命中)。
#
# 主要用 Win32 WindowFromPoint+GetClassName(經 Reflection.Emit 宣告 P/Invoke，
# 此 IronPython 無 ctypes)；若取不到視窗(回 0)再退回 UIAutomation。

NATIVE_KEY = "BAF_Native_U32"
BROWSER_CLASS_HINTS = (u"treeview", u"#32768", u"menu")


def _native():
    """建立(並快取於 AppDomain)含 WindowFromPoint/GetClassNameW/GetParent 的型別。"""
    u = AppDomain.CurrentDomain.GetData(NATIVE_KEY)
    if u is not None:
        return u
    import System
    from System import IntPtr, Int32, Int64, Type
    from System.Reflection import (AssemblyName, MethodAttributes,
                                   CallingConventions, TypeAttributes)
    from System.Reflection.Emit import AssemblyBuilderAccess
    from System.Runtime.InteropServices import CallingConvention, CharSet
    from System.Text import StringBuilder
    ab = AppDomain.CurrentDomain.DefineDynamicAssembly(
        AssemblyName("BafNativeU32"), AssemblyBuilderAccess.Run)
    mod = ab.DefineDynamicModule("M")
    tb = mod.DefineType("U32", TypeAttributes.Public)
    attrs = (MethodAttributes.Public | MethodAttributes.Static |
             MethodAttributes.PinvokeImpl)
    SB = clr.GetClrType(StringBuilder)
    # x64：POINT(兩個 int32) 以 by-value 傳遞 = 單一 64-bit；低 32=x、高 32=y
    tb.DefinePInvokeMethod("WindowFromPoint", "user32.dll", attrs,
        CallingConventions.Standard, IntPtr, System.Array[Type]([Int64]),
        CallingConvention.Winapi, CharSet.Auto)
    tb.DefinePInvokeMethod("GetClassNameW", "user32.dll", attrs,
        CallingConventions.Standard, Int32,
        System.Array[Type]([IntPtr, SB, Int32]),
        CallingConvention.Winapi, CharSet.Unicode)
    tb.DefinePInvokeMethod("GetParent", "user32.dll", attrs,
        CallingConventions.Standard, IntPtr, System.Array[Type]([IntPtr]),
        CallingConvention.Winapi, CharSet.Auto)
    u = tb.CreateType()
    AppDomain.CurrentDomain.SetData(NATIVE_KEY, u)
    return u


def _ptr_zero(h):
    try:
        return h is None or h.ToInt64() == 0
    except Exception:
        return True


def _win32_classes_under_cursor():
    """游標下視窗往上的類別名稱清單；取不到(非互動桌面/失敗)回 None。"""
    try:
        import System
        from System import Int64
        from System.Text import StringBuilder
        clr.AddReference("System.Windows.Forms")
        from System.Windows.Forms import Control
        u = _native()
        wfp = u.GetMethod("WindowFromPoint")
        gcn = u.GetMethod("GetClassNameW")
        gp = u.GetMethod("GetParent")
        pos = Control.MousePosition
        x = int(pos.X) & 0xFFFFFFFF
        y = int(pos.Y) & 0xFFFFFFFF
        packed = Int64((y << 32) | x)
        h = wfp.Invoke(None, System.Array[object]([packed]))
        if _ptr_zero(h):
            return None  # 取不到視窗 → 視為失敗，交給後備
        names, depth = [], 0
        while not _ptr_zero(h) and depth < 12:
            sb = StringBuilder(256)
            gcn.Invoke(None, System.Array[object]([h, sb, 256]))
            names.append(sb.ToString() or u"")
            h = gp.Invoke(None, System.Array[object]([h]))
            depth += 1
        return names
    except Exception:
        return None


def _browser_check_win32():
    """True/False = Win32 判定到/沒判定到；None = Win32 取不到(交給後備)。"""
    names = _win32_classes_under_cursor()
    if names is None:
        return None
    for cls in names:
        c = (cls or u"").lower()
        if any(h in c for h in BROWSER_CLASS_HINTS):
            return True
    return False


def _browser_check_uia():
    """後備：UIAutomation 取游標下元素(ControlType/ClassName)。"""
    try:
        clr.AddReference("System.Windows.Forms")
        clr.AddReference("UIAutomationClient")
        clr.AddReference("UIAutomationTypes")
        from System.Windows.Forms import Control
        from System.Windows.Automation import (
            AutomationElement, TreeWalker, ControlType)
        from System.Windows import Point
        pos = Control.MousePosition
        el = AutomationElement.FromPoint(Point(float(pos.X), float(pos.Y)))
        walker = TreeWalker.ControlViewWalker
        match_types = (ControlType.Tree, ControlType.TreeItem,
                       ControlType.Menu, ControlType.MenuItem)
        depth = 0
        while el is not None and depth < 8:
            try:
                ct = el.Current.ControlType
            except Exception:
                ct = None
            try:
                cls = (el.Current.ClassName or u"").lower()
            except Exception:
                cls = u""
            if ct in match_types:
                return True
            if any(h in cls for h in BROWSER_CLASS_HINTS):
                return True
            try:
                el = walker.GetParent(el)
            except Exception:
                el = None
            depth += 1
    except Exception:
        pass
    return False


def _opened_from_browser():
    r = _browser_check_win32()
    if r is not None:
        return r
    return _browser_check_uia()


# ---- 除錯記錄（可選）：在 hooks 資料夾放一個空檔 `_debug_on` 即啟用 ----------
# 啟用後，每次作用視圖改變會把判斷過程寫到 %TEMP%\baf_viewhook_debug.txt（覆蓋），
# 用來在無法遠端實測時，讓使用者回報「游標下到底是什麼元素、為何沒觸發」。

def _debug_enabled():
    try:
        import os
        return os.path.exists(os.path.join(os.path.dirname(__file__), "_debug_on"))
    except Exception:
        return False


def _uia_chain_desc():
    """UIAutomation：游標下元素往上 8 層的 ControlType:ClassName 串接。"""
    try:
        clr.AddReference("System.Windows.Forms")
        clr.AddReference("UIAutomationClient")
        clr.AddReference("UIAutomationTypes")
        from System.Windows.Forms import Control
        from System.Windows.Automation import AutomationElement, TreeWalker
        from System.Windows import Point
        pos = Control.MousePosition
        el = AutomationElement.FromPoint(Point(float(pos.X), float(pos.Y)))
        walker = TreeWalker.ControlViewWalker
        parts, depth = [], 0
        while el is not None and depth < 8:
            try:
                ctn = el.Current.ControlType.ProgrammaticName
            except Exception:
                ctn = u"?"
            try:
                cls = el.Current.ClassName or u""
            except Exception:
                cls = u""
            parts.append(u"{}:{}".format(ctn, cls))
            try:
                el = walker.GetParent(el)
            except Exception:
                el = None
            depth += 1
        return u" > ".join(parts) if parts else u"(none)"
    except Exception as ex:
        return u"ERR:" + unicode(ex)


def _cursor_chain_desc():
    """除錯用：Win32 與 UIAutomation 兩邊抓到的游標下視窗類別。"""
    w = _win32_classes_under_cursor()
    win = (u"WIN32=" + u" > ".join(w)) if w is not None else u"WIN32=None"
    return win + u"  ||  UIA=" + _uia_chain_desc()


def _debug_log(is_sheet, sheet, note, from_browser, newly_opened, do_show, prev, cur):
    try:
        import os
        import io
        from System import Environment
        path = os.path.join(
            Environment.GetEnvironmentVariable("TEMP") or u".",
            "baf_viewhook_debug.txt")
        lines = [
            u"is_sheet={}".format(is_sheet),
            u"sheet={}".format(
                (u"{} {}".format(sheet.SheetNumber, sheet.Name)) if is_sheet else u"-"),
            u"sheet_id={}".format(sheet.Id.IntegerValue if is_sheet else u"-"),
            u"has_note={}".format(bool(note)),
            u"from_browser={}".format(from_browser),
            u"newly_opened={}".format(newly_opened),
            u"prev_ids={}".format(sorted(prev)),
            u"cur_ids={}".format(sorted(cur)),
            u"cursor_chain={}".format(_cursor_chain_desc()),
            u"=> SHOW={}".format(do_show),
        ]
        with io.open(path, "w", encoding="utf-8") as f:
            f.write(u"\n".join(lines))
    except Exception:
        pass


# ---- 主流程 -----------------------------------------------------------------

def main():
    uidoc = revit.uidoc
    if uidoc is None:
        return

    sheet = _get_active_sheet()

    # 每次都更新「目前開啟的視圖」快照（不論是不是圖紙）
    prev = _get_prev_ids()
    cur = _open_view_ids(uidoc)
    _set_prev_ids(cur)

    is_sheet = sheet is not None
    newly_opened = bool(is_sheet and (sheet.Id.IntegerValue not in prev))
    note = _read_note(sheet) if is_sheet else None

    # 觸發條件（任一成立就跳）：
    #   1. 從專案瀏覽器開圖 → 游標在瀏覽器樹上(雙擊) 或 快顯選單上(右鍵→開啟)；
    #      不管該分頁是新開、已開、或關掉後重開都算。
    #   2. 這次才新開啟的分頁（先前不在開啟集合中）→ 雙保險，涵蓋偵測不到游標時。
    # 純粹在已開分頁之間切換：游標在分頁列、不經樹/選單，且分頁早已開 → 不跳。
    debug = _debug_enabled()
    from_browser = _opened_from_browser() if (is_sheet and (note or debug)) else False
    do_show = bool(is_sheet and note and (from_browser or newly_opened))

    if debug:
        _debug_log(is_sheet, sheet, note, from_browser, newly_opened, do_show, prev, cur)

    if do_show:
        _show_note_window(sheet, note)


if __name__ == "__main__":
    main()
