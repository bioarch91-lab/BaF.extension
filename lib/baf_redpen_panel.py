# -*- coding: utf-8 -*-
"""BaF 修正備註 Dockable 側邊面板（互動版）。

常駐側邊，切到哪張圖就顯示該圖「修正備註」逐條列出，每條可：
  - 勾「完成」→ 該條前面加「已完成」標記（匯出 Google Sheet 後該條反灰）。
  - 填「回覆」→ 該條後面接「｜回覆：…」。
按「儲存」把整份修正備註重新編碼寫回 Revit 參數（走 ExternalEvent，正規 API context，
與 pyRevit 內建 dockable 工具相同做法 → 穩定，不會像事件掛勾彈窗那樣崩）。

相容 Revit 2024/2025：只用 pyRevit forms/UI/DB + WPF，無 .NET 版本相依 import。
"""

import os
import tempfile
import clr
from pyrevit import forms, DB, UI


def _dbg(msg, append=True):
    """把除錯訊息寫到 %TEMP%\\baf_panel_debug.txt（診斷儲存流程用）。"""
    try:
        import io
        p = os.path.join(tempfile.gettempdir(), "baf_panel_debug.txt")
        with io.open(p, "a" if append else "w", encoding="utf-8") as f:
            f.write(unicode(msg) + u"\n")
    except Exception:
        pass

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
from System.Windows.Controls import (
    StackPanel, CheckBox, TextBlock, TextBox, Border, Orientation)
from System.Windows import Thickness, TextWrapping, VerticalAlignment
from System.Windows.Media import SolidColorBrush, Color

NOTE_PARAM = u"修正備註"
DONE_MARK = u"已完成"
REPLY_SEP = u"｜回覆："
ITEM_SEP = u"；"

_GREY = SolidColorBrush(Color.FromRgb(150, 150, 155))
_BLACK = SolidColorBrush(Color.FromRgb(30, 30, 40))
_GREEN = SolidColorBrush(Color.FromRgb(34, 139, 34))
_RED = SolidColorBrush(Color.FromRgb(200, 60, 60))
_LBL = SolidColorBrush(Color.FromRgb(110, 115, 125))
_LINE = SolidColorBrush(Color.FromRgb(225, 228, 233))

# 面板實例（register 後由 startup.py 存起來，供 ViewActivated 事件更新）
_INSTANCE = None


def set_panel(p):
    global _INSTANCE
    _INSTANCE = p


def get_panel():
    return _INSTANCE


# ---- 修正備註字串 ↔ 條目 [{text, done, reply}] ------------------------------

def _parse_items(note):
    items = []
    if not note:
        return items
    for chunk in unicode(note).replace(u"\r", u"").replace(u"\n", ITEM_SEP).split(ITEM_SEP):
        s = chunk.strip()
        if not s:
            continue
        done = False
        if s.startswith(DONE_MARK):
            done = True
            s = s[len(DONE_MARK):].strip()
        reply = u""
        idx = s.find(REPLY_SEP)
        if idx >= 0:
            reply = s[idx + len(REPLY_SEP):].strip()
            s = s[:idx].strip()
        items.append({"text": s, "done": done, "reply": reply})
    return items


def _sanitize(t):
    return (unicode(t if t is not None else u"").replace(ITEM_SEP, u"，")
            .replace(u"｜", u"/").replace(u"\r", u" ").replace(u"\n", u" ").strip())


def _encode_items(items):
    out = []
    for it in items:
        text = _sanitize(it["text"])
        reply = _sanitize(it["reply"])
        seg = (DONE_MARK + u" " if it["done"] else u"") + text
        if reply:
            seg += REPLY_SEP + reply
        out.append(seg)
    return ITEM_SEP.join(out)


def _read_note(sheet):
    try:
        p = sheet.LookupParameter(NOTE_PARAM)
        if p is not None and p.StorageType == DB.StorageType.String:
            v = p.AsString()
            if v and v.strip():
                return v
    except Exception:
        pass
    return None


# ---- 寫回 Revit 參數：ExternalEvent（正規 context，穩定） --------------------

class _WriteHandler(UI.IExternalEventHandler):
    def __init__(self):
        self.queue = []   # [(uid, note)]

    def Execute(self, uiapp):
        try:
            doc = uiapp.ActiveUIDocument.Document
        except Exception as ex:
            _dbg(u"EXEC no-doc: {}".format(ex))
            return
        pending = self.queue
        self.queue = []
        for uid, note in pending:
            t = None
            try:
                el = doc.GetElement(uid)
                if el is None:
                    _dbg(u"EXEC el=None uid={}".format(uid))
                    continue
                p = el.LookupParameter(NOTE_PARAM)
                _dbg(u"EXEC found_param={} readonly={}".format(
                    p is not None, (p.IsReadOnly if p is not None else u"?")))
                t = DB.Transaction(doc, "BaF 更新修正備註")
                t.Start()
                setok = None
                if p is not None and not p.IsReadOnly:
                    setok = p.Set(note)
                t.Commit()
                _dbg(u"EXEC setok={} committed".format(setok))
            except Exception as ex:
                _dbg(u"EXEC ERR: {}".format(ex))
                try:
                    if t is not None:
                        t.RollBack()
                except Exception:
                    pass

    def GetName(self):
        return "BaF 更新修正備註(面板)"


class BafRedpenPanel(forms.WPFPanel):
    panel_title = u"修正備註"
    panel_id = "b2f7c1a0-8e4d-4c9a-9f21-3a6d5e8b7c40"
    panel_source = os.path.join(os.path.dirname(__file__), "baf_redpen_pane.xaml")

    def __init__(self):
        forms.WPFPanel.__init__(self)
        self._uid = None
        self._rows = []      # [{text, chk, reply}]
        # 依 pyRevit 內建 dockable 工具做法：在建構(啟動註冊)時就建立 ExternalEvent
        self._handler = _WriteHandler()
        try:
            self._event = UI.ExternalEvent.Create(self._handler)
        except Exception as ex:
            self._event = None
            _dbg(u"INIT ExternalEvent.Create 失敗: {}".format(ex))
        try:
            self.SaveButton.Click += self._on_save
        except Exception:
            pass
        self.show_sheet(None)

    # ---- 顯示 ----
    def show_sheet(self, sheet):
        try:
            self._rows = []
            try:
                self.ItemsPanel.Children.Clear()
            except Exception:
                pass
            if not isinstance(sheet, DB.ViewSheet):
                self._uid = None
                self.HeadText.Text = u""
                self.StatusText.Text = u"（切換到一張圖紙即顯示其修正備註）"
                try:
                    self.SaveButton.IsEnabled = False
                except Exception:
                    pass
                return
            self._uid = sheet.UniqueId
            self.HeadText.Text = (u"{}　{}".format(
                sheet.SheetNumber or u"", sheet.Name or u"")).strip()
            self.StatusText.Text = u""
            try:
                self.SaveButton.IsEnabled = True
            except Exception:
                pass
            items = _parse_items(_read_note(sheet))
            if not items:
                lbl = TextBlock()
                lbl.Text = u"（此圖沒有修正備註）"
                lbl.Foreground = _LBL
                lbl.TextWrapping = TextWrapping.Wrap
                lbl.FontSize = 14
                self.ItemsPanel.Children.Add(lbl)
                return
            for it in items:
                self.ItemsPanel.Children.Add(self._make_row(it))
        except Exception:
            pass

    def _make_row(self, it):
        box = Border()
        box.BorderBrush = _LINE
        box.BorderThickness = Thickness(0, 0, 0, 1)
        box.Padding = Thickness(0, 6, 0, 6)
        col = StackPanel()

        chk = CheckBox()
        chk.IsChecked = bool(it["done"])
        chk.VerticalAlignment = VerticalAlignment.Top
        ctext = TextBlock()
        ctext.Text = it["text"]
        ctext.TextWrapping = TextWrapping.Wrap
        ctext.FontSize = 15
        chk.Content = ctext
        col.Children.Add(chk)

        rr = StackPanel()
        rr.Orientation = Orientation.Horizontal
        rr.Margin = Thickness(20, 4, 0, 0)
        rlbl = TextBlock()
        rlbl.Text = u"回覆："
        rlbl.FontSize = 13
        rlbl.Foreground = _LBL
        rlbl.VerticalAlignment = VerticalAlignment.Center
        rr.Children.Add(rlbl)
        rbox = TextBox()
        rbox.Text = it["reply"]
        rbox.FontSize = 14
        rbox.MinWidth = 180
        rbox.Padding = Thickness(3, 2, 3, 2)
        rr.Children.Add(rbox)
        col.Children.Add(rr)

        box.Child = col
        self._rows.append({"text": it["text"], "chk": chk, "reply": rbox})
        return box

    # ---- 儲存 → ExternalEvent 寫回 Revit ----
    def _on_save(self, sender, args):
        if not self._uid:
            _dbg(u"SAVE 中止：self._uid 為空", append=False)
            return
        items = []
        for r in self._rows:
            items.append({"text": r["text"],
                          "done": bool(r["chk"].IsChecked),
                          "reply": r["reply"].Text or u""})
        note = _encode_items(items)
        _dbg(u"SAVE uid={} note={}".format(self._uid, repr(note)), append=False)
        try:
            if self._event is None:
                self._handler = _WriteHandler()
                self._event = UI.ExternalEvent.Create(self._handler)
            self._handler.queue.append((self._uid, note))
            raised = self._event.Raise()
            _dbg(u"SAVE Raise()={}".format(raised))
            self.StatusText.Foreground = _GREEN
            self.StatusText.Text = u"✔ 已送出儲存（下次匯出：完成反灰、回覆帶上）"
        except Exception as ex:
            _dbg(u"SAVE ERR: {}".format(ex))
            self.StatusText.Foreground = _RED
            self.StatusText.Text = u"✗ 儲存失敗：{}".format(ex)
