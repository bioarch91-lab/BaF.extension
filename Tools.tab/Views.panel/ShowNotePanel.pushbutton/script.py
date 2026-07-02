# -*- coding: utf-8 -*-
"""開啟／顯示「修正備註」Dockable 側邊面板。

面板由擴充功能根目錄的 startup.py 在 Revit 啟動時註冊(dockable pane 只能在啟動時
註冊)。本按鈕用 forms.open_dockable_panel 把它顯示出來 —— 不小心關掉後可重新開啟。

相容 Revit 2024/2025（只用 pyRevit forms API）。
"""

from pyrevit import forms

# 必須與 lib/baf_redpen_panel.py 的 panel_id 一致
PANEL_ID = "b2f7c1a0-8e4d-4c9a-9f21-3a6d5e8b7c40"

try:
    forms.open_dockable_panel(PANEL_ID)
except Exception as ex:
    forms.alert(
        u"無法開啟「修正備註」面板：\n{}\n\n"
        u"若面板從未出現過，請將 Revit 完全關閉再重新開啟"
        u"（面板需在 Revit 啟動時註冊，只按 pyRevit Reload 不夠）。".format(ex))
