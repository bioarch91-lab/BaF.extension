# -*- coding: utf-8 -*-
"""pyRevit 啟動腳本：註冊「修正備註」Dockable 側邊面板，並訂閱 ViewActivated
在切換視圖時更新面板內容。

dockable pane 必須在應用程式啟動(pyRevit 載入)時註冊 → 放在 startup.py。
面板是常駐、正規 context，可穩定顯示/互動。
（浮動彈窗 hooks/view-activated.py 目前仍保留，與此面板並存。）

相容性：只用 pyRevit forms + Revit API 事件，Revit 2024/2025 皆可。
"""

from pyrevit import forms, HOST_APP, DB, script

import baf_redpen_panel

logger = script.get_logger()


def _on_view_activated(sender, args):
    try:
        panel = baf_redpen_panel.get_panel()
        if panel is None:
            return
        view = getattr(args, "CurrentActiveView", None)
        panel.show_sheet(view if isinstance(view, DB.ViewSheet) else None)
    except Exception:
        pass


try:
    if not forms.is_registered_dockable_panel(baf_redpen_panel.BafRedpenPanel):
        _panel = forms.register_dockable_panel(
            baf_redpen_panel.BafRedpenPanel, default_visible=True)
        baf_redpen_panel.set_panel(_panel)
        # 切換視圖時更新面板
        try:
            HOST_APP.uiapp.ViewActivated += _on_view_activated
        except Exception as ev_ex:
            logger.debug("BaF: 無法訂閱 ViewActivated：{}".format(ev_ex))
    else:
        logger.debug("BaF: 修正備註面板已註冊，略過。")
except Exception as ex:
    logger.debug("BaF: 註冊修正備註面板失敗：{}".format(ex))
