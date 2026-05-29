# -*- coding: utf-8 -*-
"""
視圖對齊工具 (Align Viewports Across Sheets)
================================================
功能:
  以使用者選定的「樣板圖紙」為基準，將其他目標圖紙上
  「同類型視圖」的 Viewport 位置對齊到樣板上的相同座標。

對齊邏輯:
  Viewport 的「匹配鍵 (matching key)」預設為:
      (View Family, View Type Name)
  例如所有「平面圖 / 1-100 平面」會被視為同一類視圖。
  你可以在 build_match_key() 裡自行調整匹配邏輯。

使用方式:
  1. 先在某張圖紙上把 Viewport 排到理想位置 (作為樣板)
  2. 執行此腳本
  3. 第一個對話框: 選擇樣板圖紙
  4. 第二個對話框: 多選要套用的目標圖紙
  5. 完成

作者: 九典聯合 / BIM 工具
"""

from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()


# ---------------------------------------------------------------------------
# 工具函式
# ---------------------------------------------------------------------------

def get_all_sheets(document):
    """取得專案中所有圖紙 (排除 placeholder sheets)。"""
    sheets = DB.FilteredElementCollector(document) \
        .OfClass(DB.ViewSheet) \
        .ToElements()
    return [s for s in sheets if not s.IsPlaceholder]


def build_match_key(viewport, document):
    """
    為一個 Viewport 建立匹配鍵。
    預設使用 (ViewFamily, ViewType Name)。
    如果你想用其他規則 (例如 View Name 前綴、自訂參數)，改這裡。
    """
    view = document.GetElement(viewport.ViewId)
    if view is None:
        return None

    view_type = document.GetElement(view.GetTypeId())
    view_type_name = view_type.get_Parameter(
        DB.BuiltInParameter.ALL_MODEL_TYPE_NAME
    ).AsString() if view_type else "Unknown"

    family = str(view.ViewType)  # FloorPlan / CeilingPlan / Elevation ...

    return (family, view_type_name)


def collect_viewport_positions(sheet, document):
    """
    回傳 dict: { match_key: BoxCenter (XYZ) }
    若同一張圖紙上有多個相同 match_key 的 Viewport，
    只保留第一個 (通常代表配置不規範，會在 log 提示)。
    """
    positions = {}
    duplicates = []

    vp_ids = sheet.GetAllViewports()
    for vp_id in vp_ids:
        vp = document.GetElement(vp_id)
        key = build_match_key(vp, document)
        if key is None:
            continue
        if key in positions:
            duplicates.append(key)
            continue
        positions[key] = vp.GetBoxCenter()

    return positions, duplicates


def align_sheet_to_template(target_sheet, template_positions, document):
    """
    將目標圖紙上的 Viewport 對齊到 template_positions 中對應的座標。
    回傳 (moved_count, skipped_keys)。
    """
    moved = 0
    skipped = []

    vp_ids = target_sheet.GetAllViewports()
    for vp_id in vp_ids:
        vp = document.GetElement(vp_id)
        key = build_match_key(vp, document)
        if key is None or key not in template_positions:
            skipped.append((vp_id, key))
            continue

        current_center = vp.GetBoxCenter()
        target_center = template_positions[key]
        translation = target_center - current_center

        # 若位移量極小則略過 (避免不必要的交易紀錄)
        if translation.GetLength() < 1e-6:
            continue

        DB.ElementTransformUtils.MoveElement(document, vp_id, translation)
        moved += 1

    return moved, skipped


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main():
    sheets = get_all_sheets(doc)
    if not sheets:
        forms.alert("專案中沒有可用的圖紙。", exitscript=True)

    # --- 包一層 wrapper 讓選單能顯示「圖號 - 圖名」---
    class SheetOption(object):
        def __init__(self, sheet):
            self.sheet = sheet
            self.label = "{} - {}".format(sheet.SheetNumber, sheet.Name)

    # --- 1. 選擇樣板圖紙 ---
    sheet_options = [SheetOption(s) for s in sorted(sheets, key=lambda s: s.SheetNumber)]
    template_opt = forms.SelectFromList.show(
        sheet_options,
        name_attr='label',
        multiselect=False,
        title='選擇「樣板圖紙」(以這張的視圖位置為基準)',
        button_name='設為樣板'
    )
    if not template_opt:
        script.exit()
    template_sheet = template_opt.sheet

    template_positions, dup_keys = collect_viewport_positions(template_sheet, doc)
    if not template_positions:
        forms.alert("樣板圖紙上沒有可用的 Viewport。", exitscript=True)

    if dup_keys:
        output.print_md(
            "⚠️ **樣板圖紙上有重複的視圖類型**，下列鍵值只會採用第一個 Viewport 的位置：\n"
        )
        for k in dup_keys:
            output.print_md("- `{}`".format(k))

    # --- 2. 選擇目標圖紙 (可多選) ---
    candidate_options = [SheetOption(s) for s in sorted(sheets, key=lambda s: s.SheetNumber)
                         if s.Id != template_sheet.Id]
    target_opts = forms.SelectFromList.show(
        candidate_options,
        name_attr='label',
        multiselect=True,
        title='選擇要套用對齊的「目標圖紙」(可多選)',
        button_name='套用對齊'
    )
    if not target_opts:
        script.exit()
    target_sheets = [opt.sheet for opt in target_opts]

    # --- 3. 執行對齊 (一個 Transaction 包整批) ---
    total_moved = 0
    report_lines = []

    with revit.Transaction("Align Viewports Across Sheets"):
        for sheet in target_sheets:
            moved, skipped = align_sheet_to_template(sheet, template_positions, doc)
            total_moved += moved
            report_lines.append(
                "- **{}** ({}): 移動 {} 個 Viewport，{} 個無對應視圖類型"
                .format(sheet.SheetNumber, sheet.Name, moved, len(skipped))
            )

    # --- 4. 輸出報告 ---
    output.print_md("# ✅ 視圖對齊完成")
    output.print_md(
        "樣板: **{} - {}**".format(template_sheet.SheetNumber, template_sheet.Name)
    )
    output.print_md("總共移動 **{}** 個 Viewport\n".format(total_moved))
    output.print_md("## 各圖紙處理結果")
    for line in report_lines:
        output.print_md(line)


if __name__ == '__main__':
    main()
