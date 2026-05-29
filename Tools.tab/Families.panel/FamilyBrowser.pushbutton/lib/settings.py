# -*- coding: utf-8 -*-
"""
族群瀏覽器 - 設定檔讀寫
================================
記錄使用者設定的「族群來源」（目前 Phase 1 只支援資料夾來源）。

資料結構（JSON）：
{
    "version": 1,
    "sources": [
        {
            "id": "uuid",
            "name": "本案族群庫",
            "type": "folder",          // Phase 2 會加入 "rvt"
            "path": "D:/projects/abc/Families",
            "recursive": true,
            "added_at": "2026-01-01T12:00:00"
        },
        ...
    ]
}

註：IronPython 2.7 內建 open() 預設用 ASCII 寫檔，遇到中文（路徑、名稱）會爆，
故一律用 io.open + UTF-8；讀檔用 utf-8-sig 容忍 BOM。
"""

import io
import json
import os
import uuid
import datetime

SETTINGS_VERSION = 1


def default_settings():
    return {
        "version": SETTINGS_VERSION,
        "sources": [],
    }


def load_settings(filepath):
    """從檔案讀取設定，失敗時回傳預設值。"""
    if not filepath or not os.path.isfile(filepath):
        return default_settings()
    try:
        with io.open(filepath, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return default_settings()
        defaults = default_settings()
        for key, val in defaults.items():
            if key not in data:
                data[key] = val
        return data
    except Exception:
        return default_settings()


def save_settings(filepath, data):
    """寫入設定檔。回傳 (success, error_message)"""
    if not filepath:
        return False, "尚未指定設定檔路徑"
    try:
        dir_path = os.path.dirname(filepath)
        if dir_path and not os.path.isdir(dir_path):
            os.makedirs(dir_path)
        text = json.dumps(data, indent=2, ensure_ascii=False)
        if isinstance(text, bytes):
            text = text.decode("utf-8")
        with io.open(filepath, "w", encoding="utf-8") as f:
            f.write(text)
        return True, None
    except Exception as ex:
        return False, str(ex)


def add_folder_source(settings, name, path, recursive=True):
    """加入一個資料夾來源。回傳新建立的 source dict。"""
    src = {
        "id": str(uuid.uuid4()),
        "name": name,
        "type": "folder",
        "path": path,
        "recursive": bool(recursive),
        "added_at": datetime.datetime.now().isoformat(),
    }
    settings["sources"].append(src)
    return src


def remove_source(settings, source_id):
    """移除來源（不刪除任何檔案）。回傳是否成功。"""
    before = len(settings["sources"])
    settings["sources"] = [s for s in settings["sources"] if s["id"] != source_id]
    return len(settings["sources"]) != before
