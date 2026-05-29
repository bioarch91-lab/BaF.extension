# -*- coding: utf-8 -*-
"""
設定檔讀寫
================
資料結構（JSON）：
{
    "version": 1,
    "documents": [
        {
            "id": "doc-uuid-1",
            "name": "主需求書",
            "path": "D:/projects/abc/spec.pdf",
            "added_at": "2024-01-01T12:00:00"
        },
        ...
    ],
    "bookmarks": [],   // Phase 2
    "notes": []        // Phase 2
}
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
        "documents": [],
        "bookmarks": [],
        "notes": [],
    }


def load_settings(filepath):
    """從檔案讀取設定，失敗時回傳預設值。"""
    if not filepath or not os.path.isfile(filepath):
        return default_settings()
    try:
        # 用 utf-8-sig 讀，容忍可能存在的 BOM，並正確解析中文
        with io.open(filepath, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
        # 簡單版本檢查
        if not isinstance(data, dict):
            return default_settings()
        # 補全缺漏的欄位
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
        # 確保目錄存在
        dir_path = os.path.dirname(filepath)
        if dir_path and not os.path.isdir(dir_path):
            os.makedirs(dir_path)
        # IronPython 2.7 的內建 open() 預設用 ASCII 寫檔，遇到中文會爆
        # ('ascii' codec can't encode...)，所以改用 io.open 明確指定 UTF-8。
        text = json.dumps(data, indent=2, ensure_ascii=False)
        if isinstance(text, bytes):   # Py2 下 json 可能回傳 str(bytes)
            text = text.decode("utf-8")
        with io.open(filepath, "w", encoding="utf-8") as f:
            f.write(text)
        return True, None
    except Exception as ex:
        return False, str(ex)


def add_document(settings, name, path):
    """加入一份文件。回傳新建立的 document dict。"""
    doc = {
        "id": str(uuid.uuid4()),
        "name": name,
        "path": path,
        "added_at": datetime.datetime.now().isoformat(),
    }
    settings["documents"].append(doc)
    return doc


def remove_document(settings, doc_id):
    """移除文件。回傳是否成功。"""
    before = len(settings["documents"])
    settings["documents"] = [d for d in settings["documents"] if d["id"] != doc_id]
    return len(settings["documents"]) != before
