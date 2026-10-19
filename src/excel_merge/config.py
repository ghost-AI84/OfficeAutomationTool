# -*- coding: utf-8 -*-
"""用户设置持久化。

存到 %APPDATA%\\ExcelMergeTool\\settings.json（Windows），
其它平台退回 ~/.config/ExcelMergeTool/settings.json。
读写都做了容错：文件损坏、权限不足、键名缺失都不会影响启动。
"""

import json
import os
import sys
from pathlib import Path

__all__ = ["DEFAULTS", "settings_path", "load_settings", "save_settings"]

APP_DIR_NAME = "ExcelMergeTool"
FILE_NAME = "settings.json"

DEFAULTS = {
    # 界面
    "last_dir": "",
    "output_dir": "",
    "settings_version": 2,
    # 合并参数
    "header_mode": "auto",       # auto / fixed / perfile
    "header_row": 1,
    # 默认按子表分组合并：工作簿里有多个子表时，同名/相似子表各成一组
    "sheet": "all",
    "merge_header_rows": True,   # 多行表头自动合并
    "drop_repeated_headers": True,
    "drop_summary_rows": True,
    "fill_merged_cells": True,
    "source_enabled": True,
    "source_position": "last",
    "source_name": "数据来源",
    "unify": True,
    "autofit": True,
    "recursive": False,
    "fuzzy": True,
    "case_sensitive": False,
    "strip_punct": False,
    "drop_empty_rows": True,
    "drop_empty_cols": False,
    "date_in": "",
    "date_out": "yyyy-mm-dd",
    "sort_by_name": False,
}


def settings_path():
    """返回设置文件路径（不保证父目录存在）。"""
    base = os.environ.get("APPDATA")
    if not base and sys.platform == "win32":
        base = str(Path.home() / "AppData" / "Roaming")
    if base:
        return Path(base) / APP_DIR_NAME / FILE_NAME
    return Path.home() / ".config" / APP_DIR_NAME / FILE_NAME


def load_settings():
    """读取设置；任何异常都回退到默认值。"""
    path = settings_path()
    data = dict(DEFAULTS)
    try:
        with open(str(path), "r", encoding="utf-8") as f:
            stored = json.load(f)
    except (OSError, ValueError):
        return data
    if isinstance(stored, dict):
        for key, default in DEFAULTS.items():
            value = stored.get(key, default)
            # 类型对不上就丢弃，避免脏数据让界面初始化崩溃
            if isinstance(default, bool):
                data[key] = bool(value)
            elif isinstance(default, int):
                try:
                    data[key] = int(value)
                except (TypeError, ValueError):
                    data[key] = default
            else:
                data[key] = str(value)

        # v2：默认改成"所有子表（按名字分组）"。
        # 旧版本存下来的 0/空值多半是没改过默认，一并迁移，否则体验不到分组。
        if int(stored.get("settings_version", 1) or 1) < 2:
            if data.get("sheet") in ("0", ""):
                data["sheet"] = "all"
            data["settings_version"] = 2
    return data


def save_settings(settings):
    """原子写入设置文件；失败时静默返回 False。"""
    path = settings_path()
    tmp = path.with_name(path.name + ".part")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        payload = {k: settings[k] for k in DEFAULTS if k in settings}
        with open(str(tmp), "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        os.replace(str(tmp), str(path))
        return True
    except OSError:
        return False
