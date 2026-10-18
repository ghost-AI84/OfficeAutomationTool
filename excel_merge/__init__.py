# -*- coding: utf-8 -*-
"""Excel 批量合并工具 —— 核心包。

模块划分：
    logic   合并引擎（只依赖 pandas / numpy，不含任何界面代码）
    config  用户设置持久化（%APPDATA%\\ExcelMergeTool\\settings.json）
    ui      Tkinter 图形界面
"""

__version__ = "2.0.0"
__all__ = ["logic", "config", "ui"]
