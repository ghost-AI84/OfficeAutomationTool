# -*- coding: utf-8 -*-
"""启动入口：Excel 批量合并工具。

这是一个"入口文件"：只负责把 src/ 加进模块搜索路径，然后启动界面。
工具本身的代码在 src/excel_merge/ 里；PyInstaller 也用这个文件打包。

注意：入口文件名不能叫 excel_merge.py，否则 PyInstaller 会把它当成
excel_merge 模块，和 src 里的包重名，导致打包时找不到 excel_merge.ui。

直接运行：
    python apps/merge_excel.py
"""

import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_merge.ui import main  # noqa: E402  (必须在 sys.path 调整之后)

if __name__ == "__main__":
    sys.exit(main())
