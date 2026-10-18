# -*- coding: utf-8 -*-
"""Excel 批量合并工具 —— 启动入口。

这里是 PyInstaller 的构建入口（见 .github/workflows/build.yml），只负责
引入包并启动界面；合并逻辑在 excel_merge/logic.py，界面在 excel_merge/ui.py。

直接运行：
    python merge_gui.py
"""

import sys

from excel_merge.ui import main

if __name__ == "__main__":
    sys.exit(main())
