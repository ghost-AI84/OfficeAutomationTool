# -*- coding: utf-8 -*-
"""界面测试：布局、校验、设置持久化、日志裁剪。

需要可用的图形环境；无法创建窗口时自动跳过（返回 0），
所以可以安全地放在没有桌面的 CI 上。
"""
import os
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

# 沙箱/CI 上不写用户真实配置，改用临时目录
os.environ["APPDATA"] = tempfile.mkdtemp(prefix="merge-appdata-")

import tkinter as tk
from tkinter import ttk

import pandas as pd

from excel_merge import config
from excel_merge.logic import MergeOptions
from excel_merge.ui import MergeApp, enable_dpi_awareness


def main():
    enable_dpi_awareness()
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print("SKIP: no usable display (%s)" % e)
        return 0

    try:
        ttk.Style().theme_use("clam")
    except tk.TclError:
        pass

    app = MergeApp(root)
    root.update_idletasks()
    print("  scale=%.2f  minsize=%s" % (app.scale, root.minsize()))

    # 1. 最小尺寸下关键控件必须可见
    mw, mh = root.minsize()
    root.geometry("%dx%d" % (mw, mh))
    root.update_idletasks()
    root.update()
    assert app.run_btn.winfo_ismapped(), "run button not visible at min size"
    assert app.cancel_btn.winfo_ismapped(), "cancel button not visible at min size"
    assert app.progress.winfo_ismapped(), "progress bar not visible at min size"
    assert app.log_text.winfo_height() > 40, "log squeezed to %d" % app.log_text.winfo_height()
    print("  min size keeps buttons visible    OK")

    # 2. 行号输入框随模式启用/禁用
    app.v_header_mode.set("auto")
    app._sync_header_entry()
    root.update_idletasks()
    assert str(app.header_entry.cget("state")) == "disabled"
    app.v_header_mode.set("fixed")
    app._sync_header_entry()
    root.update_idletasks()
    assert str(app.header_entry.cget("state")) == "normal"
    print("  header entry follows mode         OK")

    # 3. 参数收集
    app.files = [str(ROOT / "apps" / "merge_excel.py")]
    app.v_output.set(str(Path(tempfile.gettempdir()) / "x.xlsx"))
    opts, warns = app._collect_options()
    assert isinstance(opts, MergeOptions)
    assert opts.header_row == 1 and opts.sheet_name == 0
    assert opts.merge_header_rows == "auto"
    assert opts.drop_repeated_headers and opts.drop_summary_rows
    print("  collect_options -> MergeOptions   OK")

    # 3b. 逐文件指定表头行
    app.v_header_mode.set("perfile")
    app.header_overrides = {str(ROOT / "apps" / "merge_excel.py"): 3}
    opts_pf, _ = app._collect_options()
    assert opts_pf.header_row == {str(ROOT / "apps" / "merge_excel.py"): 3}, opts_pf.header_row
    app.v_header_mode.set("fixed")
    app._sync_header_entry()
    print("  per-file header overrides         OK")

    # 3c. 工作表 auto
    app.v_sheet.set("auto")
    opts_auto, _ = app._collect_options()
    assert opts_auto.sheet_name == "auto", opts_auto.sheet_name
    app.v_sheet.set("0")
    print("  sheet=auto accepted               OK")

    # 3d. 表头调整对话框：能打开、能检测、能回填
    import time
    from excel_merge.ui import HeaderDialog
    demo = Path(tempfile.mkdtemp(prefix="hdr-dlg-")) / "demo.xlsx"
    pd.DataFrame([["销售明细", None], [None, None],
                  ["姓名", "金额"], ["张三", "100"]]).to_excel(demo, index=False, header=False)
    dlg = HeaderDialog(app, [str(demo)], 0, {})
    deadline = time.time() + 15
    while time.time() < deadline and not dlg.tree.get_children():
        root.update()
        time.sleep(0.05)
    assert dlg.tree.get_children(), "dialog did not populate"
    dlg._ok()
    root.update()
    assert app.header_overrides, app.header_overrides
    assert app.v_header_mode.get() == "perfile"
    print("  header dialog detects and applies OK")

    # 4. 非法输出后缀
    app.v_output.set(str(Path(tempfile.gettempdir()) / "x.xls"))
    try:
        app._collect_options()
        raise AssertionError(".xls was accepted")
    except ValueError:
        pass
    print("  reject .xls output                OK")

    # 5. 非法强制类型
    app.v_output.set(str(Path(tempfile.gettempdir()) / "x.xlsx"))
    app.type_text.insert("1.0", "金额=data\n")
    try:
        app._collect_options()
        raise AssertionError("bad type was accepted")
    except ValueError:
        pass
    app.type_text.delete("1.0", "end")
    print("  reject bad force type             OK")

    # 6. 填类型但没勾统一格式 -> 自动勾上并提示
    app.type_text.insert("1.0", "金额=number\n")
    app.v_unify.set(False)
    _, warns2 = app._collect_options()
    assert app.v_unify.get() is True
    assert any("统一数据格式" in w for w in warns2), warns2
    app.type_text.delete("1.0", "end")
    print("  auto-enable unify for types       OK")

    # 7. 顺序上下移动并保持选中
    app.files = ["C:\\a.xlsx", "C:\\b.xlsx", "C:\\c.xlsx"]
    app._refresh_files()
    app.file_listbox.selection_set(2)
    app.move_up()
    assert app.files == ["C:\\a.xlsx", "C:\\c.xlsx", "C:\\b.xlsx"], app.files
    assert app.file_listbox.selection_includes(1)
    print("  move up keeps selection           OK")

    # 8. 设置持久化
    app.settings["last_dir"] = str(ROOT)
    app.v_source_name.set("来源文件")
    app._persist_settings()
    reloaded = config.load_settings()
    assert reloaded["source_name"] == "来源文件", reloaded["source_name"]
    assert Path(config.settings_path()).exists()
    print("  settings round-trip               OK")

    # 9. 日志行数上限
    for i in range(4200):
        app._append_log("line %d" % i)
    lines = int(app.log_text.index("end-1c").split(".")[0])
    assert lines <= 4100, lines
    print("  log line cap                      OK")

    root.destroy()
    print("gui tests passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
