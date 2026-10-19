# -*- coding: utf-8 -*-
"""Tkinter 界面。

相对旧版的改动：
  - 参数收进 Notebook 三个标签页，窗口再小也不会把"开始合并"挤出可视区
  - 启动时声明 DPI 感知（Win7 走旧 API 回退），不再被系统位图拉伸
  - 文件顺序可上移/下移，主表头 = 列表第一个文件（改动真正生效）
  - 增加取消按钮；关窗时若仍在合并会先询问
  - 日志批量插入并限制行数，避免长任务卡顿
  - 输出后缀、覆盖、强制类型都会校验并提示
  - 参数与上次输出目录写入 settings.json
"""

import ctypes
import os
import queue
import subprocess
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from .config import load_settings, save_settings
from .logic import (
    ALLOWED_TYPES,
    OUTPUT_EXTS,
    SUPPORTED_EXTS,
    MergeCancelled,
    MergeOptions,
    inspect_headers,
    merge_files,
)

__all__ = ["MergeApp", "enable_dpi_awareness", "main"]

POLL_MS = 60
LOG_MAX_LINES = 4000

FILE_TYPES = [
    ("Excel / CSV", "*.xlsx *.xlsm *.xls *.xlsb *.ods *.csv"),
    ("Excel 文件", "*.xlsx *.xlsm *.xls *.xlsb"),
    ("CSV 文件", "*.csv"),
    ("所有文件", "*.*"),
]

# 工作表下拉选项：显示文字 -> 内部取值
SHEET_CHOICES = [
    ("第一个工作表", "0"),
    ("自动选择有数据的表", "auto"),
    ("所有子表（按名字分组）", "all"),
]


def sheet_label(code):
    code = str(code)
    for label, value in SHEET_CHOICES:
        if value == code:
            return label
    return code


def sheet_code(text):
    text = str(text).strip()
    for label, value in SHEET_CHOICES:
        if text == label:
            return value
    return text


def enable_dpi_awareness():
    """声明 DPI 感知。必须在创建 Tk 窗口之前调用。

    返回 1 = 系统级（Win8.1+），2 = 旧 API（Win7），0 = 未启用。
    """
    if sys.platform != "win32":
        return 0
    try:
        # Win8.1 及以上
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        return 1
    except (AttributeError, OSError):
        pass
    try:
        # Win7 回退
        ctypes.windll.user32.SetProcessDPIAware()
        return 2
    except (AttributeError, OSError):
        return 0


class MergeApp:
    def __init__(self, root):
        self.root = root
        self.settings = load_settings()

        self.files = []
        self.msg_queue = queue.Queue()
        self.worker = None
        self.cancel_event = threading.Event()
        self._last_scan_ok = True
        self._poll_id = None

        self.scale = self._detect_scale()
        self._setup_window()
        self._init_vars()
        self._build_ui()
        self._sync_header_entry()
        self.root.bind("<Destroy>", self._on_destroy, add="+")
        self._poll_queue()

    # ------------------------------------------------------------ 初始化
    def _detect_scale(self):
        try:
            return max(1.0, self.root.winfo_fpixels("1i") / 96.0)
        except tk.TclError:
            return 1.0

    def _px(self, value):
        return int(round(value * self.scale))

    def _setup_window(self):
        self.root.title("Excel 批量合并工具")
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w = min(self._px(1020), int(sw * 0.92))
        h = min(self._px(760), int(sh * 0.84))
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 3)
        self.root.geometry("%dx%d+%d+%d" % (w, h, x, y))
        # 内容区（日志以上）约 430px；再按屏幕裁剪，避免在小屏 + 高缩放时
        # 最小尺寸反而超过屏幕可用区域
        min_w = min(self._px(880), max(int(sw * 0.9), 320))
        min_h = min(self._px(620), max(int(sh * 0.75), 260))
        self.root.minsize(min_w, min_h)

    def _init_vars(self):
        s = self.settings
        self.v_header_mode = tk.StringVar(value=s["header_mode"])
        self.v_header_row = tk.StringVar(value=str(s["header_row"]))
        self.v_sheet = tk.StringVar(value=sheet_label(s["sheet"]))

        self.v_source = tk.BooleanVar(value=s["source_enabled"])
        self.v_source_pos = tk.StringVar(value=s["source_position"])
        self.v_source_name = tk.StringVar(value=s["source_name"])

        self.v_unify = tk.BooleanVar(value=s["unify"])
        self.v_autofit = tk.BooleanVar(value=s["autofit"])
        self.v_recursive = tk.BooleanVar(value=s["recursive"])
        self.v_sort_name = tk.BooleanVar(value=s["sort_by_name"])
        self.v_merge_rows = tk.BooleanVar(value=s["merge_header_rows"])
        self.v_drop_repeat = tk.BooleanVar(value=s["drop_repeated_headers"])
        self.v_drop_summary = tk.BooleanVar(value=s["drop_summary_rows"])
        self.v_fill_merged = tk.BooleanVar(value=s["fill_merged_cells"])

        # 逐文件表头行（键是文件绝对路径，值是 1 基行号）
        self.header_overrides = {}

        self.v_fuzzy = tk.BooleanVar(value=s["fuzzy"])
        self.v_case = tk.BooleanVar(value=s["case_sensitive"])
        self.v_punct = tk.BooleanVar(value=s["strip_punct"])
        self.v_drop_rows = tk.BooleanVar(value=s["drop_empty_rows"])
        self.v_drop_cols = tk.BooleanVar(value=s["drop_empty_cols"])

        self.v_date_in = tk.StringVar(value=s["date_in"])
        self.v_date_out = tk.StringVar(value=s["date_out"])
        self.v_output = tk.StringVar(value="")

        self.v_status = tk.StringVar(value="就绪")

    # ------------------------------------------------------------ 构建界面
    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        body = ttk.Frame(self.root)
        body.grid(row=0, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(4, weight=1)   # 只有日志区吃掉多余空间

        self._build_file_frame(body)
        self._build_options(body)
        self._build_output_row(body)
        self._build_run_row(body)
        self._build_log(body)

    def _build_file_frame(self, parent):
        frame = ttk.LabelFrame(parent, text=" 1. 选择要合并的文件 ")
        frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        frame.columnconfigure(0, weight=1)

        bar = ttk.Frame(frame)
        bar.grid(row=0, column=0, sticky="ew", padx=6, pady=6)

        self._file_buttons = []
        for text, cmd in (
            ("添加文件", self.add_files),
            ("添加文件夹", self.add_folder),
            ("移除选中", self.remove_selected),
            ("清空列表", self.clear_files),
        ):
            btn = ttk.Button(bar, text=text, command=cmd)
            btn.pack(side="left", padx=2)
            self._file_buttons.append(btn)

        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=6)

        self.btn_up = ttk.Button(bar, text="上移", width=6, command=self.move_up)
        self.btn_up.pack(side="left", padx=2)
        self.btn_down = ttk.Button(bar, text="下移", width=6, command=self.move_down)
        self.btn_down.pack(side="left", padx=2)
        self._file_buttons.extend([self.btn_up, self.btn_down])

        self.count_label = ttk.Label(bar, text="已选 0 个文件")
        self.count_label.pack(side="right", padx=6)

        wrap = ttk.Frame(frame)
        wrap.grid(row=1, column=0, sticky="ew", padx=6)
        wrap.columnconfigure(0, weight=1)

        self.file_listbox = tk.Listbox(wrap, selectmode="extended", height=5,
                                       activestyle="none", exportselection=False)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=sb.set)
        self.file_listbox.grid(row=0, column=0, sticky="ew")
        sb.grid(row=0, column=1, sticky="ns")

        self.template_label = ttk.Label(
            frame, text="主表头模板：—（列表第一个文件决定列顺序）", foreground="#666")
        self.template_label.grid(row=2, column=0, sticky="w", padx=10, pady=(4, 6))

    def _build_options(self, parent):
        nb = ttk.Notebook(parent)
        nb.grid(row=1, column=0, sticky="ew", padx=10, pady=5)

        self._build_basic_tab(nb)
        self._build_match_tab(nb)
        self._build_advanced_tab(nb)

    def _build_basic_tab(self, nb):
        tab = ttk.Frame(nb)
        nb.add(tab, text=" 基础 ")

        row = ttk.Frame(tab)
        row.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Label(row, text="表头所在行：").pack(side="left")
        ttk.Radiobutton(row, text="自动检测", variable=self.v_header_mode,
                        value="auto", command=self._sync_header_entry).pack(side="left")
        ttk.Radiobutton(row, text="固定行号", variable=self.v_header_mode,
                        value="fixed", command=self._sync_header_entry).pack(side="left", padx=(0, 4))
        self.header_entry = ttk.Entry(row, textvariable=self.v_header_row, width=6)
        self.header_entry.pack(side="left")
        ttk.Label(row, text="（从 1 开始）", foreground="#666").pack(side="left", padx=4)
        ttk.Radiobutton(row, text="逐文件指定", variable=self.v_header_mode,
                        value="perfile", command=self._sync_header_entry).pack(side="left", padx=(12, 4))
        self.btn_header_dialog = ttk.Button(row, text="检测并调整…",
                                            command=self.open_header_dialog)
        self.btn_header_dialog.pack(side="left")

        row = ttk.Frame(tab)
        row.pack(fill="x", padx=8, pady=4)
        ttk.Label(row, text="工作表：").pack(side="left")
        self.sheet_combo = ttk.Combobox(row, textvariable=self.v_sheet, width=26,
                                        values=[lbl for lbl, _ in SHEET_CHOICES])
        self.sheet_combo.pack(side="left", padx=4)
        ttk.Label(row, text="（选“所有子表”会把同名子表合并到各自的工作表；也可直接输入表名）",
                  foreground="#666").pack(side="left")

        row = ttk.Frame(tab)
        row.pack(fill="x", padx=8, pady=4)
        ttk.Checkbutton(row, text="添加来源列", variable=self.v_source).pack(side="left")
        ttk.Label(row, text="位置：").pack(side="left", padx=(8, 0))
        ttk.Radiobutton(row, text="末列", variable=self.v_source_pos,
                        value="last").pack(side="left")
        ttk.Radiobutton(row, text="首列", variable=self.v_source_pos,
                        value="first").pack(side="left", padx=(0, 8))
        ttk.Label(row, text="列名：").pack(side="left")
        ttk.Entry(row, textvariable=self.v_source_name, width=10).pack(side="left", padx=4)

        row = ttk.Frame(tab)
        row.pack(fill="x", padx=8, pady=(4, 8))
        ttk.Checkbutton(row, text="统一数据格式", variable=self.v_unify).pack(side="left")
        ttk.Checkbutton(row, text="自动调整列宽（xlsx）",
                        variable=self.v_autofit).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(row, text="文件夹递归",
                        variable=self.v_recursive).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(row, text="按文件名排序",
                        variable=self.v_sort_name).pack(side="left", padx=(12, 0))

    def _build_match_tab(self, nb):
        tab = ttk.Frame(nb)
        nb.add(tab, text=" 表头匹配 ")

        row = ttk.Frame(tab)
        row.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Checkbutton(row, text="启用容错匹配（同义词、空格、全半角、不可见字符）",
                        variable=self.v_fuzzy).pack(side="left")

        row = ttk.Frame(tab)
        row.pack(fill="x", padx=8, pady=4)
        ttk.Checkbutton(row, text="区分大小写", variable=self.v_case).pack(side="left")
        ttk.Checkbutton(row, text="忽略标点",
                        variable=self.v_punct).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(row, text="丢弃全空行",
                        variable=self.v_drop_rows).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(row, text="丢弃全空列",
                        variable=self.v_drop_cols).pack(side="left", padx=(12, 0))

        row = ttk.Frame(tab)
        row.pack(fill="x", padx=8, pady=4)
        ttk.Checkbutton(row, text="自动合并多行表头",
                        variable=self.v_merge_rows).pack(side="left")
        ttk.Checkbutton(row, text="填充合并单元格表头（xlsx）",
                        variable=self.v_fill_merged).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(row, text="丢弃重复表头行",
                        variable=self.v_drop_repeat).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(row, text="丢弃末尾合计行",
                        variable=self.v_drop_summary).pack(side="left", padx=(12, 0))

        ttk.Label(
            tab, foreground="#666", wraplength=self._px(760), justify="left",
            text="表头行不在同一行、前面有标题或空行、两级表头（分组名是合并单元格）、"
                 "分页导出夹着重复表头、末尾有合计行——这些都会自动处理。"
                 "想确认或手工改，用「检测并调整…」。",
        ).pack(fill="x", padx=10, pady=(4, 8))

    def _build_advanced_tab(self, nb):
        tab = ttk.Frame(nb)
        nb.add(tab, text=" 高级 ")

        pane = ttk.Frame(tab)
        pane.pack(fill="x", padx=8, pady=(8, 4))
        pane.columnconfigure(0, weight=1)
        pane.columnconfigure(1, weight=1)

        left = ttk.LabelFrame(pane, text=" 强制列类型（列名=类型） ")
        left.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.type_text = tk.Text(left, height=4, width=28, wrap="none")
        self.type_text.pack(fill="both", expand=True, padx=3, pady=3)

        right = ttk.LabelFrame(pane, text=" 列名重命名（旧名=新名） ")
        right.grid(row=0, column=1, sticky="ew", padx=(4, 0))
        self.map_text = tk.Text(right, height=4, width=28, wrap="none")
        self.map_text.pack(fill="both", expand=True, padx=3, pady=3)

        row = ttk.Frame(tab)
        row.pack(fill="x", padx=8, pady=4)
        ttk.Label(row, text="输入日期格式：").pack(side="left")
        ttk.Entry(row, textvariable=self.v_date_in, width=16).pack(side="left", padx=4)
        ttk.Label(row, text="输出日期格式：").pack(side="left", padx=(12, 0))
        ttk.Entry(row, textvariable=self.v_date_out, width=16).pack(side="left", padx=4)

        ttk.Label(
            tab, foreground="#666",
            text="类型可选：text / number / datetime / bool；留空则自动推断。"
                 "日期格式留空表示自动识别，例如 %Y-%m-%d。",
        ).pack(fill="x", padx=10, pady=(0, 8))

    def _build_output_row(self, parent):
        frame = ttk.LabelFrame(parent, text=" 2. 输出设置 ")
        frame.grid(row=2, column=0, sticky="ew", padx=10, pady=5)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="输出文件：").grid(row=0, column=0, sticky="w", padx=(8, 0), pady=6)
        ttk.Entry(frame, textvariable=self.v_output).grid(
            row=0, column=1, sticky="ew", padx=4, pady=6)
        ttk.Button(frame, text="浏览...", command=self.choose_output).grid(
            row=0, column=2, padx=(0, 4), pady=6)
        ttk.Button(frame, text="打开目录", command=self.open_output_folder).grid(
            row=0, column=3, padx=(0, 8), pady=6)

    def _build_run_row(self, parent):
        frame = ttk.Frame(parent)
        frame.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
        frame.columnconfigure(2, weight=1)

        self.run_btn = ttk.Button(frame, text="开始合并", command=self.start_merge)
        self.run_btn.grid(row=0, column=0, padx=(0, 4))

        self.cancel_btn = ttk.Button(frame, text="取消", command=self.cancel_merge,
                                     state="disabled")
        self.cancel_btn.grid(row=0, column=1, padx=(0, 8))

        self.progress = ttk.Progressbar(frame, mode="determinate", maximum=1, value=0)
        self.progress.grid(row=0, column=2, sticky="ew", padx=4)

        ttk.Label(frame, textvariable=self.v_status, width=18,
                  anchor="e").grid(row=0, column=3, sticky="e")

    def _build_log(self, parent):
        frame = ttk.LabelFrame(parent, text=" 运行日志 ")
        frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=(5, 10))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.log_text = ScrolledText(frame, height=8, state="disabled", wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.log_text.tag_config("error", foreground="#c0392b")
        self.log_text.tag_config("ok", foreground="#1e8449")
        self.log_text.tag_config("info", foreground="#2c3e50")
        self.log_text.tag_config("warn", foreground="#b9770e")
        self.log_text.tag_config("dim", foreground="#7f8c8d")

        self._append_log(
            "提示：列表第一个文件决定主表头与列顺序；文件顺序可用「上移 / 下移」调整。", "dim")

    # ------------------------------------------------------------ 小工具
    def _sync_header_entry(self):
        state = "normal" if self.v_header_mode.get() == "fixed" else "disabled"
        try:
            self.header_entry.config(state=state)
        except tk.TclError:
            pass

    def _sorted_selection(self):
        return sorted(self.file_listbox.curselection())

    def _refresh_files(self):
        self.file_listbox.delete(0, "end")
        for path in self.files:
            self.file_listbox.insert("end", path)
        self.count_label.config(text="已选 " + str(len(self.files)) + " 个文件")
        if self.files:
            self.template_label.config(
                text="主表头模板：" + Path(self.files[0]).name + "（决定列顺序）")
        else:
            self.template_label.config(text="主表头模板：—（列表第一个文件决定列顺序）")

    def _set_running(self, running):
        self.run_btn.config(state="disabled" if running else "normal")
        self.cancel_btn.config(state="normal" if running else "disabled")
        for btn in self._file_buttons:
            btn.config(state="disabled" if running else "normal")

    def _append_log(self, msg, tag="info"):
        self._append_logs([(msg, tag)])

    def _append_logs(self, entries):
        if not entries:
            return
        try:
            at_bottom = self.log_text.yview()[1] > 0.999
        except tk.TclError:
            at_bottom = True
        self.log_text.configure(state="normal")
        for msg, tag in entries:
            if tag == "info":
                if msg.startswith("[跳过]") or "注意：" in msg:
                    tag = "warn"
            self.log_text.insert("end", msg + "\n", tag)
        self._trim_log()
        if at_bottom:
            self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _trim_log(self):
        total = int(self.log_text.index("end-1c").split(".")[0])
        if total > LOG_MAX_LINES:
            self.log_text.delete("1.0", "%d.0" % (total - LOG_MAX_LINES + 1))

    # ------------------------------------------------------------ 文件操作
    def add_files(self):
        initial = self.settings.get("last_dir") or None
        paths = filedialog.askopenfilenames(
            title="选择要合并的文件", filetypes=FILE_TYPES, initialdir=initial)
        if paths:
            self.settings["last_dir"] = str(Path(paths[0]).parent)
        self._add_paths(paths)

    def add_folder(self):
        initial = self.settings.get("last_dir") or None
        folder = filedialog.askdirectory(title="选择文件夹", initialdir=initial)
        if not folder:
            return
        self.settings["last_dir"] = folder
        recursive = self.v_recursive.get()
        self._append_log("正在扫描文件夹：%s（%s）" % (
            folder, "含子文件夹" if recursive else "仅当前层"), "dim")
        self._set_scanning(True)

        def scan():
            try:
                root_dir = Path(folder)
                it = root_dir.rglob("*") if recursive else root_dir.iterdir()
                found = []
                for p in it:
                    try:
                        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS \
                                and not p.name.startswith("~$"):
                            found.append(p)
                    except OSError:
                        continue
                self.msg_queue.put(("files", found, folder))
            except Exception as e:  # 权限、路径消失等
                self.msg_queue.put(("scan_error", str(e)))

        threading.Thread(target=scan, daemon=True).start()

    def _set_scanning(self, scanning):
        state = "disabled" if scanning else "normal"
        for btn in self._file_buttons:
            btn.config(state=state)
        self.v_status.set("正在扫描..." if scanning else "就绪")

    def _add_paths(self, paths, quiet=False):
        added = 0
        for p in paths:
            f = Path(p)
            if f.suffix.lower() not in SUPPORTED_EXTS:
                continue
            if f.name.startswith("~$"):
                continue
            try:
                resolved = str(f.resolve())
            except OSError:
                continue
            if resolved in self.files:
                continue
            self.files.append(resolved)
            added += 1
        self._refresh_files()
        if added:
            self._suggest_output()
        elif not quiet:
            self._append_log("没有新增文件（已存在或格式不支持）", "warn")

    def _suggest_output(self):
        if self.v_output.get().strip() or not self.files:
            return
        first = Path(self.files[0])
        self.v_output.set(str(first.parent / (first.stem + "_合并.xlsx")))

    def remove_selected(self):
        idxs = self._sorted_selection()
        if not idxs:
            return
        for i in reversed(idxs):
            del self.files[i]
        self._refresh_files()

    def clear_files(self):
        self.files.clear()
        self._refresh_files()

    def move_up(self):
        new_idxs = []
        for i in self._sorted_selection():
            j = max(0, i - 1)
            if j != i:
                self.files[j], self.files[i] = self.files[i], self.files[j]
            new_idxs.append(j)
        self._refresh_files()
        for j in new_idxs:
            if 0 <= j < len(self.files):
                self.file_listbox.selection_set(j)

    def move_down(self):
        new_idxs = []
        for i in reversed(self._sorted_selection()):
            j = min(len(self.files) - 1, i + 1)
            if j != i:
                self.files[j], self.files[i] = self.files[i], self.files[j]
            new_idxs.append(j)
        self._refresh_files()
        for j in new_idxs:
            if 0 <= j < len(self.files):
                self.file_listbox.selection_set(j)

    def choose_output(self):
        initial_dir = self.settings.get("output_dir") or self.settings.get("last_dir") or None
        path = filedialog.asksaveasfilename(
            title="保存合并结果",
            defaultextension=".xlsx",
            initialfile="合并结果.xlsx",
            initialdir=initial_dir,
            filetypes=[("Excel 文件", "*.xlsx"), ("CSV 文件", "*.csv")],
        )
        if path:
            self.v_output.set(path)
            self.settings["output_dir"] = str(Path(path).parent)

    def open_output_folder(self):
        target = self.v_output.get().strip()
        if target:
            folder = Path(target).parent
        else:
            folder = Path(self.settings.get("last_dir") or ".")
        if not folder.exists():
            self._append_log("目录不存在：%s" % folder, "warn")
            return
        try:
            if sys.platform == "win32":
                os.startfile(str(folder))  # noqa: 仅 Windows 存在
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except OSError as e:
            self._append_log("无法打开目录：%s" % e, "error")

    # ------------------------------------------------------------ 参数收集
    @staticmethod
    def _parse_kv(text):
        result = {}
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                k, v = line.split("=", 1)
                k, v = k.strip(), v.strip()
                if k and v:
                    result[k] = v
        return result

    def _collect_options(self):
        """校验输入并生成 MergeOptions；返回 (options, warnings)。"""
        warnings = []

        if not self.files:
            raise ValueError("请先添加要合并的文件")

        out_path = self.v_output.get().strip()
        if not out_path:
            raise ValueError("请先设置输出文件路径")
        suffix = Path(out_path).suffix.lower()
        if suffix not in OUTPUT_EXTS:
            raise ValueError("输出文件只支持 .xlsx 或 .csv 格式")

        mode = self.v_header_mode.get()
        if mode == "auto":
            header_row = None
        elif mode == "perfile":
            # 没调整过就退回自动检测
            header_row = dict(self.header_overrides) or None
        else:
            try:
                header_row = int(self.v_header_row.get().strip())
                if header_row < 1:
                    raise ValueError
            except ValueError:
                raise ValueError("表头行号必须是大于等于 1 的整数")

        sheet_raw = sheet_code(self.v_sheet.get())
        if sheet_raw.lower() == "auto":
            sheet_name = "auto"
        elif sheet_raw in ("", "0"):
            sheet_name = 0
        else:
            try:
                sheet_name = int(sheet_raw)
            except ValueError:
                sheet_name = sheet_raw

        force_types = self._parse_kv(self.type_text.get("1.0", "end"))
        bad_types = {k: v for k, v in force_types.items()
                     if v.strip().lower() not in ALLOWED_TYPES}
        if bad_types:
            raise ValueError(
                "强制列类型的取值只能是 %s，以下无法识别：%s"
                % (" / ".join(ALLOWED_TYPES),
                   ", ".join("%s=%s" % (k, v) for k, v in bad_types.items())))
        force_types = {k: v.strip().lower() for k, v in force_types.items()}

        unify = self.v_unify.get()
        if force_types and not unify:
            # 填了类型却关掉统一格式，等于类型被静默丢弃 —— 自动打开并告知
            self.v_unify.set(True)
            unify = True
            warnings.append("已自动勾选「统一数据格式」：填写了强制列类型。")

        column_name_map = self._parse_kv(self.map_text.get("1.0", "end"))
        if column_name_map:
            targets = list(column_name_map.values())
            dupes = sorted({t for t in targets if targets.count(t) > 1})
            if dupes:
                warnings.append("列名重命名后有重复目标名：%s，输出会出现同名列"
                                % ", ".join(dupes))

        options = MergeOptions(
            output_path=out_path,
            header_row=header_row,
            sheet_name=sheet_name,
            add_source_column=self.v_source.get(),
            source_column_name=self.v_source_name.get().strip() or "数据来源",
            source_position=self.v_source_pos.get(),
            drop_empty_rows=self.v_drop_rows.get(),
            drop_empty_cols=self.v_drop_cols.get(),
            column_name_map=column_name_map,
            unify_formats=unify,
            date_format=self.v_date_in.get().strip() or None,
            output_date_format=self.v_date_out.get().strip() or None,
            force_types=force_types,
            autofit_columns=self.v_autofit.get(),
            fuzzy_header_match=self.v_fuzzy.get(),
            header_case_sensitive=self.v_case.get(),
            header_strip_punct=self.v_punct.get(),
            sort_by_name=self.v_sort_name.get(),
            merge_header_rows="auto" if self.v_merge_rows.get() else 0,
            drop_repeated_headers=self.v_drop_repeat.get(),
            drop_summary_rows=self.v_drop_summary.get(),
            fill_merged_header_cells=self.v_fill_merged.get(),
        )
        return options, warnings

    # ------------------------------------------------------------ 运行
    def start_merge(self):
        try:
            options, warnings = self._collect_options()
        except ValueError as e:
            messagebox.showwarning("提示", str(e))
            return

        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo("提示", "任务正在进行中，请稍候 ...")
            return

        out_path = Path(options.output_path)
        if out_path.exists():
            if not messagebox.askyesno(
                    "确认覆盖", "输出文件已存在：\n%s\n\n继续将覆盖它。" % out_path):
                return

        for msg in warnings:
            self._append_log(msg, "warn")

        self._set_running(True)
        self.progress["value"] = 0
        self.progress["maximum"] = max(len(self.files), 1)
        self.v_status.set("0 / %d" % len(self.files))

        self._append_log("=" * 60, "dim")
        self._append_log("开始合并 %d 个文件 ..." % len(self.files), "info")
        self._append_log("工作表：%s" % self.v_sheet.get(), "dim")

        self.cancel_event = threading.Event()
        files = list(self.files)
        self.worker = threading.Thread(
            target=self._run_merge, args=(files, options, self.cancel_event), daemon=True)
        self.worker.start()

    def cancel_merge(self):
        if self.worker is None or not self.worker.is_alive():
            return
        self.cancel_event.set()
        self.cancel_btn.config(state="disabled")
        self.v_status.set("正在取消...")
        self._append_log("已请求取消，正在等待当前文件处理完 ...", "warn")

    def _run_merge(self, files, options, cancel_event):
        try:
            merge_files(
                files, options,
                log_callback=lambda m: self.msg_queue.put(("log", m, "info")),
                progress_callback=lambda c, t: self.msg_queue.put(("progress", c, t)),
                cancel_event=cancel_event,
            )
        except MergeCancelled:
            self.msg_queue.put(("cancelled", ""))
        except Exception as e:
            self.msg_queue.put(("error", str(e), traceback.format_exc()))
        else:
            self.msg_queue.put(("done", options.output_path))

    # ------------------------------------------------------------ 消息循环
    def _poll_queue(self):
        try:
            if not self.root.winfo_exists():
                return
            items = []
            try:
                while True:
                    items.append(self.msg_queue.get_nowait())
            except queue.Empty:
                pass
            if items:
                self._handle_messages(items)
            self._poll_id = self.root.after(POLL_MS, self._poll_queue)
        except tk.TclError:
            # 窗口已经销毁，停止轮询
            pass

    def _stop_poll(self):
        if self._poll_id is not None:
            try:
                self.root.after_cancel(self._poll_id)
            except tk.TclError:
                pass
            self._poll_id = None

    def _on_destroy(self, event):
        if event.widget is self.root:
            self._stop_poll()

    def _handle_messages(self, items):
        logs = []
        progress = None
        terminal = None

        for item in items:
            kind = item[0]
            if kind == "log":
                logs.append((item[1], item[2]))
            elif kind == "progress":
                progress = (item[1], item[2])
            elif kind == "files":
                self._set_scanning(False)
                self._add_paths(item[1], quiet=True)
                logs.append(("扫描完成，匹配到 %d 个文件" % len(item[1]), "dim"))
            elif kind == "scan_error":
                self._set_scanning(False)
                logs.append(("扫描文件夹失败：%s" % item[1], "error"))
            elif kind in ("done", "error", "cancelled"):
                terminal = item

        self._append_logs(logs)
        if progress is not None:
            cur, total = progress
            self.progress["maximum"] = max(total, 1)
            self.progress["value"] = cur
            self.v_status.set("%d / %d" % (cur, total))

        if terminal is not None:
            self._finish(terminal)

    def _finish(self, item):
        kind = item[0]
        self._set_running(False)

        if kind == "done":
            path = item[1]
            self.progress["value"] = self.progress["maximum"]
            self.v_status.set("完成")
            self._append_log("已生成文件：" + str(path), "ok")
            self._persist_settings()
            if messagebox.askyesno("完成", "合并完成！\n\n输出文件：\n%s\n\n是否打开所在文件夹？" % path):
                self._open_folder(path)
        elif kind == "cancelled":
            self.progress["value"] = 0
            self.v_status.set("已取消")
            self._append_log("合并已取消，未生成输出文件。", "warn")
        else:
            msg, tb = item[1], item[2]
            self.progress["value"] = 0
            self.v_status.set("出错")
            self._append_log("出错：" + msg, "error")
            self._append_log(tb, "error")
            messagebox.showerror("出错", msg)

    def _open_folder(self, path):
        try:
            folder = Path(path).parent
            if sys.platform == "win32":
                os.startfile(str(folder))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except OSError:
            pass

    # ------------------------------------------------------------ 表头调整
    def open_header_dialog(self):
        if not self.files:
            messagebox.showinfo("提示", "请先添加要合并的文件")
            return
        sheet_raw = sheet_code(self.v_sheet.get())
        sheet_name = 0 if sheet_raw in ("", "0") else sheet_raw
        HeaderDialog(self, list(self.files), sheet_name, self.header_overrides)

    def apply_header_overrides(self, overrides):
        self.header_overrides = dict(overrides)
        self.v_header_mode.set("perfile")
        self._sync_header_entry()
        if overrides:
            rows = sorted(set(overrides.values()))
            self._append_log("已按文件指定表头行：共 %d 个文件，行号 %s"
                             % (len(overrides), rows), "ok")

    # ------------------------------------------------------------ 关闭
    def _persist_settings(self):
        s = self.settings
        s["header_mode"] = self.v_header_mode.get()
        try:
            s["header_row"] = int(self.v_header_row.get().strip())
        except ValueError:
            s["header_row"] = 1
        s["sheet"] = sheet_code(self.v_sheet.get()) or "0"
        s["source_enabled"] = bool(self.v_source.get())
        s["source_position"] = self.v_source_pos.get()
        s["source_name"] = self.v_source_name.get().strip() or "数据来源"
        s["unify"] = bool(self.v_unify.get())
        s["autofit"] = bool(self.v_autofit.get())
        s["recursive"] = bool(self.v_recursive.get())
        s["fuzzy"] = bool(self.v_fuzzy.get())
        s["case_sensitive"] = bool(self.v_case.get())
        s["strip_punct"] = bool(self.v_punct.get())
        s["drop_empty_rows"] = bool(self.v_drop_rows.get())
        s["drop_empty_cols"] = bool(self.v_drop_cols.get())
        s["date_in"] = self.v_date_in.get().strip()
        s["date_out"] = self.v_date_out.get().strip()
        s["sort_by_name"] = bool(self.v_sort_name.get())
        s["merge_header_rows"] = bool(self.v_merge_rows.get())
        s["drop_repeated_headers"] = bool(self.v_drop_repeat.get())
        s["drop_summary_rows"] = bool(self.v_drop_summary.get())
        s["fill_merged_cells"] = bool(self.v_fill_merged.get())
        out = self.v_output.get().strip()
        if out:
            s["output_dir"] = str(Path(out).parent)
        save_settings(s)

    def _on_close(self):
        if self.worker is not None and self.worker.is_alive():
            if not messagebox.askyesno(
                    "确认退出",
                    "合并仍在进行中。退出会中断任务，输出文件不会生成。\n\n确定退出吗？"):
                return
            self.cancel_event.set()
            self.worker.join(timeout=3)
        self._persist_settings()
        self.root.destroy()


def apply_default_fonts(root):
    """优先用系统里的中文字体，避免默认回退到宋体显得发粗。"""
    import tkinter.font as tkfont
    try:
        families = set(tkfont.families(root))
    except tk.TclError:
        return
    for family in ("Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", "Arial"):
        if family not in families:
            continue
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(name).configure(family=family)
            except tk.TclError:
                pass
        return


class HeaderDialog(tk.Toplevel):
    """逐文件表头行：显示自动检测结果，允许手工改正。"""

    def __init__(self, app, files, sheet_name, overrides):
        tk.Toplevel.__init__(self, app.root)
        self.app = app
        self.files = files
        self.sheet_name = sheet_name
        self.initial = dict(overrides)
        self.rows = {}
        self.queue = queue.Queue()

        self.title("每文件表头行")
        self.transient(app.root)
        self.geometry("%dx%d" % (app._px(780), app._px(440)))
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        ttk.Label(self, foreground="#555",
                  text="下面是从各文件自动检测到的表头行；选中一行，在底部改成正确的行号。"
                  ).grid(row=0, column=0, sticky="w", padx=10, pady=(10, 4))

        wrap = ttk.Frame(self)
        wrap.grid(row=1, column=0, sticky="nsew", padx=10)
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(wrap, columns=("file", "row", "reason"),
                                 show="headings", height=10)
        self.tree.heading("file", text="文件")
        self.tree.column("file", width=app._px(260))
        self.tree.heading("row", text="表头行")
        self.tree.column("row", width=app._px(70), anchor="center")
        self.tree.heading("reason", text="依据")
        self.tree.column("reason", width=app._px(380))
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        bar = ttk.Frame(self)
        bar.grid(row=2, column=0, sticky="ew", padx=10, pady=8)
        ttk.Label(bar, text="改为第").pack(side="left")
        self.v_row = tk.StringVar()
        ttk.Entry(bar, textvariable=self.v_row, width=6).pack(side="left", padx=4)
        ttk.Label(bar, text="行").pack(side="left")
        ttk.Button(bar, text="应用到选中", command=self._apply_selected).pack(side="left", padx=6)
        ttk.Button(bar, text="全部设为该行", command=self._apply_all).pack(side="left")
        ttk.Button(bar, text="确定", command=self._ok).pack(side="right")
        ttk.Button(bar, text="取消", command=self.destroy).pack(side="right", padx=6)

        self.status = ttk.Label(self, text="正在检测…", foreground="#666")
        self.status.grid(row=3, column=0, sticky="w", padx=10, pady=(0, 8))

        self.tree.bind("<Double-1>", self._on_double)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._poll()
        threading.Thread(target=self._scan, daemon=True).start()

    def _scan(self):
        for path in self.files:
            try:
                info = inspect_headers(path, self.sheet_name)
                row = info.row + 1
                reason = info.reason
            except Exception as e:
                row, reason = 1, "读取失败：%s" % e
            self.queue.put((path, row, reason))
        self.queue.put(None)

    def _poll(self):
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return
        try:
            while True:
                item = self.queue.get_nowait()
                if item is None:
                    self.status.config(text="检测完成，共 %d 个文件" % len(self.files))
                    continue
                path, row, reason = item
                base = self.initial.get(path)
                if base:
                    row = int(base)
                self.rows[path] = int(row)
                self.tree.insert("", "end", iid=path,
                                 values=(Path(path).name, row, reason))
        except queue.Empty:
            pass
        self.after(80, self._poll)

    def _on_double(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.v_row.set(str(self.rows.get(item, "")))

    def _read_row(self):
        try:
            row = int(self.v_row.get().strip())
            if row < 1:
                raise ValueError
        except ValueError:
            messagebox.showwarning("提示", "行号必须是大于等于 1 的整数", parent=self)
            return None
        return row

    def _apply_selected(self):
        row = self._read_row()
        if row is None:
            return
        for item in self.tree.selection():
            self.rows[item] = row
            values = list(self.tree.item(item, "values"))
            values[1] = row
            self.tree.item(item, values=values)

    def _apply_all(self):
        row = self._read_row()
        if row is None:
            return
        for item in self.tree.get_children():
            self.rows[item] = row
            values = list(self.tree.item(item, "values"))
            values[1] = row
            self.tree.item(item, values=values)

    def _ok(self):
        self.app.apply_header_overrides(self.rows)
        self.destroy()


def main():
    enable_dpi_awareness()
    root = tk.Tk()
    apply_default_fonts(root)
    try:
        ttk.Style().theme_use("clam")
    except tk.TclError:
        pass
    app = MergeApp(root)
    root.protocol("WM_DELETE_WINDOW", app._on_close)
    root.mainloop()


if __name__ == "__main__":  # pragma: no cover
    main()
