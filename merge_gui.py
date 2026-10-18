# -*- coding: utf-8 -*-
"""
Excel 批量合并工具（图形界面版）

表头对齐策略：
  - 以第一个文件的表头为准，列顺序不变
  - 后续文件表头若与主表头仅存在"格式差异"（空格、全角半角、大小写、
    换行、不可见字符等），自动视为同一列并对齐到主表头
  - 日志会打印每一条对齐记录，方便核对
  - 新增"数据来源"列，记录数据来自哪个文件
  - 输出 xlsx 后自动调整列宽

依赖：
  pip install pandas openpyxl xlrd pyxlsb odfpy
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import queue
import traceback
import datetime
from pathlib import Path
import re
import unicodedata
import pandas as pd
import numpy as np


# =====================================================================
#                         一、核心合并逻辑
# =====================================================================

ENGINE_MAP = {
    ".xlsx": "openpyxl",
    ".xlsm": "openpyxl",
    ".xls":  "xlrd",
    ".xlsb": "pyxlsb",
    ".ods":  "odf",
}
SUPPORTED_EXTS = tuple(ENGINE_MAP.keys()) + (".csv",)

NA_TOKENS = {
    "", "-", "--", "—", "/", "\\", "na", "n/a", "null", "none",
    "nan", "nil", "无", "空", "暂无", "#n/a", "#value!", "#div/0!",
}
TRUE_TOKENS  = {"是", "有", "y", "yes", "true", "t", "√", "✓", "对"}
FALSE_TOKENS = {"否", "无", "n", "no", "false", "f", "×", "✗", "错"}

# 不可见字符（零宽空格、BOM、软连字符等），匹配时全部剔除
_INVISIBLE_CHARS = "\u200b\u200c\u200d\u2060\ufeff\u00ad\u180e"

# 各类空白字符，匹配时统一转成普通空格
_WHITESPACE_CHARS = (
    "\u3000"   # 全角空格
    "\u00a0"   # 不换行空格
    "\u2000\u2001\u2002\u2003\u2004\u2005\u2006"
    "\u2007\u2008\u2009\u200a"
    "\u202f"   # 窄不换行空格
    "\u205f"   # 数学中等空格
    "\t\r\n\x0b\x0c"
)


def _noop(*args, **kwargs):
    pass


def _match_key(name, case_sensitive=False, strip_punct=False):
    """
    生成用于表头匹配的键。
    处理内容：剔除不可见字符、各类空白统一为普通空格、全角转半角、
            连续空格压缩、大小写、可选标点。
    仅用于匹配，不改变输出表头。
    """
    if name is None:
        return ""
    if isinstance(name, float) and pd.isna(name):
        return ""

    s = str(name)

    # 1. 剔除不可见字符
    for ch in _INVISIBLE_CHARS:
        s = s.replace(ch, "")

    # 2. 各类空白统一为普通空格
    for ch in _WHITESPACE_CHARS:
        s = s.replace(ch, " ")

    # 3. 全角转半角（数字、字母、常见标点；汉字不变）
    s = unicodedata.normalize("NFKC", s)

    # 4. 连续空白压缩成一个空格，去首尾
    s = re.sub(r"\s+", " ", s).strip()

    # 5. 大小写
    if not case_sensitive:
        s = s.lower()

    # 6. 可选：去掉中英文常见标点
    if strip_punct:
        s = re.sub(r"[，。、；：！？（）【】《》\"'`,.;:!?()\[\]{}<>]", "", s)

    return s


def _to_halfwidth(s):
    return unicodedata.normalize("NFKC", str(s))


def _clean_text(s):
    s = _to_halfwidth(s)
    for ch in _WHITESPACE_CHARS:
        s = s.replace(ch, " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _to_number(v):
    if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
        return v
    if not isinstance(v, str):
        return v
    s = _clean_text(v)
    if not s:
        return np.nan

    is_pct = s.endswith("%")
    if is_pct:
        s = s[:-1].strip()

    s = re.sub(r"[¥$€£￥]", "", s)
    s = re.sub(r"(元|人民币|rmb|usd|美元|万元|万|元整)$", "", s, flags=re.I)
    s = s.replace(",", "").replace("，", "").strip()

    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1].strip()

    try:
        num = float(s)
    except ValueError:
        return v

    if is_pct:
        num /= 100.0
    if neg:
        num = -num
    if num.is_integer() and not is_pct:
        return int(num)
    return num


def _to_bool(v):
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)) and not pd.isna(v):
        if v == 1:
            return True
        if v == 0:
            return False
        return v
    if isinstance(v, str):
        t = _clean_text(v).lower()
        if t in TRUE_TOKENS:
            return True
        if t in FALSE_TOKENS:
            return False
    return v


def _safe_to_datetime(s, fmt=None):
    try:
        if fmt:
            return pd.to_datetime(s, errors="coerce", format=fmt)
        return pd.to_datetime(s, errors="coerce", format="mixed")
    except (TypeError, ValueError):
        return pd.to_datetime(s, errors="coerce")


def _infer_col_type(series):
    non_na = series.dropna()
    if len(non_na) == 0:
        return "text"

    if pd.api.types.is_datetime64_any_dtype(series):
        return "datetime"

    bool_ok = sum(
        1 for v in non_na
        if isinstance(v, bool)
        or (isinstance(v, str) and _clean_text(v).lower() in (TRUE_TOKENS | FALSE_TOKENS))
    )
    if bool_ok / len(non_na) >= 0.9:
        str_vals = [v for v in non_na if isinstance(v, str)]
        if str_vals or all(isinstance(v, bool) for v in non_na):
            return "bool"

    num_ok = 0
    for v in non_na:
        if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
            num_ok += 1
        elif isinstance(v, str):
            r = _to_number(v)
            if not isinstance(r, str):
                num_ok += 1
    if num_ok / len(non_na) >= 0.9:
        return "number"

    if non_na.map(lambda x: isinstance(x, str)).all():
        sample = non_na.head(200)
        if not sample.str.fullmatch(r"-?\d+(\.\d+)?").all():
            parsed = _safe_to_datetime(sample)
            if parsed.notna().mean() >= 0.8:
                return "datetime"

    return "text"


def _convert_column(series, dtype, date_format=None):
    s = series.map(
        lambda v: np.nan if (isinstance(v, str) and _clean_text(v).lower() in NA_TOKENS) else v
    )
    if dtype == "number":
        return s.map(_to_number).astype("Float64")
    if dtype == "bool":
        return s.map(_to_bool).astype("boolean")
    if dtype == "datetime":
        return _safe_to_datetime(s, fmt=date_format)
    return s.map(lambda v: _clean_text(v) if isinstance(v, str) else v)


def normalize_dataframe(df, date_format=None, force_types=None):
    df = df.copy()
    force_types = force_types or {}
    for col in df.columns:
        if col in force_types:
            df[col] = _convert_column(df[col], force_types[col], date_format)
            continue
        if pd.api.types.is_numeric_dtype(df[col]) and not pd.api.types.is_bool_dtype(df[col]):
            continue
        dtype = _infer_col_type(df[col])
        df[col] = _convert_column(df[col], dtype, date_format)
    return df


def _read_raw(path, sheet_name=0):
    if sheet_name is None:
        sheet_name = 0
    ext = path.suffix.lower()

    if ext == ".csv":
        last_err = None
        for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
            try:
                return pd.read_csv(path, header=None, encoding=enc, dtype=object)
            except UnicodeDecodeError as e:
                last_err = e
        raise last_err or ValueError("CSV 编码无法识别")

    if ext not in ENGINE_MAP:
        raise ValueError("不支持的文件格式：" + ext)

    return pd.read_excel(
        path, sheet_name=sheet_name, header=None,
        engine=ENGINE_MAP[ext], dtype=object,
    )


def _detect_header_row(df, max_scan=20):
    n = min(len(df), max_scan)
    if n == 0:
        return 0
    counts = []
    for i in range(n):
        row = df.iloc[i]
        nn = int(row.notna().sum())
        vals = [v for v in row if pd.notna(v)]
        sl = sum(1 for v in vals if isinstance(v, str) and v.strip())
        counts.append((nn, sl))
    max_nn = max(c[0] for c in counts)
    threshold = max(2, int(max_nn * 0.6))
    for i, (nn, sl) in enumerate(counts):
        if nn >= threshold and sl >= max(1, nn * 0.5):
            return i
    return max(range(n), key=lambda i: (counts[i][0], -i))


# ---------------------------------------------------------------- 列宽自适应
def _cell_display_width(v):
    if v is None:
        return 0
    if isinstance(v, datetime.datetime):
        if v.time() == datetime.time(0, 0, 0):
            s = v.strftime("%Y-%m-%d")
        else:
            s = v.strftime("%Y-%m-%d %H:%M")
    elif isinstance(v, datetime.date):
        s = v.strftime("%Y-%m-%d")
    elif isinstance(v, float) and v.is_integer():
        s = str(int(v))
    else:
        s = str(v)

    if "\n" in s:
        s = max(s.split("\n"), key=len)
    s = s.replace("\t", " ")

    w = 0
    for ch in s:
        w += 2 if ord(ch) > 127 else 1
    return w


def _autofit_excel_columns(path, sheet_name="汇总",
                           min_width=8, max_width=60,
                           padding=3.0, sample_rows=2000):
    try:
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
    except ImportError:
        return

    try:
        wb = load_workbook(path)
    except Exception:
        return

    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.active

    n_rows = min(ws.max_row, sample_rows + 1)
    n_cols = ws.max_column

    for c in range(1, n_cols + 1):
        letter = get_column_letter(c)
        max_len = 0
        for r in range(1, n_rows + 1):
            v = ws.cell(row=r, column=c).value
            w = _cell_display_width(v)
            if w > max_len:
                max_len = w
        width = max(min_width, min(max_width, max_len + padding))
        ws.column_dimensions[letter].width = width

    try:
        wb.save(path)
    except Exception:
        pass


# -------------------------------------------------------------- 合并主函数
def merge_files(
    file_paths,
    output_path,
    header_row=1,
    sheet_name=0,
    add_source_column=True,
    source_column_name="数据来源",
    source_position="last",
    drop_empty_rows=True,
    drop_empty_cols=False,
    column_name_map=None,
    unify_formats=True,
    date_format=None,
    output_date_format=None,
    force_types=None,
    autofit_columns=True,

    # 表头容错匹配
    fuzzy_header_match=True,        # 是否启用容错匹配
    header_case_sensitive=False,    # 匹配时是否区分大小写
    header_strip_punct=False,       # 匹配时是否忽略标点

    log_callback=None,
    progress_callback=None,
):
    log = log_callback or _noop
    progress = progress_callback or _noop

    out_path = Path(output_path)
    out_resolved = out_path.resolve()

    files = []
    for p in file_paths:
        f = Path(p)
        if not f.is_file() or f.name.startswith("~$"):
            continue
        if f.resolve() == out_resolved:
            continue
        files.append(f)
    files.sort(key=lambda p: p.name)

    if not files:
        raise ValueError("没有可用的输入文件")

    total = len(files)
    progress(0, total)

    # 主表头（原始显示名）与对应匹配键
    master_cols = []
    master_keys = []

    def make_key(name):
        if not fuzzy_header_match:
            return str(name)   # 精确匹配
        return _match_key(name,
                          case_sensitive=header_case_sensitive,
                          strip_punct=header_strip_punct)

    frames, failed = [], []
    for idx, f in enumerate(files, start=1):
        try:
            raw = _read_raw(f, sheet_name=sheet_name)
        except Exception as e:
            failed.append((f.name, repr(e)))
            log("[跳过] " + f.name + " 读取失败：" + str(e))
            progress(idx, total)
            continue

        if len(raw) == 0:
            log("[跳过] " + f.name + " 内容为空")
            progress(idx, total)
            continue

        # 表头行
        if header_row is None:
            h = _detect_header_row(raw)
        elif isinstance(header_row, dict):
            h = int(header_row.get(f.name, header_row.get(f.stem, 1))) - 1
        else:
            h = int(header_row) - 1
        h = max(0, min(h, max(len(raw) - 1, 0)))

        header_vals = raw.iloc[h].tolist()
        cols = []
        for j, name in enumerate(header_vals):
            if name is None or name == "" or (isinstance(name, float) and pd.isna(name)):
                cols.append("列" + str(j + 1))
            else:
                cols.append(name)

        # 同一表内重名去重
        seen = {}
        final_cols = []
        for c in cols:
            key = str(c)
            if key in seen:
                seen[key] += 1
                final_cols.append(str(c) + "_" + str(seen[key]))
            else:
                seen[key] = 0
                final_cols.append(c)

        body = raw.iloc[h + 1:].reset_index(drop=True)
        body.columns = final_cols

        # ============ 表头对齐 ============
        if not master_cols:
            # 第一个文件：作为主表头
            master_cols = list(final_cols)
            master_keys = [make_key(c) for c in final_cols]
            log("[" + str(idx) + "/" + str(total) + "] 已读取 " + f.name +
                "  表头行=" + str(h + 1) +
                "  数据=" + str(body.shape[0]) + " 行 x " + str(body.shape[1]) + " 列")
            log("  该文件作为主表头，共 " + str(len(master_cols)) + " 列")
        else:
            key_to_master = {}
            for i, k in enumerate(master_keys):
                if k and k not in key_to_master:
                    key_to_master[k] = master_cols[i]

            rename_map = {}
            matched_diff = []     # 记录"字面不同但被匹配为同一列"的表头
            extra_cols = []
            for c in final_cols:
                k = make_key(c)
                if k and k in key_to_master:
                    target = key_to_master[k]
                    if str(c) != str(target):
                        rename_map[c] = target
                        matched_diff.append((str(c), str(target)))
                else:
                    extra_cols.append(c)

            if rename_map:
                body = body.rename(columns=rename_map)

            log("[" + str(idx) + "/" + str(total) + "] 已读取 " + f.name +
                "  表头行=" + str(h + 1) +
                "  数据=" + str(body.shape[0]) + " 行 x " + str(body.shape[1]) + " 列")

            if matched_diff:
                log("  表头对齐（格式差异，视为同一列）：")
                for src, dst in matched_diff:
                    log("    " + repr(src) + "  ->  " + repr(dst))

            if extra_cols:
                log("  注意：出现主表头之外的新列 " + str(extra_cols) +
                    "，已追加到表头末尾")
                master_cols = master_cols + extra_cols
                master_keys = master_keys + [make_key(c) for c in extra_cols]

        # 按主表头顺序重排（缺的列补 NaN）
        body = body.reindex(columns=master_cols)

        if drop_empty_rows:
            body = body.dropna(how="all").reset_index(drop=True)

        if drop_empty_cols:
            keep = [c for c in body.columns if not body[c].isna().all()]
            body = body[keep]
            master_cols = keep
            master_keys = [make_key(c) for c in keep]

        # 数据格式统一
        if unify_formats and len(body) > 0:
            force = {}
            for k, v in (force_types or {}).items():
                kk = make_key(k)
                for c in body.columns:
                    if make_key(c) == kk:
                        force[c] = v
            body = normalize_dataframe(body, date_format=date_format, force_types=force)

        # 来源列
        if add_source_column:
            if source_position == "first":
                body.insert(0, source_column_name, f.stem)
            else:
                body[source_column_name] = f.stem

        frames.append(body)
        progress(idx, total)

    if not frames:
        raise RuntimeError("所有文件都读取失败，无法合并")

    log("正在合并数据 ...")
    merged = pd.concat(frames, ignore_index=True, sort=False)

    if column_name_map:
        merged = merged.rename(columns=column_name_map)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    log("正在写入文件：" + str(out_path))

    if out_path.suffix.lower() == ".csv":
        merged.to_csv(out_path, index=False, encoding="utf-8-sig")
    else:
        with pd.ExcelWriter(out_path, engine="openpyxl",
                            datetime_format=output_date_format or "yyyy-mm-dd") as writer:
            merged.to_excel(writer, index=False, sheet_name="汇总")
        if autofit_columns:
            log("正在调整列宽 ...")
            _autofit_excel_columns(out_path, sheet_name="汇总")

    log("完成：" + str(len(frames)) + " 个文件 / 共 " +
        str(len(merged)) + " 行 x " + str(merged.shape[1]) + " 列")
    log("最终表头：" + str(list(merged.columns)))
    if failed:
        log("有 " + str(len(failed)) + " 个文件读取失败：" + str([n for n, _ in failed]))

    return merged


# =====================================================================
#                         二、图形界面
# =====================================================================

class MergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 批量合并工具")
        self.root.geometry("940x820")
        self.root.minsize(840, 720)

        self.files = []
        self.msg_queue = queue.Queue()
        self.worker = None

        self._build_ui()
        self._poll_queue()

    def _build_ui(self):
        # 1. 文件区
        file_frame = ttk.LabelFrame(self.root, text=" 1. 选择要合并的文件 ")
        file_frame.pack(fill="both", expand=False, padx=10, pady=(10, 5))

        btns = ttk.Frame(file_frame)
        btns.pack(fill="x", padx=6, pady=6)
        ttk.Button(btns, text="添加文件", command=self.add_files).pack(side="left", padx=2)
        ttk.Button(btns, text="添加文件夹", command=self.add_folder).pack(side="left", padx=2)
        ttk.Button(btns, text="移除选中", command=self.remove_selected).pack(side="left", padx=2)
        ttk.Button(btns, text="清空列表", command=self.clear_files).pack(side="left", padx=2)
        self.count_label = ttk.Label(btns, text="已选 0 个文件")
        self.count_label.pack(side="right", padx=6)

        list_wrap = ttk.Frame(file_frame)
        list_wrap.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.file_listbox = tk.Listbox(list_wrap, selectmode="extended", height=6,
                                       activestyle="none")
        sb = ttk.Scrollbar(list_wrap, orient="vertical", command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=sb.set)
        self.file_listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        ttk.Label(file_frame,
                  text="提示：合并后的表头以列表中第一个文件的表头为准，原列顺序不变。",
                  foreground="#666").pack(anchor="w", padx=10, pady=(0, 6))

        # 2. 参数区
        param_frame = ttk.LabelFrame(self.root, text=" 2. 参数设置 ")
        param_frame.pack(fill="x", padx=10, pady=5)

        row1 = ttk.Frame(param_frame)
        row1.pack(fill="x", padx=6, pady=4)
        ttk.Label(row1, text="表头所在行：").pack(side="left")
        self.header_mode = tk.StringVar(value="auto")
        ttk.Radiobutton(row1, text="自动检测", variable=self.header_mode,
                        value="auto").pack(side="left", padx=4)
        ttk.Radiobutton(row1, text="固定行号", variable=self.header_mode,
                        value="fixed").pack(side="left", padx=4)
        self.header_entry = ttk.Entry(row1, width=6)
        self.header_entry.insert(0, "1")
        self.header_entry.pack(side="left")
        ttk.Label(row1, text="（从 1 开始数）").pack(side="left", padx=4)

        row2 = ttk.Frame(param_frame)
        row2.pack(fill="x", padx=6, pady=4)
        ttk.Label(row2, text="工作表：").pack(side="left")
        self.sheet_entry = ttk.Entry(row2, width=16)
        self.sheet_entry.insert(0, "0")
        self.sheet_entry.pack(side="left", padx=4)
        ttk.Label(row2, text="（0 表示第一个；也可填工作表名）").pack(side="left")

        # 表头容错匹配
        row_fuzzy = ttk.LabelFrame(param_frame, text=" 表头容错匹配 ")
        row_fuzzy.pack(fill="x", padx=6, pady=4)

        fr1 = ttk.Frame(row_fuzzy)
        fr1.pack(fill="x", padx=4, pady=3)
        self.opt_fuzzy = tk.BooleanVar(value=True)
        ttk.Checkbutton(fr1,
                        text="启用（忽略多余空格、全角半角、换行、不可见字符等）",
                        variable=self.opt_fuzzy).pack(side="left")
        self.opt_case = tk.BooleanVar(value=False)
        ttk.Checkbutton(fr1, text="区分大小写", variable=self.opt_case).pack(side="left", padx=(12, 0))
        self.opt_punct = tk.BooleanVar(value=False)
        ttk.Checkbutton(fr1, text="忽略标点", variable=self.opt_punct).pack(side="left", padx=(12, 0))

        # 来源列 + 格式统一
        row3 = ttk.Frame(param_frame)
        row3.pack(fill="x", padx=6, pady=4)
        self.opt_source = tk.BooleanVar(value=True)
        self.opt_unify = tk.BooleanVar(value=True)
        self.opt_recursive = tk.BooleanVar(value=False)
        self.opt_autofit = tk.BooleanVar(value=True)
        self.source_pos = tk.StringVar(value="last")

        ttk.Checkbutton(row3, text="添加“数据来源”列",
                        variable=self.opt_source).pack(side="left", padx=(0, 6))
        ttk.Label(row3, text="位置：").pack(side="left")
        ttk.Radiobutton(row3, text="末列", variable=self.source_pos,
                        value="last").pack(side="left")
        ttk.Radiobutton(row3, text="首列", variable=self.source_pos,
                        value="first").pack(side="left", padx=(0, 12))
        ttk.Checkbutton(row3, text="统一数据格式",
                        variable=self.opt_unify).pack(side="left", padx=(0, 12))

        row3b = ttk.Frame(param_frame)
        row3b.pack(fill="x", padx=6, pady=4)
        ttk.Checkbutton(row3b, text="自动调整列宽（仅 xlsx）",
                        variable=self.opt_autofit).pack(side="left", padx=(0, 12))
        ttk.Checkbutton(row3b, text="文件夹递归读取",
                        variable=self.opt_recursive).pack(side="left")

        row4 = ttk.Frame(param_frame)
        row4.pack(fill="x", padx=6, pady=4)
        ttk.Label(row4, text="输出日期格式：").pack(side="left")
        self.date_fmt_entry = ttk.Entry(row4, width=18)
        self.date_fmt_entry.insert(0, "yyyy-mm-dd")
        self.date_fmt_entry.pack(side="left", padx=4)
        ttk.Label(row4, text="（如 yyyy年mm月dd日）").pack(side="left")

        row5 = ttk.Frame(param_frame)
        row5.pack(fill="both", expand=False, padx=6, pady=4)

        left = ttk.LabelFrame(row5, text="强制列类型（每行：列名=类型）")
        left.pack(side="left", fill="both", expand=True, padx=(0, 4))
        self.type_text = tk.Text(left, height=4, width=30)
        self.type_text.pack(fill="both", expand=True, padx=3, pady=3)

        right = ttk.LabelFrame(row5, text="列名重命名（每行：旧名=新名，留空则不改）")
        right.pack(side="left", fill="both", expand=True, padx=(4, 0))
        self.map_text = tk.Text(right, height=4, width=30)
        self.map_text.pack(fill="both", expand=True, padx=3, pady=3)

        ttk.Label(param_frame,
                  text="类型可选：text / number / datetime / bool。示例：身份证号=text",
                  foreground="#666").pack(anchor="w", padx=10, pady=(0, 6))

        # 3. 输出区
        out_frame = ttk.LabelFrame(self.root, text=" 3. 输出设置 ")
        out_frame.pack(fill="x", padx=10, pady=5)
        row = ttk.Frame(out_frame)
        row.pack(fill="x", padx=6, pady=6)
        ttk.Label(row, text="输出文件：").pack(side="left")
        self.out_entry = ttk.Entry(row)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row, text="浏览...", command=self.choose_output).pack(side="left")

        # 4. 运行区
        run_frame = ttk.Frame(self.root)
        run_frame.pack(fill="x", padx=10, pady=5)
        self.run_btn = ttk.Button(run_frame, text="开始合并", command=self.start_merge)
        self.run_btn.pack(side="left")
        self.progress = ttk.Progressbar(run_frame, mode="determinate", maximum=1, value=0)
        self.progress.pack(side="left", fill="x", expand=True, padx=8)

        # 5. 日志区
        log_frame = ttk.LabelFrame(self.root, text=" 运行日志 ")
        log_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        self.log_text = ScrolledText(log_frame, height=12, state="disabled", wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)
        self.log_text.tag_config("error", foreground="#c0392b")
        self.log_text.tag_config("ok", foreground="#1e8449")
        self.log_text.tag_config("info", foreground="#2c3e50")
        self.log_text.tag_config("warn", foreground="#b9770e")

    # ------------------------------------------------- 文件操作
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择要合并的文件",
            filetypes=[
                ("Excel / CSV", "*.xlsx *.xlsm *.xls *.xlsb *.ods *.csv"),
                ("Excel 文件", "*.xlsx *.xlsm *.xls *.xlsb"),
                ("CSV 文件", "*.csv"),
                ("所有文件", "*.*"),
            ],
        )
        self._add_paths(paths)

    def add_folder(self):
        folder = filedialog.askdirectory(title="选择文件夹")
        if not folder:
            return
        folder_path = Path(folder)
        if self.opt_recursive.get():
            candidates = [p for p in folder_path.rglob("*") if p.is_file()]
        else:
            candidates = [p for p in folder_path.iterdir() if p.is_file()]
        valid = [p for p in candidates
                 if p.suffix.lower() in SUPPORTED_EXTS and not p.name.startswith("~$")]
        self._add_paths(valid)

    def _add_paths(self, paths):
        added = 0
        for p in paths:
            f = Path(p)
            if f.suffix.lower() not in SUPPORTED_EXTS:
                continue
            if f.name.startswith("~$"):
                continue
            s = str(f.resolve())
            if s in self.files:
                continue
            self.files.append(s)
            self.file_listbox.insert("end", s)
            added += 1
        self.count_label.config(text="已选 " + str(len(self.files)) + " 个文件")
        if added == 0:
            messagebox.showinfo("提示", "没有新增文件（可能已存在或格式不支持）")

    def remove_selected(self):
        idxs = list(self.file_listbox.curselection())
        for i in reversed(idxs):
            self.file_listbox.delete(i)
            del self.files[i]
        self.count_label.config(text="已选 " + str(len(self.files)) + " 个文件")

    def clear_files(self):
        self.files.clear()
        self.file_listbox.delete(0, "end")
        self.count_label.config(text="已选 0 个文件")

    def choose_output(self):
        path = filedialog.asksaveasfilename(
            title="保存合并结果",
            defaultextension=".xlsx",
            initialfile="合并结果.xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("CSV 文件", "*.csv")],
        )
        if path:
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, path)

    # ------------------------------------------------- 参数收集
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
        if not self.files:
            raise ValueError("请先添加要合并的文件")

        out_path = self.out_entry.get().strip()
        if not out_path:
            raise ValueError("请先设置输出文件路径")

        mode = self.header_mode.get()
        if mode == "auto":
            header_row = None
        else:
            try:
                header_row = int(self.header_entry.get().strip())
                if header_row < 1:
                    raise ValueError
            except ValueError:
                raise ValueError("表头行号必须是大于等于 1 的整数")

        sheet_raw = self.sheet_entry.get().strip()
        if sheet_raw == "" or sheet_raw == "0":
            sheet_name = 0
        else:
            try:
                sheet_name = int(sheet_raw)
            except ValueError:
                sheet_name = sheet_raw

        column_name_map = self._parse_kv(self.map_text.get("1.0", "end"))
        force_types = self._parse_kv(self.type_text.get("1.0", "end"))
        date_fmt = self.date_fmt_entry.get().strip() or None

        return dict(
            output_path=out_path,
            header_row=header_row,
            sheet_name=sheet_name,
            add_source_column=self.opt_source.get(),
            source_position=self.source_pos.get(),
            unify_formats=self.opt_unify.get(),
            output_date_format=date_fmt,
            column_name_map=column_name_map or None,
            force_types=force_types or None,
            autofit_columns=self.opt_autofit.get(),
            fuzzy_header_match=self.opt_fuzzy.get(),
            header_case_sensitive=self.opt_case.get(),
            header_strip_punct=self.opt_punct.get(),
        )

    # ------------------------------------------------- 运行
    def start_merge(self):
        try:
            opts = self._collect_options()
        except ValueError as e:
            messagebox.showwarning("提示", str(e))
            return

        if self.worker and self.worker.is_alive():
            messagebox.showinfo("提示", "任务正在进行中，请稍候 ...")
            return

        self.run_btn.config(state="disabled")
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.files)
        self._append_log("=" * 60, "info")
        self._append_log("开始合并 " + str(len(self.files)) + " 个文件 ...", "info")

        files = list(self.files)
        self.worker = threading.Thread(
            target=self._run_merge, args=(files, opts), daemon=True
        )
        self.worker.start()

    def _run_merge(self, files, opts):
        try:
            merge_files(
                file_paths=files,
                output_path=opts["output_path"],
                header_row=opts["header_row"],
                sheet_name=opts["sheet_name"],
                add_source_column=opts["add_source_column"],
                source_position=opts["source_position"],
                unify_formats=opts["unify_formats"],
                output_date_format=opts["output_date_format"],
                column_name_map=opts["column_name_map"],
                force_types=opts["force_types"],
                autofit_columns=opts["autofit_columns"],
                fuzzy_header_match=opts["fuzzy_header_match"],
                header_case_sensitive=opts["header_case_sensitive"],
                header_strip_punct=opts["header_strip_punct"],
                log_callback=lambda m: self.msg_queue.put(("log", m, "info")),
                progress_callback=lambda c, t: self.msg_queue.put(("progress", c, t)),
            )
            self.msg_queue.put(("done", opts["output_path"], None))
        except Exception as e:
            self.msg_queue.put(("error", str(e), traceback.format_exc()))

    def _poll_queue(self):
        try:
            while True:
                kind = self.msg_queue.get_nowait()
                if kind[0] == "log":
                    _, msg, tag = kind
                    self._append_log(msg, tag)
                elif kind[0] == "progress":
                    _, cur, total = kind
                    self.progress["maximum"] = max(total, 1)
                    self.progress["value"] = cur
                elif kind[0] == "done":
                    _, path, _ = kind
                    self._append_log("已生成文件：" + path, "ok")
                    self.run_btn.config(state="normal")
                    messagebox.showinfo("完成", "合并完成！\n\n输出文件：\n" + path)
                elif kind[0] == "error":
                    _, msg, tb = kind
                    self._append_log("出错：" + msg, "error")
                    self._append_log(tb, "error")
                    self.run_btn.config(state="normal")
                    messagebox.showerror("出错", msg)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_queue)

    def _append_log(self, msg, tag="info"):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


# =====================================================================
#                              启动入口
# =====================================================================

def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        style.theme_use("clam")
    except tk.TclError:
        pass
    app = MergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()