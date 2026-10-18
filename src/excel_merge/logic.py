# -*- coding: utf-8 -*-
"""
Excel 合并引擎（与界面完全解耦，可单独调用 / 单独测试）。

表头对齐策略：
  - 以第一个文件的表头为准，列顺序不变（文件顺序由调用方决定，引擎不再重排）
  - 后续文件表头若与主表头仅存在"格式差异"（空格、全角半角、大小写、
    换行、不可见字符等），自动视为同一列并对齐到主表头
  - 智能同义词识别（如"身份证"与"身份证号码"自动对齐）
  - 每条对齐记录都会回传给日志回调，方便核对
  - 可选的"数据来源"列，记录数据来自哪个文件
  - 输出 xlsx 后自动调整列宽
  - 先写临时文件再原子替换，中途失败不会留下半截的输出文件
"""

import datetime
import os
import re
import unicodedata
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

try:  # dataclasses 在 3.7+ 是标准库；保留回退只是为了让报错更清楚
    from dataclasses import dataclass, field
except ImportError:  # pragma: no cover
    raise RuntimeError("需要 Python 3.7 及以上版本（dataclasses）")

__all__ = [
    "MergeOptions",
    "MergeCancelled",
    "merge_files",
    "SUPPORTED_EXTS",
    "ALLOWED_TYPES",
]

ENGINE_MAP = {
    ".xlsx": "openpyxl",
    ".xlsm": "openpyxl",
    ".xls": "xlrd",
    ".xlsb": "pyxlsb",
    ".ods": "odf",
}
SUPPORTED_EXTS = tuple(ENGINE_MAP.keys()) + (".csv",)

# 输出只支持这两种：其余后缀会让 openpyxl 直接抛异常
OUTPUT_EXTS = (".xlsx", ".csv")

NA_TOKENS = {
    "", "-", "--", "—", "/", "\\", "na", "n/a", "null", "none",
    "nan", "nil", "无", "空", "暂无", "#n/a", "#value!", "#div/0!",
}
TRUE_TOKENS = {"是", "有", "y", "yes", "true", "t", "√", "✓", "对"}
FALSE_TOKENS = {"否", "无", "n", "no", "false", "f", "×", "✗", "错"}

# 表头智能探测关键词库（增加权重，降低大标题误判）
HEADER_KEYWORDS = {
    "日期", "时间", "编号", "姓名", "名称", "金额", "数量", "单价", "部门",
    "状态", "手机", "电话", "地址", "id", "name", "date", "amount", "备注", "分类",
}

# 表头同义词映射（左边是变体，右边是标准词）
SYNONYM_MAP = {
    "身份证": "证件号",
    "身份证号码": "证件号",
    "证件号码": "证件号",
    "手机": "电话",
    "手机号": "电话",
    "联系电话": "电话",
    "金额(元)": "金额",
    "总金额": "金额",
    "订单日期": "日期",
    "交易时间": "日期",
}

# 强制列类型允许的取值
ALLOWED_TYPES = ("text", "number", "datetime", "bool")

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

_PUNCT_RE = re.compile(r"[，。、；：！？（）【】《》\"'`,.;:!?()\[\]{}<>]")
_SPACE_RE = re.compile(r"\s+")
_NUMBER_RE = re.compile(r"-?\d+(\.\d+)?")


class MergeCancelled(Exception):
    """用户主动取消了合并。"""


@dataclass
class MergeOptions:
    """一次合并所需的全部参数。

    以前这些参数散在 merge_files 的 20 个关键字参数、界面的 dict 和
    线程入口三处，改一个要同步三处；现在只有这一个定义。
    """

    output_path: str
    header_row: object = 1          # int 表示固定行号；None 表示自动检测
    sheet_name: object = 0          # int 序号或工作表名
    add_source_column: bool = True
    source_column_name: str = "数据来源"
    source_position: str = "last"   # "last" / "first"
    drop_empty_rows: bool = True
    drop_empty_cols: bool = False
    column_name_map: dict = field(default_factory=dict)
    unify_formats: bool = True
    date_format: object = None      # 读取时的日期格式
    output_date_format: object = "yyyy-mm-dd"
    force_types: dict = field(default_factory=dict)
    autofit_columns: bool = True
    fuzzy_header_match: bool = True
    header_case_sensitive: bool = False
    header_strip_punct: bool = False
    sort_by_name: bool = False      # 是否按文件名排序；默认保持调用方给的顺序
    # 表头容错
    merge_header_rows: object = "auto"   # "auto" / 0 / 1 / 2… 多行表头合并成几行
    drop_repeated_headers: bool = True   # 丢弃数据区里重复出现的表头行
    drop_summary_rows: bool = True       # 丢弃表格末尾的合计/小计行
    fill_merged_header_cells: bool = True  # 合并单元格的表头向右填充（仅 xlsx）


@dataclass
class HeaderInfo:
    """一张表的表头识别结果。"""

    row: int                 # 主表头行（0 基）
    extra_rows: int = 0      # 主表头上方还有几行属于表头（分组表头）
    confidence: float = 0.0
    reason: str = ""

    @property
    def first(self):
        return max(0, self.row - self.extra_rows)

    @property
    def rows(self):
        return list(range(self.first, self.row + 1))


def _noop(*args, **kwargs):
    pass


def _match_key(name, case_sensitive=False, strip_punct=False):
    """
    生成用于表头匹配的键。
    处理内容：剔除不可见字符、空白统一、全角转半角、同义词映射、大小写、标点。
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

    # 3. 全角转半角
    s = unicodedata.normalize("NFKC", s)

    # 4. 连续空白压缩成一个空格，去首尾
    s = _SPACE_RE.sub(" ", s).strip()

    # 5. 大小写
    if not case_sensitive:
        s = s.lower()

    # 6. 可选：去掉中英文常见标点
    if strip_punct:
        s = _PUNCT_RE.sub("", s)

    # 7. 同义词智能替换（如 身份证号 -> 证件号）
    if s in SYNONYM_MAP:
        s = SYNONYM_MAP[s]

    return s


def _to_halfwidth(s):
    return unicodedata.normalize("NFKC", str(s))


def _clean_text(s):
    s = _to_halfwidth(s)
    for ch in _WHITESPACE_CHARS:
        s = s.replace(ch, " ")
    s = _SPACE_RE.sub(" ", s).strip()
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

    # 处理中文单位
    multiplier = 1
    if s.endswith("千万"):
        multiplier = 10000000
        s = s[:-2]
    elif s.endswith("万") or s.endswith("萬"):
        multiplier = 10000
        s = s[:-1]
    elif s.endswith("亿"):
        multiplier = 100000000
        s = s[:-1]

    s = re.sub(r"[¥$€£￥]", "", s)
    s = re.sub(r"(元|人民币|rmb|usd|美元|元整)$", "", s, flags=re.I)
    s = s.replace(",", "").replace("，", "").strip()

    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1].strip()

    try:
        num = float(s) * multiplier
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
    """修复了 pandas 1.5.3 不支持 format='mixed' 的兼容性问题。

    pandas 2.x 在无法推断格式时会逐行回退到 dateutil 并对每一行发
    UserWarning；推断日期本来就是在文本列上试探，这类警告只会刷屏，
    这里统一静音。
    """
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            if fmt:
                return pd.to_datetime(s, errors="coerce", format=fmt)
            # 尝试自动推断（兼容旧版 pandas）
            return pd.to_datetime(s, errors="coerce")
    except Exception:
        return pd.to_datetime(s, errors="coerce")


def _infer_col_type(series):
    """优化了推断逻辑，减少 map 循环，利用 pandas 原生提速"""
    non_na = series.dropna()
    if len(non_na) == 0:
        return "text"

    if pd.api.types.is_datetime64_any_dtype(series):
        return "datetime"

    # 1. 布尔值判断
    bool_mask = non_na.astype(str).str.lower().str.strip().isin(TRUE_TOKENS | FALSE_TOKENS)
    if bool_mask.mean() >= 0.9:
        return "bool"

    # 2. 数字判断（结合了中文单位和 NA 过滤）
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

    # 3. 日期判断
    if non_na.map(lambda x: isinstance(x, str)).all():
        sample = non_na.head(200)
        if not sample.str.fullmatch(_NUMBER_RE).all():
            parsed = _safe_to_datetime(sample)
            if parsed.notna().mean() >= 0.8:
                return "datetime"

    return "text"


def _convert_column(series, dtype, date_format=None):
    # 先统一将 NA 字符串替换为真正的 np.nan
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


def _detect_header_row(df, max_scan=30):
    """旧接口：只返回行号。保留给外部调用，内部用 detect_header()。"""
    return detect_header(df, max_scan=max_scan).row


def _is_blank_value(v):
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    if isinstance(v, str) and not v.strip():
        return True
    return False


def _row_profile(df, i):
    """(非空数, 短文本数, 关键词命中, 最长文本长度)"""
    row = df.iloc[i]
    vals = [v for v in row if not _is_blank_value(v)]
    texts = [_clean_text(v) for v in vals]
    n = len(vals)
    short = sum(1 for t in texts if len(t) <= 24 and not _NUMBER_RE.fullmatch(t))
    keywords = sum(1 for t in texts if any(kw in t.lower() for kw in HEADER_KEYWORDS))
    longest = max((len(t) for t in texts), default=0)
    return n, short, keywords, longest


def _header_score(profile):
    n, short, keywords, longest = profile
    score = n * 1.0 + short * 2.0 + keywords * 8.0
    if n and short / n < 0.5:
        score -= n            # 数字为主的行不像表头
    if longest > 40:
        score -= 6            # 长句更像标题/说明
    return score


def detect_header(df, max_scan=30):
    """识别表头行，并判断上方是否还有一行分组表头。

    返回 HeaderInfo：row 是主表头行（0 基），extra_rows 是它上方还有几行
    属于表头，confidence 是"比第二名高多少"的比值，reason 写进日志。
    """
    n = min(len(df), max_scan)
    if n == 0:
        return HeaderInfo(0, 0, 0.0, "空表")

    profiles = [(i, _row_profile(df, i)) for i in range(n)]
    scored = [(i, _header_score(p)) for i, p in profiles if p[0] >= 1]
    if not scored:
        return HeaderInfo(0, 0, 0.0, "没有非空行")

    best_score = max(s for _, s in scored)
    # 分数接近时取最靠前的一行：标题/说明行往往分数明显低，不会进这个区间
    near = [i for i, s in scored if s >= best_score * 0.95]
    row = min(near)
    second = max((s for i, s in scored if i != row), default=0)
    confidence = 1.0 if second <= 0 else min(1.0, best_score / second)

    # 多行表头：上一行只有少数几个短标签（分组名），且命中关键词
    extra = 0
    if row > 0:
        above = profiles[row - 1][1]
        below = profiles[row][1]
        if 2 <= above[0] < below[0] and above[3] <= 12 and above[2] >= 1:
            extra = 1

    prof = profiles[row][1]
    reason = "非空 %d 格 / 关键词 %d 个 / 最长 %d 字 / 置信度 %.2f" % (
        prof[0], prof[2], prof[3], confidence)
    return HeaderInfo(row, extra, confidence, reason)


def _merged_cells_map(path, sheet_name, max_row):
    """读取合并单元格，把左上角的值铺满整个合并区域（只取前 max_row 行）。"""
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        return {}
    try:
        from openpyxl import load_workbook
    except ImportError:
        return {}
    try:
        wb = load_workbook(str(path), data_only=True)
        try:
            if isinstance(sheet_name, str) and sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
            else:
                idx = int(sheet_name or 0)
                ws = wb.worksheets[idx] if 0 <= idx < len(wb.worksheets) else wb.worksheets[0]
            out = {}
            for rng in ws.merged_cells.ranges:
                if rng.min_row > max_row:
                    continue
                value = ws.cell(row=rng.min_row, column=rng.min_col).value
                for r in range(rng.min_row, min(rng.max_row, max_row) + 1):
                    for c in range(rng.min_col, rng.max_col + 1):
                        out[(r - 1, c - 1)] = value
            return out
        finally:
            wb.close()
    except Exception:
        return {}


def inspect_headers(path, sheet_name=0):
    """给界面用：返回某个文件的表头识别结果。"""
    p = Path(path)
    sheet = _resolve_sheet(p, sheet_name)
    raw = _read_raw(p, sheet_name=sheet)
    if len(raw) == 0:
        return HeaderInfo(0, 0, 0.0, "空表")
    return detect_header(raw)


def _row_has_merge(merged, row):
    return any(r == row for r, _ in merged)


def _looks_like_header_row(df, i):
    """判断第 i 行是否像"子表头"：短文本为主、不是长句。"""
    if i < 0 or i >= len(df):
        return False
    n, short, _keywords, longest = _row_profile(df, i)
    if n < 2 or longest > 20:
        return False
    return short / float(n) >= 0.8


def _combine_headers(raw, rows, merged=None):
    """把多行表头按列拼成一个名字，空列返回 None。"""
    width = raw.shape[1]
    names = []
    for j in range(width):
        parts = []
        for r in rows:
            if merged and (r, j) in merged:
                v = merged[(r, j)]
            elif r < len(raw):
                v = raw.iat[r, j]
            else:
                v = None
            if _is_blank_value(v):
                continue
            text = _clean_text(v)
            if text and text not in parts:
                parts.append(text)
        names.append("_".join(parts) if parts else None)
    return names


def _resolve_sheet(path, sheet_name, log=None):
    """sheet_name == "auto" 时挑第一个有数据的工作表。"""
    if sheet_name != "auto":
        return sheet_name
    ext = path.suffix.lower()
    if ext == ".csv" or ext not in ENGINE_MAP:
        return 0
    try:
        xl = pd.ExcelFile(str(path), engine=ENGINE_MAP[ext])
        names = list(xl.sheet_names)
    except Exception:
        return 0
    for name in names:
        try:
            probe = pd.read_excel(str(path), sheet_name=name, header=None, nrows=30,
                                  engine=ENGINE_MAP[ext], dtype=object)
        except Exception:
            continue
        if len(probe) and int(probe.notna().sum().sum()) >= 3:
            if log:
                log("  自动选择工作表：" + str(name))
            return name
    return 0


def _drop_repeated_header_rows(body, log):
    """数据区里重复出现的表头行（分页导出常见）要丢掉。"""
    if body.empty:
        return body
    hits = pd.Series(0, index=body.index, dtype="int64")
    for col in body.columns:
        key = _match_key(col)
        hits = hits + body[col].map(
            lambda v: int((not _is_blank_value(v)) and _match_key(v) == key)
        ).astype("int64")
    nonblank = body.notna().sum(axis=1).clip(lower=1)
    mask = (hits >= 2) & (hits.astype("float64") / nonblank >= 0.8)
    count = int(mask.sum())
    if count:
        log("  已丢弃 %d 行重复表头" % count)
        return body[~mask].reset_index(drop=True)
    return body


SUMMARY_KEYWORDS = ("合计", "小计", "总计", "汇总", "累计", "total", "subtotal")


def _drop_summary_rows(body, log, tail=5):
    """表格末尾的合计/小计行不是数据，丢掉。只检查最后几行，避免误删。"""
    if body.empty:
        return body
    keep = [True] * len(body)
    start = max(0, len(body) - tail)
    for pos in range(start, len(body)):
        row = body.iloc[pos]
        vals = [v for v in row if not _is_blank_value(v)]
        if not vals:
            continue
        first = _clean_text(vals[0]).lower()
        if any(first == kw or first.startswith(kw) for kw in SUMMARY_KEYWORDS):
            keep[pos] = False
    dropped = keep.count(False)
    if dropped:
        log("  已丢弃 %d 行合计/汇总行" % dropped)
        return body[keep].reset_index(drop=True)
    return body


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


def _unique_columns(cols):
    """同一表内重名去重：重复的追加 _2、_3 …"""
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
    return final_cols


def merge_files(files, options, log_callback=None, progress_callback=None,
                cancel_event=None):
    """按 options 合并 files，返回合并后的 DataFrame。

    files 的顺序会被原样保留（除非 options.sort_by_name），第一个成功读取的
    文件决定主表头。
    """
    if not isinstance(options, MergeOptions):
        raise TypeError("options 必须是 MergeOptions 实例")

    log = log_callback or _noop
    progress = progress_callback or _noop

    def check_cancel():
        if cancel_event is not None and cancel_event.is_set():
            raise MergeCancelled("用户取消了合并")

    out_path = Path(options.output_path)
    out_resolved = out_path.resolve()

    candidates = []
    for p in files:
        f = Path(p)
        if not f.is_file() or f.name.startswith("~$"):
            continue
        if f.resolve() == out_resolved:
            continue
        candidates.append(f)

    # 关键修正：默认不再按文件名重排，用户排的顺序就是主表头顺序
    if options.sort_by_name:
        candidates.sort(key=lambda p: p.name)

    if not candidates:
        raise ValueError("没有可用的输入文件")

    total = len(candidates)
    progress(0, total)
    check_cancel()

    master_cols = []
    master_keys = []

    def make_key(name):
        if not options.fuzzy_header_match:
            return str(name)   # 精确匹配
        return _match_key(
            name,
            case_sensitive=options.header_case_sensitive,
            strip_punct=options.header_strip_punct,
        )

    force_types = options.force_types or {}

    frames, failed = [], []
    for idx, f in enumerate(candidates, start=1):
        check_cancel()
        sheet = _resolve_sheet(f, options.sheet_name, log)
        try:
            raw = _read_raw(f, sheet_name=sheet)
        except Exception as e:
            failed.append((f.name, repr(e)))
            log("[跳过] " + f.name + " 读取失败：" + str(e))
            progress(idx, total)
            continue

        if len(raw) == 0:
            log("[跳过] " + f.name + " 内容为空")
            progress(idx, total)
            continue

        # ---- 表头行：每个文件可以不一样 ----
        info = None
        if options.header_row is None:
            info = detect_header(raw)
            h = info.row
            log("  自动检测表头：第 %d 行（%s）" % (h + 1, info.reason))
        elif isinstance(options.header_row, dict):
            value = options.header_row.get(str(f))
            if value is None:
                value = options.header_row.get(f.name)
            if value is None:
                value = options.header_row.get(f.stem, 1)
            h = int(value) - 1
        else:
            h = int(options.header_row) - 1
        h = max(0, min(h, max(len(raw) - 1, 0)))

        # ---- 多行表头 ----
        merged = {}
        if options.fill_merged_header_cells:
            merged = _merged_cells_map(f, sheet, min(len(raw), h + 3))

        before = after = 0
        if options.merge_header_rows == "auto":
            if info is not None and info.extra_rows:
                before = info.extra_rows
            elif merged and _row_has_merge(merged, h) and _looks_like_header_row(raw, h + 1):
                # 分组表头（合并单元格）+ 下一行子表头
                after = 1
        else:
            before = max(0, int(options.merge_header_rows) - 1)
        before = min(before, h)
        after = min(after, max(len(raw) - h - 2, 0))
        header_rows = list(range(h - before, h + after + 1))
        header_label = (str(h + 1) if len(header_rows) == 1
                        else "%d-%d" % (h - before + 1, h + after + 1))

        names = _combine_headers(raw, header_rows, merged)
        cols = []
        for j, name in enumerate(names):
            cols.append(name if name else "列" + str(j + 1))

        final_cols = _unique_columns(cols)

        body = raw.iloc[header_rows[-1] + 1:].reset_index(drop=True)
        body.columns = final_cols

        # ============ 表头对齐 ============
        if not master_cols:
            # 第一个文件：作为主表头
            master_cols = list(final_cols)
            master_keys = [make_key(c) for c in final_cols]
            log("[" + str(idx) + "/" + str(total) + "] 已读取 " + f.name +
                "  表头行=" + header_label +
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
                "  表头行=" + header_label +
                "  数据=" + str(body.shape[0]) + " 行 x " + str(body.shape[1]) + " 列")

            if matched_diff:
                log("  表头对齐（智能匹配，视为同一列）：")
                for src, dst in matched_diff:
                    log("    " + repr(src) + "  ->  " + repr(dst))

            if extra_cols:
                log("  注意：出现主表头之外的新列 " + str(extra_cols) +
                    "，已追加到表头末尾")
                master_cols = master_cols + extra_cols
                master_keys = master_keys + [make_key(c) for c in extra_cols]

        # 按主表头顺序重排（缺的列补 NaN）
        body = body.reindex(columns=master_cols)

        if options.drop_empty_rows:
            body = body.dropna(how="all").reset_index(drop=True)

        # 分页导出会在中间重复表头；末尾常有合计行——都不是数据
        if options.drop_repeated_headers:
            body = _drop_repeated_header_rows(body, log)
        if options.drop_summary_rows:
            body = _drop_summary_rows(body, log)

        if options.drop_empty_cols:
            keep = [c for c in body.columns if not body[c].isna().all()]
            body = body[keep]
            master_cols = keep
            master_keys = [make_key(c) for c in keep]

        # 数据格式统一
        if options.unify_formats and len(body) > 0:
            force = {}
            for k, v in force_types.items():
                kk = make_key(k)
                for c in body.columns:
                    if make_key(c) == kk:
                        force[c] = v
            body = normalize_dataframe(body, date_format=options.date_format, force_types=force)

        # 来源列
        if options.add_source_column:
            name = options.source_column_name or "数据来源"
            if name in body.columns:      # 避免与原表同名列冲突
                body = body.drop(columns=[name])
            if options.source_position == "first":
                body.insert(0, name, f.stem)
            else:
                body[name] = f.stem

        frames.append(body)
        progress(idx, total)

    if not frames:
        raise RuntimeError("所有文件都读取失败，无法合并")

    check_cancel()
    log("正在合并数据 ...")
    merged = pd.concat(frames, ignore_index=True, sort=False)

    if options.column_name_map:
        merged = merged.rename(columns=options.column_name_map)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    log("正在写入文件：" + str(out_path))

    # 先写 .part 再原子替换：中途失败或取消不会留下半截文件
    tmp_path = out_path.with_name(out_path.name + ".part")
    try:
        if out_path.suffix.lower() == ".csv":
            merged.to_csv(tmp_path, index=False, encoding="utf-8-sig")
        else:
            with pd.ExcelWriter(tmp_path, engine="openpyxl",
                                datetime_format=options.output_date_format or "yyyy-mm-dd") as writer:
                merged.to_excel(writer, index=False, sheet_name="汇总")
            if options.autofit_columns:
                log("正在调整列宽 ...")
                _autofit_excel_columns(tmp_path, sheet_name="汇总")
        check_cancel()
        os.replace(str(tmp_path), str(out_path))
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass

    log("完成：" + str(len(frames)) + " 个文件 / 共 " +
        str(len(merged)) + " 行 x " + str(merged.shape[1]) + " 列")
    log("最终表头：" + str(list(merged.columns)))
    if failed:
        log("有 " + str(len(failed)) + " 个文件读取失败：" + str([n for n, _ in failed]))

    return merged
