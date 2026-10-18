# -*- coding: utf-8 -*-
"""合并引擎测试：顺序、表头对齐、类型、取消、原子写入。

不依赖界面，直接 `python tests/test_logic.py` 运行。
"""
import sys
import tempfile
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import pandas as pd

from excel_merge.logic import MergeCancelled, MergeOptions, merge_files

WORK = Path(tempfile.mkdtemp(prefix="merge-test-"))


def make(name, headers, rows):
    path = WORK / name
    pd.DataFrame(rows, columns=headers).to_excel(path, index=False)
    return str(path)


def run(files, **kw):
    opts = MergeOptions(output_path=str(WORK / kw.pop("out", "out.xlsx")), **kw)
    logs = []
    df = merge_files(files, opts, log_callback=logs.append)
    return df, logs


def test_order_defines_master_header():
    """用户列表顺序决定主表头，即使文件名排序相反。"""
    f1 = make("z_first.xlsx", ["姓名", "身份证号码", "金额(元)"],
              [["张三", "110101199001011234", "1,200.50"]])
    f2 = make("a_second.xlsx", [" 姓名 ", "证件号码", "金额"],
              [["李四", "310101199002022345", "800"]])
    f3 = make("m_third.xlsx", ["姓名", "备注"], [["王五", "老客户"]])

    df, _ = run([f1, f2, f3], out="ordered.xlsx")
    assert list(df.columns)[:3] == ["姓名", "身份证号码", "金额(元)"], list(df.columns)
    assert "备注" in df.columns
    assert df["身份证号码"].notna().sum() == 2
    assert list(df["数据来源"]) == ["z_first", "a_second", "m_third"]
    print("  order + synonym + extra column   OK")


def test_sort_by_name():
    f1 = make("z_first2.xlsx", ["姓名", "身份证号码"], [["张三", "1"]])
    f2 = make("a_second2.xlsx", ["姓名", "证件号码"], [["李四", "2"]])
    df, _ = run([f1, f2], out="sorted.xlsx", sort_by_name=True)
    assert "证件号码" in df.columns and "身份证号码" not in df.columns, list(df.columns)
    print("  sort_by_name flips master header  OK")


def test_cancel_leaves_no_output():
    f = make("cancel.xlsx", ["姓名"], [["张三"]])
    ev = threading.Event()
    ev.set()
    out = WORK / "cancelled.xlsx"
    try:
        merge_files([f], MergeOptions(output_path=str(out)), cancel_event=ev)
        raise AssertionError("cancel was not honoured")
    except MergeCancelled:
        pass
    assert not out.exists()
    assert not (WORK / "cancelled.xlsx.part").exists()
    print("  cancel + no leftovers             OK")


def test_csv_and_source_first():
    f = make("csv.xlsx", ["姓名", "金额"], [["张三", "100"]])
    df, _ = run([f], out="out.csv", source_position="first")
    assert list(df.columns)[0] == "数据来源", list(df.columns)
    assert (WORK / "out.csv").exists()
    assert not (WORK / "out.csv.part").exists()
    print("  csv output + source first         OK")


def test_force_types():
    f = make("types.xlsx", ["金额(元)", "身份证号码"],
             [["1,200.50", "110101199001011234"]])
    df, _ = run([f], out="types_out.xlsx",
                force_types={"金额(元)": "number", "身份证号码": "text"})
    assert str(df["金额(元)"].dtype) in ("Float64", "float64"), df["金额(元)"].dtype
    assert df["金额(元)"].iloc[0] == 1200.5, df["金额(元)"].iloc[0]
    print("  force column types                OK")


def test_auto_header_detection():
    path = WORK / "with_title.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        pd.DataFrame([["2026 年销售报表", None, None],
                      [None, None, None],
                      ["姓名", "手机号", "金额"],
                      ["赵六", "13800000000", "300"]]).to_excel(w, index=False, header=False)
    df, logs = run([str(path)], out="detected.xlsx", header_row=None)
    assert list(df.columns)[:3] == ["姓名", "手机号", "金额"], list(df.columns)
    assert any("自动检测表头" in m for m in logs)
    print("  auto header detection             OK")


def test_drop_empty_rows_and_cols():
    f = make("blanks.xlsx", ["姓名", "空列", "金额"],
             [["钱七", None, "100"], [None, None, None], ["孙八", None, "200"]])
    df, _ = run([f], out="blanks_out.xlsx")
    assert len(df) == 2 and "空列" in df.columns
    df2, _ = run([f], out="blanks_out2.xlsx", drop_empty_cols=True)
    assert "空列" not in df2.columns, list(df2.columns)
    print("  drop empty rows / columns         OK")


def test_source_column_name_clash():
    f = make("clash.xlsx", ["姓名", "数据来源"], [["周九", "原值"]])
    df, _ = run([f], out="clash_out.xlsx")
    assert list(df["数据来源"]) == ["clash"], df["数据来源"].tolist()
    print("  source column name clash          OK")


TESTS = [
    test_order_defines_master_header,
    test_sort_by_name,
    test_cancel_leaves_no_output,
    test_csv_and_source_first,
    test_force_types,
    test_auto_header_detection,
    test_drop_empty_rows_and_cols,
    test_source_column_name_clash,
]


def main():
    print("logic tests (%d)" % len(TESTS))
    for fn in TESTS:
        fn()
    print("logic tests passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
