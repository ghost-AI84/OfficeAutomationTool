# OfficeAutomationTool

办公小工具合集。**一个工具 = 一个入口文件 + 一个代码包 + 一个 exe。**

## 我该看哪里？

| 你想做什么 | 打开 |
|---|---|
| **直接使用工具** | `releases\` 里的 exe，双击即可（Win7 及以上 / 32 位） |
| **改代码、加功能** | `src\` |
| **加一个新工具** | 见下面「加第二个工具」，照抄三步 |
| **看某个工具怎么启动** | `apps\`，一个工具一个文件 |

## 现有的工具

| 工具 | 入口 | 成品 |
|---|---|---|
| Excel 批量合并 | `apps\merge_excel.py` | `releases\ExcelMergeTool-2.0.0-win7-x86.exe` |

功能：把多个 `.xlsx / .xlsm / .xls / .xlsb / .ods / .csv` 合成一张表；
表头以列表第一个文件为准，空格/全半角/大小写/同义词自动对齐；
可加「数据来源」列；数字日期自动识别，也可强制指定列类型。

### 表头不在同一行怎么办

默认就是**每个文件单独检测**表头行，以下情况会自动处理：

- 前面有标题行、空行、说明文字 → 自动跳过
- 两级表头，分组名是合并单元格 → 自动拼成 `上半年_1月` 这样的列名
- 分页导出时数据中间又出现一次表头 → 自动丢弃
- 末尾的「合计 / 小计 / 总计 / 汇总」行 → 自动丢弃
- 第一个工作表是空的说明页 → 工作表填 `auto`，自动挑有数据的表
- 一个工作簿里有 `1月 / 2月 / 3月` 多个子表 → 工作表填 `all`，**按子表名分组**：
  各文件的「附件1」合并成一张、「附件2」合并成另一张，输出成多个工作表，
  不会把内容不同的子表混在一起；空子表和说明页自动跳过

自动判断不可能 100% 正确。点「检测并调整…」能看到**每个文件被识别成第几行**，
以及判断依据，然后逐行改正；改正后按「确定」，这次合并就按你指定的行来。

> 注意：装 exe 的文件夹如果在 `D:\ExcelMergeTool` 里，exe 会继承 DSH 工作区的
> 低完整性标签，双击会报 "cannot create temporary directory"。把 exe 复制到
> 工作区外面（例如桌面）再运行即可。

## 目录结构

```
OfficeAutomationTool\
├── releases\                  成品 exe，双击即用（不进 git）
│   └── ExcelMergeTool-2.0.0-win7-x86.exe
├── apps\                      每个工具一个启动入口
│   └── merge_excel.py
├── src\                       每个工具一个代码包
│   └── excel_merge\
│       ├── logic.py           合并引擎（无界面，可单独调用）
│       ├── ui.py              图形界面
│       └── config.py          设置存到 %APPDATA%\ExcelMergeTool\
├── tests\                     测试（逻辑 / 界面 / 端到端）
├── scripts\                   本地脚本：build.ps1、test.ps1
├── dist\                      本地构建的临时输出（不进 git）
├── .github\workflows\         云端自动构建
├── requirements.txt           运行依赖
├── requirements-dev.txt       打包依赖（pyinstaller）
└── README.md
```

## 加第二个工具

1. **写代码**：建 `src\新工具\`，里面放 `__init__.py` 和你的模块。
2. **加入口**：复制 `apps\merge_excel.py` 为 `apps\你的工具.py`，
   把最后一行改成 `from 你的工具.ui import main`。
   ⚠️ 入口文件名**不要**和 `src\` 里的包同名（比如入口叫 `foo.py`、包也叫
   `foo`），PyInstaller 会把它当成模块而不是包，打包时报
   `missing module named 'foo.xxx'`。
3. **加测试**：在 `tests\` 放 `test_新工具.py`；
   在 `.github\workflows\build.yml` 里复制一段打包步骤，把
   `--name` 和入口路径换掉。

`releases\` 里按工具名放成品 exe，一眼能看出哪个是哪个。

## 从源码运行

需要 Python 3.8 及以上：

```bat
pip install -r requirements.txt
python apps\merge_excel.py
```

## 构建

本地一条命令（默认先跑测试）：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/build.ps1
```

产物在 `dist\ExcelMergeTool.exe`。加 `-SkipTests` 可跳过测试。

云端：push 到 `main` 会自动构建，去 Actions 页面下载，或者：

```bat
gh run download --name ExcelMergeTool-win7 --dir dist\win7
```

## 测试

```powershell
powershell -ExecutionPolicy Bypass -File scripts/test.ps1
```

三个脚本分别覆盖合并引擎、界面布局与校验、界面端到端流程。
没有图形环境的机器上，界面测试会自动跳过而不是失败。
