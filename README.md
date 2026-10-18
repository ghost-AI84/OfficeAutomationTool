# OfficeAutomationTool

办公自动化小工具集合。目前只有一个工具：**Excel 批量合并**（图形界面）。

## 现在能做什么

- 把多个 `.xlsx / .xlsm / .xls / .xlsb / .ods / .csv` 合成一张表
- 表头以**列表里第一个文件**为准，列顺序不变；后续文件的表头只差空格、
  全半角、大小写、不可见字符时自动对齐，还认得「身份证 / 身份证号码」这类同义词
- 可选「数据来源」列，记住每行来自哪个文件
- 自动识别数字、日期、布尔；可强制指定列类型
- 输出 xlsx 后自动调整列宽

## 直接用（不用装 Python）

`releases/` 目录里的 exe 双击即可，Win7 及以上、32 位系统都能跑。

## 从源码运行

需要 Python 3.8 及以上：

```bat
pip install -r requirements.txt
python merge_gui.py
```

## 目录结构

```
OfficeAutomationTool/
├── merge_gui.py            启动入口（PyInstaller 也用它）
├── excel_merge/            程序本体
│   ├── logic.py            合并引擎，不含界面，可单独调用
│   ├── ui.py               图形界面
│   └── config.py           设置保存到 %APPDATA%\ExcelMergeTool\
├── tests/                  自动化测试
├── scripts/                本地常用脚本
├── releases/               下载下来的成品 exe（不进 git）
├── dist/                   本地构建输出（不进 git）
└── .github/workflows/      云端自动构建
```

## 构建

本地一条命令（会先跑测试）：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/build.ps1
```

产物在 `dist\OfficeAutomationTool.exe`。

云端：push 到 `main` 就会自动构建，去 Actions 页面下载，或者：

```bat
gh run download --name OfficeAutomationTool-win7 --dir dist\win7
```

## 测试

```powershell
powershell -ExecutionPolicy Bypass -File scripts/test.ps1
```

三个脚本分别覆盖合并引擎、界面布局与校验、界面端到端流程。
没有图形环境的机器上，界面测试会自动跳过而不是失败。

## 以后加新工具

两种组织方式，按需要选：

**A. 一个工具一个仓库**（推荐起步）
新工具新建仓库，把 `scripts/`、`.github/workflows/build.yml`、
`.gitattributes` 连同 README 一起复制过去。好处是每个工具独立发版、
互不影响，CI 也最简单。

**B. 多个工具共用一个仓库**（工具间要共享代码时）
在根目录建 `tools/<工具名>/`，每个工具自带入口和 `tests/`，
`build.yml` 用 matrix 或分步分别打包。共享逻辑放 `common/`。
代价是 CI 和工作流都要跟着改，工具少的时候不划算。
