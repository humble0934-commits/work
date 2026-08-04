# work 备份
使用codex制作的日报周报月报登记程序、周报日期修复程序。周报源码丢失，周报日期填写错误，需要使用周报日期修复程序。
# 日报
# Excel日报自动处理助手

Windows 桌面批处理工具，使用 `openpyxl` 修改 XLSX，使用 `tkinter` 提供 GUI。
当前版本包含日报自动处理、Excel数据清理分列和XLS数据清理转换；不包含周报/月报汇总入口，
也不会加载 `report_summary.py`。

## 处理流程

1. 用户填写处理日期。程序动态计算今日和昨日工作表名，例如 `2026/8/1` 对应今日 `8.1`、昨日 `7.31`。
2. 以销售文件名创建销售工作表，例如 `1.xlsx` 创建工作表 `1`，完整复制单元格、样式、行高、列宽、合并区域、冻结窗格及隐藏属性。
3. 严格使用 `workbook.copy_worksheet(昨日工作表)` 生成今日副本；先复制、再删除已有今日表、最后重命名，并补充复制 openpyxl 默认遗漏的条件格式和数据验证。禁止逐个复制模板单元格。
4. 使用 `openpyxl.formula.Tokenizer` 修改公式中的工作表 RANGE 标记，不拆解 VLOOKUP 字符串；兼容带单引号、不带单引号以及 IFERROR 外层。
5. 使用 `openpyxl.formula.Tokenizer` 读取 VLOOKUP 的结构化参数，从引用的销售数据区域直接取得精确匹配结果作为排序键，不调用公式计算引擎。
6. 昨日销售额列按处理日期动态引用前一天工作表，例如处理 `7.31` 时引用 `'7.30'`；
   如果前一天工作表不存在，则保留最近历史模板的格式并将昨日销售额数据行写为 `0`。
7. 排序时移动完整数据行，公式保持为公式，相对引用随新位置调整。
8. 昨日和今日工作表中的公式均保持为公式，不进行公式数值化。
9. 以原文件名保存到用户选择的输出目录。
10. 排序后从 `B2` 开始按 `1, 2, 3...` 填充至汇总行上一行；所有除法公式保留原表达式并统一为 `=IFERROR((...),"0")`。

## 重要限制

程序不依赖 Microsoft Excel、LibreOffice 或其他外部计算软件。今日销售额排序仅支持精确匹配的 VLOOKUP（第四参数为 `FALSE`），查找值必须是当前工作表的单元格引用，数据区域必须引用工作簿内的销售数据工作表，返回列必须是直接数值。公式本身不会被替换成数字。

`openpyxl` 不计算任意 Excel 公式。本程序不会把昨日或今日工作表公式转换成数值。

数据区若存在跨行合并单元格，程序会拒绝排序，以免破坏模板。VBA、ActiveX、外部链接缓存、形状或某些第三方扩展对象不属于本工具承诺的保留范围。请始终保留原始文件备份。

## 源码运行

```bat
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
```

## 打包 EXE

双击 `build_exe.bat`，或运行：

```bat
pyinstaller --onefile --windowed --hidden-import=openpyxl --hidden-import=openpyxl.styles --hidden-import=openpyxl.utils --name "Excel日报自动处理助手" main.py
```

输出文件：

```text
dist\Excel日报自动处理助手.exe
```

## 使用步骤

1. 填写处理日期，例如 `2026/8/1`。
2. 点击“选择多个品牌日报 XLSX”。
3. 选择销售数据文件，例如 `1.xlsx`。
4. 选择独立的输出目录（不能与源文件目录相同）。
5. 点击“开始处理”。
6. 完成后用 Excel 打开输出文件，检查今日、昨日、销售数据工作表和汇总结果。


# 月报
# Excel月报自动汇总助手

Windows 10/11 月报汇总工具，使用 Python 3.12、tkinter 和 openpyxl。

当前版本只保留月报汇总功能，不包含周报标签页、周报文件选择、周报日期范围、
周报工作表复制、周报 SUMIF、周报排序或周报汇总公式代码。

## 功能

- 选择多个品牌日报 XLSX 文件。
- 选择多个品牌月报 XLSX 文件。
- 输入目标月份 `YYYY/MM`。
- 自动按品牌模糊匹配日报和月报。
- 在 Windows 桌面生成处理后的月报，并保持原月报文件名。
- 保留现有月报字段识别、公式生成、排序、条件格式、颜色、边框、合并单元格、
  行高和列宽处理逻辑。

## 源码运行

```bat
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
```

## 打包

运行：

```bat
build_exe.bat
```

输出：

```text
dist\Excel月报自动汇总助手.exe
```
