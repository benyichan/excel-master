---
name: excel-master
title: excel-master
description: 从 DataFrame 生成/美化摩根士丹利标准格式的 Excel 报表。纯 openpyxl，零 xlwings，一次保存线性流程。支持 beautify（保留公式只改格式）。
---

# excel-master

## 强制约束：每次与用户交互时必须问的2个问题

**适用范围：** 用户要求「美化表格」「出个Excel」「生成报表」「导出Excel」「保存为xlsx」等涉及本 skill 的对话场景。

**自动化调用豁免：** 上层脚本（千川投流、定时任务等）程序化调用 `make_excel`/`beautify` 时不走此约束，直接传参或走默认值。

### 执行流程

与用户确认需求后，先问这2个问题再调用核心脚本：

**问题1：表头在哪几行？**
> 表头的开始和结束行是？比如"第1行"或"第1行到第2行"

- 用户回答 → 传给 `interactive_make_excel.py --freeze-rows N`
- 用户说"跳过"/没提供 → 不传 `--freeze-rows`，核心脚本自动推断

**问题2：用什么配色方案？**
> 配色选哪个？水蓝（默认） / 深海蓝 / 墨玉绿 / 陨石灰蓝 / 勃艮第红 / 珊瑚橙

- 用户选择 → 传给 `interactive_make_excel.py --theme 主题名`
- 用户说"默认"/没选 → 不传 `--theme`，默认水蓝

### 调用方式

```bash
# 用户回答了2个问题 → 传参
python scripts/interactive_make_excel.py beautify 报表.xlsx --freeze-rows 3 --theme coral

# 用户跳过 → 不传参，走自动推断+水蓝
python scripts/interactive_make_excel.py beautify 报表.xlsx

# 从 CSV 生成
python scripts/interactive_make_excel.py make 数据.csv 输出.xlsx --theme deep-navy
```

核心脚本 `make_excel.py` **不直接**与用户交互，只接收参数。wrapper 是对话场景的唯一入口。

## 什么时候用

用户要求"出个Excel"、"导出报表"、"生成表格"、"保存为xlsx"、"美化/格式化这个Excel"时加载本 skill。

## 前置依赖

```
pip install pandas openpyxl>=3.0
```

xlwings **不需要**——本 skill 不做外部进程调用（xlwings 存盘会吞 openpyxl 的格式），列宽用字符估算法。

## 快速使用

```python
from make_excel import make_excel, beautify
import pandas as pd

# 从零生成（DataFrame → 摩根系 Excel）
df = pd.read_csv('data.csv')
make_excel(df, '报表.xlsx')

# 多 sheet
make_excel([('汇总', df_summary), ('明细', df_detail)], '报表.xlsx')

# 切换色系主题（默认: default 水蓝）
make_excel(df, '报表-深海蓝.xlsx', theme='deep-navy')
make_excel(df, '报表-珊瑚橙.xlsx', theme='coral')

# 自定义数字格式（覆盖默认的 #,##0.00 等）
make_excel(df, '报表.xlsx', fmt_override={'money': '#,##0.0', 'pct': '0.0%'})

# 冻结表头行数控制（默认自动推断）
make_excel(df, '报表.xlsx', freeze_rows=2)      # 冻前2行
make_excel(df, '报表.xlsx', freeze_rows=0)      # 不冻结

# beautify — 只改格式不改数据（保留公式）
beautify('已有报表.xlsx')                    # 原地美化，自动备份
beautify('已有报表.xlsx', '美化后.xlsx')      # 另存美化（不备份）

# beautify + 手动指定列类型
beautify('订单表.xlsx', col_types={'订单号': 'text', '金额': 'money'})

# beautify + 切换主题
beautify('已有报表.xlsx', '报表-墨玉绿.xlsx', theme='jade')

# beautify + 关闭自动备份
beautify('临时表.xlsx', backup=False)

# beautify + 自定义数字格式
beautify('已有报表.xlsx', '美化后.xlsx', fmt_override={'money': '#,##0'})

# beautify + 冻结表头
beautify('已有报表.xlsx', freeze_rows=3)      # 冻前3行
```

## 命令行

```bash
python scripts/make_excel.py 数据.csv 输出.xlsx
python scripts/make_excel.py --beautify 已有报表.xlsx 美化后.xlsx
python scripts/make_excel.py --beautify 已有报表.xlsx 美化后.xlsx --theme coral
python scripts/make_excel.py --beautify 已有报表.xlsx 美化后.xlsx --freeze-rows 3
python scripts/make_excel.py 数据.csv 深海蓝报表.xlsx --theme deep-navy
```

## 色系主题

通过 `theme` 参数切换配色，共 6 套：

| 主题名 | 中文 | 表头 | 合计行 | 特点 |
|--------|------|------|--------|------|
| `default` | 水蓝 | #4472C4 白字 | #D9E2F3 | 经典摩根系默认 |
| `deep-navy` | 深海蓝 | #1F4E79 白字 | #D6E4F0 | 更深沉专业 |
| `jade` | 墨玉绿 | #375623 白字 | #E2EFDA | 清爽自然 |
| `slate` | 陨石灰蓝 | #404040 白字 | #D9D9D9 | 现代极简 |
| `burgundy` | 勃艮第红 | #843C0C 白字 | #FCE4D6 | 暖色调高级感 |
| `coral` | 珊瑚橙 | #D84B4B 白字 | #FDE8E8 | 活泼明亮 |

用法：

```python
make_excel(df, '输出.xlsx', theme='deep-navy')
beautify('输入.xlsx', '输出.xlsx', theme='coral')
```

## 参考文件

- `references/type-inference-rules.md` — 列类型推断关键词规则和优先级
- `references/implementation-checklist.md` — 交付前逐项验证清单

> Excel 照相机截图流程和指定列打码见同分类下的 `excel-screenshot-blur` skill。

## 摩根系标准 — 9 大原则（来自《为什么精英都是Excel控》）

### 已自动化的原则

| # | 原则 | 实现 |
|---|------|------|
| 1 | 行高 18 | 所有数据行 height=18 |
| 2 | 英文字体 Arial，字号统一 11 | 所有单元格 Font(name='Arial', size=11)。文本列默认黑，数字列叠加颜色 |
| 3 | 数字千分撇区隔 | money → #,##0.00，number → #,##0，pct → 0.00% |
| 7 | 框线上下粗中间虚线无竖线 | TOP_BORDER / MID_BORDER / BOTTOM_BORDER，left/right=None |
| 8 | 文字左对齐，数字右对齐 + 表头全部右对齐 | LEFT_ALIGN / RIGHT_ALIGN，表头统一右对齐 |
| 9 | 不从 A1 开始 | **标题在 B2，Row 1 留空**，A 列留空；make_excel 从 B2 开始写；beautify 保留原始数据位置，样式从 B 列开始。**注意**：手工构建的自定义布局（非 make_excel）也需遵守——Row 1 必须空白，标题行放 B2 |
| — | 隐藏网格线 | ws.sheet_view.showGridLines = False |
| — | 冻结表头 | ws.freeze_panes = freeze_cell，4 种策略：参数指定 / 检测 header_row / B2 起始 / 默认首行 |
| — | 右侧空白列（宽 3） | 数据区域最后一列右侧加一栏，宽度 3 |
| — | 数字颜色：手动输入=蓝，公式=黑 | make_excel（无公式）数字全蓝；beautify 公式检测 → 黑，值 → 蓝 |
| — | 水蓝表头(#4472C4)白字粗体 | HDR_FILL + HDR_FONT + HDR_ALIGN |

### 需要人工控制的原则（不自动化，但要知道）

这些依赖业务上下文和数据层级关系，自动化会做出错误判断。做表时提醒用户注意。

| # | 原则 | 为什么不自动做 |
|---|------|---------------|
| 4 | 项目下细项要缩排（空白列列宽1） | 需要知道数据层级关系（哪些行是父级、哪些是子级），无法自动推断 |
| 5 | 单位自成一栏（元/个/%单独一列） | 单位信息通常混在列名或值中，抽离为独立列会改变数据结构 |
| 6 | 同一层级栏宽统一 | 需要知道哪几列属于同一层级。当前用字符估算法单独调宽 |
| — | 空单元格填 N/A / N/M | 内容层面的决策，不知道哪些是"不需要填"哪些是"漏填" |
| — | 隐藏行/列用群组而非隐藏 | 群组需要知道哪些行/列属于同一逻辑组，无法自动判断 |
| — | 绿色字体（跨工作表引用） | 需要检测公式中的跨表引用，格式是 `=Sheet2!A1`——可以加但当前未实现 |
| — | 仅3种颜色、淡色背景 | 设计原则，不是代码能控制的 |
| — | 边输入边格式 | 工作习惯，不是输出结果 |

## 设计原则

1. **零 xlwings** — 不调 Excel 外部进程。xlwings 保存时会丢掉 openpyxl 的边框/颜色/数字格式，之前为了 autofit 做了三步修补循环（写→保存→xlwings打开→读列宽→写回→再保存），这是架构级别的脆弱性。改为字符估算法后，一次保存完事。
2. **一次保存** — wb.save() 只调一次，没有三步修补循环。格式始终完整。
3. **纯函数式** — 同样的输入永远产生同样的输出。无外部状态，无随机因素。
4. **单文件** — 一个 script 解决所有，不搞模块拆分配置继承。要加功能也用函数追加在文件末尾。
5. **基座 skill 职责分离** — excel-master 是「基座」skill，提供可配置的选项（如 `theme` 参数），高层脚本（如千川投流）按需选择。**不动基座只改上层脚本**——基座的配色、行为应作为参数暴露，应用层直接传参调用。保持基座与应用层的职责分离。

## Pitfalls

### xlwings 吞格式

xlwings 保存时会丢弃 openpyxl 设置的边框/颜色/填充。本 skill 已废弃 xlwings 方案，改用纯 openpyxl 字符估算法。

### 公式检测双重保险（2026-06-09 修复）

`_beautify_worksheet()` 中公式检测原为 `cell.value.startswith('=')` 单重判断，升级为：
```
is_formula = (isinstance(val, str) and val.startswith('=')) or cell.data_type == 'f'
```
`cell.data_type == 'f'` 是 openpyxl 标记公式的官方字段，作为第二保险。make_excel 模式无需此检测（pandas 不产生公式单元格）。

### 右侧空白列 end_col 变量未定义（2026-06-08/09 修复）

`_apply_styles()` 和 `_beautify_worksheet()` 两处曾使用未定义变量 `end_col` 计算右侧空白列，导致 `NameError`。正确的值应从 `start_col + ncols - 1` 推算。已在两处统一修复。后续修改时注意不要重新引入。

### A列格式遗漏（2026-06-13 修复）

`_beautify_worksheet()` 中 `_detect_data_range()` 硬编码 `start_col=2`，导致 beautify 时 B~F 列均应用了完整摩根系格式，但 A 列（如果有"部门/名称/合计"等数据）只有宽度估算，缺少表头蓝底白字、数据 Arial 11、框线和对齐。修复：在数据格式化块后新增 `# 4.5 A列格式` 段，检测 A1 有值时补齐全部格式（表头→蓝底白字右对齐、数据→Arial 11 左对齐、合计行→浅蓝底粗体），边框循环也扩展为 `border_start=1 if A1有值 else start_col`。`_apply_styles()`（make_excel 模式）不受影响——A 列始终留空。

### 类型推断关键词优先级：pct > date > text > money

列名含有"毛利率""完成率"等以"率"结尾的词时，如果 money 关键词先匹配，会被错误归类为 money。当前用 `率$` 正则限定词尾"率"字，并在 money 中用 `毛利(?!率)` 排除误匹配。详见 `references/type-inference-rules.md`。

### 数值/百分比列长公式限宽15

数值/百分比列（money/number/pct）列宽估算时，如果检测到该列有超过 15 字符的公式（如 `=Sheet2!B2/Sheet1!D3`），列宽固定为 15 而非按公式字符串字符宽度估算。原因：公式是中间计算过程，不是展示给终端用户看的内容，不需要完整显示。

beautify 和 make_excel 两套入口都实现了此逻辑。

### 数字颜色规则

摩根系规范：手动输入的数字用蓝色（可调的价值动因），公式计算结果用黑色。

- **make_excel 模式**：数据来自 pandas，都是值（无公式），所以数字列全部蓝色。这个假定是"从零输入的源数据都是手动填的"——如果源数据有公式计算过的值，会误标为蓝色。要准确区分，用 beautify 从原表加载。
- **beautify 模式**：双重检测公式——`cell.value` 是否以 `=` 开头 **或** `cell.data_type == 'f'`。检测到公式→黑色，否则→蓝色。
- **关键约束**：beautify 只改单元格的 font / fill / alignment / number_format / border / row_height / column_width，**不改 `cell.value`**。公式原样残留，不转值。

### theme 无效时静默回退（P2）

`_build_theme_styles()` 中 `THEMES.get(theme_name, THEMES['default'])` 在 theme 不存在时静默回退到 default，没有任何日志或警告。调用方以为用了指定主题（如 coral），实际是 default。

**修复**：在 `make_excel()` 和 `beautify()` 入口中校验 theme：
```python
if theme not in THEMES:
    raise ValueError(f"未知主题 '{theme}'，可选: {sorted(THEMES.keys())}")
```

### _detect_data_range() 的 A 列依赖假设

`_detect_data_range()` 中 `row[0].row` 获取行号、`any(cell.value is not None for cell in row)` 检查整行是否有数据——这个逻辑在标准 B2 起始布局下是正确的（`any()` 扫描所有列）。但如果工作表的前 N 行有零散的非表头数据（如备注行、空行混杂注释），函数可能误判 header_row。**beautify 时如果遇到 A 列全空且其他列有数据的表格，使用前确认首行确实是表头。**

### `_beautify_worksheet` 中 `is_summary` 前向引用（2026-06-13 修复）

4.5 段（A列格式补全）引用了 `is_summary`，但该变量定义在 5 段（合计行特殊格式）。当 A 列有数据时，4.5 段先于 5 段执行 → `UnboundLocalError`。

**修复**：将 `is_summary` 的检测逻辑（A列末行含合计关键词？B列末行兜底？）提取到 4.5 段之前，4.5 和 5 段共用。`_apply_styles()`（make_excel 模式）不受影响——其 `is_summary` 定义在使用之后，无此问题。
