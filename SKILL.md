---
name: excel-master
title: excel-master
description: "从 DataFrame 生成/美化摩根士丹利标准格式的 Excel 报表。纯 openpyxl，零 xlwings，一次保存线性流程。支持 beautify（保留公式只改格式）和 make_excel（从零生成）两种模式。12 套色系主题（水蓝/深海蓝/墨玉绿/陨石灰蓝/勃艮第红/珊瑚橙/樱花粉/暖阳橙/薰衣草紫/抹茶绿/蜜桃/雾蓝紫），条件格式自适应，类型推断（pct 关键词正则 `率$|占比|...` / date / text / money）带三段式 fallback 恢复表，11 条反例黑名单。当用户说「出个Excel」「美化表格」「导出报表」「生成xlsx」「格式化表格」「把数据变成Excel」「摩根士丹利格式」「水蓝色表头」「做表」「整理Excel」「美化一下」「表格太难看了帮我调调」时使用。"
---

# excel-master

## 强制约束：每次与用户交互时必须问的2个问题

**适用范围：** 用户要求「美化表格」「出个Excel」「生成报表」「导出Excel」「保存为xlsx」等涉及本 skill 的对话场景。

**自动化调用豁免：** 上层脚本（千川投流、定时任务等）程序化调用 `make_excel`/`beautify` 时不走此约束，直接传参或走默认值。

### 执行流程（Step-by-Step）

**Step 1：读取数据 → 了解表结构**
- **输入**：用户提供的文件路径（xlsx/csv）或 DataFrame 描述
- **处理**：加载文件前 5 行，确定表范围、列名、数据类型
- **输出**：文件路径 + 列名列表 + 预估表头行位置

**Step 2：确认表头位置 → 确定 freeze_rows 参数**
- **输入**：Step 1 的输出 + 用户回答
- **处理**：问用户"表头在哪几行？"——如"第1行"或"第1行到第2行"
- **输出**：`freeze_rows=N`（用户指定）或自动推断（用户跳过）

**Step 3：确认配色 → 确定 theme 参数**
- **输入**：Step 2 的输出 + 用户选择
- **处理**：展示 12 套主题分类让用户选：
  - 经典商务：水蓝(default) / 深海蓝 / 墨玉绿 / 陨石灰蓝
  - 暖色高级：勃艮第红 / 珊瑚橙
  - 青春活力：樱花粉 / 暖阳橙 / 薰衣草紫 / 抹茶绿 / 蜜桃 / 雾蓝紫
- **输出**：`theme=主题名`（用户指定）或 `default`（用户跳过）

> **🔴 CHECKPOINT · 🛑 STOP** — Step 2 和 Step 3 都得到明确回答了？用户说"跳过"不算"未确认"。确认后再进入 Step 4。

**Step 4：PRE-FLIGHT 检查（beautify 模式）**
- **输入**：文件路径
- **处理**：验证文件存在/合法、第 1 个非空行是表头、A 列是否有数据、列名有无"率"误判风险
- **输出**：检查结果（通过/阻断原因）

**Step 5：调用核心脚本 → 生成 xlsx**
- **输入**：`interactive_make_excel.py make/beautify 文件路径 [--freeze-rows N] [--theme T]`
- **处理**：执行脚本，纯 openpyxl 一次保存
- **输出**：生成的 xlsx 文件路径

**Step 6：TYPE CHECK → 验证类型推断**
- **输入**：Step 5 生成的 xlsx 文件
- **处理**：打开文件，逐列检查 `number_format` 是否符合预期，尤其是：
  - 小数值列（万元单位）未被误判为 pct
  - 关键词未命中的列（列名含糊）推断是否正确
- **输出**：类型检查结果。有误则用 `col_types` 覆盖后重跑

**Step 7：DELIVERY GATE → 交付前全量验证**
- **输入**：最终 xlsx 文件
- **处理**：按 DELIVERY GATE 检查清单逐项验证
- **输出**：交付文件给用户

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

通过 `theme` 参数切换配色，共 12 套，分三大类：

### 经典商务系（冷色/中性）

| 主题名 | 中文 | 表头 | 合计行 | 数据字体 | 特点 |
|--------|------|------|--------|---------|------|
| `default` | 水蓝 | #4472C4 白字 | #D9E2F3 | 深蓝 | 经典摩根系默认 |
| `deep-navy` | 深海蓝 | #1F4E79 白字 | #D6E4F0 | 深蓝 | 更沉稳专业 |
| `jade` | 墨玉绿 | #375623 白字 | #E2EFDA | 深茶绿 | 清爽自然 |
| `slate` | 陨石灰蓝 | #404040 白字 | #D9D9D9 | 深灰 | 现代极简 |

### 暖色高级感系

| 主题名 | 中文 | 表头 | 合计行 | 数据字体 | 特点 |
|--------|------|------|--------|---------|------|
| `burgundy` | 勃艮第红 | #843C0C 白字 | #FCE4D6 | 深褐 | 酒红高级感 |
| `coral` | 珊瑚橙 | #D84B4B 白字 | #FDE8E8 | 深红褐 | 活泼明亮 |

### 青春活力系（新增）

| 主题名 | 中文 | 表头 | 合计行 | 数据字体 | 特点 |
|--------|------|------|--------|---------|------|
| `sakura` | 樱花粉 | #D94F70 白字 | #FDE8EE | 深浆果色 | 温柔不失活力 |
| `warm-sun` | 暖阳橙 | #E8843A 白字 | #FEF0E6 | 深琥珀 | 阳光有能量 |
| `lavender` | 薰衣草紫 | #8B7EC8 白字 | #EDEAF8 | 深紫 | 知性温婉 |
| `matcha` | 抹茶绿 | #7BA23F 白字 | #F0F5E6 | 深茶绿 | 清新自然 |
| `peach` | 蜜桃 | #E8897C 白字 | #FEF0ED | 深桃红褐 | 甜美不腻 |
| `misty` | 雾蓝紫 | #6B7FB5 白字 | #EDF0F8 | 深蓝紫 | 现代感冷色调 |

用法：

```python
make_excel(df, '输出.xlsx', theme='sakura')
beautify('输入.xlsx', '输出.xlsx', theme='warm-sun')
```

> **🔴 CHECKPOINT · 🛑 TYPE CHECK** — beautify 类型推断完成后，检查输出文件：
> 1. 所有列的 `number_format` 是否符合预期（尤其是含小数值的列）
> 2. 如果 `_infer_col_type` 返回 `'unknown'`，必须用 `col_types` 强制指定
> 3. 文本列（订单号/编码等）是否左对齐，而非被当作数字右对齐
> 4. 关键词没命中的列（列名含糊如"字段A""数据1"）手动检查一次
>
> 类型推断永远可以靠 `col_types={'列名': '类型'}` 覆盖。不要接受推断结果不加验证。

### 条件格式自适应

从 v2.1 开始，`beautify` 会自动检测已有文件中的 **colorScale 色阶条件格式**，并将其最大色（高值色）替换为当前主题的表头色，使条件格式与整体配色统一。

- 支持 2 色/3 色色阶（替换最后一个颜色）
- 数据条(dataBar)、图标集(iconSet)等其他条件格式不受影响
- make_excel 模式新建文件无历史条件格式，不执行此步骤

## 参考文件

- `references/type-inference-rules.md` — 列类型推断关键词规则和优先级
- `references/implementation-checklist.md` — 交付前逐项验证清单
- `references/dual-header-format.py` — 双表头/多数据块布局手工格式脚本
- `references/camera-screenshot-white-bg.md` — Excel 照相机截图白底修正方案
- `test-prompts.json` — 类型推断/美化/万元单位的测试用例（含预期结果）

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
| — | 数字颜色：手动输入=主题色，公式=黑色 | beautify 模式：公式→`formula_font`（黑），值→`data_font`（主题色，如深蓝/深褐/深紫取决于主题）。make_excel 模式无公式，数字列全用 `data_font` |
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

## 失败模式与恢复（三段式 Fallback 表）

> 所有失败模式按「触发条件 → 一线修复 → 兜底方案」三层结构编码。
> 遇到报错或异常输出时，先查本表定位触发条件，按一线修复操作，仍失败走兜底。

| # | 触发条件 | 一线修复 | 兜底方案 |
|---|---------|---------|---------|
| 1 | beautify 后 A 列（部门/名称/合计）缺少蓝底白字和框线 | `_beautify_worksheet()` 已自动检测 A1 是否有值，有则补齐 A 列格式。确认 A1 非空再跑 beautify | 若 A1 有值但仍缺格式，用 `references/dual-header-format.py` 手动指定 A 列范围补充 |
| 2 | 列名含"毛利率"误判为 money 而非 pct | 列名关键词有优先级 `pct > date > text > money`，`率$` 优先匹配。确认列名以"率"结尾 | 用 `col_types={'毛利率': 'pct'}` 强制指定 |
| 3 | 含长公式的数值列（如 `=Sheet2!B2/Sheet1!D3`）列宽过大 | 自动检测公式超 15 字符时列宽固定为 15。确认该列确有长公式 | 手动 `ws.column_dimensions[col].width = 15` |
| 4 | theme 参数写错（如 `corol` 而非 `coral`） | make_excel/beautify 入口已校验 `if theme not in THEMES: raise ValueError`，会报错中断而非静默回退 | 检查 theme 名拼写，可用 `sorted(THEMES.keys())` 查看完整列表 |
| 5 | beautify 时首行不是表头（有备注行/空行在前面） | `_detect_data_range()` 从第 1 个非空行开始。跑前确认第 1 个非空行就是表头 | 用 `make_excel()` 从 DataFrame 重建，不走 beautify |
| 6 | 工作表名超 31 字符或含非法字符 `\\/ * ? : [ ]` | 代码自动截断 31 字符并替换非法字符为 `_`。检查截断后是否与其他表名冲突 | 手动重命名其中一个 sheet，避免前 31 字符重复 |
| 7 | beautify 遇到双表头布局（Row2 大标题 + Row3 列头 + 后续第二组列头） | beautify 不支持纵向双表头。立即回退到自动备份 `*_backup_*.xlsx` | 用 openpyxl 手工脚本逐段控制格式，见 `references/dual-header-format.py` |
| 8 | beautify 遇到 Formula 对象导致列宽失控（DataTableFormula/ArrayFormula） | 代码已有 `_cell_display_text()` 保护，不会出现内存地址撑宽列宽。若列宽仍异常，确认编辑器版本 ≥ 2026-06-22 | 手动 `ws.column_dimensions[col].width = 15` |
| 9 | beautify 中 `is_summary` 报 `UnboundLocalError` | 2026-06-13 已修复：从 4.5 段前提取检测逻辑，与 5 段共用。确认版本 ≥ 该日期 | 回退备份，确保 `is_summary` 在引用前定义 |
| 10 | CSV 调用 `python3` 而非 `python` | Windows 上 `python3` 不可用。CLI 示例均使用 `python` | 用 `python` 替换 `python3` |
| 11 | 万元单位（0~1 的 float）被误判为 pct | 2026-06-22 已改为：百分比只靠列名关键词推断，float 一律判 money。确认该列列名不含"率/占比/百分比" | 用 `col_types={'列名': 'money'}` 强制指定 |
| 12 | beautify 后公式颜色配错（应为黑色但显示蓝色） | 2026-06-22 修复：用 `_cell_display_text(cell).startswith('=')` 统一检测公式。确认非 `DataTableFormula` 对象 | 手动设置 `cell.font = DATA_FONT_BLACK` |

### 代码级失败恢复（运行时）

| 失败场景 | 自动恢复机制 |
|---------|-------------|
| beautify 修改出错 | 自动生成 `*_backup_*.xlsx` 备份，可回退 |
| theme 不存在 | 显式抛出 `ValueError`（不会静默回退） |
| 类型推断不准 | 提供 `col_types_override` / `fmt_override` 两重手动覆盖参数 |
| 数据范围检测偏移 | `freeze_rows` 参数可手动指定冻结行，绕过自动推断 |

## 反例与黑名单

当你使用本 skill 时，以下模式是经过反复迭代验证的"不要做的事"：

| # | 反模式 | 为什么不要 | 替代做法 |
|---|--------|-----------|---------|
| 1 | 用 xlwings 保存 | xlwings 会吞掉 openpyxl 设置的边框/颜色/填充，导致三步修补循环的脆弱架构 | 用纯 openpyxl 字符估算法，一次保存完成 |
| 2 | 在 beautify 中改 cell.value | beautify 只负责格式，改数据会破坏原始表的公式/值完整性 | beautify 只改 font/fill/alignment/number_format/border/row_height/column_width |
| 3 | 多色相颜色方案当深浅对比 | 不同色相之间肉眼优先比较色相差异，忽略明度对比。RGB 值看代码是深红→浅绿，用户看到的是红绿不同 | 用同一色系的 HSL lightness 渐变，单蓝色 `hsl(212, 75%, light%)` |
| 4 | 忽略 A 列格式 | beautify 默认从 B 列开始，如果 A 列有数据（部门/名称/合计），会缺少表头蓝底白字和框线 | beautify 时自动检测 A1 是否有值，有则从 A 列开始应用完整摩根系格式 |
| 5 | 跳过类型推断关键词优先级 | 列名含"毛利"但实际是"毛利率"（百分比类型），被 money 规则先匹配 | 用 `率$` 限定词尾，money 中用 `毛利(?!率)` 排除误匹配 |
| 6 | beautify 前不确认首行是表头 | `_detect_data_range()` 在首行有零散非表头数据时可能误判 header_row | beautify 时如果 A 列有空行或备注行混杂，先确认首行确实是表头 |
| 7 | 直接用 `#,##0.00` 做公式列列宽 | 公式字符串（如 `=Sheet2!B2/Sheet1!D3`）按字符宽度估算会超长，用户根本不需要看到完整公式 | 检测到公式超 15 字符时，列宽固定为 15 |
| 8 | 不验证 theme 参数有效性 | 写错的 theme 名（如 `coral` 拼成 `corol`）会静默回退到 default 水蓝，完全没有出错提示 | 在入口处显式校验 `if theme not in THEMES: raise ValueError` |
| 9 | 用 beautify 处理双表头/多数据块布局 | beautify 只做首行检测+单一连续数据区域，Row3+实际列头被跳过、空行区域被加框线 | 恢复备份，用手工脚本精确控制每段格式 |
| 10 | 在 beautify 中对 Formula 对象用 str() 提取文本 | DataTableFormula/ArrayFormula 的 str() 返回内存地址（如 `<openpyxl.worksheet.formula.DataTableFormula at 0x...>`），导致列宽被撑到 MAX_COL_WIDTH | 用 `_cell_display_text(cell)` 安全提取：常规值返回自身、公式对象取 `.value`、DataTable 返回空 |
| 11 | 对万元/千元单位的小数值表使用 beautify | **[已修复]** beautify 此前有值分析 0~1 → pct 的回退逻辑，会误判万元单位。2026-06-22 已改为百分比只靠列名关键词，值分析不再回退 pct | 如果列名不含 pct 关键词（率/占比等），beautify 现在会正确判为 money，不再需要绕行 |

> **🔴 CHECKPOINT · 🛑 DELIVERY GATE** — 文件交付给用户前，逐项验证：
> 1. 打开 xlsx，目测所有列的格式是否基本对（尤其是小数列和百分数列不分岔）
> 2. 检查 A1 是否空白无样式（摩根标准：A列留空或已补齐格式）
> 3. 边框：上下粗 → 中间虚线 → 下粗，无竖线
> 4. 表头：蓝底白字（或对应主题色），全部右对齐
> 5. 数据字体 Arial 11，数字千分位两位小数
> 6. 网格线已隐藏，冻结窗格到位
> 7. 公式列字体颜色是否正确（公式→黑色，值→蓝色）
>
> 以上全部通过后再交给用户。**交付后被用户指出格式问题，说明没有过这道门。**
