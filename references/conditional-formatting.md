# openpyxl 条件格式 — 颜色阶 / 数据条 / 图标集

## 适用场景

摩根系规范管的是静态格式（边框、字体、对齐、数字格式），条件格式用于数据驱动型报表——让数值大小、分布状态通过颜色/图标一目了然，同时不丢失原始数值。

## 导入

```python
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, IconSetRule, CellIsRule
```

## 绿-白色阶（finance 常用）

高值→绿色，低值→白色。适合"越大越好"类指标（消耗、GMV、转化数）。

```python
from openpyxl.utils import get_column_letter

# 数值最大=绿(00B050)，最小=白(FFFFFF)
ws.conditional_formatting.add(
    f"A2:Z10",
    ColorScaleRule(
        start_type='min', start_color='FFFFFF',
        end_type='max', end_color='00B050'
    )
)
```

### 关键：排除合计列

条件格式只作用于纯数据区域，**合计/汇总列不参与色阶**，否则合计列的高值会拉平整体色阶分布，让日常数据失去区分度。

```python
# Section 1: 单行数据，列 2~(date_count+1) 是日期列，最后一列(date_count+2)是合计
sec1_data_row = header_row + 1
start_col = 2
end_col = date_count + 1  # 不含合计列
ws.conditional_formatting.add(
    f"{get_column_letter(start_col)}{sec1_data_row}:{get_column_letter(end_col)}{sec1_data_row}",
    ColorScaleRule(start_type='min', start_color='FFFFFF', end_type='max', end_color='00B050')
)

# Section 2: 多行明细，列 3~(date_count+2) 是日期列，最后一列(date_count+3)是合计
data_start_row = header_row + 1
data_end_row = last_data_row
start_col = 3
end_col = date_count + 2  # 不含合计列
ws.conditional_formatting.add(
    f"{get_column_letter(start_col)}{data_start_row}:{get_column_letter(end_col)}{data_end_row}",
    ColorScaleRule(start_type='min', start_color='FFFFFF', end_type='max', end_color='00B050')
)
```

## 红-白-绿色阶（双向对比）

红→白→绿三色，适合"偏离目标值"类指标：红色=偏差大，白色=刚好，绿色=超额完成。

```python
ws.conditional_formatting.add(
    "A2:Z10",
    ColorScaleRule(
        start_type='min', start_color='FF9999',
        mid_type='percentile', mid_value=50, mid_color='FFFFFF',
        end_type='max', end_color='99FF99'
    )
)
```

## 红-白色阶（越大越差）

高值→红色，低值→白色。适合"越小越好"类指标（退款率、损耗率、缺货天数）。

```python
ws.conditional_formatting.add(
    "A2:Z10",
    ColorScaleRule(
        start_type='min', start_color='FFFFFF',
        end_type='max', end_color='FF4444'
    )
)
```

## 数据条

数值本身映射为彩色条带长度，适合在紧凑表格中快速读取相对大小。

```python
ws.conditional_formatting.add(
    "B2:B20",
    DataBarRule(start_type='min', end_type='max',
                color='4472C4',    # 水蓝，和摩根系表头一致
                showValue=True)    # 同时显示数值
)
```

注意：数据条和单元格数值颜色/字体是叠加的。如果数字列已经有蓝色字体（make_excel 自动设的），数据条颜色会盖在下面但数值仍是蓝色。如果想数据条下方显示黑色数值，先清空格子颜色再设 DataBarRule。

## 图标集

适合三态分类（增/稳/降 或 优/中/差）。

```python
from openpyxl.formatting.rule import FormatObject

ws.conditional_formatting.add(
    "B2:B20",
    IconSetRule(icon_style='3TrafficLights1',   # 红黄绿灯
                type='percentile', values=[0, 33, 67])
)
```

常用 icon_style 值：`3TrafficLights1`（红黄灯）/ `3Arrows`（上下平箭头）/ `3Symbols`（✓✗）

## 注意事项

1. **ColorScaleRule vs ColorScale** — `from openpyxl.formatting.rule import ColorScaleRule` 是快捷构造器，用 start_type/end_color 等关键字参数。不推荐手拼 ColorScale + ConditionalFormattingList。
2. **范围不跨空白区** — 用 `"A1:B5"` 字符串指定范围，不要传 Python range。openpyxl 内部用 CoordinateSequence 解析。
3. **多条件格式叠加** — openpyxl 的 conditionformatting 是列表追加，多个规则按添加顺序优先级从高到低。同时有 ColorScaleRule 和 DataBarRule 在同一区域时，后添加的优先渲染。
4. **排除空白单元格** — 默认 blank cells 不参与色阶（openpyxl 行为）。如果想改变，需要额外参数：`ColorScaleRule(..., auto=False)`。
5. **排除合计列的重要性** — 这是最常见的坑。合计值往往是日期的 N 倍，如果参与色阶，所有日期列都变成浅绿（无区分度）。必须在构造 range 时手动截断合计列。
6. **冻品窗格 vs 条件格式** — 条件格式范围里包含冻结窗格起始行没关系，但不要用条件格式给冻品行（表头）上色——用 PatternFill 在写表头时直接设，条件格式只用于数据区域。

## 主题色搭配的条件格式（色阶跟随配色）

当工作簿应用不同配色主题（水蓝/深海蓝/墨玉绿/陨石灰蓝/勃艮第红/珊瑚橙）时，条件格式的色阶终点颜色应**跟随主题色**，而不是全部用绿色。

```
主题      表头色    色阶终点
水蓝      #4472C4   #4472C4
深海蓝    #1F4E79   #1F4E79
墨玉绿    #375623   #375623
陨石灰蓝  #404040   #404040
勃艮第红  #843C0C   #843C0C
珊瑚橙    #D84B4B   #D84B4B
```

实现模式：先在 openpyxl 中修改表头/合计行的 PatternFill 颜色，然后删掉旧的 CF 规则，用匹配的主题色新建 ColorScaleRule：

```python
from openpyxl.formatting.rule import ColorScaleRule

# 清除旧规则
ws.conditional_formatting._cf_rules.clear()

# 添加主题匹配的新规则
ws.conditional_formatting.add(
    "B4:I4",
    ColorScaleRule(start_type='min', start_color='FFFFFF',
                   end_type='max', end_color=theme_hex)
)
```

注意：`theme_hex` 必须与表头色一致，这样整张表的视觉重心是统一的。旧规则被清理是因为 openpyxl 的 CF 规则是列表追加，不清除的话新旧规则叠加会导致混乱。
