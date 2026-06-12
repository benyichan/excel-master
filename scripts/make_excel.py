#!/usr/bin/env python3
"""
make_excel — 从 DataFrame 生成摩根士丹利标准格式的 Excel 报表。
beautify — 对已有 Excel 文件只改格式不改数据（保留公式和值）。

用法:
    from make_excel import make_excel, beautify
    make_excel(df, 'output.xlsx')
    beautify('input.xlsx')              # 原地美化
    beautify('input.xlsx', 'out.xlsx')  # 另存美化

    python make_excel.py data.csv output.xlsx
    python make_excel.py --beautify input.xlsx output.xlsx

依赖: pandas, openpyxl（openpyxl>=3.0，无 xlwings）
设计: 写数据 → 上格式 → 估列宽 → 一次保存，线性流程不回头。
"""

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path
from shutil import copy2
from typing import Union, Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ═══════════════════════════════════════════════════════════════════════════
# 格式常量 — 摩根士丹利标准
# ═══════════════════════════════════════════════════════════════════════════

# 框线（只有上下，无竖线）
THICK = Side(style='medium', color='000000')
DASHED = Side(style='dashed', color='808080')
NO = Side(style=None)

TOP_BORDER = Border(top=THICK, bottom=DASHED, left=NO, right=NO)
MID_BORDER = Border(top=DASHED, bottom=DASHED, left=NO, right=NO)
BOTTOM_BORDER = Border(top=DASHED, bottom=THICK, left=NO, right=NO)

# 表头：水蓝底白字
HDR_FILL = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
HDR_FONT = Font(name='Arial', size=11, bold=True, color='FFFFFF')
HDR_ALIGN = Alignment(horizontal='right', vertical='center')

# 数据字体颜色（摩根系：蓝=手动输入，黑=公式结果）
DATA_FONT_BLUE = Font(name='Arial', size=11, color='0000FF')
DATA_FONT_BLACK = Font(name='Arial', size=11, color='000000')

# 合计行样式
SUM_FILL = PatternFill(start_color='D9E2F3', end_color='D9E2F3', fill_type='solid')
SUM_FONT = Font(name='Arial', size=11, bold=True)

# 合计行关键词（当最后一行的首列文本匹配时，视为合计行）
SUM_KEYWORDS = ('合计', '小计', '总计', '汇总', 'sum', 'total', 'subtotal', '小計', '合計')

# 数据对齐
LEFT_ALIGN = Alignment(horizontal='left', vertical='center')
RIGHT_ALIGN = Alignment(horizontal='right', vertical='center')

# 数字格式
FMT = {
    'money': '#,##0.00',
    'number': '#,##0',
    'pct': '0.00%',
    'date': 'yyyy/mm/dd',
    'text': '@',
    'unknown': '@',
}

# 布局
ROW_HEIGHT = 18
A_COL_WIDTH = 2
MIN_COL_WIDTH = 10
MAX_COL_WIDTH = 50


# ═══════════════════════════════════════════════════════════════════════════
# 色系主题
# ═══════════════════════════════════════════════════════════════════════════

THEMES = {
    'default': {
        'label': '水蓝',
        'header_fill': '4472C4',
        'header_font_color': 'FFFFFF',
        'sum_fill': 'D9E2F3',
        'sum_font_bold': True,
        'sum_font_color': '000000',
        'data_font_blue': '0000FF',
        'data_font_black': '000000',
        'desc': '经典水蓝（原摩根系默认）',
    },
    'deep-navy': {
        'label': '深海蓝',
        'header_fill': '1F4E79',
        'header_font_color': 'FFFFFF',
        'sum_fill': 'D6E4F0',
        'sum_font_bold': True,
        'sum_font_color': '000000',
        'data_font_blue': '0000FF',
        'data_font_black': '000000',
        'desc': '深海蓝表头，更沉稳专业',
    },
    'jade': {
        'label': '墨玉绿',
        'header_fill': '375623',
        'header_font_color': 'FFFFFF',
        'sum_fill': 'E2EFDA',
        'sum_font_bold': True,
        'sum_font_color': '000000',
        'data_font_blue': '0000FF',
        'data_font_black': '000000',
        'desc': '墨绿表头，清爽自然',
    },
    'slate': {
        'label': '陨石灰蓝',
        'header_fill': '404040',
        'header_font_color': 'FFFFFF',
        'sum_fill': 'D9D9D9',
        'sum_font_bold': True,
        'sum_font_color': '000000',
        'data_font_blue': '0000FF',
        'data_font_black': '000000',
        'desc': '深灰表头 + 蓝色色阶，现代极简',
    },
    'burgundy': {
        'label': '勃艮第红',
        'header_fill': '843C0C',
        'header_font_color': 'FFFFFF',
        'sum_fill': 'FCE4D6',
        'sum_font_bold': True,
        'sum_font_color': '000000',
        'data_font_blue': '0000FF',
        'data_font_black': '000000',
        'desc': '酒红表头，暖色调高级感',
    },
    'coral': {
        'label': '珊瑚橙',
        'header_fill': 'D84B4B',
        'header_font_color': 'FFFFFF',
        'sum_fill': 'FDE8E8',
        'sum_font_bold': True,
        'sum_font_color': '000000',
        'data_font_blue': '0000FF',
        'data_font_black': '000000',
        'desc': '珊瑚红表头，活泼明亮',
    },
}

AVAILABLE_THEMES = ', '.join(sorted(THEMES.keys()))


def _validate_theme(theme_name: str):
    """校验主题名，无效则抛 ValueError。"""
    if theme_name not in THEMES:
        raise ValueError(
            f"未知主题 '{theme_name}'。可用主题: {sorted(THEMES.keys())}"
        )


def _build_theme_styles(theme_name: str):
    """根据主题名构建 PatternFill / Font 对象。"""
    _validate_theme(theme_name)  # 保险：二次校验
    t = THEMES[theme_name]
    return {
        'header_fill': PatternFill(start_color=t['header_fill'], end_color=t['header_fill'], fill_type='solid'),
        'header_font': Font(name='Arial', size=11, bold=True, color=t['header_font_color']),
        'sum_fill': PatternFill(start_color=t['sum_fill'], end_color=t['sum_fill'], fill_type='solid'),
        'sum_font': Font(name='Arial', size=11, bold=t['sum_font_bold'], color=t['sum_font_color']),
        'data_font_blue': Font(name='Arial', size=11, color=t['data_font_blue']),
        'data_font_black': Font(name='Arial', size=11, color=t['data_font_black']),
    }


# ═══════════════════════════════════════════════════════════════════════════
# 列类型推断
# ═══════════════════════════════════════════════════════════════════════════

def _infer_column_type(name: str, series: pd.Series) -> str:
    """根据列名关键词 + 值特征推断列类型。"""
    n = name.lower()

    # ── 关键词匹配（优先级：pct > date > text > money）──
    # "率"优先检测，避免"毛利率"被 money 关键词误匹配
    if re.search(r'(率$|占比|百分比|增长率|利润率|达成率|rate$|ratio$|margin|pct|percent)', n):
        return 'pct'
    if re.search(r'(日期|时间|年月|期间|年份|月份|date|time|year|month)', n):
        return 'date'
    if re.search(r'(编号|id|单号|编码|电话|手机|备注|说明|名称|地址|描述|号码|订单号|负责人|姓名|联系人|部门|岗位|phone|email|code|desc|address)', n):
        return 'text'
    if re.search(r'(金额|价格|收入|成本|费用|毛利额|净利|合计|总额|售价|单价|预算|支出|毛利(?!率)|返点|佣金|折扣|补贴|周转|本金|amount|price|cost|revenue|total|budget|expense|commission|discount|rebate|profit|spending)', n):
        return 'money'

    # ── 值特征 ──
    clean = series.dropna()
    if len(clean) == 0:
        return 'unknown'

    try:
        if pd.api.types.is_datetime64_any_dtype(clean):
            return 'date'

        if pd.api.types.is_numeric_dtype(clean):
            sample = clean.head(20)
            if pd.api.types.is_float_dtype(sample):
                # 0~1 之间的浮点 → 百分比
                if (sample >= 0).all() and (sample <= 1).all() and sample.nunique() > 2:
                    return 'pct'
                return 'money'
            return 'number'

        # 长数字字符串 → text
        str_sample = clean.astype(str).head(20)
        if str_sample.str.match(r'^\d{12,}$').any():
            return 'text'
    except Exception:
        pass

    return 'unknown'


# ═══════════════════════════════════════════════════════════════════════════
# 列宽估算
# ═══════════════════════════════════════════════════════════════════════════

def _char_width(text: str) -> int:
    """中文字符计2，拉丁/数字计1。"""
    return sum(2 if ord(ch) > 127 else 1 for ch in str(text))


def _estimate_column_width(header: str, series: pd.Series) -> int:
    """估算列宽：取表头和数据中最大字符宽度 + padding。"""
    max_w = _char_width(header)
    for v in series.dropna():
        try:
            w = _char_width(str(v))
            if w > max_w:
                max_w = w
        except Exception:
            pass
    return max(MIN_COL_WIDTH, min(max_w + 2, MAX_COL_WIDTH))


# ═══════════════════════════════════════════════════════════════════════════
# 格式引擎
# ═══════════════════════════════════════════════════════════════════════════

def _apply_styles(ws, df: pd.DataFrame, theme: str = 'default', fmt_override: Optional[dict] = None):
    """对已写入数据的工作表应用摩根系格式。

    fmt_override : dict, optional
        自定义数字格式，如 {'money': '#,##0.0', 'pct': '0.0%'}。
        合并到默认 FMT 之上，同名覆盖。
    """
    s = _build_theme_styles(theme)
    # 合并自定义数字格式
    effective_fmt = dict(FMT)
    if fmt_override:
        effective_fmt.update(fmt_override)
    nrows = 1 + len(df)   # 表头 + 数据行
    ncols = len(df.columns)
    start_col = 2         # B 列

    # 1. 行高 + A列宽 + A列格式（如果A列有数据）
    for r in range(1, nrows + 1):
        ws.row_dimensions[r].height = ROW_HEIGHT
    a_has_data = any(ws.cell(row=r, column=1).value is not None for r in range(2, nrows + 1))
    if a_has_data:
        a_max = _char_width(str(ws.cell(row=1, column=1).value or ''))
        for r in range(2, nrows + 1):
            v = ws.cell(row=r, column=1).value
            if v is not None:
                w = _char_width(str(v))
                if w > a_max:
                    a_max = w
        ws.column_dimensions['A'].width = max(MIN_COL_WIDTH, min(a_max + 2, MAX_COL_WIDTH))
        # A列表头 + 数据格式（类型推断，右对齐/蓝色）
        ws.cell(row=1, column=1).font = s['header_font']
        ws.cell(row=1, column=1).fill = s['header_fill']
        ws.cell(row=1, column=1).alignment = HDR_ALIGN
        # A 列类型推断（与 B~K 列一致）
        _a_col_name = str(df.columns[0]) if len(df.columns) > 0 else ''
        _a_col_type = _infer_column_type(_a_col_name, df.iloc[:, 0]) if len(df.columns) > 0 else 'unknown'
        _a_is_num = _a_col_type in ('money', 'number', 'pct')
        _a_align = RIGHT_ALIGN if _a_is_num else LEFT_ALIGN
        _a_nf = effective_fmt.get(_a_col_type)
        for r in range(2, nrows + 1):
            cell = ws.cell(row=r, column=1)
            cell.alignment = _a_align
            if _a_nf:
                cell.number_format = _a_nf
            if _a_is_num:
                cell.font = s['data_font_blue']
            else:
                cell.font = Font(name='Arial', size=11)
    else:
        ws.column_dimensions['A'].width = A_COL_WIDTH

    # 2. 列宽估算 + 类型推断（一次扫描）
    col_types = {}
    for ci in range(ncols):
        col_name = df.columns[ci]
        col_letter = get_column_letter(ci + start_col)
        col_type = _infer_column_type(str(col_name), df.iloc[:, ci])
        col_types[ci] = col_type

        est = _estimate_column_width(str(col_name), df.iloc[:, ci])
        ws.column_dimensions[col_letter].width = est

    # 3. 表头（第1行）
    for ci in range(ncols):
        cell = ws.cell(row=1, column=ci + start_col)
        cell.font = s['header_font']
        cell.fill = s['header_fill']
        cell.alignment = HDR_ALIGN

    # 4. 数据行（统一 Arial 11 + 数字格式 + 对齐 + 字体颜色）
    for ci in range(ncols):
        ct = col_types.get(ci, 'unknown')
        is_num = ct in ('money', 'number', 'pct')
        align = RIGHT_ALIGN if is_num else LEFT_ALIGN
        nf = effective_fmt.get(ct)

        for ri in range(2, nrows + 1):
            cell = ws.cell(row=ri, column=ci + start_col)
            # 所有单元格统一 Arial 11
            cell.font = s['data_font_blue'] if is_num else Font(name='Arial', size=11)
            cell.alignment = align
            if nf:
                cell.number_format = nf

    # 5. 合计行检测与特殊格式（最后一行的首列含合计关键词）
    last_val = ws.cell(row=nrows, column=start_col).value
    is_summary = (
        nrows > 1
        and last_val is not None
        and any(kw in str(last_val) for kw in SUM_KEYWORDS)
    )
    if is_summary:
        for ci in range(ncols):
            cell = ws.cell(row=nrows, column=ci + start_col)
            cell.font = s['sum_font']
            cell.fill = s['sum_fill']

    # 6. 边框（无竖线，上下粗中间虚线）
    for ci in range(ncols):
        col = ci + start_col
        for ri in range(1, nrows + 1):
            cell = ws.cell(row=ri, column=col)
            if ri == 1:
                cell.border = TOP_BORDER
            elif ri == nrows:
                cell.border = BOTTOM_BORDER
            else:
                cell.border = MID_BORDER
    # A列边框（如果有数据）
    if a_has_data:
        for ri in range(1, nrows + 1):
            cell = ws.cell(row=ri, column=1)
            if ri == 1:
                cell.border = TOP_BORDER
            elif ri == nrows:
                cell.border = BOTTOM_BORDER
            else:
                cell.border = MID_BORDER

    # 7. 右侧空白列（宽 3）
    last_data_col = start_col + ncols - 1
    right_col = get_column_letter(last_data_col + 1)
    ws.column_dimensions[right_col].width = 3

    # 8. 网格线 + 冻结
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = 'A2'


# ═══════════════════════════════════════════════════════════════════════════
# 写数据
# ═══════════════════════════════════════════════════════════════════════════

def _write_data(ws, df: pd.DataFrame):
    """从 B2 开始写入（第1行 = 表头，A列留空）。"""
    for ci, col_name in enumerate(df.columns):
        ws.cell(row=1, column=ci + 2, value=col_name)
    for ri, (_, row) in enumerate(df.iterrows()):
        for ci, value in enumerate(row):
            ws.cell(row=ri + 2, column=ci + 2, value=value)


# ═══════════════════════════════════════════════════════════════════════════
# 主入口
# ═══════════════════════════════════════════════════════════════════════════

def make_excel(data, output_path: str, sheet_name: str = 'Sheet1', theme: str = 'default',
               fmt_override: Optional[dict] = None) -> str:
    """生成摩根系标准格式 Excel。

    参数
    ----
    data : DataFrame | list[tuple[str, DataFrame]]
        单 sheet 传 df，多 sheet 传 [(名称, df), ...]。
    output_path : str
        输出文件路径。
    sheet_name : str
        单 sheet 模式的工作表名称。
    theme : str
        色系主题名。可用: """ + AVAILABLE_THEMES + """。默认 'default'。
    fmt_override : dict, optional
        自定义数字格式，如 {'money': '#,##0.0', 'pct': '0.0%'}。
        合并到默认 FMT 之上，同名覆盖。

    返回
    ----
    str : 输出文件的绝对路径。
    """
    _validate_theme(theme)

    # 统一为多 sheet 格式
    if isinstance(data, pd.DataFrame):
        sheets = [(sheet_name, data)]
    else:
        sheets = list(data)

    wb = Workbook()
    for name in list(wb.sheetnames):
        del wb[name]

    for idx, (name, df) in enumerate(sheets):
        if df is None or (hasattr(df, 'empty') and df.empty):
            continue
        safe_name = re.sub(r'[\\/*?:\[\]]', '_', str(name))[:31]
        ws = wb.create_sheet(title=safe_name, index=idx)
        _write_data(ws, df)
        _apply_styles(ws, df, theme=theme, fmt_override=fmt_override)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    return str(out.absolute())


# ═══════════════════════════════════════════════════════════════════════════
# beautify — 只改格式不改数据
# ═══════════════════════════════════════════════════════════════════════════

def _detect_data_range(ws):
    """检测工作表的实际数据范围。返回 (header_row, data_start, data_end, col_end) 或 None。

    规则：从第1行开始找第一个非空行作为表头行；
    从表头行之后找最后一个非空行作为数据结束；
    跳过列A（索引列通常为空）。
    """
    # 先确定最后一列 — 扫描所有行
    last_col = 0
    header_row = None
    last_data_row = 0

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row or 1, max_col=ws.max_column or 1):
        r = row[0].row
        has_data = any(cell.value is not None for cell in row)
        if has_data and header_row is None:
            header_row = r
        if has_data:
            last_data_row = r
        for cell in row:
            if cell.value is not None and cell.column > last_col:
                last_col = cell.column

    if header_row is None:
        return None  # 空表

    data_start = header_row + 1
    data_end = max(last_data_row, data_start - 1)  # 至少包含表头
    # 数据范围从 B 列（2）开始，跳过 A 列
    start_col = 2
    end_col = max(start_col, last_col)
    return header_row, data_start, data_end, start_col, end_col


def _sample_ws_column(ws, col, start_row, end_row, max_samples=100):
    """从工作表列中采样非空值，返回列表。"""
    samples = []
    for r in range(start_row, end_row + 1):
        v = ws.cell(row=r, column=col).value
        if v is not None:
            samples.append(v)
            if len(samples) >= max_samples:
                break
    return samples


def _infer_col_type_from_ws(header_value, samples):
    """从工作表列推断类型（不依赖 pandas Series）。"""
    name = str(header_value) if header_value is not None else ''
    n = name.lower()

    # 关键词匹配（与 _infer_column_type 一致）
    if re.search(r'(率$|占比|百分比|增长率|利润率|达成率|rate$|ratio$|margin|pct|percent)', n):
        return 'pct'
    if re.search(r'(日期|时间|年月|期间|年份|月份|date|time|year|month)', n):
        return 'date'
    if re.search(r'(编号|id|单号|编码|电话|手机|备注|说明|名称|地址|描述|号码|订单号|负责人|姓名|联系人|部门|岗位|phone|email|code|desc|address)', n):
        return 'text'
    if re.search(r'(金额|价格|收入|成本|费用|毛利额|净利|合计|总额|售价|单价|预算|支出|毛利(?!率)|返点|佣金|折扣|补贴|周转|本金|amount|price|cost|revenue|total|budget|expense|commission|discount|rebate|profit|spending)', n):
        return 'money'

    if not samples:
        return 'unknown'

    # 值采样
    numeric_samples = [v for v in samples if isinstance(v, (int, float))]
    str_samples = [str(v) for v in samples if not isinstance(v, (int, float))]

    # 检查日期
    from datetime import datetime, date
    date_count = sum(1 for v in samples if isinstance(v, (datetime, date)))
    if date_count >= len(samples) * 0.5:
        return 'date'

    if numeric_samples:
        # 浮点检查百分比
        float_samples = [v for v in numeric_samples if isinstance(v, float)]
        if float_samples:
            all_between_0_1 = all(0 < v < 1 for v in float_samples)
            unique_gt_2 = len(set(float_samples)) > 2
            if all_between_0_1 and unique_gt_2:
                return 'pct'
            return 'money'
        return 'number'

    # 长数字字符串
    if any(re.match(r'^\d{12,}$', s) for s in str_samples):
        return 'text'

    return 'unknown'


def _estimate_col_width_from_cells(ws, col, header_row, data_end):
    """从工作表的表头和单元格值估算列宽。"""
    max_w = _char_width(str(ws.cell(row=header_row, column=col).value or ''))
    for r in range(header_row + 1, data_end + 1):
        v = ws.cell(row=r, column=col).value
        if v is not None:
            try:
                w = _char_width(str(v))
                if w > max_w:
                    max_w = w
            except Exception:
                pass
    return max(MIN_COL_WIDTH, min(max_w + 2, MAX_COL_WIDTH))


def _beautify_worksheet(ws, col_types_override: Optional[dict] = None, theme: str = 'default',
                        fmt_override: Optional[dict] = None):
    """对单个工作表应用摩根系格式（不改变任何单元格的值）。

    参数
    ----
    col_types_override : dict, optional
        强制指定列的类型，如 {'订单号': 'text', '金额': 'money'}。
        列名匹配，覆盖自动推断结果。
    theme : str
        色系主题名。可用: """ + AVAILABLE_THEMES + """。默认 'default'。
    fmt_override : dict, optional
        自定义数字格式，如 {'money': '#,##0.0', 'pct': '0.0%'}。
        合并到默认 FMT 之上，同名覆盖。
    """
    s = _build_theme_styles(theme)
    # 合并自定义数字格式
    effective_fmt = dict(FMT)
    if fmt_override:
        effective_fmt.update(fmt_override)
    dr = _detect_data_range(ws)
    if dr is None:
        return  # 空表跳过
    header_row, data_start, data_end, start_col, end_col = dr

    # 1. 行高 + A列宽（采样A列数据估算宽度）
    for r in range(header_row, data_end + 1):
        ws.row_dimensions[r].height = ROW_HEIGHT
    a_max = _char_width(str(ws.cell(row=header_row, column=1).value or ''))
    a_has_data = False
    for r in range(data_start, data_end + 1):
        v = ws.cell(row=r, column=1).value
        if v is not None:
            a_has_data = True
            try:
                w = _char_width(str(v))
                if w > a_max:
                    a_max = w
            except Exception:
                pass
    if a_has_data:
        ws.column_dimensions['A'].width = max(MIN_COL_WIDTH, min(a_max + 2, MAX_COL_WIDTH))
        # A列表头 + 数据格式
        ws.cell(row=header_row, column=1).font = s['header_font']
        ws.cell(row=header_row, column=1).fill = s['header_fill']
        ws.cell(row=header_row, column=1).alignment = HDR_ALIGN
        # A 列类型推断（与 B~K 列一致）
        a_header_val = ws.cell(row=header_row, column=1).value
        a_samples = _sample_ws_column(ws, 1, data_start, data_end)
        a_ct = _infer_col_type_from_ws(a_header_val, a_samples)
        a_is_num = a_ct in ('money', 'number', 'pct')
        a_align = RIGHT_ALIGN if a_is_num else LEFT_ALIGN
        a_nf = effective_fmt.get(a_ct)
        for r in range(data_start, data_end + 1):
            cell = ws.cell(row=r, column=1)
            cell.alignment = a_align
            if a_nf:
                cell.number_format = a_nf
            if a_is_num:
                val = cell.value
                is_formula = (isinstance(val, str) and val.startswith('=')) or cell.data_type == 'f'
                if is_formula:
                    cell.font = s['data_font_black']
                else:
                    cell.font = s['data_font_blue']
            else:
                cell.font = Font(name='Arial', size=11)
    else:
        ws.column_dimensions['A'].width = A_COL_WIDTH

    # 2. 列宽估算 + 类型推断（支持手动覆盖）
    col_types = {}
    col_headers = {}  # 列号 → 列名，用于 override 匹配
    for c in range(start_col, end_col + 1):
        col_letter = get_column_letter(c)
        header_val = ws.cell(row=header_row, column=c).value
        col_headers[c] = str(header_val) if header_val is not None else ''

        # 手动覆盖优先
        if col_types_override and col_headers[c] in col_types_override:
            ct = col_types_override[col_headers[c]]
        else:
            samples = _sample_ws_column(ws, c, data_start, data_end)
            ct = _infer_col_type_from_ws(header_val, samples)
        col_types[c] = ct

        est = _estimate_col_width_from_cells(ws, c, header_row, data_end)
        ws.column_dimensions[col_letter].width = est

    # 3. 表头格式
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=header_row, column=c)
        cell.font = s['header_font']
        cell.fill = s['header_fill']
        cell.alignment = HDR_ALIGN

    # 4. 数据行格式（统一 Arial 11 + 对齐 + 数字格式 + 字体颜色）
    for c in range(start_col, end_col + 1):
        ct = col_types.get(c, 'unknown')
        is_num = ct in ('money', 'number', 'pct')
        align = RIGHT_ALIGN if is_num else LEFT_ALIGN
        nf = effective_fmt.get(ct)

        for r in range(data_start, data_end + 1):
            cell = ws.cell(row=r, column=c)
            cell.alignment = align
            if nf:
                cell.number_format = nf
            # 所有单元格统一 Arial 11
            if is_num:
                val = cell.value
                is_formula = (isinstance(val, str) and val.startswith('=')) or cell.data_type == 'f'
                if is_formula:
                    cell.font = s['data_font_black']  # 公式=黑，不改公式值
                else:
                    cell.font = s['data_font_blue']   # 手动输入=蓝
            else:
                cell.font = Font(name='Arial', size=11)  # 文本=Arial 11

    # 5. 合计行检测与特殊格式
    # 检测最后一行的首列（列A或第一数据列）是否含合计关键词
    first_val = ws.cell(row=data_end, column=1).value  # A 列
    if not first_val or not any(kw in str(first_val) for kw in SUM_KEYWORDS):
        first_val = ws.cell(row=data_end, column=start_col).value  # B 列兜底
    is_summary = (
        data_end > header_row
        and first_val is not None
        and any(kw in str(first_val) for kw in SUM_KEYWORDS)
    )
    if is_summary:
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=data_end, column=c)
            cell.font = s['sum_font']
            cell.fill = s['sum_fill']

    # 6. 边框
    data_rows = data_end - header_row + 1  # 表头 + 数据总行数
    for c in range(start_col, end_col + 1):
        for ri, r in enumerate(range(header_row, data_end + 1)):
            cell = ws.cell(row=r, column=c)
            if ri == 0:
                cell.border = TOP_BORDER
            elif ri == data_rows - 1:
                cell.border = BOTTOM_BORDER
            else:
                cell.border = MID_BORDER
    # A列边框（如果有数据）
    if a_has_data:
        for ri, r in enumerate(range(header_row, data_end + 1)):
            cell = ws.cell(row=r, column=1)
            if ri == 0:
                cell.border = TOP_BORDER
            elif ri == data_rows - 1:
                cell.border = BOTTOM_BORDER
            else:
                cell.border = MID_BORDER

    # 7. 右侧空白列（宽 3）
    ncols = end_col - start_col + 1
    last_data_col = start_col + ncols - 1
    right_col = get_column_letter(last_data_col + 1)
    ws.column_dimensions[right_col].width = 3

    # 8. 网格线 + 冻结（表头在哪行就冻结在哪行下方）
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = f'A{header_row + 1}'


def beautify(
    input_path: str,
    output_path: Optional[str] = None,
    *,
    col_types: Optional[dict] = None,
    backup: bool = True,
    theme: str = 'default',
    fmt_override: Optional[dict] = None,
) -> str:
    """美化已有 Excel 文件，只改格式不改数据（保留公式和值）。

    参数
    ----
    input_path : str
        输入的 xlsx 文件路径。
    output_path : str, optional
        输出路径。不传则原地覆盖原文件。
    col_types : dict, optional
        手动指定列类型，如 {'订单号': 'text', '金额': 'money'}。
        覆盖自动类型推断。可选值: money/number/pct/date/text。
    backup : bool, default True
        原地覆盖前是否自动备份。备份文件名为 <原文件名>_backup_YYYYMMDD_HHMMSS.xlsx。
    theme : str
        色系主题名。可用: """ + AVAILABLE_THEMES + """。默认 'default'。
    fmt_override : dict, optional
        自定义数字格式，如 {'money': '#,##0.0', 'pct': '0.0%'}。
        合并到默认 FMT 之上，同名覆盖。

    返回
    ----
    str : 输出文件的绝对路径。
    """
    _validate_theme(theme)
    from openpyxl import load_workbook

    out = output_path or input_path

    # 自动备份
    if backup and output_path is None:
        inp = Path(input_path)
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        bak = inp.with_stem(f'{inp.stem}_backup_{ts}')
        copy2(input_path, str(bak))
        print(f'备份: {bak}')

    wb = load_workbook(input_path)  # data_only=False 默认，公式保留

    for ws in wb.worksheets:
        _beautify_worksheet(ws, col_types_override=col_types, theme=theme, fmt_override=fmt_override)

    out_path = Path(out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    return str(out_path.absolute())


def main():
    parser = argparse.ArgumentParser(description='生成/美化摩根系标准格式 Excel')
    parser.add_argument('input', help='输入文件（CSV 或 --beautify 时传 xlsx）')
    parser.add_argument('output', nargs='?', default=None,
                        help='输出路径（默认: 桌面/<输入文件名>.xlsx）')
    parser.add_argument('--sheet-name', default='Sheet1', help='工作表名称（仅 CSV 模式）')
    parser.add_argument('--beautify', action='store_true',
                        help='美化已有 xlsx（只改格式不改数据，保留公式）')
    parser.add_argument('--theme', default='default', choices=list(THEMES.keys()),
                        help=f"色系主题 (default: default)。可选: {AVAILABLE_THEMES}")
    args = parser.parse_args()

    if args.beautify:
        path = beautify(args.input, args.output, theme=args.theme)
        print(f'已美化: {path}')
        return

    if args.output is None:
        in_path = Path(args.input)
        args.output = str(Path.home() / 'Desktop' / in_path.with_suffix('.xlsx').name)

    df = pd.read_csv(args.input)
    path = make_excel(df, args.output, sheet_name=args.sheet_name, theme=args.theme)
    print(f'已生成: {path}')


if __name__ == '__main__':
    main()
