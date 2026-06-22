# 交付前验证清单

## 格式完整性

- [ ] A 列：若含数据，表头蓝底白字右对齐 + 数据 Arial 11 左对齐 + 合计行浅蓝底粗体；若为空则保持列宽 2
- [ ] A1 空白或已补齐格式（摩根标准：A 列不留脏样式）
- [ ] 边框：数据区域上下粗（`medium`）、中间虚线（`dashed`）、无竖线
- [ ] 多数据块表格检查：首个非空列偏移 >2 的区块已各自独立 TOP/MID/BOTTOM 框线
- [ ] 表头：蓝底白字（或当前主题色），全部右对齐
- [ ] 数据字体：全部 Arial 11，数字列叠加颜色（公式黑 / 值主题色），文本列默认黑
- [ ] 千分位符格式：money → `#,##0.00`，number → `#,##0`，pct → `0.00%`
- [ ] 行高 18 已应用
- [ ] 隐藏网格线（`showGridLines = False`）
- [ ] 右侧空白列（宽 3）已添加
- [ ] 冻结窗格已设置（参数指定 / 自动检测 / B2 起始 / 默认首行）

## beautify 专项

- [ ] 公式保留：单元格值未被改动（beautify 只改 font/fill/alignment/number_format/border/row_height/column_width）
- [ ] 公式颜色双重检测：`_cell_display_text(cell).startswith('=')` or `cell.data_type == 'f'` → 黑色，否则 → 主题色
- [ ] DataTableFormula / ArrayFormula 对象列宽正常（`_cell_display_text` 保护生效）
- [ ] `_detect_data_range()` 正确识别表头行。如果首行有备注/空行，确认第 1 个非空行确实是表头

## 类型推断验证（2026-06-22 策略：百分比只靠列名关键词）

- [ ] 关键列手动验证一次 `number_format`：
  - 小数值列（万元/千元单位）→ `#,##0.00`，不是 `0.00%`
  - 含"率/占比/百分比"关键词的列 → `0.00%`
  - 日期列 → `yyyy/mm/dd`（或对应 date 格式）
  - 文本列（订单号/编码/证书号）→ `@`，左对齐
- [ ] `_infer_column_type` 如果返回 `'unknown'`，已用 `col_types` 强制指定

## 条件格式自适应

- [ ] colorScale 色阶的最大色已替换为当前主题表头色
- [ ] 其他条件格式（dataBar、iconSet）未被动到

## 主题校验

- [ ] theme 参数存在，不静默回退（`if theme not in THEMES: raise ValueError`）
- [ ] 12 套可选主题：水蓝 / 深海蓝 / 墨玉绿 / 陨石灰蓝 / 勃艮第红 / 珊瑚橙 / 樱花粉 / 暖阳橙 / 薰衣草紫 / 抹茶绿 / 蜜桃 / 雾蓝紫

## 数据逻辑

- [ ] 合计行 = 明细之和
- [ ] 占比 = 部分/整体（非 Excel 四舍五入误差累积）
- [ ] 示例表数据经得起推演
