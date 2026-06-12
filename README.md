<div align="center">
  <img src="skill-card.png" alt="excel-master skill card" width="800">
</div>

<br>

# excel-master

Hermes Agent 技能 —— 从 DataFrame 生成/美化摩根士丹利标准格式的 Excel 报表。

纯 openpyxl，零 xlwings，一次保存。

## 这是什么

一个 Hermes 技能，安装在 `skills/data-science/excel-master/` 下。AI Agent（大天二）在接到「出个 Excel」「美化报表」「导出 xlsx」等指令时，调用 `make_excel.py` 生成符合《为什么精英都是Excel控》九大原则的格式。

主要能力：

- **make_excel** — 从 DataFrame 生成摩根系标准格式（B2 起写、水蓝表头、千分位、框线上下粗中间细、隐藏网格线）
- **beautify** — 美化已有 xlsx，只改边框/字体/填充/对齐/数字格式/行高/列宽，不碰单元格的值和公式
- **6 套色系主题** — 表头/合计行/数据字体颜色一键切换

## 色系主题

| 主题名 | 表头 | 合计行 |
|--------|------|--------|
| `default` | 水蓝 `#4472C4` | 浅蓝 `#D9E2F3` |
| `deep-navy` | 深海蓝 `#1F4E79` | 浅蓝 `#D6E4F0` |
| `jade` | 墨玉绿 `#375623` | 浅绿 `#E2EFDA` |
| `slate` | 陨石灰蓝 `#404040` | 浅灰 `#D9D9D9` |
| `burgundy` | 勃艮第红 `#843C0C` | 浅杏 `#FCE4D6` |
| `coral` | 珊瑚橙 `#D84B4B` | 浅粉 `#FDE8E8` |

## 作为独立脚本使用

Hermes 之外的场景也能用这个脚本。环境需要 `pandas` + `openpyxl`：

```bash
pip install pandas openpyxl
python make_excel.py 数据.csv 输出.xlsx
python make_excel.py --beautify 已有报表.xlsx 美化后.xlsx --theme coral
```

Python import：

```python
from make_excel import make_excel, beautify

# 生成
make_excel(df, '报表.xlsx', theme='deep-navy')

# 美化已有文件（保留公式）
beautify('原始表.xlsx', '美化表.xlsx', theme='slate')
```

## 摩根系九大原则

来自《为什么精英都是Excel控》——行高18、Arial 11、千分位、上下粗框线中间虚线无竖线、数字右对齐文字左对齐、不从A1开始、隐藏网格线、冻结首行。详见 `SKILL.md`。

## 更新日志

### 2026-06-12
- 修复 beautify 冻结窗格位置错误：freeze_panes 从硬编码 A2 改为跟随表头实际位置
- 修复 blur-excel-column.py 健壮性：win32com/PIL 改为延迟导入，缺失依赖时友好报错
- 修复 A 列格式覆盖不全：beautify 和 make_excel 的 A 列数据现在会做类型推断，应用正确的数字格式/对齐/颜色
- 扩充业务关键词：pct 和 money 类增加返点/佣金/折扣/补贴/commission/discount/rebate 等常用词
- 扩充合计行关键词：增加 subtotal、汇总
- 代码可维护性：修复步骤编号冲突、消除 THEMES.get() 死代码、_detect_data_range 变量重命名消除歧义
- 文档同步：更新 SKILL.md 过时描述，新增公式保护说明和跨平台限制说明
- 补充 README 更新日志

### 2026-06-10
- 初始发布，支持 make_excel（DataFrame → 摩根系 Excel）、beautify（美化不改数据）、6 套色系主题一键切换
- 架构决策：彻底废弃 xlwings，改用纯 openpyxl 字符估算法，一次保存保证格式完整
- 修复 `end_col` 未定义变量导致 NameError（_apply_styles 和 _beautify_worksheet 两处）
- 修复类型推断优先级（pct > date > text > money），解决"毛利率"误匹配为 money 的问题
- 全表字体统一：所有单元格强制 Font(name='Arial', size=11)
- 确立 beautify 只改格式不改数据原则

## License

MIT
