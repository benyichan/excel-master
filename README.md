# excel-master

**Hermes Agent 技能** — 从 DataFrame 生成摩根士丹利标准格式的 Excel 报表。

纯 openpyxl，零 xlwings，一次保存线性流程。支持 beautify（保留公式只改格式）、6 套色系主题、自定义数字格式、可配置表头冻结。

---

## 核心能力

- **make_excel** — DataFrame → 摩根系 Excel（Arial 11、水蓝表头、千分位、框线上下粗中间细、B2起始）
- **beautify** — 美化已有 Excel，只改格式不改数据，保留公式和值
- **6 套色系主题** — 水蓝（默认）、深海蓝、墨玉绿、陨石灰蓝、勃艮第红、珊瑚橙
- **智能列宽** — 非空列下限10上限50自适应，数值/百分比列长公式限宽15
- **可配置表头冻结** — 参数指定/自动检测/B2起始/默认首行，4种策略
- **智能类型推断** — 通过列名关键词 + 值采样自动识别 money / number / pct / date / text
- **手动列类型覆盖** — `col_types` 参数手动指定
- **自定义数字格式** — `fmt_override` 参数覆盖默认格式
- **多 sheet 支持** — 单 DataFrame 或 `[(名称, df), ...]` 列表
- **自动备份** — beautify 原地覆盖前自动备份

---

## 作为 Hermes Agent 技能使用

### 安装依赖

```bash
pip install pandas openpyxl>=3.0
```

### Python API

```python
from make_excel import make_excel, beautify
import pandas as pd

# 从零生成
df = pd.read_csv('data.csv')
make_excel(df, '报表.xlsx')

# 多 sheet
make_excel([
    ('汇总', df_summary),
    ('明细', df_detail),
], '报表.xlsx')

# 切换主题
make_excel(df, '报表-深海蓝.xlsx', theme='deep-navy')

# 自定义数字格式
make_excel(df, '报表.xlsx', fmt_override={'money': '#,##0.0'})

# 冻结表头（默认自动推断）
make_excel(df, '报表.xlsx', freeze_rows=2)   # 冻前2行
make_excel(df, '报表.xlsx', freeze_rows=0)   # 不冻结

# 美化已有文件（保留公式）
beautify('已有报表.xlsx')               # 原地美化，自动备份
beautify('已有报表.xlsx', '美化后.xlsx')  # 另存美化

# 美化 + 手动指定列类型
beautify('订单表.xlsx', col_types={'订单号': 'text', '金额': 'money'})

# 美化 + 冻结表头
beautify('报表.xlsx', freeze_rows=3)     # 冻前3行

# 美化 + 自定义格式
beautify('报表.xlsx', fmt_override={'pct': '0.0%'})
```

### 命令行

```bash
# 生成
python scripts/make_excel.py 数据.csv 输出.xlsx

# 美化
python scripts/make_excel.py --beautify 已有报表.xlsx 美化后.xlsx

# 带参数
python scripts/make_excel.py --beautify 报表.xlsx 美化后.xlsx --theme coral --freeze-rows 3
```

---

## 对话场景（agent 交互层）

当 agent 在对话中为用户处理 Excel 时，使用 `interactive_make_excel.py` 作为入口：

```bash
# 用户回答了表头行数和配色
python scripts/interactive_make_excel.py beautify 报表.xlsx --freeze-rows 3 --theme coral

# 用户跳过 → 自动推断+水蓝
python scripts/interactive_make_excel.py beautify 报表.xlsx

# 从 CSV 生成
python scripts/interactive_make_excel.py make 数据.csv 输出.xlsx --theme deep-navy
```

自动化脚本（定时任务、上层工作流）直接调用 `make_excel()` / `beautify()`，不走此入口。

---

## 色系主题

| 主题名 | 中文 | 表头 | 合计行 | 特点 |
|--------|------|------|--------|------|
| `default` | 水蓝 | #4472C4 白字 | #D9E2F3 | 经典摩根系默认 |
| `deep-navy` | 深海蓝 | #1F4E79 白字 | #D6E4F0 | 更深沉专业 |
| `jade` | 墨玉绿 | #375623 白字 | #E2EFDA | 清爽自然 |
| `slate` | 陨石灰蓝 | #404040 白字 | #D9D9D9 | 现代极简 |
| `burgundy` | 勃艮第红 | #843C0C 白字 | #FCE4D6 | 暖色调高级感 |
| `coral` | 珊瑚橙 | #D84B4B 白字 | #FDE8E8 | 活泼明亮 |

---

## 摩根系 9 大原则

源自《为什么精英都是Excel控》：

1. 行高 18
2. 英文字体 Arial，字号统一 11
3. 数字千分位区隔（#,##0.00）
4. 项目下细项缩排（需人工处理）
5. 单位自成一栏（需人工处理）
6. 同一层级栏宽统一（需人工处理）
7. 框线上下粗中间细虚线，无竖线
8. 文字左对齐，数字右对齐，表头全部右对齐
9. 不从 A1 开始（B2 起始，A 列留空）

---

## 相关参考

- `SKILL.md` — Hermes skill 完整文档（含强制约束和 Pitfalls）
- `references/implementation-checklist.md` — 交付前逐格验证清单
- `references/type-inference-rules.md` — 列类型推断规则
- `scripts/make_excel.py` — 核心脚本
- `scripts/interactive_make_excel.py` — agent 交互层

> Excel 照相机截图和指定列打码见同分类下的 `excel-screenshot-blur` skill。

## License

MIT
