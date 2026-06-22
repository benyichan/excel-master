# excel-master

**Hermes Agent 技能** — 从 DataFrame 生成/美化摩根士丹利标准格式的 Excel 报表。

纯 openpyxl，零 xlwings，一次保存线性流程。支持 beautify（保留公式只改格式）和 make_excel（从零生成）两种模式。

---

## 核心能力

- **make_excel** — DataFrame → 摩根系 Excel（Arial 11、水蓝表头、千分位、框线上下粗中间细虚线无竖线、B2 起始、隐藏网格线）
- **beautify** — 美化已有 Excel，只改格式不改数据，保留公式和值，自动备份
- **12 套色系主题** — 经典商务（水蓝/深海蓝/墨玉绿/陨石灰蓝）、暖色高级（勃艮第红/珊瑚橙）、青春活力（樱花粉/暖阳橙/薰衣草紫/抹茶绿/蜜桃/雾蓝紫）
- **条件格式自适应** — beautify 自动将 colorScale 色阶最大色替换为当前主题表头色
- **智能类型推断** — 通过列名关键词（pct `率$` > date > text > money）识别列类型，浮点一律判 money 防万元单位误判
- **可配置表头冻结** — 参数指定/自动检测/B2 起始/默认首行，4 种策略
- **自定义数字格式** — `fmt_override` 参数覆盖默认格式
- **公式颜色区分** — 公式→黑色，手动输入→蓝色（摩根系标准）
- **三段式 fallback 恢复** — 12 个已知失败模式 + 4 个运行时自动恢复

## 快速使用

```bash
# 安装依赖
pip install pandas openpyxl>=3.0

# CLI：从 CSV 生成
python scripts/make_excel.py 数据.csv 输出.xlsx

# CLI：美化已有文件
python scripts/make_excel.py --beautify 已有报表.xlsx 美化后.xlsx

# CLI：带主题参数
python scripts/make_excel.py --beautify 报表.xlsx 美化后.xlsx --theme coral --freeze-rows 3
```

```python
from make_excel import make_excel, beautify
import pandas as pd

# 从零生成
df = pd.read_csv('data.csv')
make_excel(df, '报表.xlsx')

# 多 sheet
make_excel([('汇总', df_summary), ('明细', df_detail)], '报表.xlsx')

# 切换主题
make_excel(df, '报表-深海蓝.xlsx', theme='deep-navy')

# 美化已有文件（保留公式）
beautify('已有报表.xlsx')                    # 原地美化，自动备份
beautify('已有报表.xlsx', '美化后.xlsx')       # 另存

# 美化 + 手动指定列类型
beautify('订单表.xlsx', col_types={'订单号': 'text', '金额': 'money'})
```

## 色系主题

| 主题名 | 中文 | 表头色 | 分类 |
|--------|------|--------|------|
| `default` | 水蓝 | #4472C4 | 经典商务 |
| `deep-navy` | 深海蓝 | #1F4E79 | 经典商务 |
| `jade` | 墨玉绿 | #375623 | 经典商务 |
| `slate` | 陨石灰蓝 | #404040 | 经典商务 |
| `burgundy` | 勃艮第红 | #843C0C | 暖色高级 |
| `coral` | 珊瑚橙 | #D84B4B | 暖色高级 |
| `sakura` | 樱花粉 | #D94F70 | 青春活力 |
| `warm-sun` | 暖阳橙 | #E8843A | 青春活力 |
| `lavender` | 薰衣草紫 | #8B7EC8 | 青春活力 |
| `matcha` | 抹茶绿 | #7BA23F | 青春活力 |
| `peach` | 蜜桃 | #E8897C | 青春活力 |
| `misty` | 雾蓝紫 | #6B7FB5 | 青春活力 |

## 执行流程

1. **Step 1** — 读数据，了解表结构（列名、数据类型、范围）
2. **Step 2** — 确认表头行位置，确定 `freeze_rows`
3. **Step 3** — 选择配色，确定 `theme`
4. **Step 4** — PRE-FLIGHT 检查（文件合法、首行是表头、A 列数据、类型推断风险）
5. **Step 5** — 调用核心脚本生成 xlsx
6. **Step 6** — TYPE CHECK 验证类型推断
7. **Step 7** — DELIVERY GATE 全量验证后交付

## 类型推断

**核心原则：百分比只靠列名关键词判断。** 值分析阶段所有 float 先判为 money。

优先级：pct（`率$`/`占比`/`百分比`/`rate$`/`ratio$`）> date > text > money（含 `毛利(?!率)` 防误匹配）

误判时可用 `col_types={'列名': '类型'}` 强制覆盖。

## 设计原则

1. **零 xlwings** — 纯 openpyxl，一次保存，不调 Excel 外部进程
2. **纯函数式** — 同样输入永远同样输出
3. **单文件** — `make_excel.py` 一个文件解决所有
4. **基座职责分离** — 配色/行为作为参数暴露，应用层直接传参

## 参考文档

- `SKILL.md` — 完整 skill 文档（含强制约束、三段式 fallback 表、11 条反例黑名单）
- `references/type-inference-rules.md` — 列类型推断规则
- `references/implementation-checklist.md` — 交付前逐项验证清单
- `references/dual-header-format.py` — 双表头/多数据块布局手工格式脚本

## License

MIT
