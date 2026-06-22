# 列类型推断规则

## 核心原则：百分比只靠列名判断

百分比类型（pct）**仅通过列名关键词匹配触发**，不依赖值范围分析。因为 0~1 之间的浮点数可能是万元/千元单位的金额（如 0.84 万元），不一定是百分数。值分析阶段的所有 float 都先判为 money。

## 关键词优先级（高→低）

1. **pct（百分比）** — 以"率"结尾（毛利率、完成率、增长率）或含 rate/ratio
2. **date（日期）** — 日期、date、年、month、年月
3. **text（文本）** — 姓名、地址、电话、邮箱、备注、编码、phone、email、code、desc、address
4. **money（金额）** — 金额、收入、成本、利润、费用、预算、price、cost、revenue、total、budget、expense
5. **number（数字）** — 以上都不匹配时的兜底

## pct 检测的防误匹配

列名含"毛利"时容易与 money 冲突。用 `毛利(?!率)` 正则排除"毛利率"：
- "毛利率" → pct ✓
- "毛利额" → money ✓
- "毛利" → money ✓（不含"率"）

## 英文关键词对照

| 类型 | 英文关键词 |
|------|-----------|
| pct | rate, ratio |
| date | date, year, month |
| text | phone, email, code, desc, address, name |
| money | amount, price, cost, revenue, total, budget, expense |

## `_infer_col_type()` 实现逻辑

逐条匹配关键词列表，命中即返回。匹配顺序 = 优先级顺序。兜底返回 'number'。

## 扩展方式

在 `_infer_col_type()` 的关键词元组中追加新词即可。需要新增类型时，在函数顶部添加新 If 分支，并补充对应的 `number_format` 映射。
