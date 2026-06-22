#!/usr/bin/env python3
"""
interactive_make_excel — agent 交互层，处理用户输入后再调核心脚本。

职责
----
1. 接收 agent 从用户那里问来的 freeze_rows 和 theme 参数
2. 用户提供了 → 透传给核心 make_excel/beautify
3. 用户没提供/跳过 → 走默认（自动推断+水蓝），不阻塞

为什么单独一个文件？
- 核心脚本 make_excel.py 保持纯函数、零交互，自动化调用直接 import
- 这个 wrapper 只处理「人机对话→传参」这一层逻辑，可独立修改

用法（agent 在对话中收集完用户信息后调用）
--------
# 美化已有文件，用户指定了表头行和配色
python interactive_make_excel.py beautify 输入.xlsx --freeze-rows 3 --theme sakura

# 美化，用户没提供信息→自动推断+水蓝
python interactive_make_excel.py beautify 输入.xlsx

# 从 CSV 生成，用户指定了配色
python interactive_make_excel.py make 数据.csv 输出.xlsx --theme deep-navy
"""

import argparse
import sys
from pathlib import Path

# 核心脚本与 wrapper 同目录
SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

from make_excel import make_excel, beautify


def main():
    parser = argparse.ArgumentParser(
        description='Agent wrapper: 美化/生成 Excel（收集用户输入后调核心）'
    )
    parser.add_argument('action', choices=['make', 'beautify'],
                        help='make=从零生成, beautify=美化已有')
    parser.add_argument('input', help='输入文件（CSV -> make, xlsx -> beautify）')
    parser.add_argument('output', nargs='?', default=None,
                        help='输出路径。不传则 beautify 原地覆盖，make 输出到桌面')
    parser.add_argument('--freeze-rows', type=int, default=None,
                        help='冻结表头行数。不传则自动推断（检测 header_row / A2）')
    parser.add_argument('--theme', default=None,
                        help='配色主题。不传则 default（水蓝）')
    parser.add_argument('--col-types', nargs='*', default=None,
                        help='列类型覆盖，格式: 订单号:text 金额:money')

    args = parser.parse_args()
    theme = args.theme if args.theme else 'default'

    # ── beautify ──
    if args.action == 'beautify':
        col_types_dict = None
        if args.col_types:
            col_types_dict = {}
            for pair in args.col_types:
                if ':' in pair:
                    k, v = pair.split(':', 1)
                    col_types_dict[k] = v

        result = beautify(
            args.input, args.output,
            theme=theme,
            freeze_rows=args.freeze_rows,
            col_types=col_types_dict,
        )
        print(f'已美化: {result}')
        return

    # ── make ──
    import pandas as pd
    df = pd.read_csv(args.input)

    if not args.output:
        args.output = str(Path.home() / 'Desktop' / Path(args.input).with_suffix('.xlsx').name)

    result = make_excel(
        df, args.output,
        theme=theme,
        freeze_rows=args.freeze_rows,
    )
    print(f'已生成: {result}')


if __name__ == '__main__':
    main()
