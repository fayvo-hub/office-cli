"""replace — 在区域内查找并替换文本(可选正则、区分大小写、公式内替换)。"""

from __future__ import annotations

import argparse
import re

from .. import ioplan, xlutil
from ..cli import add_file_arg, add_sheet_arg
from ..errors import CliError

NAME = "replace"
HELP = "查找替换文本(--find/--replace,支持 --regex/--match-case)"
DESCRIPTION = """在指定区域(默认整表数据区)查找并替换单元格文本。

用法示例:
  office replace -f f.xlsx --find 张三 --replace 李四
  office replace -f f.xlsx --find '2000' --replace '2001' --range C2:C100
  office replace -f f.xlsx --find '^2024' --replace '2025' --regex
  office replace -f f.xlsx --find '未结' --replace '已结' --match-case

说明:
- 只替换文本单元格;公式单元格默认跳过(加 --in-formulas 也会替换
  公式里的文本片段,公式引用本身不变)
- --match-case 区分大小写(默认不区分);--regex 用正则(默认按字面文本)
- 替换结果若以 "=" 开头会被强制存为文本,不会变成公式
- 默认范围是整表有数据区域;超大表格请用 --range 收窄避免扫描过慢
- 操作不可撤销,建议先备份
输出 JSON: {"ok": true, "file": "...", "cells_changed": N, "occurrences": N}
"""


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp)
    add_sheet_arg(sp)
    sp.add_argument("--find", required=True, metavar="TEXT", help="查找内容")
    sp.add_argument("--replace", default="", metavar="TEXT",
                    help="替换为(默认空串=删除)")
    sp.add_argument("--range", metavar="REF", help="限定区域,如 A1:E100")
    sp.add_argument("--regex", action="store_true", help="把 --find 当正则")
    sp.add_argument("--match-case", action="store_true", help="区分大小写")
    sp.add_argument("--in-formulas", action="store_true",
                    help="同时替换公式中的文本片段")


def run(args: argparse.Namespace) -> dict:
    plan = ioplan.excel_plan(args.file)
    wb = xlutil.open_workbook(plan.read_path)
    ws = xlutil.choose_sheet(wb, args.sheet)

    dims = xlutil.data_dimensions(ws)
    if args.range:
        r1, c1, r2, c2 = xlutil.parse_range(args.range, ws)
    elif dims is None:
        return {"ok": True, "file": plan.write_path, "sheet": ws.title,
                "range": None, "cells_changed": 0, "occurrences": 0,
                "skipped_formulas": 0}
    else:
        r1, c1, r2, c2 = dims

    if (r2 - r1 + 1) * (c2 - c1 + 1) > 3_000_000:
        raise CliError("bad_args", "区域过大(超过 300 万格),请用 --range 收窄")

    flags = 0 if args.match_case else re.IGNORECASE
    pattern = re.compile(re.escape(args.find) if not args.regex
                         else args.find, flags)

    changed = occurrences = skipped = 0
    broken = 0
    for row in xlutil.iter_area(ws, r1, c1, r2, c2):
        for cell in row:
            v = cell.value
            if v is None or not isinstance(v, str) or not v:
                continue
            # 公式判定用 data_type(以 '=' 开头但被 --literal 存的文本 data_type='s',不算)
            is_formula = cell.data_type == "f"
            if is_formula and not args.in_formulas:
                skipped += 1
                continue
            if not pattern.search(v):
                continue
            new = pattern.sub(args.replace, v)
            if is_formula:
                # 公式:替换后必须仍是 '=' 开头的公式形态,否则保留原值并记录
                if not new.startswith("="):
                    broken += 1
                    continue
                cell.value = new
            elif new.startswith("="):
                # 先赋值再强制 data_type='s'——value setter 会把 '=' 开头当公式
                cell.value = new
                cell.data_type = "s"
            else:
                cell.value = new
            changed += 1
            occurrences += len(pattern.findall(v))

    xlutil.save_workbook_atomic(wb, plan.write_path)
    _up = {"upgraded_from": plan.upgraded_from} if plan.upgraded_from else {}
    warnings = []
    if skipped:
        warnings.append(f"跳过 {skipped} 个公式单元格(如需替换公式文本请加 --in-formulas)")
    if broken:
        warnings.append(f"{broken} 个公式因替换会丢失 '=' 前缀已保持原值")
    if occurrences:
        warnings.append("Excel 打开后会重算公式,当前值以文件为准")
    return {"ok": True, "file": plan.write_path, **_up, "sheet": ws.title,
            "range": xlutil.area_label(r1, c1, r2, c2),
            "cells_changed": changed, "occurrences": occurrences,
            "warnings": warnings}
