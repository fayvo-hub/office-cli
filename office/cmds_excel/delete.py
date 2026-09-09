"""delete — 删除整行/整列(其后数据自动上移/左移)。"""

from __future__ import annotations

import argparse
import re

from .. import ioplan, xlutil
from ..cli import add_file_arg, add_sheet_arg
from ..errors import CliError

NAME = "delete"
HELP = "删除整行/整列,如 --rows 3-5 或 --cols B:D"
DESCRIPTION = """删除整行或整列,其后内容自动移位(Excel 打开后同效果)。

用法示例:
  office delete -f f.xlsx --rows 3-5     # 删除第 3~5 行
  office delete -f f.xlsx --rows 2       # 只删第 2 行
  office delete -f f.xlsx --cols B:D     # 删除 B~D 三列

说明与限制:
- 公式引用不会随删除更新,被删区域的单元格被引用时公式将报错
- 删除带涉及合并单元格/表格对象/自动筛选区域时会先报错拒绝,
  请先用 merge --unmerge / layout --unfilter 清理后再删除
- 操作不可撤销,建议先备份
输出 JSON: {"ok": true, "file": "...", "deleted": "rows|cols", "range": "3-5"}
"""


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp)
    add_sheet_arg(sp)
    g = sp.add_mutually_exclusive_group(required=True)
    g.add_argument("--rows", metavar="RANGE", help="删除行,如 3-5 或 2")
    g.add_argument("--cols", metavar="RANGE", help="删除列,如 B:D 或 C")


def run(args: argparse.Namespace) -> dict:
    plan = ioplan.excel_plan(args.file)
    wb = xlutil.open_workbook(plan.read_path)
    ws = xlutil.choose_sheet(wb, args.sheet)
    axis, r1, r2, label = _range(args)

    xlutil.guard_rowcol_shift(ws, axis, r1, "删除")
    count = r2 - r1 + 1
    if axis == "rows":
        ws.delete_rows(r1, count)
    else:
        ws.delete_cols(r1, count)

    xlutil.save_workbook_atomic(wb, plan.write_path)
    _up = {"upgraded_from": plan.upgraded_from} if plan.upgraded_from else {}
    return {"ok": True, "file": plan.write_path, **_up, "sheet": ws.title,
            "deleted": axis, "range": label, "count": count,
            "warnings": ["公式引用不会随删除更新,请核对涉及公式的区域"]}


# ---------------------------------------------------------------------------

def _range(args) -> tuple[str, int, int, str]:
    if args.rows is not None:
        m = re.fullmatch(r"(\d+)(?:-(\d+))?", args.rows.strip())
        if not m:
            raise CliError("bad_args", f"--rows 格式应为 '3-5' 或 '3',收到 {args.rows!r}")
        a, b = int(m.group(1)), int(m.group(2) or m.group(1))
        if not 1 <= a <= b <= xlutil.MAX_ROWS:
            raise CliError("bad_args", f"--rows 需在 1..{xlutil.MAX_ROWS} 且起≤止")
        return "rows", a, b, args.rows.strip()
    text = args.cols.strip().upper()
    m = re.fullmatch(r"([A-Z]+)(?::([A-Z]+))?", text)
    if not m:
        raise CliError("bad_args", f"--cols 格式应为 'B:D' 或 'C',收到 {args.cols!r}")
    a, b = xlutil.col_to_idx(m.group(1)), xlutil.col_to_idx(m.group(2) or m.group(1))
    if a > b:
        raise CliError("bad_args", f"--cols 起止颠倒: {text}")
    if b > xlutil.MAX_COLS:
        raise CliError("bad_args", f"--cols 超出范围: {text}")
    return "cols", a, b, text
